/* ParticipantTable — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 11711-12888 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { EventService, SPRegistration } from '../../../services/EventService';
import { PersonContactHover } from '../../PersonContactHover';
import { formatDate, translateStatus } from '../../../utils/eventStatus';
import { Check, ChevronDown, ChevronUp, Columns, FileText, Pencil, UserMinus, X } from '../../Icons';
// v31.3: Gemeinsamer Klassensatz (docs/ui-leitfaden.md, Abschnitt 5b) — die
// Tabelle hatte bis dahin jeden Hover als Inline-Style-Notlösung
// (onMouseEnter-State für den Aufklapp-Knopf, Farbwechsel per DOM-Zugriff).
import { cx, ensureDexUiStyles } from '../../dexUi';
import { isEventOver } from '../../../utils/eventFormat';
import { selfCancelLocked } from '../../../utils/cancelPolicy';
import { externalInvitationEmail } from '../../../services/EmailTemplates';
import { buildUnsentEmlDraft, downloadEml } from '../../../utils/emlDraft';
import { SplitMergeToggle } from '../../admin/ActionsMenu';
import { DeloitteEvent } from '../../../types';

export interface ParticipantTableProps {
  activeRegs: SPRegistration[];
  allEvents: DeloitteEvent[];
  attachmentsByReg: Record<number, { fileName: string; serverRelativeUrl: string; }[]>;
  availableColumns: { id: string; label: string; alwaysVisible?: boolean; }[];
  colToggleHover: boolean;
  columnOrder: string[];
  computeRoommatePairs: (rows: SPRegistration[]) => Array<[SPRegistration, SPRegistration]>;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  duplicateEmails: Set<string>;
  eventServiceRef: EventService;
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
  getRoommateInfo: (reg: {    ParticipantEmail?: string;}) => { partnerName: string; partnerEmail: string; mutual: boolean; };
  handleSort: (col: string) => void;
  hasRoommateColumn: boolean;
  hiddenColumns: string[];
  hideColumn: (id: string) => void;
  highlightMatch: (text: unknown) => React.ReactNode;
  inactiveAccounts: string[];
  isDe: boolean;
  isSplitCapacity: boolean;
  moveColumn: (id: string, direction: -1 | 1) => void;
  openEditModal: (reg: SPRegistration) => void;
  orgPastLock: boolean;
  parentEventForSelected: DeloitteEvent;
  parentRegsByEmail: Record<string, SPRegistration>;
  performStandardCancel: (reg: SPRegistration) => Promise<void>;
  personalColsCollapsed: boolean;
  query: string;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setAttachmentsModalReg: React.Dispatch<React.SetStateAction<SPRegistration>>;
  setColToggleHover: React.Dispatch<React.SetStateAction<boolean>>;
  setDupCancelReg: React.Dispatch<React.SetStateAction<SPRegistration>>;
  setParticipantDetail: React.Dispatch<React.SetStateAction<{ name: string; email: string; jobTitle: string; location: string; company: string; department: string; phone: string; status: string; tid: number; }>>;
  setPersonalColsCollapsed: React.Dispatch<React.SetStateAction<boolean>>;
  setShowColumnPicker: React.Dispatch<React.SetStateAction<boolean>>;
  setSplitParticipantsView: React.Dispatch<React.SetStateAction<"split" | "merged">>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  showColumn: (id: string) => void;
  showColumnPicker: boolean;
  showMatches: boolean;
  sortIcon: (col: string) => string;
  splitParticipantsView: "split" | "merged";
  stripLocPrefix: (loc: string) => string;
}

export const ParticipantTable: React.FC<ParticipantTableProps> = (p) => {
  // v31.3: `colToggleHover`/`setColToggleHover` bleiben in der Schnittstelle,
  // werden aber nicht mehr gelesen — der Aufklapp-Knopf hat seinen Hover jetzt
  // über `dex-ui-chip` (Leitfaden 1.3: kein onMouseEnter-State für reine Optik).
  const { activeRegs, allEvents, attachmentsByReg, availableColumns, columnOrder, computeRoommatePairs, confirmDialog, duplicateEmails, eventServiceRef, getRoommateInfo, handleSort, hasRoommateColumn, hiddenColumns, hideColumn, highlightMatch, inactiveAccounts, isDe, isSplitCapacity, moveColumn, openEditModal, orgPastLock, parentEventForSelected, parentRegsByEmail, performStandardCancel, personalColsCollapsed, query, registrations, reloadRegistrations, selectedEvent, setAttachmentsModalReg, setDupCancelReg, setParticipantDetail, setPersonalColsCollapsed, setShowColumnPicker, setSplitParticipantsView, showAlert, showColumn, showColumnPicker, showMatches, sortIcon, splitParticipantsView, stripLocPrefix } = p;
  // Idempotent — Modal und WizardFormShell rufen es ebenfalls; hier nötig, weil
  // die Tabelle auch ohne offenes Modal gerendert wird.
  ensureDexUiStyles();
  // v31.3: Für die Fußzeile „N von M": M ist die ungefilterte Zahl aktiver
  // Anmeldungen (dieselben Stati wie `activeRegs` in AdminPage, dort aber
  // schon durch die Suche gefiltert). Nur Anzeige — kein Filter-State.
  const activeTotal = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length;
  // v31.3: Sortier-Pfeil als eigene Klasse — `sortIcon` liefert weiterhin den
  // String (Sortier-State bleibt beim Aufrufer), hier wird er nur gerahmt.
  const sortMark = (col: string): React.ReactNode => {
    const s = sortIcon(col).trim();
    return s ? <span className="dex-ui-table-sort" aria-hidden="true">{s}</span> : null;
  };
  const isSorted = (col: string): boolean => sortIcon(col).trim().length > 0;
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
  return (
          /* v17.13: overflowX: 'auto' entfernt — der scrollbare Wrapper
             hat die sticky-thead-Berechnung gebrochen (sticky relative zum
             Scroll-Container statt zum Window). Tabelle lässt die Karte
             jetzt horizontal überlaufen, was bei vielen Spalten zu einer
             Scrollbar AM AUSSEREN Container (SP-Page) führt — Sticky-
             thead funktioniert dort einwandfrei. */
          <div style={{ overflowX: 'visible' }}>
            {(() => {
              // v6.17: Spaltenkonfiguration — Header und Body-Zellen werden dynamisch
              // anhand `columnOrder` (+ `hiddenColumns`) gerendert. So kann der User
              // Spalten ein-/ausblenden und per Pfeilen umsortieren. Die Render-Logik
              // selbst (Sort-Buttons, Badges, Custom-Field-Anzeige etc.) bleibt gleich,
              // nur die Iteration ist umgebaut.
              const visibleColumnIds = columnOrder.filter(id => hiddenColumns.indexOf(id) < 0);
              // v23.33: Eingeklappt = die Personen-Spalten zu EINER „person"-
              // Spalte (Foto + zweizeilig) zusammenfassen. Die synthetische
              // 'person'-Spalte ersetzt die erste sichtbare Personen-Spalte,
              // die übrigen entfallen.
              const PERSONAL_IDS = ['anrede', 'vorname', 'nachname', 'email', 'jobTitle', 'location', 'company'];
              const effectiveColumnIds = (() => {
                if (!personalColsCollapsed) return visibleColumnIds;
                const out: string[] = [];
                let inserted = false;
                for (const cid of visibleColumnIds) {
                  if (PERSONAL_IDS.indexOf(cid) >= 0) {
                    if (!inserted) { out.push('person'); inserted = true; }
                    continue;
                  }
                  out.push(cid);
                }
                if (!inserted) out.unshift('person');
                return out;
              })();

              const sortableCols: Record<string, 'id' | 'anrede' | 'vorname' | 'nachname' | 'email' | 'status' | 'date'> = {
                id: 'id', anrede: 'anrede', vorname: 'vorname', nachname: 'nachname', email: 'email', status: 'status', date: 'date',
              };

              const hideButton = (id: string): React.ReactNode => {
                const col = availableColumns.find(c => c.id === id);
                if (!col || col.alwaysVisible) return null;
                // v31.3: Symbol-Knopf mit rotem Hover über die Klasse statt
                // per DOM-Farbwechsel; kleiner als der Standard, weil er im
                // Kopf neben dem Sortier-Klickziel steht.
                return (
                  <button
                    type="button"
                    className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                    onClick={(e) => { e.stopPropagation(); hideColumn(id); }}
                    aria-label={isDe ? `Spalte ${col.label} ausblenden` : `Hide column ${col.label}`}
                    title={isDe ? 'Spalte ausblenden' : 'Hide column'}
                    style={{ width: 18, height: 18, marginLeft: 4, verticalAlign: 'middle' }}
                  >
                    <X size={11} />
                  </button>
                );
              };

              const renderHeader = (id: string): React.ReactNode => {
                // v15.3: lange Spalten-Überschriften (Custom-Field-Labels wie
                // „Please check if you have marked all parts you want to attend
                // and confirm") brechen jetzt um statt mit Ellipsis abgeschnitten
                // zu werden. Begrenzte maxWidth + Wortumbruch — der Header bleibt
                // lesbar ohne dass der User hovern muss.
                // v15.4.1: wordBreak:'break-word' war zu aggressiv (Edge
                // brach kurze Wörter wie „Vorname" → „Vorna\nme"). Jetzt
                // overflowWrap:'break-word' — Umbruch nur an Wort-Grenzen
                // oder wenn ein einzelnes Wort breiter als die Spalte ist.
                // v31.3: Sticky, Grund, Rand und Innenabstand kommen aus
                // `dex-ui-table--compact` in `dex-ui-table-wrap--sticky`
                // (renderTable, v24.96: eigener Scroll-Container). Hier bleibt
                // nur, was die Klasse nicht weiß: Umbruch und Breite.
                const baseStyle: React.CSSProperties = {
                  whiteSpace: 'normal',
                  overflowWrap: 'break-word',
                  hyphens: 'auto',
                  maxWidth: 180,
                  verticalAlign: 'top',
                  lineHeight: 1.3,
                };
                // v31.3: Sortierbare Kopfzelle — Zeiger und Hover über
                // `is-sortable`, die sortierte Spalte grün über `is-sorted`.
                const sortableClass = (col: string): string => cx('is-sortable', isSorted(col) && 'is-sorted');
                // v23.33: eingeklappte „Teilnehmer"-Spalte (Foto + zweizeilig).
                if (id === 'person') {
                  // v30.60: Dieselbe beschriftete Pille wie in der
                  // konsolidierten Matrix (v30.21). Diese Tabelle — und damit
                  // jedes Event mit geteilten Gruppen — hatte weiterhin den
                  // kleinen runden Knopf direkt neben dem Sortier-Klickziel:
                  // ein leicht versetzter Klick sortierte, statt aufzuklappen.
                  // Genau der Befund, der die Matrix schon umgebaut hat; zwei
                  // Tabellen mit zwei Bedienungen für dieselbe Sache sind
                  // zudem die Sucherei, die eine Vereinheitlichung erspart.
                  // v31.3: Die Pille ist ein `dex-ui-chip` — Hover aus der Klasse.
                  return (
                    <th key="person" className={sortableClass('nachname')} style={{ ...baseStyle, whiteSpace: 'nowrap' }} onClick={() => handleSort('nachname')}>
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 4 }}>
                        <button
                          type="button"
                          className="dex-ui-chip"
                          onClick={(e) => { e.stopPropagation(); setPersonalColsCollapsed(false); }}
                          title={isDe ? 'Vorname, Nachname, E-Mail, Job Title, Standort und Unternehmen als eigene Spalten anzeigen' : 'Show first/last name, email, job title, location and company as separate columns'}
                          style={{ padding: '2px 10px', fontSize: '0.7rem', textTransform: 'none', letterSpacing: 0 }}
                        >» {isDe ? 'Aufklappen' : 'Expand'}</button>
                        <span>{isDe ? 'Teilnehmer' : 'Participant'}{sortMark('nachname')}</span>
                      </div>
                    </th>
                  );
                }
                const sortable = sortableCols[id];
                if (sortable) {
                  return (
                    <th
                      key={id}
                      className={cx(sortableClass(sortable), id === 'id' && 'is-num')}
                      style={baseStyle}
                      onClick={() => handleSort(sortable)}
                    >
                      {/* v30.60: „Vorname" trägt den Zuklapp-Knopf als Pille in
                          einer eigenen Zeile — Gegenstück zum Aufklappen oben. */}
                      {id === 'vorname' ? (
                        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 4 }}>
                          <button
                            type="button"
                            className="dex-ui-chip"
                            onClick={(e) => { e.stopPropagation(); setPersonalColsCollapsed(true); }}
                            title={isDe ? 'Personen-Spalten einklappen (nur Foto + Name)' : 'Collapse personal columns (photo + name only)'}
                            style={{ padding: '2px 10px', fontSize: '0.7rem', textTransform: 'none', letterSpacing: 0 }}
                          >« {isDe ? 'Zuklappen' : 'Collapse'}</button>
                          <span>{isDe ? 'Vorname' : 'First name'}{sortMark(sortable)}{hideButton(id)}</span>
                        </div>
                      ) : (
                        <>
                          {id === 'id' ? '#' : id === 'anrede' ? (isDe ? 'Anrede' : 'Salutation') : id === 'nachname' ? (isDe ? 'Nachname' : 'Last name') : id === 'email' ? (isDe ? 'E-Mail' : 'Email') : id === 'status' ? 'Status' : (isDe ? 'Registriert am' : 'Registered on')}
                          {sortMark(sortable)}
                          {hideButton(id)}
                        </>
                      )}
                    </th>
                  );
                }
                if (id === 'jobTitle') return <th key={id} className={sortableClass('jobTitle')} style={baseStyle} onClick={() => handleSort('jobTitle')}>Job Title{sortMark('jobTitle')}{hideButton(id)}</th>;
                if (id === 'location') return <th key={id} className={sortableClass('location')} style={baseStyle} onClick={() => handleSort('location')}>{isDe ? 'Standort' : 'Location'}{sortMark('location')}{hideButton(id)}</th>;
                if (id === 'company') return <th key={id} style={baseStyle}>{isDe ? 'Unternehmen' : 'Company'}{hideButton(id)}</th>;
                if (id === 'starterType') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? "Starter-Typ: Durchstarter oder Funstarter. Wird bei der Anmeldung gewählt und steuert die Split-Kapazität + Warteliste. Der eigentliche Startblock steht in der Custom-Field-Spalte 'Start block'." : "Starter type: Durchstarter or Funstarter. Chosen at registration and controls the split capacity + waitlist. The actual start block is in the custom field column 'Start block'."}>
                      {isDe ? 'Starter-Typ' : 'Starter type'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'startnummer') {
                  return (
                    <th key={id} className="is-num" style={baseStyle} title={isDe ? 'Die offizielle Startnummer des Veranstalters, eingelesen über „Startnummern importieren".' : 'The official bib number from the organiser, imported via "Import bib numbers".'}>
                      {isDe ? 'Startnummer' : 'Bib number'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'promotedDate') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Zeitpunkt des Nachrückens — gesetzt sobald der Teilnehmer von der Warteliste in den Aktiv-Bereich promotet wurde. Leer für Personen die sich direkt angemeldet haben.' : 'Time of promotion — set as soon as the participant was promoted from the waitlist into the active area. Empty for people who registered directly.'}>
                      {isDe ? 'Nachgerückt am' : 'Promoted on'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'replaced') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Die abgemeldete Person, deren Platz diese Person übernommen hat. Nur gesetzt für nachgerückte Personen.' : 'The cancelled person whose seat this person took. Only set for promoted people.'}>
                      {isDe ? 'Hat ersetzt' : 'Replaced'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'replacedBy') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Wer nach der Abmeldung dieses Teilnehmers den Platz übernommen hat. Nur gesetzt für abgemeldete Personen, deren Cancel einen Promote ausgelöst hat.' : 'Who took the seat after this participant cancelled. Only set for cancelled people whose cancellation triggered a promotion.'}>
                      {isDe ? 'Ersetzt durch' : 'Replaced by'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'registeredBy') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Selbst = der Teilnehmer hat sich selbst registriert. Ansonsten Name des Users, der die Registrierung durchgeführt hat.' : 'Self = the participant registered themselves. Otherwise the name of the user who performed the registration.'}>
                      {isDe ? 'Registriert von' : 'Registered by'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'team') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Team-Name des Teilnehmers (falls Team-Anmeldung aktiv).' : 'Team name of the participant (if team registration is active).'}>
                      Team{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'roommate') {
                  // v11.56: Label dynamisch aus availableColumns nehmen (entstammt dem
                  // ersten roommate-/user-Feld der Custom-Field-Definition) statt
                  // hartcodiertem „Zimmerpartner".
                  const roommateCol = availableColumns.find(c => c.id === 'roommate');
                  const roommateLabel = roommateCol?.label || 'Zimmerpartner';
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Ausgewählter User-Picker-Wert aus diesem Feld. Match = beide haben sich gegenseitig ausgewählt.' : 'Selected user-picker value from this field. Match = both selected each other.'}>
                      {roommateLabel}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'action') {
                  // v31.3: Aktionen rechtsbündig in der letzten Spalte (Leitfaden
                  // 5b) und endlich zweisprachig — „Aktion" stand fest deutsch.
                  return <th key={id} className="is-num">{isDe ? 'Aktionen' : 'Actions'}</th>;
                }
                // v14.11: pastel A = event-level (parent) fields, pastel B = sub-event-specific fields.
                // Pastel-Hintergrund nur im Sub-Event-Detail-View (parentEventForSelected gesetzt),
                // sonst neutraler Hintergrund wie bisher.
                const inSubEventDetail = !!parentEventForSelected;
                const pastelAHeader: React.CSSProperties = inSubEventDetail ? { background: 'rgba(0, 118, 168, 0.15)' } : {};
                const pastelBHeader: React.CSSProperties = inSubEventDetail ? { background: 'rgba(255, 191, 0, 0.18)' } : {};
                if (id.indexOf('cfp-') === 0) {
                  const cfId = id.substring(4);
                  const field = (parentEventForSelected?.eventSpecificFields || []).find(f => f.id === cfId);
                  if (!field) return null;
                  const label = field.label || '';
                  return (
                    <th key={id} className={sortableClass(id)} onClick={() => handleSort(id)} style={{ ...baseStyle, textTransform: 'none', letterSpacing: 0, ...pastelAHeader }} title={`${label} — ${isDe ? 'Hauptevent-Feld' : 'main-event field'}`}>
                      {label}{sortMark(id)}
                      {hideButton(id)}
                    </th>
                  );
                }
                if (id.indexOf('cf-') === 0) {
                  const cfId = id.substring(3);
                  const field = (selectedEvent?.eventSpecificFields || []).find(f => f.id === cfId);
                  if (!field) return null;
                  const label = field.label || '';
                  return (
                    <th key={id} className={sortableClass(id)} onClick={() => handleSort(id)} style={{ ...baseStyle, textTransform: 'none', letterSpacing: 0, ...pastelBHeader }} title={inSubEventDetail ? `${label} — ${isDe ? 'Sub-Event-Feld' : 'sub-event field'}` : label}>
                      {label}{sortMark(id)}
                      {hideButton(id)}
                    </th>
                  );
                }
                return null;
              };

              const renderCell = (id: string, reg: SPRegistration, i: number): React.ReactNode => {
                if (id === 'id') {
                  // v26.31: Beim Filtern die laufende Treffer-Nr. „#n" voranstellen
                  // und die echte TeilnehmerID (Platz) in Klammern zeigen — analog zur
                  // Warteliste; ohne Filter unverändert nur die TeilnehmerID.
                  const idCell = query
                    ? `#${i + 1}${reg.TeilnehmerID ? ` (#${reg.TeilnehmerID})` : ''}`
                    : (reg.TeilnehmerID || (i + 1));
                  return <td key={id} className="is-num" style={{ color: 'var(--dex-gray-500)' }}>{idCell}</td>;
                }
                // v23.33: eingeklappte „Teilnehmer"-Zelle — Foto + zweizeilig
                // (Name fett, darunter „Position • Standort" ohne Länder-Präfix).
                if (id === 'person') {
                  const vn = reg.Vorname || ((reg.ParticipantName || '').split(' ')[0] || '');
                  let nn = reg.Nachname || '';
                  if (!nn && reg.ParticipantName) { const p = reg.ParticipantName.trim().split(/\s+/); if (p.length > 1) nn = p.slice(1).join(' '); }
                  const fullName = `${vn} ${nn}`.trim() || reg.ParticipantEmail || '-';
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const jt = String((reg as any).JobTitle || '');
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const loc = stripLocPrefix(String((reg as any).Location || ''));
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const comp = String((reg as any).Company || '');
                  const sub = [jt, loc, comp].filter(Boolean).join(' • ');
                  const email = reg.ParticipantEmail || '';
                  return (
                    <td key="person">
                      {/* v31.3: Personen-Zelle nach Leitfaden 5b. Der Name ist ein
                          echter Knopf mit Hover (öffnet das Detail) — vorher ein
                          div mit cursor:pointer, dem man das Klicken nicht ansah.
                          Die Zeile selbst bleibt bewusst NICHT klickbar: Wer eine
                          E-Mail zum Kopieren markiert, würde sonst das Detail öffnen. */}
                      <div className="dex-ui-person">
                        {/* v24.56: Foto-Hover zeigt Kontaktkarte (E-Mail + Teams),
                            wie bei den Organizern auf der Anmeldeseite. */}
                        <PersonContactHover email={email} name={fullName} size={30} subline={sub} isDe={isDe} />
                        <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0, lineHeight: 1.25, alignItems: 'flex-start' }}>
                          <button
                            type="button"
                            className="dex-ui-textbtn dex-ui-textbtn--muted dex-ui-person-name"
                            style={{ padding: '1px 6px', margin: '0 -6px', fontSize: 'inherit' }}
                            title={isDe ? 'Details anzeigen' : 'Show details'}
                            // eslint-disable-next-line @typescript-eslint/no-explicit-any
                            onClick={() => setParticipantDetail({ name: fullName, email, jobTitle: jt, location: String((reg as any).Location || ''), company: comp, department: String((reg as any).Department || ''), phone: String((reg as any).Phone || ''), status: reg.Status || '', tid: reg.TeilnehmerID || null })}
                          >
                            {highlightMatch(fullName)}
                          </button>
                          {sub && <span className="dex-ui-person-sub">{highlightMatch(sub)}</span>}
                        </div>
                      </div>
                    </td>
                  );
                }
                if (id === 'anrede') {
                  return <td key={id} style={{ color: 'var(--dex-gray-500)' }}>{reg.Anrede || '-'}</td>;
                }
                if (id === 'vorname') {
                  // Fallback für Alt-Daten: erstes Wort aus ParticipantName.
                  const v = reg.Vorname || ((reg.ParticipantName || '').split(' ')[0] || '');
                  return <td key={id} style={{ fontWeight: 500 }}>{v ? highlightMatch(v) : '-'}</td>;
                }
                if (id === 'nachname') {
                  // Fallback für Alt-Daten: alles ausser dem ersten Wort als Nachname.
                  let n = reg.Nachname || '';
                  if (!n && reg.ParticipantName) {
                    const parts = reg.ParticipantName.trim().split(/\s+/);
                    if (parts.length > 1) n = parts.slice(1).join(' ');
                  }
                  return <td key={id} style={{ fontWeight: 500 }}>{n ? highlightMatch(n) : '-'}</td>;
                }
                if (id === 'email') {
                  return <td key={id} style={{ color: 'var(--dex-gray-600)' }}>{highlightMatch(reg.ParticipantEmail)}</td>;
                }
                if (id === 'jobTitle') {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const jt = String((reg as any).JobTitle || '');
                  return <td key={id} style={{ color: 'var(--dex-gray-600)' }}>{jt ? highlightMatch(jt) : '-'}</td>;
                }
                if (id === 'location') {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const lc = String((reg as any).Location || '');
                  return <td key={id} style={{ color: 'var(--dex-gray-600)' }}>{lc ? highlightMatch(lc) : '-'}</td>;
                }
                if (id === 'company') {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const cmp = String((reg as any).Company || '');
                  return <td key={id} style={{ color: 'var(--dex-gray-600)' }}>{cmp ? highlightMatch(cmp) : '-'}</td>;
                }
                if (id === 'starterType') {
                  return (
                    <td key={id}>
                      {(() => {
                        // Tatsächlicher Startblock (StarterType) + Wunsch (PreferredStarterType).
                        // Wenn beide identisch: nur einen anzeigen. Wenn unterschiedlich (z.B. per
                        // Fallback-Dialog auf anderen Typ umgestiegen): Wunsch in Klammern daneben.
                        const actual = reg.StarterType || '';
                        const pref = reg.PreferredStarterType || '';
                        if (!actual && !pref) return <span style={{ color: 'var(--dex-gray-400)' }}>—</span>;
                        if (actual && pref && actual !== pref) {
                          return <span>{actual} <span style={{ color: 'var(--dex-gray-500)' }}>({isDe ? 'Wunsch' : 'preferred'}: {pref})</span></span>;
                        }
                        if (actual) return <span>{actual}</span>;
                        // v19.12: StarterType ist leer. Bei AKTIVEN (angemeldeten/
                        // eingecheckten) Personen ist die effektive Gruppe der Wunsch
                        // — der Nachrück-Flow hat den StarterType beim Promoten nur
                        // nicht gesetzt. Solche Personen NEHMEN ihren Wunsch-Platz ein,
                        // also plain die Gruppe zeigen (NICHT „Wunsch:"). „Wunsch:"
                        // bleibt den Warteliste-Personen vorbehalten (dort ist die
                        // Gruppe wirklich noch nicht zugewiesen).
                        const isWaitlist = reg.Status === 'Warteliste';
                        return <span>{isWaitlist ? `${isDe ? 'Wunsch' : 'preferred'}: ${pref}` : pref}</span>;
                      })()}
                    </td>
                  );
                }
                if (id === 'status') {
                  // v26.47: Externe Anmeldung mit offener Datenschutz-Rückmeldung
                  // (ConsentReview='Pending') — oranger Badge statt des normalen
                  // Status, solange die Person noch aktiv (nicht abgemeldet) ist.
                  if (reg.ConsentReview === 'Pending' && reg.Status !== 'Abgemeldet') {
                    return (
                      <td key={id}>
                        <span className="dex-ui-pill dex-ui-pill--orange" title={isDe ? 'Angemeldet — die Datenschutz-Rückmeldung der externen Person steht noch aus' : 'Registered — the external person’s privacy confirmation is still pending'}>
                          {isDe ? 'Angemeldet · Rückmeldung offen' : 'Registered · response pending'}
                        </span>
                      </td>
                    );
                  }
                  // v31.3: Farbige Pille je Status (Leitfaden 5b) statt grün/grau.
                  return (
                    <td key={id}>
                      <span className={statusPillClass(reg.Status)}>
                        {translateStatus(reg.Status, isDe)}
                      </span>
                    </td>
                  );
                }
                if (id === 'date') {
                  // v27.12 (Feedback Datenschutz-Review): Zeilen, die nicht über
                  // die App angelegt wurden (z.B. direkt in der SharePoint-
                  // Liste), haben kein RegistrationDate — dann den SP-Erstell-
                  // Zeitstempel (Created) als Fallback zeigen statt Leere.
                  const regDate = (reg.RegistrationDate || '').trim() ? reg.RegistrationDate : (reg.Created || '');
                  return <td key={id} style={{ color: 'var(--dex-gray-500)' }}>{regDate ? formatDate(regDate) : '—'}</td>;
                }
                if (id === 'startnummer') {
                  // v30.48: Startnummer aus dem Veranstalter-Rücklauf.
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const bib = String((reg as any).Startnummer || '').trim();
                  return <td key={id} className="is-num" style={{ fontFamily: 'monospace', fontWeight: bib ? 700 : 400, color: bib ? 'var(--dex-gray-800)' : 'var(--dex-gray-300)' }}>{bib || '—'}</td>;
                }
                if (id === 'promotedDate') {
                  // v17.15: „Nachgerückt am" — gesetzt beim Promote
                  // von Warteliste → Angemeldet. Leer für Personen die
                  // sich direkt in den Aktiv-Bereich angemeldet haben.
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const v = (reg as any).PromotedDate as string | undefined;
                  return <td key={id} style={{ color: v ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-300)' }}>{v ? formatDate(v) : '—'}</td>;
                }
                if (id === 'replaced') {
                  // v17.15: „Ersetzt" — die Person, deren Cancel diesen
                  // Promote ausgelöst hat. Wenn die Person in den
                  // aktuellen registrations gefunden wird, zeigen wir den
                  // Namen — sonst fallback auf die rohe E-Mail.
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const email = ((reg as any).ReplacedParticipantEmail as string | undefined) || '';
                  if (!email) return <td key={id} style={{ color: 'var(--dex-gray-300)' }}>—</td>;
                  const other = registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === email.toLowerCase());
                  const label = other ? ((other.Vorname || '') + ' ' + (other.Nachname || '')).trim() || other.ParticipantName || email : email;
                  return <td key={id} style={{ color: 'var(--dex-gray-700)' }} title={email}>{label}</td>;
                }
                if (id === 'replacedBy') {
                  // v17.15: „Ersetzt durch" — die Person die nach Cancel
                  // dieses Eintrags den Platz übernommen hat. Spiegelbild
                  // von „Ersetzt".
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const email = ((reg as any).ReplacedByParticipantEmail as string | undefined) || '';
                  if (!email) return <td key={id} style={{ color: 'var(--dex-gray-300)' }}>—</td>;
                  const other = registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === email.toLowerCase());
                  const label = other ? ((other.Vorname || '') + ' ' + (other.Nachname || '')).trim() || other.ParticipantName || email : email;
                  return <td key={id} style={{ color: 'var(--dex-green-dark, #4a7c1f)' }} title={email}>{label}</td>;
                }
                if (id === 'joinOrder') {
                  // v17.9 (deprecated): joinOrder-Spalte seit v17.10 entfernt.
                  return <td key={id}>—</td>;
                }
                if (id === 'registeredBy') {
                  return (
                    <td key={id} style={{ color: 'var(--dex-gray-600)' }}>
                      {(() => {
                        const actorEmail = (reg.RegisteredByEmail || '').toLowerCase();
                        const participantEmail = (reg.ParticipantEmail || '').toLowerCase();
                        if (!actorEmail) {
                          // v27.12 (Feedback Datenschutz-Review): Zeile wurde
                          // nicht über die App angelegt (RegisteredBy* leer) —
                          // dann den SP-Zeilen-Autor als Fallback zeigen.
                          const authorEmail = (reg.Author?.EMail || '').toLowerCase();
                          const authorName = (reg.Author?.Title || '').trim();
                          if (!authorEmail && !authorName) return <span style={{ color: 'var(--dex-gray-400)' }}>-</span>;
                          if (authorEmail && authorEmail === participantEmail) {
                            return <span style={{ color: 'var(--dex-green-dark)' }}>{isDe ? 'Selbst' : 'Self'}</span>;
                          }
                          return (
                            <span
                              title={isDe ? `${reg.Author?.EMail || ''} — aus den SharePoint-Metadaten (Zeile wurde nicht über die App angelegt)` : `${reg.Author?.EMail || ''} — from SharePoint metadata (row was not created via the app)`}
                              style={{ color: 'var(--dex-gray-500)', fontStyle: 'italic' }}
                            >
                              {authorName || reg.Author?.EMail}
                            </span>
                          );
                        }
                        if (actorEmail === participantEmail) {
                          return <span style={{ color: 'var(--dex-green-dark)' }}>{isDe ? 'Selbst' : 'Self'}</span>;
                        }
                        return (
                          <span title={reg.RegisteredByEmail || ''} style={{ color: 'var(--dex-orange)' }}>
                            {reg.RegisteredByName || reg.RegisteredByEmail}
                          </span>
                        );
                      })()}
                    </td>
                  );
                }
                if (id === 'team') {
                  // v16.1: Team-Name + Lead-Markierung. Wenn der TN in
                  // keinem Team ist, „—" anzeigen.
                  const tName = (reg.TeamName || '').trim();
                  const inTeam = !!reg.TeamId;
                  if (!inTeam) return <td key={id} style={{ color: 'var(--dex-gray-400)' }}>—</td>;
                  return (
                    <td key={id} style={{ color: 'var(--dex-gray-700)' }}>
                      {tName ? `„${tName}"` : <span style={{ color: 'var(--dex-gray-500)' }}>{isDe ? 'ohne Namen' : 'unnamed'}</span>}
                      {reg.TeamLead && (
                        <span className="dex-ui-pill dex-ui-pill--green" style={{ marginLeft: 6, padding: '1px 7px', fontSize: '0.66rem' }} title={isDe ? 'Team-Lead' : 'Team lead'}>Lead</span>
                      )}
                    </td>
                  );
                }
                if (id === 'roommate') {
                  return (
                    <td key={id}>
                      {(() => {
                        const info = getRoommateInfo(reg);
                        if (!info) return <span style={{ color: 'var(--dex-gray-300)' }}>-</span>;
                        const photoEmail = (info.partnerEmail || '').trim();
                        return (
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                            {/* v26.43: Foto mit Kontaktkarte (Hover/Tap) — wie die
                                anderen Personen-Fotos in der Tabelle; vorher ein
                                nacktes <img> ohne Mouse-over. */}
                            {photoEmail && (
                              <PersonContactHover email={photoEmail} name={info.partnerName} size={24} isDe={isDe} />
                            )}
                            <span>{info.partnerName}</span>
                            {info.mutual && (
                              <span
                                className="dex-ui-pill dex-ui-pill--green"
                                style={{ marginLeft: 2, padding: '1px 7px', fontSize: '0.68rem' }}
                                title={isDe ? 'Beide haben sich gegenseitig als Zimmerpartner ausgewählt' : 'Both selected each other as roommates'}
                              >
                                Match
                              </span>
                            )}
                          </span>
                        );
                      })()}
                    </td>
                  );
                }
                // v14.11: cfp-* sind Parent-Event-Custom-Fields (Pastel A) im
                // Sub-Event-Detail-View. Wert kommt entweder aus reg.CustomData
                // (Sub-Events erben i.d.R. die Parent-Felder via Wizard-Copy)
                // oder, falls leer, aus dem SP-Internal-Name-Property.
                const inSubEventDetailCell = !!parentEventForSelected;
                const pastelACell: React.CSSProperties = inSubEventDetailCell ? { background: 'rgba(0, 118, 168, 0.08)' } : {};
                const pastelBCell: React.CSSProperties = inSubEventDetailCell ? { background: 'rgba(255, 191, 0, 0.10)' } : {};
                if (id.indexOf('cfp-') === 0) {
                  const cfId = id.substring(4);
                  const field = (parentEventForSelected?.eventSpecificFields || []).find(f => f.id === cfId);
                  if (!field) return null;
                  // v15.14: Werte für Parent-Custom-Fields kommen primär aus
                  // der Parent-Event-Registrierung der Person (lookup per
                  // ParticipantEmail in parentRegsByEmail) — die Sub-Event-
                  // Registrierung enthält diese Antworten i.d.R. nicht. Nur
                  // wenn keine Parent-Reg existiert, fallen wir auf die Sub-
                  // Event-Daten zurück.
                  const emailKey = (reg.ParticipantEmail || '').toLowerCase().trim();
                  const parentReg = emailKey ? parentRegsByEmail[emailKey] : undefined;
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const spName = (field as any).spInternalName || '';
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  let val: any = undefined;
                  if (parentReg) {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    val = spName ? (parentReg as any)[spName] : undefined;
                    if ((val === undefined || val === null || val === '') && parentReg.CustomData) {
                      try {
                        const cd = JSON.parse(parentReg.CustomData);
                        val = cd[field.id];
                      } catch { /* no-op */ }
                    }
                  }
                  if (val === undefined || val === null || val === '') {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    val = spName ? (reg as any)[spName] : undefined;
                    if ((val === undefined || val === null || val === '') && reg.CustomData) {
                      try {
                        const cd = JSON.parse(reg.CustomData);
                        val = cd[field.id];
                      } catch { /* no-op */ }
                    }
                  }
                  let display: React.ReactNode = '-';
                  if (val !== undefined && val !== null && val !== '') {
                    if (field.type === 'checkbox') {
                      const truthy = val === true || val === 'true' || val === 1 || val === '1';
                      display = <span style={{ color: truthy ? 'var(--dex-green-dark)' : 'var(--dex-gray-400)' }}>{truthy ? '✓' : '–'}</span>;
                    } else if (field.type === 'select' && field.multi) {
                      display = highlightMatch(String(val).split(' | ').map(s => s.trim()).filter(Boolean).join(', '));
                    } else {
                      display = highlightMatch(String(val));
                    }
                  }
                  return (
                    <td key={id} style={{ color: 'var(--dex-gray-700)', whiteSpace: 'nowrap', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', ...pastelACell }} title={String(val || '')}>
                      {display}
                    </td>
                  );
                }
                if (id.indexOf('cf-') === 0) {
                  const cfId = id.substring(3);
                  const field = (selectedEvent?.eventSpecificFields || []).find(f => f.id === cfId);
                  if (!field) return null;
                  // v19.2: Dokument-Felder haben keinen Spaltenwert — die Datei
                  // liegt als Attachment. In der Spalte einen Download-Link (oder
                  // mehrere) zeigen, statt „-".
                  if (field.type === 'document') {
                    const att = attachmentsByReg[reg.Id] || [];
                    const prefix = `dxf-${(field.id || '').replace(/[^a-zA-Z0-9]/g, '')}--`;
                    const docs = att.filter(a => a.fileName.startsWith(prefix));
                    const pretty = (fn: string): string => fn
                      .replace(/^dxf-[a-zA-Z0-9]+--\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}_/, '')
                      .replace(/^dxf-[a-zA-Z0-9]+--/, '');
                    return (
                      <td key={id} style={{ whiteSpace: 'nowrap', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', ...pastelBCell }}>
                        {docs.length === 0 ? (
                          <span style={{ color: 'var(--dex-gray-400)' }}>–</span>
                        ) : (
                          docs.map((d, i) => (
                            <a
                              key={d.fileName}
                              href={d.serverRelativeUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              title={pretty(d.fileName)}
                              style={{ color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'underline', marginRight: i < docs.length - 1 ? 8 : 0, display: 'inline-flex', alignItems: 'center', gap: 3 }}
                            >
                              <FileText size={12} />{docs.length > 1 ? `${isDe ? 'Datei' : 'File'} ${i + 1}` : (isDe ? 'Datei' : 'File')}
                            </a>
                          ))
                        )}
                      </td>
                    );
                  }
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const spName = (field as any).spInternalName || '';
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  let val: any = spName ? (reg as any)[spName] : undefined;
                  if ((val === undefined || val === null || val === '') && reg.CustomData) {
                    try {
                      const cd = JSON.parse(reg.CustomData);
                      val = cd[field.id];
                    } catch { /* no-op */ }
                  }
                  let display: React.ReactNode = '-';
                  if (val !== undefined && val !== null && val !== '') {
                    if (field.type === 'checkbox') {
                      const truthy = val === true || val === 'true' || val === 1 || val === '1';
                      display = <span style={{ color: truthy ? 'var(--dex-green-dark)' : 'var(--dex-gray-400)' }}>{truthy ? '✓' : '–'}</span>;
                    } else if (field.type === 'select' && field.multi) {
                      // v7.11: Mehrfachauswahl wird " | "-getrennt gespeichert.
                      // In der Admin-Tabelle als Komma-Liste anzeigen, damit
                      // der Spalten-Inhalt sauberer scanbar ist.
                      display = highlightMatch(String(val).split(' | ').map(s => s.trim()).filter(Boolean).join(', '));
                    } else {
                      display = highlightMatch(String(val));
                    }
                  }
                  return (
                    <td key={id} style={{ color: 'var(--dex-gray-700)', whiteSpace: 'nowrap', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', ...pastelBCell }} title={String(val || '')}>
                      {display}
                    </td>
                  );
                }
                if (id === 'action') {
                  const att = attachmentsByReg[reg.Id] || [];
                  // v26.11: Ist das Event vorbei, geht es nur noch um die
                  // Anwesenheits-Nachpflege — KEIN „Bearbeiten" mehr, sondern
                  // No-Show / Einchecken / stille Abmeldung (ohne E-Mail).
                  const eventOver = !!selectedEvent && isEventOver(selectedEvent);
                  // v29.25: Selbst-Abmeldung nach der Frist gesperrt (Organizer-
                  // Option) — ab da pflegt der Organizer die Anwesenheit, deshalb
                  // steht neben angemeldeten Teilnehmern zusätzlich „No-Show".
                  // Flag liegt bei Sub-Events auf dem Parent.
                  const cancelLockActive = !eventOver && !!selectedEvent
                    && selfCancelLocked(selectedEvent, selectedEvent.parentEventId ? allEvents.find(pe => pe.id === selectedEvent.parentEventId) : undefined);
                  const doCheckIn = async (): Promise<void> => {
                    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                    await eventServiceRef.checkInParticipant(selectedEvent.subsiteUrl, reg.Id);
                    await reloadRegistrations();
                  };
                  const doCheckOut = async (): Promise<void> => {
                    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                    await eventServiceRef.checkOutParticipant(selectedEvent.subsiteUrl, reg.Id);
                    await reloadRegistrations();
                  };
                  // v29.25: Echter No-Show-Status (wie auf der Check-in-Seite,
                  // v23.28) — nicht zu verwechseln mit dem „No-Show"-Knopf der
                  // Nach-Event-Pflege, der nur den Check-in zurücknimmt.
                  // Ältere Teilnehmerlisten kennen die Choice nicht (HTTP 400).
                  const doMarkNoShow = async (): Promise<void> => {
                    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                    const ok = await eventServiceRef.markNoShowParticipant(selectedEvent.subsiteUrl, reg.Id);
                    if (!ok) {
                      showAlert(isDe
                        ? 'No-Show konnte nicht gesetzt werden. Bei Events, die vor v23.28 angelegt wurden, kennt die Teilnehmerliste den Status „No-Show" noch nicht.'
                        : 'Could not set no-show. For events created before v23.28 the attendee list does not know the “No-Show” status yet.',
                        { variant: 'error' });
                      return;
                    }
                    await reloadRegistrations();
                  };
                  // v31.3: Aktionen nach Leitfaden 5b — Symbol-Knöpfe für
                  // Bearbeiten, Dateien und Abmelden (rot), die Anwesenheits-
                  // Knöpfe bleiben beschriftet: „Einchecken", „Auschecken",
                  // „No-Show" und „Zurücksetzen" sind Zustandswörter, die kein
                  // Symbol eindeutig trägt. Kein `btn-primary` mehr je Zeile
                  // (Leitfaden 1.5: ein Primär-Knopf je Ansicht); Einchecken
                  // trägt stattdessen das Häkchen.
                  const smBtn = 'btn btn-secondary dex-ui-btn-sm';
                  const checkInLabel = <><Check size={12} /> {isDe ? 'Einchecken' : 'Check in'}</>;
                  const hasFileColumn = selectedEvent?.allowAttendeeUpload || (selectedEvent?.eventSpecificFields || []).some(f => f.type === 'document');
                  return (
                    <td key={id} className="is-actions">
                      <div style={{ display: 'inline-flex', gap: 4, flexWrap: 'wrap', justifyContent: 'flex-end', alignItems: 'center' }}>
                      {/* „Bearbeiten" nur, solange das Event noch nicht vorbei ist. */}
                      {!eventOver && (
                        <button
                          type="button"
                          className="dex-ui-iconbtn"
                          aria-label={isDe ? 'Bearbeiten' : 'Edit'}
                          title={isDe ? 'Teilnehmer-Daten bearbeiten' : 'Edit attendee data'}
                          onClick={() => openEditModal(reg)}
                        >
                          <Pencil size={15} />
                        </button>
                      )}
                      {/* v11.0: Anhang-Button — wenn das Event den Teilnehmer-
                          Upload erlaubt ODER ein Dokument-Custom-Feld hat (v19.0).
                          Zeigt Counter wenn mind. eine Datei hochgeladen wurde. */}
                      {hasFileColumn && (
                        <button
                          type="button"
                          className="dex-ui-iconbtn"
                          style={att.length > 0 ? { width: 'auto', padding: '0 8px', borderRadius: 16, gap: 3, fontSize: '0.74rem', fontWeight: 700 } : undefined}
                          aria-label={isDe ? 'Dateien' : 'Files'}
                          title={att.length > 0
                            ? (isDe ? `${att.length} hochgeladene Datei(en) anzeigen` : `Show ${att.length} uploaded file(s)`)
                            : (isDe ? 'Hochgeladene Dateien anzeigen — noch keine Datei' : 'Show uploaded files — none yet')}
                          onClick={() => setAttachmentsModalReg(reg)}
                        >
                          <FileText size={15} />
                          {att.length > 0 ? att.length : null}
                        </button>
                      )}
                      {eventOver ? (
                        <>
                          {/* Nach dem Event: Anwesenheit explizit pflegen. */}
                          <button
                            type="button"
                            className={smBtn}
                            disabled={reg.Status === 'Eingecheckt'}
                            title={isDe ? 'Als anwesend markieren' : 'Mark as attended'}
                            onClick={doCheckIn}
                          >
                            {checkInLabel}
                          </button>
                          <button
                            type="button"
                            className={smBtn}
                            disabled={reg.Status !== 'Eingecheckt'}
                            title={isDe ? 'Als nicht erschienen markieren' : 'Mark as no-show'}
                            onClick={doCheckOut}
                          >
                            No-Show
                          </button>
                        </>
                      ) : cancelLockActive ? (
                        <>
                          {/* v29.25: Abmelde-Sperre aktiv — ab jetzt pflegt der
                              Organizer die Anwesenheit: Einchecken/Auschecken wie
                              bisher, daneben der echte No-Show-Status (v23.28,
                              wie auf der Check-in-Seite). */}
                          {reg.Status === 'Eingecheckt' ? (
                            <button type="button" className={smBtn} title={isDe ? 'Check-in zurücknehmen' : 'Undo check-in'} onClick={doCheckOut}>
                              {isDe ? 'Auschecken' : 'Check out'}
                            </button>
                          ) : (
                            <button type="button" className={smBtn} title={isDe ? 'Als anwesend markieren' : 'Mark as attended'} onClick={doCheckIn}>
                              {checkInLabel}
                            </button>
                          )}
                          <button
                            type="button"
                            className={smBtn}
                            disabled={reg.Status === 'No-Show' || reg.Status === 'Abgemeldet'}
                            title={isDe
                              ? 'Als nicht erschienen markieren. Die Selbst-Abmeldung ist für Teilnehmer aktuell deaktiviert — wer absagt, wird hier abgemeldet oder als No-Show markiert.'
                              : 'Mark as a no-show. Self-cancellation is currently disabled for attendees — cancel people here or mark them as no-shows.'}
                            onClick={() => { void doMarkNoShow(); }}
                          >
                            No-Show
                          </button>
                          {reg.Status === 'No-Show' && (
                            <button
                              type="button"
                              className={smBtn}
                              title={isDe ? 'No-Show zurücknehmen (Status zurück auf „Angemeldet")' : 'Undo no-show (status back to “registered”)'}
                              onClick={doCheckOut}
                            >
                              {isDe ? 'Zurücksetzen' : 'Reset'}
                            </button>
                          )}
                        </>
                      ) : (
                        reg.Status === 'Eingecheckt' ? (
                          <button type="button" className={smBtn} title={isDe ? 'Check-in zurücknehmen' : 'Undo check-in'} onClick={doCheckOut}>
                            {isDe ? 'Auschecken' : 'Check out'}
                          </button>
                        ) : (
                          <button type="button" className={smBtn} title={isDe ? 'Als anwesend markieren' : 'Mark as attended'} onClick={doCheckIn}>
                            {checkInLabel}
                          </button>
                        )
                      )}
                      {/* Abmelden: nach dem Event still & ohne E-Mail (v22.22) —
                          deshalb auch für Organizer eigener Events freigegeben. */}
                      {(eventOver || !orgPastLock) && (
                      <button
                        type="button"
                        /* v31.21: Beschrifteter Knopf statt nacktem Icon
                           (Nutzer-Fragen 11.09.2026: „warum ist der Mülleimer
                           das Symbol für Abmelden?" und „warum gibt es keinen
                           Abmelden-Button?").
                           Beide Fragen haben dieselbe Ursache: Neben „No-Show",
                           „Einchecken" und „Auschecken" — drei beschrifteten
                           Knöpfen — stand die vierte, folgenreichste Aktion als
                           stummes Symbol. Ein Icon ohne Wort liest sich in
                           dieser Reihe nicht als Knopf, sondern als Zierrat;
                           und ausgerechnet der Mülleimer sagt „löschen",
                           während die Abmeldung nichts löscht (die Zeile bleibt
                           als `Abgemeldet` stehen und ist reaktivierbar). */
                        className={smBtn}
                        /* Dieselbe Kachelform wie die Nachbarn, aber in Rot:
                           `btn-danger` wäre ein dunkler Vollton und damit der
                           lauteste Knopf der Zeile — 426-mal untereinander.
                           Die Abmeldung ist folgenreich, aber nicht der
                           Hauptweg. Keine neue Klasse in `dexUi.ts`, weil es
                           genau eine Stelle ist und keinen eigenen
                           Hover-Zustand braucht (den liefert `btn-secondary`). */
                        style={{ color: 'var(--dex-red, #da291c)' }}
                        aria-label={eventOver ? (isDe ? 'Abmelden (ohne E-Mail)' : 'Cancel (no email)') : (isDe ? 'Abmelden' : 'Cancel registration')}
                        title={eventOver
                          ? (isDe ? 'Abmelden — ohne E-Mail, das Event ist vorbei' : 'Cancel registration — no email, the event is over')
                          : (isDe ? 'Abmelden — mit Rückfrage' : 'Cancel registration — asks first')}
                        onClick={async () => {
                          if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                          // v23.2: Doppel-Anmeldung? Statt direkt abzumelden das
                          // Duplikat-Modal öffnen (still löschen vs. normal abmelden).
                          if (duplicateEmails.has((reg.ParticipantEmail || '').trim().toLowerCase())) { setDupCancelReg(reg); return; }
                          const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                          // v22.22: Vergangenes Event → stille Abmeldung (keine
                          // Abmelde-Mail, keine Outlook-Absage, kein Nachrücken,
                          // kein ID-Reorder). Der Confirm sagt das explizit.
                          const confirmMsg = eventOver
                            ? (isDe
                              ? `${name} (${reg.ParticipantEmail}) wirklich abmelden?\n\nDas Event liegt in der Vergangenheit — die Abmeldung läuft still: Es gehen keine Abmelde-Mail und keine Outlook-Absage raus, und es rückt niemand von der Warteliste nach.`
                              : `Really cancel ${name} (${reg.ParticipantEmail})?\n\nThe event is in the past — the cancellation runs silently: no cancellation email, no Outlook removal, and nobody is promoted from the waitlist.`)
                            : (isDe ? `${name} (${reg.ParticipantEmail}) wirklich abmelden?` : `Really cancel ${name} (${reg.ParticipantEmail})?`);
                          if (!(await confirmDialog(confirmMsg, { danger: true, confirmLabel: isDe ? 'Abmelden' : 'Cancel registration' }))) return;
                          await performStandardCancel(reg);
                        }}
                      >
                        <UserMinus size={14} /> {isDe ? 'Abmelden' : 'Cancel'}
                      </button>
                      )}
                      {/* v26.47: Externe Anmeldung mit offener Datenschutz-
                          Rückmeldung — die App kann keine externen Adressen
                          anmailen, deshalb lädt die anmeldende Person die
                          Einladung als .eml-Entwurf herunter und verschickt sie
                          selbst; die Rückmeldung wird danach hier bestätigt. */}
                      {reg.ConsentReview === 'Pending' && (() => {
                        const fullName = `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantName;
                        return (
                          <>
                            <button
                              type="button"
                              className={smBtn}
                              style={{ color: '#b35a00' }}
                              title={isDe
                                ? 'Einladungs-Mail als .eml-Entwurf herunterladen — in Outlook öffnen und selbst an die externe Person senden.'
                                : 'Download the invitation email as an .eml draft — open it in Outlook and send it to the external person yourself.'}
                              onClick={() => {
                                if (!selectedEvent) return;
                                const mailDe = (selectedEvent.emailLanguage || 'EN').toUpperCase() === 'DE';
                                const { subject, body } = externalInvitationEmail(
                                  fullName,
                                  selectedEvent.title,
                                  reg.RegisteredByName || '',
                                  mailDe,
                                  { startDate: selectedEvent.startDate, endDate: selectedEvent.endDate, location: selectedEvent.location }
                                );
                                const eml = buildUnsentEmlDraft({
                                  to: [reg.ParticipantEmail],
                                  cc: ['no_reply.events@deloitte.de', ...Array.from(new Set([...(selectedEvent.organizerEmails || []), ...(selectedEvent.coOrganizerEmails || [])].filter(Boolean)))],
                                  subject,
                                  html: body,
                                });
                                downloadEml('Einladung_' + (reg.ParticipantEmail || 'extern'), eml);
                              }}
                            >
                              {isDe ? 'Einladung (.eml)' : 'Invitation (.eml)'}
                            </button>
                            <button
                              type="button"
                              className={smBtn}
                              style={{ color: '#b35a00' }}
                              title={isDe
                                ? 'Bestätigen, dass die externe Person auf die Datenschutz-Einladung geantwortet hat.'
                                : 'Confirm that the external person has responded to the privacy invitation.'}
                              onClick={async () => {
                                if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                                const ok = await eventServiceRef.confirmConsentReview(
                                  selectedEvent.subsiteUrl,
                                  reg.Id,
                                  { eventId: selectedEvent.id, eventTitle: selectedEvent.title, participantName: fullName }
                                );
                                await reloadRegistrations();
                                if (ok) {
                                  showAlert(
                                    isDe ? `Datenschutz-Rückmeldung von ${fullName} bestätigt.` : `Privacy confirmation of ${fullName} recorded.`,
                                    { variant: 'success' }
                                  );
                                } else {
                                  showAlert(
                                    isDe ? 'Bestätigen fehlgeschlagen — bitte erneut versuchen.' : 'Confirmation failed — please try again.',
                                    { variant: 'error' }
                                  );
                                }
                              }}
                            >
                              {isDe ? 'Rückmeldung bestätigen' : 'Confirm response'}
                            </button>
                          </>
                        );
                      })()}
                      </div>
                    </td>
                  );
                }
                return null;
              };

              return (
                <>
                  {/* v6.17: Kontrollzeile mit Column-Picker-Button. Der Popover
                      zeigt alle verfügbaren Spalten inkl. Checkbox zum Ein-/
                      Ausblenden und Pfeilen zum Umsortieren. Die Config wird
                      pro Event in localStorage persistiert (s. useEffect oben). */}
                  {/* v31.3: EINE Werkzeugleiste über der Tabelle (Leitfaden 5a.6):
                      links die Ansicht-Wahl bei geteilten Gruppen (gehört zum
                      Inhalt), rechts die Spaltenwahl (Leitfaden: „Spaltenwahl und
                      Export rechts"). Vorher zwei getrennte, rechtsbündige Zeilen. */}
                  <div className="dex-ui-toolbar" style={{ position: 'relative' }}>
                    {isSplitCapacity && (
                      <SplitMergeToggle view={splitParticipantsView} setView={setSplitParticipantsView} isDe={isDe} />
                    )}
                    <span className="dex-ui-toolbar-spacer" />
                    <button
                      type="button"
                      className={cx('dex-ui-chip', showColumnPicker && 'is-active')}
                      aria-expanded={showColumnPicker}
                      title={isDe ? 'Spalten ein- oder ausblenden und umsortieren' : 'Show, hide and reorder columns'}
                      onClick={() => setShowColumnPicker(!showColumnPicker)}
                    >
                      <Columns size={14} /> {isDe ? 'Spalten anpassen' : 'Customize columns'}
                    </button>
                    {showColumnPicker && (
                      <div
                        className="dex-ui-card dex-ui-card--list dex-ui-fade-in"
                        style={{
                          position: 'absolute', right: 0, top: '100%', marginTop: 6,
                          boxShadow: '0 6px 20px rgba(0,0,0,0.10)',
                          // v31.10: 300 px sind auf einem 360-px-Schirm mehr, als
                          // neben den Rändern übrig bleibt — die Liste ragte dann
                          // links aus dem Bild. `maxWidth` deckelt sie auf die
                          // Fensterbreite; auf dem Rechner bleibt es bei 300.
                          width: 300, maxWidth: 'calc(100vw - 32px)',
                          zIndex: 100, maxHeight: 400, overflowY: 'auto',
                        }}
                      >
                        <div style={{ padding: '6px 8px 8px' }}>
                          <div style={{ fontSize: '0.86rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>
                            {isDe ? 'Welche Spalten zeigt die Tabelle?' : 'Which columns does the table show?'}
                          </div>
                          <div className="dex-ui-help" style={{ marginTop: 2 }}>
                            {isDe ? 'Haken = sichtbar. Die Pfeile ändern die Reihenfolge; die Auswahl bleibt je Event gespeichert.' : 'Tick = visible. The arrows change the order; the selection is saved per event.'}
                          </div>
                        </div>
                        {columnOrder.map((id, idx) => {
                          const col = availableColumns.find(c => c.id === id);
                          if (!col) return null;
                          const isHidden = hiddenColumns.indexOf(id) >= 0;
                          const isVisible = !isHidden;
                          const canMoveUp = isVisible && idx > 0 && columnOrder[idx - 1] !== undefined;
                          // "action" bleibt immer letzte → niemand darf unter "action" wandern
                          // und "action" selbst darf nicht verschoben werden.
                          const nextId = columnOrder[idx + 1];
                          const canMoveDown = isVisible && idx < columnOrder.length - 1 && id !== 'action' && nextId !== 'action';
                          const cbId = `pt-col-${id}`;
                          return (
                            <div
                              key={id}
                              className={cx('dex-ui-row', !isVisible && 'is-done')}
                              style={{ padding: '4px 8px', gap: 8 }}
                            >
                              <input
                                type="checkbox"
                                // v31.11: Ohne diese Klasse rendert das globale
                                // SCSS der App die Checkbox rund mit Punkt — das
                                // Bild fuer „nur eins davon", waehrend hier
                                // mehrere Spalten gleichzeitig sichtbar sind.
                                className="dex-ui-checkbox"
                                id={cbId}
                                checked={isVisible}
                                disabled={!!col.alwaysVisible}
                                onChange={() => {
                                  if (col.alwaysVisible) return;
                                  if (isHidden) showColumn(id); else hideColumn(id);
                                }}
                                style={{ cursor: col.alwaysVisible ? 'not-allowed' : 'pointer', margin: 0 }}
                                title={col.alwaysVisible
                                  ? (isDe ? 'Pflicht-Spalte — kann nicht ausgeblendet werden' : 'Required column — cannot be hidden')
                                  : (isHidden ? (isDe ? 'Einblenden' : 'Show') : (isDe ? 'Ausblenden' : 'Hide'))}
                              />
                              <label htmlFor={cbId} className="dex-ui-row-main" style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)', cursor: col.alwaysVisible ? 'default' : 'pointer' }}>{col.label}</label>
                              <span className="dex-ui-row-actions">
                                {/* v31.10: 26 px waren mit der Maus knapp und mit
                                    dem Finger zu wenig — die Pfeile liegen 4 px
                                    nebeneinander. Jetzt das Standardmaß der
                                    Klasse (32 px), die Zeile wächst dadurch nur
                                    um wenige Pixel. */}
                                <button
                                  type="button"
                                  className="dex-ui-iconbtn"
                                  onClick={() => moveColumn(id, -1)}
                                  disabled={!canMoveUp}
                                  aria-label={isDe ? 'Spalte nach oben' : 'Move column up'}
                                  title={isDe ? 'Nach oben' : 'Up'}
                                >
                                  <ChevronUp size={14} />
                                </button>
                                <button
                                  type="button"
                                  className="dex-ui-iconbtn"
                                  onClick={() => moveColumn(id, 1)}
                                  disabled={!canMoveDown}
                                  aria-label={isDe ? 'Spalte nach unten' : 'Move column down'}
                                  title={isDe ? 'Nach unten' : 'Down'}
                                >
                                  <ChevronDown size={14} />
                                </button>
                              </span>
                            </div>
                          );
                        })}
                        <div style={{ padding: '8px 8px 4px', display: 'flex', justifyContent: 'flex-end' }}>
                          <button
                            type="button"
                            className="btn btn-secondary dex-ui-btn-sm"
                            onClick={() => setShowColumnPicker(false)}
                          >
                            {isDe ? 'Schließen' : 'Close'}
                          </button>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* v11.98: Split-/Merged-Toggle bei Split-Kapazität.
                      Default 'split' — getrennte Tabellen pro Gruppe,
                      kleinere zuerst. */}
                  {(() => {
                    // v31.3: `footLabel` — die Gruppen-Tabellen der Split-Ansicht
                    // zählen „in dieser Gruppe", die Gesamttabelle „N von M".
                    const renderTable = (rows: SPRegistration[], indexOffset: number, footLabel?: string): React.ReactElement => {
                      // v26.44: eine normale Teilnehmer-Zeile — herausgezogen, damit
                      // die „Matches anzeigen"-Gruppierung dieselbe Zeilen-JSX
                      // wiederverwenden kann (keine Duplizierung der Zellen-Logik).
                      const renderBodyRow = (reg: SPRegistration, i: number): React.ReactElement => {
                        const isOverbook = reg.OverbookReview === 'Pending';
                            // v22.44: Inaktive Deloitte-Konten dauerhaft orange
                            // markieren (bis zur Abmeldung) — gleiche Optik wie
                            // die Überbuchungs-Markierung. inactiveAccounts kommt
                            // aus dem Konten-Aktiv-Check (nur @deloitte-Adressen).
                            const isInactiveAcct = inactiveAccounts.indexOf((reg.ParticipantEmail || '').trim().toLowerCase()) >= 0;
                            // v23.2: Doppel-Anmeldung — rote Markierung (hat Vorrang
                            // vor der orangen Überbuchungs-/Inaktiv-Markierung).
                            const isDuplicate = (reg.Status || '') !== 'Abgemeldet'
                              && duplicateEmails.has((reg.ParticipantEmail || '').trim().toLowerCase());
                            const highlight = isOverbook || isInactiveAcct;
                            const rowTitle = isDuplicate
                              ? (isDe ? 'Doppel-Anmeldung — diese Person ist mehrfach angemeldet. Über „Abmelden" lässt sich die doppelte Zeile still entfernen.' : 'Duplicate registration — this person is registered more than once. Use „Cancel" to silently remove the duplicate row.')
                              : isInactiveAcct
                              ? (isDe ? 'Kein aktives Deloitte-Konto gefunden — Person hat womöglich Deloitte verlassen. Mails/Outlook kommen ggf. nicht an.' : 'No active Deloitte account found — person may have left Deloitte. Emails/Outlook may not arrive.')
                              : isOverbook
                                ? (isDe ? 'Über Kapazität angemeldet — siehe Box „Überbuchung – zu prüfen" oben' : 'Registered over capacity — see the „Overbooking – to review" box above')
                                : undefined;
                            // v31.3: Die Markierung liegt auf den ZELLEN, nicht auf
                            // der Zeile: `dex-ui-table` färbt beim Überfahren die
                            // td-Hintergründe, und die würden eine Zeilenfarbe
                            // verdecken — eine Dublette sähe unter der Maus wie
                            // eine normale Zeile aus. Zellen mit eigenem Grund
                            // (Pastell der Custom-Felder) behalten ihn.
                            const mark = isDuplicate
                              ? { bg: 'rgba(200,0,0,0.10)', bar: 'var(--dex-red, #c00)' }
                              : highlight
                              ? { bg: 'rgba(237,139,0,0.13)', bar: 'var(--dex-orange, #ed8b00)' }
                              : null;
                            const cells = effectiveColumnIds.map(id => renderCell(id, reg, i));
                            const markedCells = mark
                              ? cells.map((c, ci) => {
                                if (!React.isValidElement(c)) return c;
                                const el = c as React.ReactElement<{ style?: React.CSSProperties }>;
                                const own = el.props.style || {};
                                return React.cloneElement(el, { style: { ...own, background: own.background || mark.bg, ...(ci === 0 ? { boxShadow: `inset 3px 0 0 ${mark.bar}` } : {}) } });
                              })
                              : cells;
                            return (
                              <tr key={reg.Id} title={rowTitle}>
                                {markedCells}
                              </tr>
                            );
                      };

                      // v26.44: „Matches anzeigen" — Anzeige-Zeilen mit eingestreuten
                      // Gruppen-Header-Zeilen: [Match 1, A, B, Match 2, C, D, …,
                      // Rest-Header, …Rest]. Paare kommen aus computeRoommatePairs
                      // über die AKTUELL gefilterten rows (Suche wirkt also weiter);
                      // die Gruppierung übersteuert solange die normale Spalten-
                      // Sortierung (der Rest-Cluster behält die aktuelle
                      // Sortierreihenfolge). Toggle aus → displayRows = null →
                      // Rendering exakt wie bisher.
                      type DisplayRow = { header: string; muted?: boolean } | { reg: SPRegistration };
                      const displayRows: DisplayRow[] | null = (() => {
                        if (!showMatches || !hasRoommateColumn) return null;
                        const pairs = computeRoommatePairs(rows);
                        const inPair = new Set<string>();
                        for (const [a, b] of pairs) {
                          inPair.add((a.ParticipantEmail || '').trim().toLowerCase());
                          inPair.add((b.ParticipantEmail || '').trim().toLowerCase());
                        }
                        const rest = rows.filter(r => !inPair.has((r.ParticipantEmail || '').trim().toLowerCase()));
                        const nameOf = (r: SPRegistration): string =>
                          `${r.Vorname || ''} ${r.Nachname || ''}`.trim() || r.ParticipantName || r.ParticipantEmail || '';
                        const out: DisplayRow[] = [];
                        pairs.forEach(([a, b], pi) => {
                          out.push({ header: `Match ${pi + 1}: ${nameOf(a)} & ${nameOf(b)}` });
                          out.push({ reg: a });
                          out.push({ reg: b });
                        });
                        out.push({
                          header: isDe
                            ? `Ohne Preferred Roommate oder Match (${rest.length})`
                            : `Without preferred roommate or match (${rest.length})`,
                          muted: true,
                        });
                        for (const r of rest) out.push({ reg: r });
                        return out;
                      })();
                      let matchRowIdx = 0;
                      // v31.3: Fußzeile „N von M" statt Zähler im Kopf (Leitfaden 5b).
                      const shown = rows.length;
                      const footText = footLabel
                        || (query
                          ? (isDe ? `${shown} von ${activeTotal} Teilnehmern · Suche „${query}“` : `${shown} of ${activeTotal} participants · search “${query}”`)
                          : (isDe ? `${shown} von ${activeTotal} Teilnehmern` : `${shown} of ${activeTotal} participants`));
                      return (
                        // v24.96: eigener Scroll-Container um die Tabelle → der
                        // thead (position:sticky top:0) klebt zuverlässig an dessen
                        // oberem Rand (CSS-sticky relativ zu DIESEM Container, nicht
                        // zum Fenster — Letzteres ist im SP-Canvas unzuverlässig).
                        // v31.3: `dex-ui-table-wrap--sticky` ist genau dieser Container
                        // (max-height 70vh, overflow auto, th sticky); der äußere
                        // `dex-ui-table-wrap` gibt Rand und Radius, die Fußzeile
                        // bleibt außerhalb des Scrollbereichs sichtbar.
                        <div className="dex-ui-table-wrap">
                        <div className="dex-ui-table-wrap--sticky">
                        <table className="dex-ui-table dex-ui-table--compact">
                          <thead>
                            <tr>
                              {effectiveColumnIds.map(id => renderHeader(id))}
                            </tr>
                          </thead>
                          <tbody>
                            {displayRows
                              ? displayRows.map((dr, di) => ('header' in dr)
                                ? (
                                  <tr key={`match-grp-${di}`}>
                                    <td
                                      colSpan={effectiveColumnIds.length}
                                      style={{
                                        fontWeight: 700, fontSize: '0.76rem', textTransform: 'uppercase', letterSpacing: '0.04em',
                                        background: dr.muted ? 'var(--dex-gray-100, #f3f4f6)' : 'rgba(134,188,37,0.10)',
                                        color: dr.muted ? 'var(--dex-gray-600)' : 'var(--dex-green-dark, #4a7c1f)',
                                      }}
                                    >
                                      {dr.header}
                                    </td>
                                  </tr>
                                )
                                : renderBodyRow(dr.reg, indexOffset + (matchRowIdx++)))
                              : rows.map((reg, i) => renderBodyRow(reg, indexOffset + i))}
                          </tbody>
                        </table>
                        </div>
                        <div className="dex-ui-table-foot">
                          <span>{footText}</span>
                          {personalColsCollapsed && (
                            <span>{isDe ? 'Name anklicken öffnet die Details' : 'Click a name to open the details'}</span>
                          )}
                        </div>
                        </div>
                      );
                    };

                    // v31.3: Der Ansicht-Umschalter sitzt seit dem Umbau in der
                    // Werkzeugleiste oben — hier nur noch die Tabelle(n).
                    if (!isSplitCapacity || splitParticipantsView === 'merged') {
                      return renderTable(activeRegs, 0);
                    }

                    // Split-View: nach Gruppe trennen (StarterType ||
                    // PreferredStarterType), kleinere Gruppe zuerst.
                    const lblA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
                    const lblB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
                    const groupA = activeRegs.filter(r => (r.StarterType || r.PreferredStarterType) === 'Durchstarter');
                    const groupB = activeRegs.filter(r => (r.StarterType || r.PreferredStarterType) === 'Funstarter');
                    const groupNone = activeRegs.filter(r => !(r.StarterType || r.PreferredStarterType));
                    const groups = [
                      { label: lblA, key: 'A', rows: groupA, cap: selectedEvent?.durchstarterCapacity || 0 },
                      { label: lblB, key: 'B', rows: groupB, cap: selectedEvent?.funstarterCapacity || 0 },
                    ].sort((x, y) => x.rows.length - y.rows.length);
                    let runningIdx = 0;
                    // v31.3: Gruppen-Kopf als `dex-ui-card-head` (Titel + Zähler),
                    // leere Gruppe als `dex-ui-empty` statt kursivem Satz.
                    const groupFoot = (n: number): string => isDe ? `${n} in dieser Gruppe` : `${n} in this group`;
                    return (
                      <>
                        {groups.map(g => {
                          const offset = runningIdx;
                          runningIdx += g.rows.length;
                          return (
                            <div key={g.key} style={{ marginBottom: 20 }}>
                              <div className="dex-ui-card-head" style={{ marginBottom: 8 }}>
                                <h4 className="dex-ui-card-head-title" style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.95rem' }}>{g.label}</h4>
                                <span className="dex-ui-card-head-meta">
                                  {g.rows.length}{g.cap > 0 ? ` / ${g.cap} ${isDe ? 'Plätze' : 'seats'}` : ''}
                                </span>
                              </div>
                              {g.rows.length === 0 ? (
                                <div className="dex-ui-empty" style={{ padding: '16px 12px' }}>
                                  {isDe ? 'Keine Teilnehmer in dieser Gruppe.' : 'No participants in this group.'}
                                </div>
                              ) : renderTable(g.rows, offset, groupFoot(g.rows.length))}
                            </div>
                          );
                        })}
                        {groupNone.length > 0 && (
                          <div style={{ marginBottom: 20 }}>
                            <div className="dex-ui-card-head" style={{ marginBottom: 8 }}>
                              <h4 className="dex-ui-card-head-title" style={{ color: 'var(--dex-gray-500)', fontSize: '0.95rem' }}>{isDe ? 'Ohne Gruppe' : 'No group'}</h4>
                              <span className="dex-ui-card-head-meta">{groupNone.length}</span>
                            </div>
                            {renderTable(groupNone, runningIdx, groupFoot(groupNone.length))}
                          </div>
                        )}
                      </>
                    );
                  })()}
                </>
              );
            })()}
          </div>
  );
};

