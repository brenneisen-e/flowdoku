/* ParticipantTable — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 11711-12888 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { EventService, SPRegistration } from '../../../services/EventService';
import { PersonContactHover } from '../../PersonContactHover';
import { formatDate, translateStatus } from '../../../utils/eventStatus';
import { FileText, Pencil } from '../../Icons';
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
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
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
  setRegistrations: React.Dispatch<React.SetStateAction<SPRegistration[]>>;
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
  const { activeRegs, allEvents, attachmentsByReg, availableColumns, colToggleHover, columnOrder, computeRoommatePairs, confirmDialog, duplicateEmails, eventServiceRef, getAllRegistrations, getRoommateInfo, handleSort, hasRoommateColumn, hiddenColumns, hideColumn, highlightMatch, inactiveAccounts, isDe, isSplitCapacity, moveColumn, openEditModal, orgPastLock, parentEventForSelected, parentRegsByEmail, performStandardCancel, personalColsCollapsed, query, registrations, selectedEvent, setAttachmentsModalReg, setColToggleHover, setDupCancelReg, setParticipantDetail, setPersonalColsCollapsed, setRegistrations, setShowColumnPicker, setSplitParticipantsView, showAlert, showColumn, showColumnPicker, showMatches, sortIcon, splitParticipantsView, stripLocPrefix } = p;
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
                return (
                  <button
                    type="button"
                    onClick={(e) => { e.stopPropagation(); hideColumn(id); }}
                    aria-label={isDe ? `Spalte ${col.label} ausblenden` : `Hide column ${col.label}`}
                    title={isDe ? 'Spalte ausblenden' : 'Hide column'}
                    style={{
                      marginLeft: 6, padding: 0, width: 16, height: 16, lineHeight: '14px',
                      border: 'none', background: 'transparent', cursor: 'pointer',
                      color: 'var(--dex-gray-400)', fontSize: '0.8rem', borderRadius: 3,
                    }}
                    onMouseEnter={(e) => { e.currentTarget.style.color = 'var(--dex-red, #c00)'; }}
                    onMouseLeave={(e) => { e.currentTarget.style.color = 'var(--dex-gray-400)'; }}
                  >
                    ✕
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
                const baseStyle: React.CSSProperties = {
                  textAlign: 'left', padding: 8,
                  whiteSpace: 'normal',
                  overflowWrap: 'break-word',
                  hyphens: 'auto',
                  maxWidth: 180,
                  verticalAlign: 'top',
                  lineHeight: 1.3,
                  // v24.96: Echtes Sticky über einen EIGENEN Scroll-Container um
                  // die Tabelle (renderTable wrappt die Tabelle in ein div mit
                  // maxHeight + overflow:auto). `top: 0` klebt dann zuverlässig am
                  // oberen Rand DIESES Containers — anders als CSS-sticky relativ
                  // zum Fenster, das im SharePoint-Canvas verrutscht.
                  position: 'sticky',
                  top: 0,
                  background: '#fff',
                  zIndex: 5,
                  borderBottom: '2px solid var(--dex-gray-200)',
                };
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
                  return (
                    <th key="person" style={{ ...baseStyle, whiteSpace: 'nowrap', userSelect: 'none', cursor: 'pointer' }} onClick={() => handleSort('nachname')}>
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 3 }}>
                        <button
                          type="button"
                          onClick={(e) => { e.stopPropagation(); setPersonalColsCollapsed(false); setColToggleHover(false); }}
                          onMouseEnter={() => setColToggleHover(true)}
                          onMouseLeave={() => setColToggleHover(false)}
                          title={isDe ? 'Vorname, Nachname, E-Mail, Job Title, Standort und Unternehmen als eigene Spalten anzeigen' : 'Show first/last name, email, job title, location and company as separate columns'}
                          style={{
                            display: 'inline-flex', alignItems: 'center', gap: 4,
                            padding: '2px 10px', borderRadius: 999,
                            border: '1px solid var(--dex-green, #86bc25)',
                            background: colToggleHover ? 'var(--dex-green, #86bc25)' : '#fff',
                            color: colToggleHover ? '#fff' : 'var(--dex-green-dark, #4a7c1f)',
                            fontSize: '0.7rem', fontWeight: 700, cursor: 'pointer', lineHeight: 1.4,
                            transition: 'background 120ms ease, color 120ms ease',
                          }}
                        >» {isDe ? 'Aufklappen' : 'Expand'}</button>
                        <span>{isDe ? 'Teilnehmer' : 'Participant'}{sortIcon('nachname')}</span>
                      </div>
                    </th>
                  );
                }
                const sortable = sortableCols[id];
                if (sortable) {
                  return (
                    <th
                      key={id}
                      style={{ ...baseStyle, cursor: 'pointer', userSelect: 'none' }}
                      onClick={() => handleSort(sortable)}
                    >
                      {/* v30.60: „Vorname" trägt den Zuklapp-Knopf als Pille in
                          einer eigenen Zeile — Gegenstück zum Aufklappen oben. */}
                      {id === 'vorname' ? (
                        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 3 }}>
                          <button
                            type="button"
                            onClick={(e) => { e.stopPropagation(); setPersonalColsCollapsed(true); setColToggleHover(false); }}
                            onMouseEnter={() => setColToggleHover(true)}
                            onMouseLeave={() => setColToggleHover(false)}
                            title={isDe ? 'Personen-Spalten einklappen (nur Foto + Name)' : 'Collapse personal columns (photo + name only)'}
                            style={{
                              display: 'inline-flex', alignItems: 'center', gap: 4,
                              padding: '2px 10px', borderRadius: 999,
                              border: '1px solid var(--dex-green, #86bc25)',
                              background: colToggleHover ? 'var(--dex-green, #86bc25)' : '#fff',
                              color: colToggleHover ? '#fff' : 'var(--dex-green-dark, #4a7c1f)',
                              fontSize: '0.7rem', fontWeight: 700, cursor: 'pointer', lineHeight: 1.4,
                              transition: 'background 120ms ease, color 120ms ease',
                            }}
                          >« {isDe ? 'Zuklappen' : 'Collapse'}</button>
                          <span>{isDe ? 'Vorname' : 'First name'}{sortIcon(sortable)}{hideButton(id)}</span>
                        </div>
                      ) : (
                        <>
                          {id === 'id' ? '#' : id === 'anrede' ? (isDe ? 'Anrede' : 'Salutation') : id === 'nachname' ? (isDe ? 'Nachname' : 'Last name') : id === 'email' ? 'Email' : id === 'status' ? 'Status' : (isDe ? 'Registriert am' : 'Registered on')}
                          {sortIcon(sortable)}
                          {hideButton(id)}
                        </>
                      )}
                    </th>
                  );
                }
                if (id === 'jobTitle') return <th key={id} style={{ ...baseStyle, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('jobTitle')}>Job Title{sortIcon('jobTitle')}{hideButton(id)}</th>;
                if (id === 'location') return <th key={id} style={{ ...baseStyle, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('location')}>{isDe ? 'Standort' : 'Location'}{sortIcon('location')}{hideButton(id)}</th>;
                if (id === 'company') return <th key={id} style={{ ...baseStyle }}>{isDe ? 'Unternehmen' : 'Company'}{hideButton(id)}</th>;
                if (id === 'starterType') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? "Starter-Typ: Durchstarter oder Funstarter. Wird bei der Anmeldung gewählt und steuert die Split-Kapazität + Warteliste. Der eigentliche Startblock steht in der Custom-Field-Spalte 'Start block'." : "Starter type: Durchstarter or Funstarter. Chosen at registration and controls the split capacity + waitlist. The actual start block is in the custom field column 'Start block'."}>
                      {isDe ? 'Starter-Typ' : 'Starter type'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'startnummer') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Die offizielle Startnummer des Veranstalters, eingelesen über „Startnummern importieren".' : 'The official bib number from the organiser, imported via "Import bib numbers".'}>
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
                  return <th key={id} style={{ textAlign: 'left', padding: 8 }}>Aktion</th>;
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
                    <th key={id} onClick={() => handleSort(id)} style={{ ...baseStyle, fontSize: '0.78rem', cursor: 'pointer', userSelect: 'none', ...pastelAHeader }} title={`${label} — ${isDe ? 'Hauptevent-Feld' : 'main-event field'}`}>
                      {label}{sortIcon(id)}
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
                    <th key={id} onClick={() => handleSort(id)} style={{ ...baseStyle, fontSize: '0.78rem', cursor: 'pointer', userSelect: 'none', ...pastelBHeader }} title={inSubEventDetail ? `${label} — ${isDe ? 'Sub-Event-Feld' : 'sub-event field'}` : label}>
                      {label}{sortIcon(id)}
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
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-400)' }}>{idCell}</td>;
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
                    <td key="person" style={{ padding: 8 }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
                        {/* v24.56: Foto-Hover zeigt Kontaktkarte (E-Mail + Teams),
                            wie bei den Organizern auf der Anmeldeseite. */}
                        <PersonContactHover email={email} name={fullName} size={30} subline={sub} isDe={isDe} />
                        <div
                          style={{ display: 'flex', flexDirection: 'column', minWidth: 0, lineHeight: 1.25, cursor: 'pointer' }}
                          title={isDe ? 'Detailinfos anzeigen' : 'Show details'}
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          onClick={() => setParticipantDetail({ name: fullName, email, jobTitle: jt, location: String((reg as any).Location || ''), company: comp, department: String((reg as any).Department || ''), phone: String((reg as any).Phone || ''), status: reg.Status || '', tid: reg.TeilnehmerID || null })}
                        >
                          <span style={{ fontWeight: 700, whiteSpace: 'nowrap' }}>{highlightMatch(fullName)}</span>
                          {sub && <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>{highlightMatch(sub)}</span>}
                        </div>
                      </div>
                    </td>
                  );
                }
                if (id === 'anrede') {
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{reg.Anrede || '-'}</td>;
                }
                if (id === 'vorname') {
                  // Fallback für Alt-Daten: erstes Wort aus ParticipantName.
                  const v = reg.Vorname || ((reg.ParticipantName || '').split(' ')[0] || '');
                  return <td key={id} style={{ padding: 8, fontWeight: 500 }}>{v ? highlightMatch(v) : '-'}</td>;
                }
                if (id === 'nachname') {
                  // Fallback für Alt-Daten: alles ausser dem ersten Wort als Nachname.
                  let n = reg.Nachname || '';
                  if (!n && reg.ParticipantName) {
                    const parts = reg.ParticipantName.trim().split(/\s+/);
                    if (parts.length > 1) n = parts.slice(1).join(' ');
                  }
                  return <td key={id} style={{ padding: 8, fontWeight: 500 }}>{n ? highlightMatch(n) : '-'}</td>;
                }
                if (id === 'email') {
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{highlightMatch(reg.ParticipantEmail)}</td>;
                }
                if (id === 'jobTitle') {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const jt = String((reg as any).JobTitle || '');
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{jt ? highlightMatch(jt) : '-'}</td>;
                }
                if (id === 'location') {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const lc = String((reg as any).Location || '');
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{lc ? highlightMatch(lc) : '-'}</td>;
                }
                if (id === 'company') {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const cmp = String((reg as any).Company || '');
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{cmp ? highlightMatch(cmp) : '-'}</td>;
                }
                if (id === 'starterType') {
                  return (
                    <td key={id} style={{ padding: 8, fontSize: '0.8rem' }}>
                      {(() => {
                        // Tatsächlicher Startblock (StarterType) + Wunsch (PreferredStarterType).
                        // Wenn beide identisch: nur einen anzeigen. Wenn unterschiedlich (z.B. per
                        // Fallback-Dialog auf anderen Typ umgestiegen): Wunsch in Klammern daneben.
                        const actual = reg.StarterType || '';
                        const pref = reg.PreferredStarterType || '';
                        if (!actual && !pref) return <span style={{ color: 'var(--dex-gray-400)' }}>—</span>;
                        if (actual && pref && actual !== pref) {
                          return <span>{actual} <span style={{ color: 'var(--dex-gray-500)' }}>(Wunsch: {pref})</span></span>;
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
                        return <span>{isWaitlist ? `Wunsch: ${pref}` : pref}</span>;
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
                      <td key={id} style={{ padding: 8 }}>
                        <span className="badge" style={{ background: '#fff3e0', color: '#b35a00' }}>
                          {isDe ? 'Angemeldet (Datenschutzrückmeldung offen)' : 'Registered (privacy confirmation pending)'}
                        </span>
                      </td>
                    );
                  }
                  return (
                    <td key={id} style={{ padding: 8 }}>
                      <span className={`badge ${reg.Status === 'Eingecheckt' ? 'badge-green' : 'badge-gray'}`}>
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
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{regDate ? formatDate(regDate) : '—'}</td>;
                }
                if (id === 'startnummer') {
                  // v30.48: Startnummer aus dem Veranstalter-Rücklauf.
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const bib = String((reg as any).Startnummer || '').trim();
                  return <td key={id} style={{ padding: 8, fontFamily: 'monospace', fontWeight: bib ? 700 : 400, color: bib ? 'var(--dex-gray-800)' : 'var(--dex-gray-300)' }}>{bib || '—'}</td>;
                }
                if (id === 'promotedDate') {
                  // v17.15: „Nachgerückt am" — gesetzt beim Promote
                  // von Warteliste → Angemeldet. Leer für Personen die
                  // sich direkt in den Aktiv-Bereich angemeldet haben.
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const v = (reg as any).PromotedDate as string | undefined;
                  return <td key={id} style={{ padding: 8, color: v ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-300)', fontSize: '0.8rem' }}>{v ? formatDate(v) : '—'}</td>;
                }
                if (id === 'replaced') {
                  // v17.15: „Ersetzt" — die Person, deren Cancel diesen
                  // Promote ausgelöst hat. Wenn die Person in den
                  // aktuellen registrations gefunden wird, zeigen wir den
                  // Namen — sonst fallback auf die rohe E-Mail.
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const email = ((reg as any).ReplacedParticipantEmail as string | undefined) || '';
                  if (!email) return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-300)' }}>—</td>;
                  const other = registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === email.toLowerCase());
                  const label = other ? ((other.Vorname || '') + ' ' + (other.Nachname || '')).trim() || other.ParticipantName || email : email;
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-700)', fontSize: '0.8rem' }} title={email}>{label}</td>;
                }
                if (id === 'replacedBy') {
                  // v17.15: „Ersetzt durch" — die Person die nach Cancel
                  // dieses Eintrags den Platz übernommen hat. Spiegelbild
                  // von „Ersetzt".
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const email = ((reg as any).ReplacedByParticipantEmail as string | undefined) || '';
                  if (!email) return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-300)' }}>—</td>;
                  const other = registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === email.toLowerCase());
                  const label = other ? ((other.Vorname || '') + ' ' + (other.Nachname || '')).trim() || other.ParticipantName || email : email;
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.8rem' }} title={email}>{label}</td>;
                }
                if (id === 'joinOrder') {
                  // v17.9 (deprecated): joinOrder-Spalte seit v17.10 entfernt.
                  return <td key={id} style={{ padding: 8 }}>—</td>;
                }
                if (id === 'registeredBy') {
                  return (
                    <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>
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
                            return <span style={{ color: 'var(--dex-green-dark)' }}>Selbst</span>;
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
                          return <span style={{ color: 'var(--dex-green-dark)' }}>Selbst</span>;
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
                  if (!inTeam) return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-400)' }}>—</td>;
                  return (
                    <td key={id} style={{ padding: 8, color: 'var(--dex-gray-700)', fontSize: '0.82rem' }}>
                      {tName ? `„${tName}"` : <span style={{ color: 'var(--dex-gray-500)' }}>ohne Namen</span>}
                      {reg.TeamLead && (
                        <span style={{ marginLeft: 6, padding: '1px 7px', background: 'var(--dex-green, #86bc25)', color: '#fff', borderRadius: 8, fontSize: '0.66rem', fontWeight: 700 }}>Lead</span>
                      )}
                    </td>
                  );
                }
                if (id === 'roommate') {
                  return (
                    <td key={id} style={{ padding: 8, fontSize: '0.8rem' }}>
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
                                className="badge"
                                style={{ marginLeft: 2, background: 'var(--dex-green)', color: '#fff', padding: '1px 6px', borderRadius: 4, fontSize: '0.7rem' }}
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
                    <td key={id} style={{ padding: 8, color: 'var(--dex-gray-700)', fontSize: '0.8rem', whiteSpace: 'nowrap', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', ...pastelACell }} title={String(val || '')}>
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
                      <td key={id} style={{ padding: 8, fontSize: '0.8rem', whiteSpace: 'nowrap', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', ...pastelBCell }}>
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
                    <td key={id} style={{ padding: 8, color: 'var(--dex-gray-700)', fontSize: '0.8rem', whiteSpace: 'nowrap', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', ...pastelBCell }} title={String(val || '')}>
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
                    const regs = await getAllRegistrations(selectedEvent.id);
                    setRegistrations(regs);
                  };
                  const doCheckOut = async (): Promise<void> => {
                    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                    await eventServiceRef.checkOutParticipant(selectedEvent.subsiteUrl, reg.Id);
                    const regs = await getAllRegistrations(selectedEvent.id);
                    setRegistrations(regs);
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
                    const regs = await getAllRegistrations(selectedEvent.id);
                    setRegistrations(regs);
                  };
                  return (
                    <td key={id} style={{ padding: 8, display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                      {/* „Bearbeiten" nur, solange das Event noch nicht vorbei ist. */}
                      {!eventOver && (
                        <button
                          className="btn btn-secondary"
                          style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                          title={isDe ? 'Teilnehmer-Daten bearbeiten' : 'Edit attendee data'}
                          onClick={() => openEditModal(reg)}
                        >
                          <Pencil size={12} /> {isDe ? 'Bearbeiten' : 'Edit'}
                        </button>
                      )}
                      {/* v11.0: Anhang-Button — wenn das Event den Teilnehmer-
                          Upload erlaubt ODER ein Dokument-Custom-Feld hat (v19.0).
                          Zeigt Counter wenn mind. eine Datei hochgeladen wurde. */}
                      {(selectedEvent?.allowAttendeeUpload || (selectedEvent?.eventSpecificFields || []).some(f => f.type === 'document')) && (
                        <button
                          className="btn btn-secondary"
                          style={{ fontSize: '0.75rem', padding: '4px 10px', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                          title={isDe ? 'Hochgeladene Dateien anzeigen' : 'Show uploaded files'}
                          onClick={() => setAttachmentsModalReg(reg)}
                        >
                          <FileText size={12} />
                          {att.length > 0 ? `${isDe ? 'Datei' : 'File'} (${att.length})` : (isDe ? 'Datei' : 'File')}
                        </button>
                      )}
                      {eventOver ? (
                        <>
                          {/* Nach dem Event: Anwesenheit explizit pflegen. */}
                          <button
                            className="btn btn-primary"
                            style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                            disabled={reg.Status === 'Eingecheckt'}
                            title={isDe ? 'Als anwesend markieren' : 'Mark as attended'}
                            onClick={doCheckIn}
                          >
                            {isDe ? 'Einchecken' : 'Check in'}
                          </button>
                          <button
                            className="btn btn-secondary"
                            style={{ fontSize: '0.75rem', padding: '4px 10px' }}
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
                            <button className="btn btn-secondary" style={{ fontSize: '0.75rem', padding: '4px 10px' }} onClick={doCheckOut}>
                              {isDe ? 'Auschecken' : 'Check out'}
                            </button>
                          ) : (
                            <button className="btn btn-primary" style={{ fontSize: '0.75rem', padding: '4px 10px' }} onClick={doCheckIn}>
                              {isDe ? 'Einchecken' : 'Check in'}
                            </button>
                          )}
                          <button
                            className="btn btn-secondary"
                            style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-gray-700, #444)' }}
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
                              className="btn btn-secondary"
                              style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                              title={isDe ? 'No-Show zurücknehmen (Status zurück auf „Angemeldet")' : 'Undo no-show (status back to “registered”)'}
                              onClick={doCheckOut}
                            >
                              {isDe ? 'Zurücksetzen' : 'Reset'}
                            </button>
                          )}
                        </>
                      ) : (
                        reg.Status === 'Eingecheckt' ? (
                          <button className="btn btn-secondary" style={{ fontSize: '0.75rem', padding: '4px 10px' }} onClick={doCheckOut}>
                            {isDe ? 'Auschecken' : 'Check out'}
                          </button>
                        ) : (
                          <button className="btn btn-primary" style={{ fontSize: '0.75rem', padding: '4px 10px' }} onClick={doCheckIn}>
                            {isDe ? 'Einchecken' : 'Check in'}
                          </button>
                        )
                      )}
                      {/* Abmelden: nach dem Event still & ohne E-Mail (v22.22) —
                          deshalb auch für Organizer eigener Events freigegeben. */}
                      {(eventOver || !orgPastLock) && (
                      <button
                        className="btn btn-secondary"
                        style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #c00)' }}
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
                        {eventOver ? (isDe ? 'Abmelden (Ohne E-Mail)' : 'Cancel (no email)') : (isDe ? 'Abmelden' : 'Cancel')}
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
                              className="btn btn-secondary"
                              style={{ fontSize: '0.75rem', padding: '4px 10px', color: '#b35a00' }}
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
                              className="btn btn-secondary"
                              style={{ fontSize: '0.75rem', padding: '4px 10px', color: '#b35a00' }}
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
                                const regs = await getAllRegistrations(selectedEvent.id);
                                setRegistrations(regs);
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
                  <div style={{ display: 'flex', justifyContent: 'flex-end', alignItems: 'center', marginBottom: 8, gap: 8, position: 'relative' }}>
                    <button
                      className="btn btn-secondary"
                      style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                      onClick={() => setShowColumnPicker(!showColumnPicker)}
                    >
                      {isDe ? 'Spalten anpassen' : 'Customize columns'}
                    </button>
                    {showColumnPicker && (
                      <div
                        style={{
                          position: 'absolute', right: 0, top: '100%', marginTop: 4,
                          background: '#fff', border: '1px solid var(--dex-gray-200)',
                          borderRadius: 8, padding: 12,
                          boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                          width: 280, zIndex: 100, maxHeight: 400, overflowY: 'auto',
                        }}
                      >
                        <div style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 8 }}>
                          {isDe ? 'Spalten verwalten' : 'Manage columns'}
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
                          return (
                            <div
                              key={id}
                              style={{
                                display: 'flex', alignItems: 'center', gap: 6,
                                padding: '4px 2px', fontSize: '0.82rem',
                                opacity: isVisible ? 1 : 0.55,
                              }}
                            >
                              <input
                                type="checkbox"
                                checked={isVisible}
                                disabled={!!col.alwaysVisible}
                                onChange={() => {
                                  if (col.alwaysVisible) return;
                                  if (isHidden) showColumn(id); else hideColumn(id);
                                }}
                                style={{ cursor: col.alwaysVisible ? 'not-allowed' : 'pointer' }}
                                title={col.alwaysVisible ? 'Pflicht-Spalte — kann nicht ausgeblendet werden' : (isHidden ? 'Einblenden' : 'Ausblenden')}
                              />
                              <span style={{ flex: 1, color: 'var(--dex-gray-700)' }}>{col.label}</span>
                              <button
                                type="button"
                                onClick={() => moveColumn(id, -1)}
                                disabled={!canMoveUp}
                                aria-label={isDe ? 'Spalte nach oben' : 'Move column up'}
                                title={isDe ? 'Nach oben' : 'Up'}
                                style={{
                                  border: 'none', background: 'transparent',
                                  cursor: canMoveUp ? 'pointer' : 'not-allowed',
                                  color: canMoveUp ? 'var(--dex-gray-600)' : 'var(--dex-gray-300)',
                                  fontSize: '0.9rem', padding: '0 4px',
                                }}
                              >
                                ↑
                              </button>
                              <button
                                type="button"
                                onClick={() => moveColumn(id, 1)}
                                disabled={!canMoveDown}
                                aria-label={isDe ? 'Spalte nach unten' : 'Move column down'}
                                title={isDe ? 'Nach unten' : 'Down'}
                                style={{
                                  border: 'none', background: 'transparent',
                                  cursor: canMoveDown ? 'pointer' : 'not-allowed',
                                  color: canMoveDown ? 'var(--dex-gray-600)' : 'var(--dex-gray-300)',
                                  fontSize: '0.9rem', padding: '0 4px',
                                }}
                              >
                                ↓
                              </button>
                            </div>
                          );
                        })}
                        <div style={{ marginTop: 8, textAlign: 'right' }}>
                          <button
                            className="btn btn-secondary"
                            style={{ fontSize: '0.72rem', padding: '3px 8px' }}
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
                    const renderTable = (rows: SPRegistration[], indexOffset: number): React.ReactElement => {
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
                            return (
                              <tr
                                key={reg.Id}
                                title={rowTitle}
                                style={{
                                  borderBottom: '1px solid var(--dex-gray-100)',
                                  ...(isDuplicate
                                    ? { background: 'rgba(200,0,0,0.10)', boxShadow: 'inset 3px 0 0 var(--dex-red, #c00)' }
                                    : highlight
                                    ? { background: 'rgba(237,139,0,0.13)', boxShadow: 'inset 3px 0 0 var(--dex-orange, #ed8b00)' }
                                    : {}),
                                }}
                              >
                                {effectiveColumnIds.map(id => renderCell(id, reg, i))}
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
                      return (
                        // v24.96: eigener Scroll-Container um die Tabelle → der
                        // thead (position:sticky top:0) klebt zuverlässig an dessen
                        // oberem Rand (CSS-sticky relativ zu DIESEM Container, nicht
                        // zum Fenster — Letzteres ist im SP-Canvas unzuverlässig).
                        <div style={{ maxHeight: '70vh', overflow: 'auto' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                          <thead>
                            <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
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
                                        padding: '6px 8px', fontWeight: 700, fontSize: '0.78rem',
                                        background: dr.muted ? 'var(--dex-gray-100, #f3f4f6)' : 'rgba(134,188,37,0.10)',
                                        color: dr.muted ? 'var(--dex-gray-600)' : 'var(--dex-green-dark, #4a7c1f)',
                                        borderBottom: '1px solid var(--dex-gray-100)',
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
                      );
                    };

                    if (!isSplitCapacity || splitParticipantsView === 'merged') {
                      return (
                        <>
                          {isSplitCapacity && (
                            <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 8 }}>
                              <SplitMergeToggle view={splitParticipantsView} setView={setSplitParticipantsView} isDe={isDe} />
                            </div>
                          )}
                          {renderTable(activeRegs, 0)}
                        </>
                      );
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
                    return (
                      <>
                        <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 8 }}>
                          <SplitMergeToggle view={splitParticipantsView} setView={setSplitParticipantsView} isDe={isDe} />
                        </div>
                        {groups.map(g => {
                          const offset = runningIdx;
                          runningIdx += g.rows.length;
                          return (
                            <div key={g.key} style={{ marginBottom: 20 }}>
                              <h4 style={{
                                margin: '0 0 8px', color: 'var(--dex-green-dark, #4a7c1f)',
                                fontSize: '0.95rem', fontWeight: 700, display: 'flex',
                                alignItems: 'baseline', gap: 8,
                              }}>
                                <span>{g.label}</span>
                                <span style={{ color: 'var(--dex-gray-500)', fontWeight: 500, fontSize: '0.85rem' }}>
                                  ({g.rows.length}{g.cap > 0 ? ` / ${g.cap}` : ''})
                                </span>
                              </h4>
                              {g.rows.length === 0 ? (
                                <p style={{ margin: '0 0 12px', fontSize: '0.82rem', color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
                                  {isDe ? 'Keine Teilnehmer in dieser Gruppe.' : 'No participants in this group.'}
                                </p>
                              ) : renderTable(g.rows, offset)}
                            </div>
                          );
                        })}
                        {groupNone.length > 0 && (
                          <div style={{ marginBottom: 20 }}>
                            <h4 style={{ margin: '0 0 8px', color: 'var(--dex-gray-500)', fontSize: '0.95rem', fontWeight: 700 }}>
                              {isDe ? 'Ohne Gruppe' : 'No group'} <span style={{ color: 'var(--dex-gray-400)', fontWeight: 500, fontSize: '0.85rem' }}>({groupNone.length})</span>
                            </h4>
                            {renderTable(groupNone, runningIdx)}
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

