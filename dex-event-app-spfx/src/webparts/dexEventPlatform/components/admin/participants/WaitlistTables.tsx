/* WaitlistTables — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 12893-13064 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich übernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3 (docs/ui-leitfaden.md 5a/5b): jede Warteliste ist eine Karte mit
 * `dex-ui-card-head`, darin eine kompakte Tabelle — Platz als Pille, Person als
 * EINE Zelle statt drei Spalten, Aktionen rechts in `is-actions` (hier richtig:
 * beide nehmen die Zeile aus der Liste). Sortierung, Gruppen-Reihenfolge und
 * beide Handler unverändert; `vorname`/`email` behalten eigene Sortier-Knöpfe.
 */
import * as React from 'react';
import { EventService, SPRegistration } from '../../../services/EventService';
import { formatDate } from '../../../utils/eventStatus';
import { isEventOver } from '../../../utils/eventFormat';
import { InfoTooltip } from '../../InfoTooltip';
import { Trash2 } from '../../Icons';
import { cx, ensureDexUiStyles } from '../../dexUi';
import { DeloitteEvent } from '../../../types';

export interface WaitlistTablesProps {
  buildCancellationMail: (ev: DeloitteEvent, reg: SPRegistration, fullName: string) => Promise<{    subject: string;    body: string;}>;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  currentUser: import("../../../types/index").User;
  eventServiceRef: EventService;
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
  isDe: boolean;
  isSplitCapacity: boolean;
  query: string;
  selectedEvent: DeloitteEvent;
  setWaitlistSortAsc: React.Dispatch<React.SetStateAction<boolean>>;
  setWaitlistSortColumn: React.Dispatch<React.SetStateAction<"date" | "location" | "pos" | "vorname" | "nachname" | "email" | "jobtitle">>;
  setWlPosModal: React.Dispatch<React.SetStateAction<{ reg: SPRegistration; currentPos: number; total: number; }>>;
  setWlPosValue: React.Dispatch<React.SetStateAction<string>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  waitlistDurch: SPRegistration[];
  waitlistFun: SPRegistration[];
  waitlistRegs: SPRegistration[];
  waitlistSortAsc: boolean;
  waitlistSortColumn: "date" | "location" | "pos" | "vorname" | "nachname" | "email" | "jobtitle";
  waitlistTruePos: Record<number, number>;
  /** v30.67: wahre Gruppengröße je Person (ungefiltert), s. AdminPage. */
  waitlistTrueTotal: Record<number, number>;
  waitlistUnassigned: SPRegistration[];
  wlPosBusy: boolean;
}

export const WaitlistTables: React.FC<WaitlistTablesProps> = (p) => {
  const { buildCancellationMail, confirmDialog, currentUser, eventServiceRef, isDe, isSplitCapacity, query, reloadRegistrations, selectedEvent, setWaitlistSortAsc, setWaitlistSortColumn, setWlPosModal, setWlPosValue, showAlert, waitlistDurch, waitlistFun, waitlistRegs, waitlistSortAsc, waitlistSortColumn, waitlistTruePos, waitlistTrueTotal, waitlistUnassigned, wlPosBusy } = p;
          // v31.3: Idempotent — ohne Modal/WizardFormShell fehlten hier sonst die
          // dex-ui-Klassen (inkl. aller :hover).
          ensureDexUiStyles();
          // Seit v6.5: bei B2Run-Split-Kapazitäten getrennte Wartelisten-Tabellen pro
          // PreferredStarterType. Ohne Split: eine einzige Warteliste wie bisher.
          // v31.3: `dotClass` statt CSS-Farbe — der Punkt vor dem Titel trägt die
          // Gruppen-Zuordnung, die früher die Überschriftfarbe hatte.
          const renderWaitlistTable = (title: string, regs: SPRegistration[], dotClass: string): React.ReactElement | null => {
            if (regs.length === 0) return null;
            // v17.8: Sortierung pro Spalte. Default 'pos' = TeilnehmerID asc
            // (FIFO-Position der Warteliste — wie vorher).
            const sortedRegs = (() => {
              const arr = regs.slice();
              const dir = waitlistSortAsc ? 1 : -1;
              const safe = (s: string | undefined): string => (s || '').toLowerCase();
              const dateMs = (s: string | undefined): number => s ? new Date(s).getTime() : Number.POSITIVE_INFINITY;
              arr.sort((a, b) => {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const anyA = a as any; const anyB = b as any;
                switch (waitlistSortColumn) {
                  case 'pos': return ((a.TeilnehmerID || 0) - (b.TeilnehmerID || 0)) * dir;
                  case 'vorname': return safe(a.Vorname).localeCompare(safe(b.Vorname), 'de') * dir;
                  case 'nachname': return safe(a.Nachname).localeCompare(safe(b.Nachname), 'de') * dir;
                  case 'email': return safe(a.ParticipantEmail).localeCompare(safe(b.ParticipantEmail)) * dir;
                  case 'jobtitle': return safe(anyA.JobTitle).localeCompare(safe(anyB.JobTitle), 'de') * dir;
                  case 'location': return safe(anyA.Location).localeCompare(safe(anyB.Location), 'de') * dir;
                  case 'date': return (dateMs(a.RegistrationDate) - dateMs(b.RegistrationDate)) * dir;
                }
                return 0;
              });
              return arr;
            })();
            const toggleSort = (k: typeof waitlistSortColumn): void => {
              if (waitlistSortColumn === k) setWaitlistSortAsc(v => !v);
              else { setWaitlistSortColumn(k); setWaitlistSortAsc(true); }
            };
            // v31.3: Pfeil nur an der sortierten Spalte (Leitfaden 5b).
            const sortMark = (k: typeof waitlistSortColumn): React.ReactNode =>
              k === waitlistSortColumn ? <span className="dex-ui-table-sort">{waitlistSortAsc ? '▲' : '▼'}</span> : null;
            const sortableTh = (k: typeof waitlistSortColumn, label: string): React.ReactElement => (
              <th className={cx('is-sortable', waitlistSortColumn === k && 'is-sorted')} onClick={() => toggleSort(k)} title={isDe ? 'Sortierung umschalten' : 'Toggle sorting'}>{label}{sortMark(k)}</th>
            );
            // v31.3: Vorname und E-Mail stecken in der Personen-Zelle, behalten
            // aber ihren Sortier-Zugriff — sonst wären zwei Sortierungen weg.
            const miniSortBtn = (k: typeof waitlistSortColumn, label: string, hint: string): React.ReactElement => (
              <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" title={hint}
                style={{ fontSize: '0.66rem', padding: '1px 6px', marginLeft: 4, textTransform: 'none', letterSpacing: 0 }}
                onClick={e => { e.stopPropagation(); toggleSort(k); }}
              >{label}{sortMark(k)}</button>
            );
            // v24.93: NICHT sticky — diese Tabellen stehen in einem
            // Wrapper mit eigener Scroll-Achse, in dem position:sticky relativ
            // zum Wrapper (statt zum Fenster) berechnet wird; der Kopf schwebte
            // dadurch mitten in der Tabelle und schnitt Zeilen ab. Sticky
            // bleibt nur bei der Teilnehmer-Tabelle (eigener renderHeader,
            // ohne Scroll-Wrapper).
            return (
              <div className="dex-ui-card" key={title} style={{ marginTop: 18 }}>
                <div className="dex-ui-card-head" style={{ marginBottom: 10 }}>
                  <h4 className="dex-ui-card-head-title">
                    <span className={cx('dex-ui-dot', dotClass)} aria-hidden="true" />
                    {title}
                    <InfoTooltip text={isDe
                      ? 'Wer auf einen frei werdenden Platz wartet. Der Platz ist der Rang nach TeilnehmerID — genau in dieser Reihenfolge rückt die App nach, und derselben folgt der Flow. „Platz ändern“ schreibt die TeilnehmerIDs um, die Person wandert also wirklich nach vorn oder hinten. Das Papierkorb-Symbol nimmt die Person von der Warteliste — mit Absage-Mail, solange das Event nicht vorbei ist.'
                      : 'Everyone waiting for a seat to free up. The position is the rank by participant ID — the app promotes in exactly that order, and so does the flow. “Change position” rewrites the participant IDs, so the person really does move up or down. The bin icon takes the person off the waitlist — with a cancellation email as long as the event is not over.'} />
                  </h4>
                  <span className="dex-ui-card-head-meta">{query ? (isDe ? `${regs.length} Treffer` : `${regs.length} matches`) : isDe ? `${regs.length} ${regs.length === 1 ? 'Person wartet' : 'Personen warten'}` : `${regs.length} ${regs.length === 1 ? 'person waiting' : 'people waiting'}`}</span>
                </div>
                <div className="dex-ui-table-wrap">
                  <table className="dex-ui-table dex-ui-table--compact">
                    <thead>
                      <tr>
                        {sortableTh('pos', isDe ? 'Platz' : 'Position')}
                        <th className={cx('is-sortable', waitlistSortColumn === 'nachname' && 'is-sorted')} onClick={() => toggleSort('nachname')} title={isDe ? 'Nach Nachname sortieren' : 'Sort by last name'}>
                          Person{sortMark('nachname')}
                          {miniSortBtn('vorname', isDe ? 'Vorname' : 'First name', isDe ? 'Nach Vorname sortieren' : 'Sort by first name')}
                          {miniSortBtn('email', isDe ? 'E-Mail' : 'Email', isDe ? 'Nach E-Mail sortieren' : 'Sort by email')}
                        </th>
                        {sortableTh('jobtitle', isDe ? 'Rolle' : 'Job title')}
                        {sortableTh('location', isDe ? 'Standort' : 'Location')}
                        {isSplitCapacity && <th>{isDe ? 'Wunsch-Gruppe' : 'Preferred group'}</th>}
                        {sortableTh('date', isDe ? 'Registriert am' : 'Registered')}
                        <th style={{ textAlign: 'right' }}>{isDe ? 'Aktion' : 'Action'}</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sortedRegs.map((reg, i) => {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const anyReg = reg as any;
                        // v17.8/v26.31: „Platz" = wahre FIFO-Position (TeilnehmerID
                        // asc) aus der UNGEFILTERTEN Warteliste (waitlistTruePos),
                        // unabhängig von Sortierung UND Suchfilter — sonst zeigte eine
                        // gefilterte Trefferliste fälschlich Platz 1, 2, … Beim Filtern
                        // steht „#n" daneben (v31.3 grau; der Platz bleibt die Pille).
                        const truePos = waitlistTruePos[reg.Id];
                        const fullName = [reg.Vorname, reg.Nachname].filter(Boolean).join(' ') || reg.ParticipantName || '—';
                        const initials = ((reg.Vorname || '').charAt(0) + (reg.Nachname || '').charAt(0)).toUpperCase()
                          || (fullName || reg.ParticipantEmail || '?').charAt(0).toUpperCase();
                        return (
                        <tr key={reg.Id}>
                          <td>
                            <span className="dex-ui-pill dex-ui-pill--orange">{truePos != null ? truePos : i + 1}</span>
                            {!!query && truePos != null && <span className="dex-ui-muted" style={{ marginLeft: 6, fontSize: '0.72rem' }} title={isDe ? 'Laufende Nummer in der Trefferliste' : 'Row number within the search results'}>#{i + 1}</span>}
                          </td>
                          <td><div className="dex-ui-person">
                            <span className="dex-ui-avatar" aria-hidden="true">{initials}</span>
                            <div style={{ minWidth: 0 }}>
                              <div className="dex-ui-person-name">{fullName}</div>
                              <div className="dex-ui-person-sub">{reg.ParticipantEmail || '—'}</div>
                            </div>
                          </div></td>
                          <td style={{ color: 'var(--dex-gray-600)' }}>{anyReg.JobTitle || '—'}</td>
                          <td style={{ color: 'var(--dex-gray-600)' }}>{anyReg.Location || '—'}</td>
                          {isSplitCapacity && <td>{reg.PreferredStarterType ? <span className="dex-ui-pill dex-ui-pill--gray">{reg.PreferredStarterType}</span> : '—'}</td>}
                          <td style={{ color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                          <td className="is-actions">
                            {/* v28.70: Wartelisten-Platz manuell setzen. Die
                                Position ist der Rang nach TeilnehmerID — genau
                                danach sortieren App-Nachrücken UND der Flow. */}
                            <button type="button" className="btn btn-secondary dex-ui-btn-sm" style={{ marginRight: 6 }} disabled={wlPosBusy}
                              onClick={() => {
                                // v30.67: `regs` ist die SUCH-gefilterte Liste — als
                                // Nenner/`max` stand bei aktiver Suche also die
                                // Trefferzahl („Platz 7 von 1"). Die wahre Gruppengröße
                                // kommt wie die wahre Position aus der ungefilterten Liste.
                                const trueTotal = waitlistTrueTotal[reg.Id];
                                setWlPosModal({ reg, currentPos: truePos != null ? truePos : (i + 1), total: trueTotal != null ? trueTotal : regs.length });
                                setWlPosValue('1');
                              }}
                            >
                              {isDe ? 'Platz ändern' : 'Change position'}
                            </button>
                            {/* v31.3: Entfernen als rotes Symbol (Leitfaden 5b) — der
                                Text steht in title/aria-label, mit dem Namen dabei. */}
                            <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                              title={isDe ? 'Von der Warteliste entfernen' : 'Remove from waitlist'}
                              aria-label={isDe ? `${fullName} von der Warteliste entfernen` : `Remove ${fullName} from the waitlist`}
                              onClick={async () => {
                                if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                                const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                                // v22.22: Vergangenes Event → stilles Entfernen
                                // (keine Abmelde-Mail, kein ID-Reorder).
                                const eventWasOver = isEventOver(selectedEvent);
                                if (!(await confirmDialog(`${name} von der Warteliste entfernen?${eventWasOver ? (isDe ? '\n\nDas Event liegt in der Vergangenheit — es geht keine Abmelde-Mail raus.' : '\n\nThe event is in the past — no cancellation email will be sent.') : ''}`, { danger: true, confirmLabel: isDe ? 'Entfernen' : 'Remove' }))) return;
                                await eventServiceRef.cancelRegistration(selectedEvent.subsiteUrl, reg.Id, `${currentUser.firstName} ${currentUser.surname}`.trim(), currentUser.email);
                                if (reg.ParticipantEmail && !selectedEvent.disableEmails && !selectedEvent.disableCancellationEmail && !eventWasOver) {
                                  const emailData = await buildCancellationMail(selectedEvent, reg, name);
                                  eventServiceRef.queueEmail(
                                    emailData.subject, reg.ParticipantEmail, name, emailData.body,
                                    'Abmeldung', selectedEvent.title, selectedEvent.id
                                  ).catch(err => console.warn('[DEX]', err));
                                }
                                if (reg.ParticipantEmail && selectedEvent.eventNumber) {
                                  eventServiceRef.removeParticipantEvent(reg.ParticipantEmail, selectedEvent.eventNumber).catch(err => console.warn('[DEX]', err));
                                }
                                if (selectedEvent.subsiteUrl && !eventWasOver) {
                                  // v30.80: geprüfter Pfad (Wiederholungen, Event-Log, Merker).
                                  const r = await eventServiceRef.queueIDReorderChecked({
                                    eventId: selectedEvent.id, eventNumber: selectedEvent.eventNumber || 0,
                                    subsiteUrl: selectedEvent.subsiteUrl, eventTitle: selectedEvent.title,
                                    cancelledName: `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantName || undefined,
                                    cancelledEmail: reg.ParticipantEmail || undefined,
                                  }, 'waitlist-remove');
                                  if (!r.ok) {
                                    showAlert(isDe
                                      ? `Entfernt, aber der Reorder-Auftrag konnte nach ${r.attempts} Versuchen nicht in die Queue geschrieben werden (HTTP ${r.status}). Die App holt ihn beim nächsten App-Start nach; sonst einmal „IDs neu vergeben" klicken.`
                                      : `Removed, but the reorder job could not be written to the queue after ${r.attempts} attempts (HTTP ${r.status}). The app retries on the next start; otherwise click "Reassign IDs" once.`,
                                      { variant: 'error' });
                                  }
                                }
                                // v30.67 (Review): gemeinsamer Nachlade-Pfad statt `[]` bei 429.
                                await reloadRegistrations();
                              }}
                            >
                              <Trash2 size={15} />
                            </button>
                          </td>
                        </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          };

          if (isSplitCapacity) {
            // v11.6: Wartelisten-Tabellen mit den frei wählbaren Gruppen-
            // Labels statt hartcodeten 'Durchstarter'/'Funstarter'.
            const wlLabelA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
            const wlLabelB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
            // v11.29: Reihenfolge respektiert splitDisplayOrderReversed
            // (gleicher Toggle wie auf Register-Page + Kapazitäts-Cards).
            // v31.3: Der Gruppen-Punkt vor dem Titel ersetzt die alte
            // Überschriftfarbe; die Titel sind jetzt zweisprachig.
            const wl = isDe ? 'Warteliste' : 'Waitlist';
            const wlA = renderWaitlistTable(`${wl} ${wlLabelA}`, waitlistDurch, 'dex-ui-dot--green');
            const wlB = renderWaitlistTable(`${wl} ${wlLabelB}`, waitlistFun, 'dex-ui-dot--orange');
            const reversed = !!selectedEvent?.splitDisplayOrderReversed;
            return (
              <>
                {reversed ? <>{wlB}{wlA}</> : <>{wlA}{wlB}</>}
                {renderWaitlistTable(isDe ? 'Warteliste ohne Gruppe' : 'Waitlist without group', waitlistUnassigned, '')}
              </>
            );
          }
          return renderWaitlistTable(isDe ? 'Warteliste' : 'Waitlist', waitlistRegs, 'dex-ui-dot--orange');
};

