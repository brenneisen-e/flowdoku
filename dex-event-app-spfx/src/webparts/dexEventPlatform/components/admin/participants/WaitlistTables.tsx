/* WaitlistTables — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 12893-13064 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { EventService, SPRegistration } from '../../../services/EventService';
import { formatDate } from '../../../utils/eventStatus';
import { isEventOver } from '../../../utils/eventFormat';
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
          // Seit v6.5: bei B2Run-Split-Kapazitäten getrennte Wartelisten-Tabellen pro
          // PreferredStarterType. Ohne Split: eine einzige Warteliste wie bisher.
          const renderWaitlistTable = (title: string, regs: SPRegistration[], accentColor: string): React.ReactElement | null => {
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
            const arrow = (k: typeof waitlistSortColumn): string => k === waitlistSortColumn ? (waitlistSortAsc ? ' ▲' : ' ▼') : '';
            const toggleSort = (k: typeof waitlistSortColumn): void => {
              if (waitlistSortColumn === k) setWaitlistSortAsc(v => !v);
              else { setWaitlistSortColumn(k); setWaitlistSortAsc(true); }
            };
            // v24.93: NICHT sticky — diese Tabellen stehen in einem
            // overflowX:'auto'-Wrapper, in dem position:sticky relativ zum
            // Wrapper (statt zum Fenster) berechnet wird; der Kopf schwebte
            // dadurch mitten in der Tabelle und schnitt Zeilen ab. Sticky
            // bleibt nur bei der Teilnehmer-Tabelle (eigener renderHeader,
            // ohne Scroll-Wrapper).
            const thClickable: React.CSSProperties = { textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none', background: '#fff', borderBottom: '2px solid var(--dex-gray-200)' };
            return (
              <React.Fragment key={title}>
                <h4 style={{ marginTop: 24, color: accentColor }}>{title} ({regs.length})</h4>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                    <thead>
                      <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                        <th style={thClickable} onClick={() => toggleSort('pos')}>Platz{arrow('pos')}</th>
                        <th style={thClickable} onClick={() => toggleSort('vorname')}>Vorname{arrow('vorname')}</th>
                        <th style={thClickable} onClick={() => toggleSort('nachname')}>Nachname{arrow('nachname')}</th>
                        <th style={thClickable} onClick={() => toggleSort('email')}>Email{arrow('email')}</th>
                        <th style={thClickable} onClick={() => toggleSort('jobtitle')}>Job Title{arrow('jobtitle')}</th>
                        <th style={thClickable} onClick={() => toggleSort('location')}>Standort{arrow('location')}</th>
                        {isSplitCapacity && <th style={{ textAlign: 'left', padding: 8 }}>Wunsch-Typ</th>}
                        <th style={thClickable} onClick={() => toggleSort('date')}>Registriert am{arrow('date')}</th>
                        <th style={{ textAlign: 'left', padding: 8 }}>Aktion</th>
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
                        // wird zusätzlich die laufende Treffer-Nummer „#n" vorangestellt.
                        const truePos = waitlistTruePos[reg.Id];
                        const posDisplay = (truePos != null)
                          ? (query ? `#${i + 1} (Platz ${truePos})` : String(truePos))
                          : String(i + 1);
                        return (
                        <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                          <td style={{ padding: 8, fontWeight: 600, color: accentColor }}>{posDisplay}</td>
                          <td style={{ padding: 8, fontWeight: 500 }}>{reg.Vorname || '-'}</td>
                          <td style={{ padding: 8, fontWeight: 500 }}>{reg.Nachname || '-'}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{anyReg.JobTitle || '-'}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{anyReg.Location || '-'}</td>
                          {isSplitCapacity && (
                            <td style={{ padding: 8, color: 'var(--dex-gray-700)' }}>
                              {reg.PreferredStarterType || '—'}
                            </td>
                          )}
                          <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                          <td style={{ padding: 8, whiteSpace: 'nowrap' }}>
                            {/* v28.70: Wartelisten-Platz manuell setzen. Die
                                Position ist der Rang nach TeilnehmerID — genau
                                danach sortieren App-Nachrücken UND der Flow. */}
                            <button
                              className="btn btn-secondary"
                              style={{ fontSize: '0.75rem', padding: '4px 10px', marginRight: 6 }}
                              disabled={wlPosBusy}
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
                            <button
                              className="btn btn-secondary"
                              style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #c00)' }}
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
                                  try {
                                    const ok = await eventServiceRef.queueIDReorder(
                                      selectedEvent.id, selectedEvent.eventNumber || 0,
                                      selectedEvent.subsiteUrl, selectedEvent.title,
                                      `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantName || undefined,
                                      reg.ParticipantEmail || undefined
                                    );
                                    if (!ok) {
                                      showAlert(isDe ? 'Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.' : 'Cancellation successful, but the ID reorder entry could not be written to the queue. Please click "Reassign IDs" once.');
                                    }
                                  } catch {
                                    showAlert(isDe ? 'Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.' : 'Cancellation successful, but the ID reorder entry could not be written to the queue. Please click "Reassign IDs" once.');
                                  }
                                }
                                // v30.67 (Review): gemeinsamer Nachlade-Pfad statt `[]` bei 429.
                                await reloadRegistrations();
                              }}
                            >
                              Entfernen
                            </button>
                          </td>
                        </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </React.Fragment>
            );
          };

          if (isSplitCapacity) {
            // v11.6: Wartelisten-Tabellen mit den frei wählbaren Gruppen-
            // Labels statt hartcodeten 'Durchstarter'/'Funstarter'.
            const wlLabelA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
            const wlLabelB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
            // v11.29: Reihenfolge respektiert splitDisplayOrderReversed
            // (gleicher Toggle wie auf Register-Page + Kapazitäts-Cards).
            const wlA = renderWaitlistTable(`Warteliste ${wlLabelA}`, waitlistDurch, 'var(--dex-green-dark, #6b9a1e)');
            const wlB = renderWaitlistTable(`Warteliste ${wlLabelB}`, waitlistFun, 'var(--dex-orange, #ff8c00)');
            const reversed = !!selectedEvent?.splitDisplayOrderReversed;
            return (
              <>
                {reversed ? <>{wlB}{wlA}</> : <>{wlA}{wlB}</>}
                {renderWaitlistTable('Warteliste ohne Gruppe', waitlistUnassigned, 'var(--dex-gray-500)')}
              </>
            );
          }
          return renderWaitlistTable('Warteliste', waitlistRegs, 'var(--dex-orange)');
};

