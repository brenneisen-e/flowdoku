/* CancelledList — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13066-13454 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { EventService, SPRegistration } from '../../../services/EventService';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { formatDate } from '../../../utils/eventStatus';
import { PersonContactHover } from '../../PersonContactHover';
import { Trash2 } from '../../Icons';
import { DeloitteEvent } from '../../../types';

export interface CancelledListProps {
  cancelledRegs: (SPRegistration & { _subsiteUrl?: string; _sectionTitle?: string; _sectionId?: string; })[];
  cancelledSortAsc: boolean;
  cancelledSortColumn: "date" | "type" | "location" | "vorname" | "nachname" | "email" | "jobtitle";
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  consolidatedChildren: DeloitteEvent[];
  eventServiceRef: EventService;
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
  hasWaitlistActivity: boolean;
  isAdmin: boolean;
  isConsolidatedMode: boolean;
  isDe: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setCancelledSortAsc: React.Dispatch<React.SetStateAction<boolean>>;
  setCancelledSortColumn: React.Dispatch<React.SetStateAction<"date" | "type" | "location" | "vorname" | "nachname" | "email" | "jobtitle">>;
  setSubRegReloadTick: React.Dispatch<React.SetStateAction<number>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  stripLocPrefix: (loc: string) => string;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
}

export const CancelledList: React.FC<CancelledListProps> = (p) => {
  const { cancelledRegs, cancelledSortAsc, cancelledSortColumn, confirmDialog, consolidatedChildren, eventServiceRef, hasWaitlistActivity, isAdmin, isConsolidatedMode, isDe, isOrganizerFor, registrations, reloadRegistrations, selectedEvent, setCancelledSortAsc, setCancelledSortColumn, setSubRegReloadTick, showAlert, stripLocPrefix, subEventRegsByEventId } = p;
          // v18.11: Abmeldungs-Liste mit denselben Spalten + Sortierung wie
          // Teilnehmer-/Warteliste. Unterscheidet proaktive Absagen
          // (CustomData _declined = „Ich nehme nicht teil", ohne vorherige
          // Anmeldung) von regulären Abmeldungen.
          const isDeclined = (reg: SPRegistration): boolean => {
            try { return !!(JSON.parse(reg.CustomData || '{}')._declined); } catch { return false; }
          };
          const safe = (s: string | undefined): string => (s || '').toLowerCase();
          const dateMs = (s: string | undefined): number => s ? new Date(s).getTime() : 0;
          const dir = cancelledSortAsc ? 1 : -1;
          const sorted = cancelledRegs.slice().sort((a, b) => {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const anyA = a as any; const anyB = b as any;
            switch (cancelledSortColumn) {
              case 'vorname': return safe(a.Vorname).localeCompare(safe(b.Vorname), 'de') * dir;
              case 'nachname': return safe(a.Nachname).localeCompare(safe(b.Nachname), 'de') * dir;
              case 'email': return safe(a.ParticipantEmail).localeCompare(safe(b.ParticipantEmail)) * dir;
              case 'jobtitle': return safe(anyA.JobTitle).localeCompare(safe(anyB.JobTitle), 'de') * dir;
              case 'location': return safe(anyA.Location).localeCompare(safe(anyB.Location), 'de') * dir;
              case 'type': return ((isDeclined(a) ? 1 : 0) - (isDeclined(b) ? 1 : 0)) * dir;
              case 'date': return (dateMs(a.CancellationDate) - dateMs(b.CancellationDate)) * dir;
            }
            return 0;
          });
          const arrow = (k: typeof cancelledSortColumn): string => k === cancelledSortColumn ? (cancelledSortAsc ? ' ▲' : ' ▼') : '';
          const toggleSort = (k: typeof cancelledSortColumn): void => {
            if (cancelledSortColumn === k) setCancelledSortAsc(v => !v);
            else { setCancelledSortColumn(k); setCancelledSortAsc(true); }
          };
          // v24.93: NICHT sticky (steht in overflowX-Wrapper, s.o.).
          const thClickable: React.CSSProperties = { textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none', background: '#fff', borderBottom: '2px solid var(--dex-gray-200)' };
          const declineCount = cancelledRegs.filter(isDeclined).length;
          // v24.82: Abmeldungen dürfen NUR bei Entwurf-Events (isFictive)
          // gelöscht werden — z.B. zum Aufräumen von Test-Anmeldungen, BEVOR
          // das Event live geht. Sobald das Event live war/ist, bleiben
          // Abmeldungen wegen der einjährigen Aufbewahrungsfrist erhalten
          // (dann kein „Löschen"-Button).
          const canDelete = !!selectedEvent && selectedEvent.isFictive === true && (isAdmin || isOrganizerFor(selectedEvent)) && !!selectedEvent.subsiteUrl;
          const deleteCancelled = async (reg: SPRegistration): Promise<void> => {
            if (!selectedEvent) return;
            // v22.59: im Klammer-Modus die Subsite der jeweiligen Sub-Section
            // nutzen (die Zeile trägt sie mit), sonst die Klammer-Subsite.
            const targetSubsite = (reg as SPRegistration & { _subsiteUrl?: string })._subsiteUrl || selectedEvent.subsiteUrl;
            if (!targetSubsite) return;
            const nm = `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantName || reg.ParticipantEmail;
            const msg = isDe
              ? `Diese abgemeldete Registrierung von „${nm}" ENDGÜLTIG löschen?\n\nDie Zeile wird komplett aus der Teilnehmerliste entfernt und kann NICHT wiederhergestellt werden. (Nützlich z.B. zum Aufräumen von Test-Anmeldungen.)`
              : `Permanently DELETE this cancelled registration of „${nm}"?\n\nThe row is removed entirely from the participant list and CANNOT be restored. (Useful e.g. for cleaning up test registrations.)`;
            if (!(await confirmDialog(msg, { danger: true, title: isDe ? 'Registrierung löschen' : 'Delete registration', confirmLabel: isDe ? 'Endgültig löschen' : 'Delete permanently' }))) return;
            const ok = await eventServiceRef.deleteRegistration(targetSubsite, reg.Id);
            if (ok) {
              try {
                await eventServiceRef.writeChangeLog({
                  action: 'RegistrationDeleted',
                  targetType: 'Participant',
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  targetId: ((reg as any).ParticipantEmail || '') + '#' + reg.Id,
                  targetName: nm,
                  eventId: selectedEvent.id,
                  eventTitle: selectedEvent.title,
                  details: { deletedStatus: reg.Status, cancellationDate: reg.CancellationDate || '' },
                });
              } catch { /* Audit best-effort */ }
              // v22.59/v22.63: im Klammer-Modus sowohl die Sub-Event-Listen
              // (Sub-Section-Abmeldungen) ALS AUCH die Klammer-Registrierungen
              // (z.B. Absagen auf der Klammer) neu laden, sonst die Event-Regs.
              if (isConsolidatedMode) {
                setSubRegReloadTick(t => t + 1);
                await reloadRegistrations();
              } else {
                await reloadRegistrations();
              }
            } else {
              // eslint-disable-next-line no-alert
              showAlert(isDe ? 'Löschen fehlgeschlagen.' : 'Delete failed.');
            }
          };
          // v22.63: Konsolidierte Abmelde-Matrix — EINE Zeile pro Person, mit
          // einem ✗ je Section (Gesamt-Event + Sub-Events), in der sich die
          // Person abgemeldet hat. Analog zur konsolidierten Anmelde-Matrix.
          if (isConsolidatedMode) {
            const sectionCols: Array<{ id: string; title: string }> = [
              ...(cancelledRegs.some(r => r._sectionId === '__parent') ? [{ id: '__parent', title: isDe ? 'Gesamt-Event' : 'Overall event' }] : []),
              ...consolidatedChildren
                .filter(ch => cancelledRegs.some(r => r._sectionId === ch.id))
                .map(ch => ({ id: ch.id, title: shortSubEventTitle(ch.title, selectedEvent!.title) })),
            ];
            type CancelRow = SPRegistration & { _subsiteUrl?: string; _sectionId?: string };
            interface CancelPerson { email: string; firstName: string; lastName: string; jobTitle: string; location: string; latest: number; declinedAny: boolean; bySection: Record<string, CancelRow> }
            const peopleMap = new Map<string, CancelPerson>();
            for (const r of cancelledRegs) {
              const key = (r.ParticipantEmail || '').toLowerCase().trim();
              if (!key) continue;
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const anyR = r as any;
              let p = peopleMap.get(key);
              if (!p) { p = { email: r.ParticipantEmail || '', firstName: '', lastName: '', jobTitle: '', location: '', latest: 0, declinedAny: false, bySection: {} }; peopleMap.set(key, p); }
              if (!p.firstName && r.Vorname) p.firstName = r.Vorname;
              if (!p.lastName && r.Nachname) p.lastName = r.Nachname;
              if (!p.jobTitle && anyR.JobTitle) p.jobTitle = anyR.JobTitle;
              if (!p.location && anyR.Location) p.location = anyR.Location;
              p.bySection[r._sectionId || '__parent'] = r;
              const tms = r.CancellationDate ? new Date(r.CancellationDate).getTime() : 0;
              if (tms > p.latest) p.latest = tms;
              if (isDeclined(r)) p.declinedAny = true;
            }
            const pdir = cancelledSortAsc ? 1 : -1;
            const people = Array.from(peopleMap.values()).sort((a, b) => {
              switch (cancelledSortColumn) {
                case 'nachname': return a.lastName.toLowerCase().localeCompare(b.lastName.toLowerCase(), 'de') * pdir;
                case 'email': return a.email.toLowerCase().localeCompare(b.email.toLowerCase()) * pdir;
                case 'date': return (a.latest - b.latest) * pdir;
                default: return a.firstName.toLowerCase().localeCompare(b.firstName.toLowerCase(), 'de') * pdir;
              }
            });
            const declinePeople = people.filter(p => p.declinedAny).length;
            const deletePerson = async (p: CancelPerson): Promise<void> => {
              if (!selectedEvent) return;
              const emailLc = p.email.toLowerCase().trim();
              // v22.66: Person ÜBERALL löschen — nicht nur die Abmelde-Zeilen,
              // sondern ALLE Zeilen dieser E-Mail über die Klammer UND alle
              // Sub-Events (inkl. der aktiven „Schatten"-Zeile auf der Klammer,
              // die beim reinen Abmelden/Löschen sonst verwaist liegen bleibt).
              // v30.68: Termine ZUERST, Klammer ZULETZT — und die Klammer nur,
              // wenn jede Termin-Zeile weg ist. Vorher stand die Klammer vorn:
              // Ein 429 bei der dritten Termin-Zeile ließ die Person auf dem
              // Termin stehen, während ihre Klammer-Zeile schon gelöscht war.
              const parentTargets: Array<{ sub: string; id: number }> = [];
              const subTargets: Array<{ sub: string; id: number }> = [];
              if (selectedEvent.subsiteUrl) {
                for (const r of registrations) {
                  if ((r.ParticipantEmail || '').toLowerCase().trim() === emailLc) parentTargets.push({ sub: selectedEvent.subsiteUrl, id: r.Id });
                }
              }
              for (const c of consolidatedChildren) {
                if (!c.subsiteUrl) continue;
                for (const r of (subEventRegsByEventId[c.id] || [])) {
                  if ((r.ParticipantEmail || '').toLowerCase().trim() === emailLc) subTargets.push({ sub: c.subsiteUrl, id: r.Id });
                }
              }
              const targets = subTargets.concat(parentTargets);
              const nm = `${p.firstName} ${p.lastName}`.trim() || p.email;
              const msg = isDe
                ? `„${nm}" wirklich überall löschen? Alle ${targets.length} Einträge dieser Person (Gesamt-Event + Sub-Events) werden endgültig entfernt und können nicht wiederhergestellt werden.`
                : `Permanently delete „${nm}" everywhere? All ${targets.length} entries of this person (overall event + sub-events) will be removed and cannot be restored.`;
              if (!(await confirmDialog(msg, { danger: true, title: isDe ? 'Person überall löschen' : 'Delete person everywhere', confirmLabel: isDe ? 'Endgültig löschen' : 'Delete permanently' }))) return;
              // v30.67: `deleteRegistration` wirft nicht, es liefert false — das
              // try/catch fing also nie etwas. Der Audit-Eintrag behauptete
              // `count: 12, everywhere: true`, auch wenn drei Zeilen auf einer
              // Subsite ohne Delete-Recht liegen blieben; die Person zählte dort
              // weiter, bekam Mails und Outlook-Termine. Zählen, was WIRKLICH
              // gelöscht wurde, und es sagen.
              let delOk = 0;
              let delFailed = 0;
              for (const t of subTargets) {
                const ok = await eventServiceRef.deleteRegistration(t.sub, t.id).catch(() => false);
                if (ok) delOk += 1; else delFailed += 1;
              }
              let parentSkipped = 0;
              if (delFailed === 0) {
                for (const t of parentTargets) {
                  const ok = await eventServiceRef.deleteRegistration(t.sub, t.id).catch(() => false);
                  if (ok) delOk += 1; else delFailed += 1;
                }
              } else {
                parentSkipped = parentTargets.length;
              }
              try {
                await eventServiceRef.writeChangeLog({ action: 'RegistrationDeleted', targetType: 'Participant', targetId: p.email, targetName: nm, eventId: selectedEvent.id, eventTitle: selectedEvent.title, details: { deletedStatus: 'Abgemeldet', count: delOk, failed: delFailed, parentSkipped, everywhere: delFailed === 0 } });
              } catch { /* */ }
              setSubRegReloadTick(t => t + 1);
              // v30.67 (Review): gemeinsamer Nachlade-Pfad. Genau hier schnappte
              // die Falle zu: „Person überall löschen" über 12 Zeilen löst die
              // Drosselung aus, der Reload danach bekam 429 → `[]` → „Noch
              // keine Teilnehmer registriert", alle Kacheln 0, kein Hinweis.
              await reloadRegistrations();
              if (delFailed > 0) {
                showAlert(isDe
                  ? `${delOk} von ${targets.length} Einträgen gelöscht — ${delFailed} konnten nicht entfernt werden (fehlende Rechte auf dem Termin oder Drosselung).${parentSkipped > 0 ? ' Die Klammer-Zeile wurde deshalb bewusst stehen gelassen.' : ''} Bitte „Organizer-Berechtigungen reparieren“ ausführen und erneut versuchen.`
                  : `${delOk} of ${targets.length} entries deleted — ${delFailed} could not be removed (missing permissions on the date or throttling).${parentSkipped > 0 ? ' The umbrella row was therefore deliberately kept.' : ''} Please run „Repair organizer permissions“ and try again.`,
                  { variant: 'error' });
              }
            };
            return (
              <>
                <h4 style={{ marginTop: 24, color: 'var(--dex-gray-400)' }}>
                  {isDe ? 'Abmeldungen' : 'Cancellations'} ({people.length})
                  {declinePeople > 0 && (
                    <span style={{ fontSize: '0.8rem', fontWeight: 400, marginLeft: 8, color: 'var(--dex-gray-500)' }}>
                      {isDe ? `davon ${declinePeople} Absage(n) ohne Anmeldung` : `incl. ${declinePeople} decline(s) without registration`}
                    </span>
                  )}
                </h4>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                    <thead>
                      <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                        <th style={thClickable} onClick={() => toggleSort('vorname')}>Vorname{arrow('vorname')}</th>
                        <th style={thClickable} onClick={() => toggleSort('nachname')}>Nachname{arrow('nachname')}</th>
                        <th style={thClickable} onClick={() => toggleSort('email')}>Email{arrow('email')}</th>
                        <th style={{ ...thClickable, cursor: 'default' }}>Job Title</th>
                        <th style={{ ...thClickable, cursor: 'default' }}>Standort</th>
                        {sectionCols.map(sc => (
                          <th key={sc.id} style={{ ...thClickable, cursor: 'default', textAlign: 'center' }} title={sc.title}>{sc.title}</th>
                        ))}
                        <th style={thClickable} onClick={() => toggleSort('date')}>{isDe ? 'Letzte Abmeldung' : 'Last cancellation'}{arrow('date')}</th>
                        {canDelete && (
                          <th style={{ ...thClickable, cursor: 'default', textAlign: 'right' }}>{isDe ? 'Löschen' : 'Delete'}</th>
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {people.map(p => (
                        <tr key={p.email} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                          <td style={{ padding: 8, fontWeight: 500 }}>{p.firstName || '-'}</td>
                          <td style={{ padding: 8, fontWeight: 500 }}>{p.lastName || '-'}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{p.email}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{p.jobTitle || '-'}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{p.location || '-'}</td>
                          {sectionCols.map(sc => {
                            const r = p.bySection[sc.id];
                            // v23.2/v23.6: In der „Gesamt-Event"-Spalte kein nacktes
                            // X, sondern dieselbe Darstellung wie in der Status-
                            // Spalte der Teilnehmerliste: blaue Pille „Absage (nicht
                            // angemeldet)" für Leute, die sich NIE angemeldet, aber
                            // hinterlegt haben, dass sie nicht teilnehmen können;
                            // rotes „Abgemeldet" für echte Abmeldungen vom Gesamt-Event.
                            if (sc.id === '__parent') {
                              return (
                                <td key={sc.id} style={{ padding: 8, textAlign: 'center' }}>
                                  {r
                                    ? (isDeclined(r)
                                        ? <span
                                            title={`${isDe ? 'Diese Person hat sich NICHT angemeldet, sondern hinterlegt, dass sie nicht am Event teilnehmen kann (Absage ohne Anmeldung).' : 'This person did NOT register but recorded that they cannot attend the event (decline without registration).'} — ${formatDate(r.CancellationDate)}`}
                                            style={{ display: 'inline-block', fontSize: '0.72rem', fontWeight: 600, padding: '2px 8px', borderRadius: 999, background: 'rgba(0,118,168,0.10)', color: 'var(--dex-blue, #0076a8)', whiteSpace: 'nowrap' }}
                                          >
                                            {isDe ? 'Absage (nicht angemeldet)' : 'Decline (never registered)'}
                                          </span>
                                        : <span
                                            title={`${isDe ? 'Diese Person war für das Gesamt-Event angemeldet und hat sich wieder abgemeldet.' : 'This person was registered for the overall event and later cancelled.'} — ${formatDate(r.CancellationDate)}`}
                                            style={{ display: 'inline-block', fontSize: '0.72rem', fontWeight: 700, color: 'var(--dex-red, #da291c)', whiteSpace: 'nowrap' }}
                                          >
                                            {isDe ? 'Abgemeldet' : 'Cancelled'}
                                          </span>)
                                    : <span style={{ color: 'var(--dex-gray-300)' }}>–</span>}
                                </td>
                              );
                            }
                            return (
                              <td key={sc.id} style={{ padding: 8, textAlign: 'center' }}>
                                {r
                                  ? <span title={`${isDeclined(r) ? (isDe ? 'Absage (nicht angemeldet)' : 'Decline (never registered)') : (isDe ? 'Abgemeldet' : 'Cancelled')} — ${formatDate(r.CancellationDate)}`} style={{ color: 'var(--dex-red, #da291c)', fontWeight: 700, fontSize: '1rem' }}>&#10007;</span>
                                  : <span style={{ color: 'var(--dex-gray-300)' }}>–</span>}
                              </td>
                            );
                          })}
                          <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{p.latest ? formatDate(new Date(p.latest).toISOString()) : '-'}</td>
                          {canDelete && (
                            <td style={{ padding: 8, textAlign: 'right' }}>
                              <button
                                type="button"
                                className="btn btn-outline"
                                style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #da291c)', borderColor: 'var(--dex-red, #da291c)' }}
                                onClick={() => { deletePerson(p).catch(() => { /* */ }); }}
                              >
                                {isDe ? 'Löschen' : 'Delete'}
                              </button>
                            </td>
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            );
          }
          // v24.82: Abmeldungen im selben Zeilen-Layout wie die aktiven
          // Anmeldungen (Foto + Name + „Position • Standort • Firma"), aber
          // alle Texte in hellem Grau, damit sie klar von den aktiven
          // Anmeldungen zu unterscheiden sind. Ein „Löschen"-Button erscheint
          // NUR bei Entwurf-Events (canDelete) — sonst bleiben Abmeldungen
          // wegen der einjährigen Aufbewahrungsfrist erhalten.
          const greyText = 'var(--dex-gray-400)';
          return (
            <>
              <h4 style={{ marginTop: 24, color: 'var(--dex-gray-400)' }}>
                {isDe ? 'Abmeldungen' : 'Cancellations'} ({cancelledRegs.length})
                {declineCount > 0 && (
                  <span style={{ fontSize: '0.8rem', fontWeight: 400, marginLeft: 8, color: 'var(--dex-gray-500)' }}>
                    {isDe ? `davon ${declineCount} Absage(n)` : `incl. ${declineCount} decline(s)`}
                  </span>
                )}
              </h4>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                      <th style={thClickable} onClick={() => toggleSort('nachname')}>{isDe ? 'Teilnehmer' : 'Attendee'}{arrow('nachname')}</th>
                      <th style={thClickable} onClick={() => toggleSort('type')}>{isDe ? 'Art' : 'Type'}{arrow('type')}</th>
                      <th style={{ ...thClickable, cursor: 'default' }}>{isDe ? 'Abgemeldet von' : 'Cancelled by'}</th>
                      {isConsolidatedMode && (
                        <th style={{ ...thClickable, cursor: 'default' }}>{isDe ? 'Sub-Event' : 'Sub-event'}</th>
                      )}
                      <th style={thClickable} onClick={() => toggleSort('date')}>{isDe ? 'Abgemeldet am' : 'Cancelled on'}{arrow('date')}</th>
                      <th style={{ ...thClickable, cursor: 'default' }}>{isDe ? 'Nach Frist?' : 'After deadline?'}</th>
                      {/* v19.4: „Wurde ersetzt durch" — die nachgerückte Person, die
                          den frei gewordenen Platz übernommen hat (vom Flow gesetzt).
                          v19.11: nur bei Events mit echter Warteliste-/Nachrück-
                          Aktivität (sonst durchgehend leer). */}
                      {hasWaitlistActivity && (
                        <th style={{ ...thClickable, cursor: 'default' }}>{isDe ? 'Wurde ersetzt durch' : 'Replaced by'}</th>
                      )}
                      {canDelete && (
                        <th style={{ ...thClickable, cursor: 'default', textAlign: 'right' }}>{isDe ? 'Löschen' : 'Delete'}</th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {sorted.map(reg => {
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const anyReg = reg as any;
                      const declined = isDeclined(reg);
                      const vn = reg.Vorname || ((reg.ParticipantName || '').split(' ')[0] || '');
                      let nn = reg.Nachname || '';
                      if (!nn && reg.ParticipantName) { const pp = reg.ParticipantName.trim().split(/\s+/); if (pp.length > 1) nn = pp.slice(1).join(' '); }
                      const fullName = `${vn} ${nn}`.trim() || reg.ParticipantEmail || '-';
                      const sub = [String(anyReg.JobTitle || ''), stripLocPrefix(String(anyReg.Location || '')), String(anyReg.Company || '')].filter(Boolean).join(' • ');
                      // v24.88: Status-Pille wieder FARBIG (blau = Absage ohne
                      // Anmeldung, rot = abgemeldet) — der Rest der Zeile bleibt grau.
                      const artLabel = declined
                        ? (isDe ? 'Absage (nicht angemeldet)' : 'Decline (never registered)')
                        : (isDe ? 'Abgemeldet' : 'Cancelled');
                      return (
                        <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                          <td style={{ padding: 8 }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
                              <PersonContactHover email={reg.ParticipantEmail || ''} name={fullName} size={30} subline={sub} isDe={isDe} />
                              <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0, lineHeight: 1.25 }}>
                                <span style={{ fontWeight: 600, color: greyText, whiteSpace: 'nowrap' }}>{fullName}</span>
                                {sub && <span style={{ fontSize: '0.78rem', color: greyText, whiteSpace: 'nowrap' }}>{sub}</span>}
                              </div>
                            </div>
                          </td>
                          <td style={{ padding: 8 }}>
                            {declined
                              ? <span style={{ fontSize: '0.72rem', fontWeight: 600, padding: '2px 8px', borderRadius: 999, background: 'rgba(0,118,168,0.10)', color: 'var(--dex-blue, #0076a8)' }}>{artLabel}</span>
                              : <span style={{ fontSize: '0.72rem', fontWeight: 600, padding: '2px 8px', borderRadius: 999, background: 'rgba(218,41,28,0.08)', color: 'var(--dex-red, #da291c)' }}>{artLabel}</span>}
                          </td>
                          {/* v24.88: „Abgemeldet von" — selbst abgemeldet vs. durch
                              jemand anderen (Audit CancelledBy*), analog „Angemeldet von". */}
                          <td style={{ padding: 8, color: greyText, fontSize: '0.8rem' }}>
                            {(() => {
                              const cby = (anyReg.CancelledByEmail || '').toLowerCase();
                              const pe = (reg.ParticipantEmail || '').toLowerCase();
                              if (!cby) return <span style={{ color: 'var(--dex-gray-300)' }}>—</span>;
                              if (cby === pe) return <span>{isDe ? 'Selbst abgemeldet' : 'Self'}</span>;
                              return <span title={anyReg.CancelledByEmail}>{anyReg.CancelledByName || anyReg.CancelledByEmail}</span>;
                            })()}
                          </td>
                          {isConsolidatedMode && (
                            <td style={{ padding: 8, color: greyText, fontSize: '0.8rem' }}>{(reg as SPRegistration & { _sectionTitle?: string })._sectionTitle || '-'}</td>
                          )}
                          <td style={{ padding: 8, color: greyText }}>{formatDate(reg.CancellationDate)}</td>
                          {/* v24.88: Markierung, wenn die Abmeldung NACH der
                              kommunizierten Abmeldefrist (lastDeregisterDate) erfolgte. */}
                          <td style={{ padding: 8, fontSize: '0.8rem' }}>
                            {(() => {
                              if (declined) return <span style={{ color: 'var(--dex-gray-300)' }}>—</span>;
                              let ev = selectedEvent;
                              if (isConsolidatedMode) {
                                const sid = (reg as SPRegistration & { _sectionId?: string })._sectionId;
                                if (sid && sid !== '__parent') { const ch = consolidatedChildren.find(c => c.id === sid); if (ch) ev = ch; }
                              }
                              const dlRaw = ev?.lastDeregisterDate;
                              if (!dlRaw || !reg.CancellationDate) return <span style={{ color: 'var(--dex-gray-300)' }}>—</span>;
                              const isLate = new Date(reg.CancellationDate).getTime() > new Date(dlRaw).getTime();
                              return isLate
                                ? <span title={`${isDe ? 'Abmeldefrist war' : 'Deadline was'}: ${formatDate(dlRaw)}`} style={{ fontSize: '0.72rem', fontWeight: 700, padding: '2px 8px', borderRadius: 999, background: 'rgba(237,139,0,0.15)', color: 'var(--dex-orange-dark, #b35a00)' }}>{isDe ? 'Nach Frist' : 'After deadline'}</span>
                                : <span style={{ color: 'var(--dex-gray-300)' }}>—</span>;
                            })()}
                          </td>
                          {hasWaitlistActivity && (
                            <td style={{ padding: 8, color: greyText, fontSize: '0.8rem' }}>
                              {(() => {
                                const email = (anyReg.ReplacedByParticipantEmail as string | undefined) || '';
                                if (!email) return <span style={{ color: 'var(--dex-gray-300)' }}>—</span>;
                                const other = registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === email.toLowerCase());
                                const label = other ? (((other.Vorname || '') + ' ' + (other.Nachname || '')).trim() || other.ParticipantName || email) : email;
                                return <span title={email}>{label}</span>;
                              })()}
                            </td>
                          )}
                          {canDelete && (
                            <td style={{ padding: 8, textAlign: 'right' }}>
                              <button
                                type="button"
                                title={isDe ? 'Registrierung endgültig löschen (nur im Entwurf möglich)' : 'Permanently delete registration (drafts only)'}
                                onClick={() => { deleteCancelled(reg).catch(() => { /* */ }); }}
                                style={{
                                  display: 'inline-flex', alignItems: 'center', gap: 5,
                                  border: '1px solid var(--dex-red, #da291c)', background: 'rgba(218,41,28,0.06)',
                                  color: 'var(--dex-red, #da291c)', borderRadius: 6, padding: '4px 9px',
                                  fontSize: '0.74rem', fontWeight: 600, cursor: 'pointer',
                                }}
                              >
                                <Trash2 size={13} /> {isDe ? 'Löschen' : 'Delete'}
                              </button>
                            </td>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </>
          );
};

