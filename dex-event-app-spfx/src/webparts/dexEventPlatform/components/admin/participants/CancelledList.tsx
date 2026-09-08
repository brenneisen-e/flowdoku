/* CancelledList — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13066-13454 des
 * Stands vor dem Schnitt). Die Logik ist seither unverändert; die Oberfläche
 * folgt seit v31.3 dem UI-Leitfaden (Karte mit klickbarer Kopfzeile,
 * dex-ui-Tabelle, Personen-Zelle, Status-Pillen). Die Anzeige-Bedingung bleibt
 * beim Aufrufer.
 */
import * as React from 'react';
import { EventService, SPRegistration } from '../../../services/EventService';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { groupSubEventTabs, stripGroupPrefix } from '../../../utils/subEventGroups';
import { formatDate } from '../../../utils/eventStatus';
import { PersonContactHover } from '../../PersonContactHover';
import { ChevronDown, Trash2, Users } from '../../Icons';
import { cx, ensureDexUiStyles } from '../../dexUi';
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
  // v31.3: Die gemeinsamen Klassen einmal ins Dokument — Inline-Styles können
  // kein :hover, und hier sind Kopfzeile, Kopfzellen und Termin-Zeilen klickbar.
  ensureDexUiStyles();
  // v30.83: Sicht der konsolidierten Abmeldungen („nach Person" / „nach Tag")
  // und der aufgeklappte Tag. Hooks VOR dem isConsolidatedMode-Zweig — der
  // hat ein eigenes return (rules-of-hooks).
  const [cancelView, setCancelView] = React.useState<'person' | 'day'>('person');
  const [openDay, setOpenDay] = React.useState<string | null>(null);
  // v31.3: Auf-/zugeklappte Karte (siehe `shell`). Startet ZU (Leitfaden 5a,
  // Punkt 7) — wer die Event-Seite öffnet, will wissen, wer da IST; der Zähler
  // im Kopf sagt, ob sich das Aufklappen lohnt.
  const [openCard, setOpenCard] = React.useState(false);
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
          const toggleSort = (k: typeof cancelledSortColumn): void => {
            if (cancelledSortColumn === k) setCancelledSortAsc(v => !v);
            else { setCancelledSortColumn(k); setCancelledSortAsc(true); }
          };
          // v24.93: NICHT sticky (steht in overflowX-Wrapper, s.o.).
          // v30.87: Dieselbe Kopfzeile wie die Anmeldungen-Tabelle — sticky im
          // eigenen Scroll-Container (Nutzer-Ansage 07.09.2026: „Abmeldungen in
          // die gleiche Struktur, Logik und Anzeige wie Anmeldungen — mit
          // Inline-Scroll").
          // v31.3: Das übernehmen jetzt `dex-ui-table-wrap--sticky` und
          // `th.is-sortable` — ein Stil-Objekt je Kopfzelle konnte keinen Hover
          // zeigen, obwohl ein Klick darauf sortiert (Leitfaden 1.3).
          const sortMark = (k: typeof cancelledSortColumn): React.ReactNode => k === cancelledSortColumn ? <span className="dex-ui-table-sort">{cancelledSortAsc ? '▲' : '▼'}</span> : null;
          const thSort = (k: typeof cancelledSortColumn, label: string): React.ReactElement => (
            <th
              className={cx('is-sortable', k === cancelledSortColumn && 'is-sorted')}
              onClick={() => toggleSort(k)}
              title={isDe ? 'Nach dieser Spalte sortieren' : 'Sort by this column'}
            >
              {label}{sortMark(k)}
            </th>
          );
          // v31.3 (Nachzug): Beide Namen stecken in EINER Personen-Zelle — die
          // Sortierung nach Vorname bleibt als kleiner Knopf im Kopf (wie in
          // `WaitlistTables`), mit gestoppter Weitergabe an die Kopfzelle.
          const miniSortBtn = (k: typeof cancelledSortColumn, label: string, hint: string): React.ReactElement => (
            <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" title={hint}
              style={{ fontSize: '0.66rem', padding: '1px 6px', marginLeft: 4, textTransform: 'none', letterSpacing: 0 }}
              onClick={e => { e.stopPropagation(); toggleSort(k); }}
            >{label}{sortMark(k)}</button>
          );
          // Ein fehlender Wert bleibt ein Strich — nie eine 0 und nie leer,
          // sonst liest sich „unbekannt" wie „nichts" (CLAUDE.md).
          const dash = <span style={{ color: 'var(--dex-gray-300)' }}>—</span>;
          // v24.88/v31.3: Farblogik der Status-Pille — blau = Absage ohne
          // vorherige Anmeldung, rot = abgemeldet. Beide Sichten nutzen sie.
          const pillCls = (declined: boolean): string => cx('dex-ui-pill', declined ? 'dex-ui-pill--blue' : 'dex-ui-pill--red');
          const declineCount = cancelledRegs.filter(isDeclined).length;
          // v24.82: Abmeldungen dürfen NUR bei Entwurf-Events (isFictive)
          // gelöscht werden — z.B. zum Aufräumen von Test-Anmeldungen, BEVOR
          // das Event live geht. Sobald das Event live war/ist, bleiben
          // Abmeldungen wegen der einjährigen Aufbewahrungsfrist erhalten
          // (dann kein „Löschen"-Button).
          const canDelete = !!selectedEvent && selectedEvent.isFictive === true && (isAdmin || isOrganizerFor(selectedEvent)) && !!selectedEvent.subsiteUrl;
          // v31.3: Beide Sichten (Klammer-Matrix und Einzel-Event) stecken in
          // derselben Karte mit klickbarer Kopfzeile — Titel, Zähler, Chevron
          // (Leitfaden 5a, Punkt 7).
          const shell = (count: number, declines: number, toolbar: React.ReactNode, body: React.ReactNode): React.ReactElement => (
            <div className={cx('dex-ui-card', 'dex-ui-card--list')} style={{ marginTop: 24 }}>
              <div
                role="button"
                tabIndex={0}
                aria-expanded={openCard}
                className="dex-ui-row"
                style={{ cursor: 'pointer', userSelect: 'none', flexWrap: 'wrap' }}
                onClick={() => setOpenCard(v => !v)}
                onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setOpenCard(v => !v); } }}
              >
                {/* v31.3 (Nachzug): echte Überschrift — sonst verliert die Sprungnavigation über Überschriften den Abschnitt. */}
                <h3 className="dex-ui-card-head-title"><Users size={18} /> {isDe ? 'Abmeldungen' : 'Cancellations'}</h3>
                <span className="dex-ui-pill dex-ui-pill--gray">{count}</span>
                {declines > 0 && (
                  <span className="dex-ui-pill dex-ui-pill--blue">
                    {isDe
                      ? `${declines} ${declines === 1 ? 'Absage' : 'Absagen'} ohne Anmeldung`
                      : `${declines} ${declines === 1 ? 'decline' : 'declines'} without registration`}
                  </span>
                )}
                <span className={cx('dex-ui-disclosure-chevron', openCard && 'is-open')} style={{ marginLeft: 'auto' }} aria-hidden="true">
                  <ChevronDown size={18} />
                </span>
              </div>
              {openCard && (
                <div style={{ padding: '0 6px 6px' }}>
                  <div className="dex-ui-muted" style={{ margin: '0 0 10px' }}>
                    {isDe
                      ? 'Wer sich abgemeldet hat — und wer von vornherein abgesagt hat. Löschen geht nur, solange das Event ein Entwurf ist.'
                      : 'Who cancelled — and who declined right away. Rows can only be deleted while the event is a draft.'}
                  </div>
                  {toolbar}
                  {body}
                </div>
              )}
            </div>
          );
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
            // v30.83: Die Abmelde-MATRIX (eine Spalte je Termin) ist bei
            // Kalender-Events unlesbar: 20 Datumsspalten, in denen 27 von 28
            // Zeilen fast nur „–" stehen, horizontales Scrollen, und die Frage
            // des Organizers („welcher Tag bröckelt?") bleibt unbeantwortet
            // (Nutzer-Befund 07.09.2026: „nicht so gelungen"). Jetzt zwei
            // Sichten: „nach Person" mit Termin-Chips je Zeile (kein Informations-
            // verlust, keine Breite) und „nach Termin" mit Anzahl, Balken und
            // aufklappbaren Namen je Tag. Die Teilnehmer-Matrix darüber bleibt —
            // dort stehen die Haken dicht, dort trägt das Raster.
            const isCal = !!selectedEvent.subEventCalendar;
            const childById = new Map<string, DeloitteEvent>();
            consolidatedChildren.forEach(c => childById.set(c.id, c));
            const fmtDay = (iso?: string): string => {
              if (!iso) return '';
              const d = new Date(iso);
              return isFinite(d.getTime()) ? d.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { weekday: 'short', day: '2-digit', month: '2-digit' }) : '';
            };
            const secLabel = (id: string): string => {
              if (id === '__parent') return isDe ? 'Gesamt-Event' : 'Overall event';
              const ch = childById.get(id);
              if (!ch) return id;
              return (isCal && fmtDay(ch.startDate)) || shortSubEventTitle(ch.title, selectedEvent.title) || ch.title;
            };
            const secOrder = new Map<string, number>();
            sectionCols.forEach((sc, i) => secOrder.set(sc.id, i));
            const chipsFor = (p: CancelPerson): React.ReactElement[] => Object.keys(p.bySection)
              .sort((a, b) => (secOrder.get(a) ?? 999) - (secOrder.get(b) ?? 999))
              .map(id => {
                const r = p.bySection[id];
                const declined = isDeclined(r);
                const label = id === '__parent'
                  ? (declined ? (isDe ? 'Absage (gesamt)' : 'Decline (overall)') : (isDe ? 'Gesamt-Event' : 'Overall event'))
                  : secLabel(id);
                return (
                  <span key={id} className={pillCls(declined)} title={`${declined ? (isDe ? 'Absage ohne Anmeldung' : 'Decline without registration') : (isDe ? 'Abgemeldet' : 'Cancelled')} — ${formatDate(r.CancellationDate)}`}>
                    {label}
                  </span>
                );
              });
            // Sicht „nach Termin": eine Zeile je Termin mit Abmeldungen, Gesamt-Event zuerst.
            const perSection = sectionCols.map(sc => {
              const rows = people.filter(p => !!p.bySection[sc.id]);
              const declinedOnly = rows.length > 0 && rows.every(p => isDeclined(p.bySection[sc.id]));
              return { id: sc.id, label: secLabel(sc.id), fullTitle: sc.id === '__parent' ? '' : ((childById.get(sc.id) || { title: '' }).title || ''), rows, declinedOnly };
            });
            const maxCount = perSection.reduce((m, s) => Math.max(m, s.rows.length), 0) || 1;
            const quietChildren = consolidatedChildren.length - sectionCols.filter(sc => sc.id !== '__parent').length;
            // Präfix-Gruppen („Day 1 - …") wie überall — nur bei Titel-, nicht bei Datums-Chips.
            const childSections = perSection.filter(s => s.id !== '__parent');
            const grouping = isCal ? { grouped: false, groups: [] as Array<{ label: string; idxs: number[] }> } : groupSubEventTabs(childSections.map(s => s.fullTitle));
            // v31.3: Die Termin-Zeile ist der Aufklapper — deshalb Zeilen-Hover
            // (`dex-ui-row`) und ein Chevron, der sich beim Öffnen dreht. Der
            // Balken nutzt `dex-ui-progress`; blau, wenn es an diesem Termin nur
            // Absagen ohne Anmeldung gab (Farblogik v24.88).
            const dayRow = (s: typeof perSection[number], shownLabel: string): React.ReactElement => {
              const open = openDay === s.id;
              return (
                // v31.3 (Nachzug): Trennlinie am Wrapper — am Knopf nahm
                // `--bordered:last-child` sie jeder zugeklappten Zeile weg.
                <div key={s.id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                  <button
                    type="button"
                    onClick={() => setOpenDay(open ? null : s.id)}
                    aria-expanded={open}
                    className={cx('dex-ui-row', 'dex-ui-rowbtn', open && 'is-active')}
                    style={{ padding: '8px 10px' }}
                  >
                    <span className="dex-ui-row-title" style={{ flex: '0 1 220px' }} title={s.fullTitle || shownLabel}>{shownLabel}</span>
                    <span className={pillCls(s.declinedOnly)}>{s.rows.length}</span>
                    <span className="dex-ui-progress" style={{ flex: '1 1 60px' }}>
                      {/* v31.3 (Nachzug): `display:block` ist Pflicht — auf einem Inline-Span wirkt width nicht, der Füllstand hatte null Größe. */}
                      <span
                        className="dex-ui-progress-bar dex-ui-progress-bar--red"
                        style={{ display: 'block', width: `${Math.max(4, Math.round((s.rows.length / maxCount) * 100))}%`, background: s.declinedOnly ? 'var(--dex-blue, #3860b2)' : undefined }}
                      />
                    </span>
                    <span aria-hidden="true" className={cx('dex-ui-disclosure-chevron', open && 'is-open')}><ChevronDown size={16} /></span>
                  </button>
                  {open && (
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, padding: '6px 10px 12px' }}>
                      {s.rows
                        .slice()
                        .sort((a, b) => a.lastName.toLowerCase().localeCompare(b.lastName.toLowerCase(), 'de'))
                        .map(p => {
                          const r = p.bySection[s.id];
                          const nm = `${p.firstName} ${p.lastName}`.trim() || p.email;
                          const decl = isDeclined(r);
                          return (
                            <span key={p.email} className={cx('dex-ui-pill', decl ? 'dex-ui-pill--blue' : 'dex-ui-pill--gray')}
                              title={`${p.email}${p.jobTitle ? ' · ' + p.jobTitle : ''}${p.location ? ' · ' + stripLocPrefix(p.location) : ''} — ${formatDate(r.CancellationDate)}`}>
                              {nm}
                              <span style={{ opacity: 0.75, fontWeight: 500 }}>· {formatDate(r.CancellationDate)}</span>
                              {decl && <span>· {isDe ? 'Absage' : 'Decline'}</span>}
                            </span>
                          );
                        })}
                    </div>
                  )}
                </div>
              );
            };
            const viewTab = (v: 'person' | 'day', label: string): React.ReactElement => (
              <button
                type="button"
                onClick={() => setCancelView(v)}
                aria-pressed={cancelView === v}
                className={cx('dex-ui-tab', cancelView === v && 'is-active')}
              >{label}</button>
            );
            // v31.3: Der Sicht-Umschalter steht als Reiter-Paar LINKS über der
            // Tabelle (Leitfaden 2a′) — vorher hing er rechts außen in der
            // Überschrift, wo er wie eine Beschriftung aussah.
            return shell(people.length, declinePeople, (
              <div className="dex-ui-toolbar">
                <div className="dex-ui-tabs">
                  {viewTab('person', isDe ? 'Nach Person' : 'By person')}
                  {viewTab('day', isCal ? (isDe ? 'Nach Tag' : 'By day') : (isDe ? 'Nach Termin' : 'By date'))}
                </div>
              </div>
            ), people.length === 0 ? (
              // Abmeldungen ohne E-Mail lassen sich keiner Person zuordnen —
              // besser ein benannter leerer Zustand als eine leere Fläche.
              <div className="dex-ui-empty">
                <div className="dex-ui-empty-title">{isDe ? 'Keine Abmeldung ist einer Person zuzuordnen' : 'No cancellation can be matched to a person'}</div>
                {isDe
                  ? 'Den Zeilen fehlt die E-Mail-Adresse — sie ist der Schlüssel dieser Übersicht.'
                  : 'The rows have no e-mail address — that is the key of this overview.'}
              </div>
            ) : cancelView === 'day' ? (
              <div>
                {perSection.filter(s => s.id === '__parent').map(s => dayRow(s, s.label))}
                {!grouping.grouped
                  ? childSections.map(s => dayRow(s, s.label))
                  : grouping.groups.map(g => {
                    const members = g.idxs.map(i => childSections[i]).filter(Boolean);
                    const sum = members.reduce((n, s) => n + s.rows.length, 0);
                    const label = g.label === 'Weitere' ? (isDe ? 'Weitere' : 'Other') : g.label;
                    return (
                      <div key={g.label} style={{ marginTop: 12 }}>
                        <div className="dex-ui-section-title">
                          {/* v31.3 (Nachzug): Gruppenname aus dem Sub-Event-Titel — die Versalien machten aus „Day 1" ein „DAY 1". */}
                          <span style={{ textTransform: 'none', letterSpacing: 0 }}>{label}</span>
                          <span style={{ textTransform: 'none', letterSpacing: 0 }}>{sum} {isDe ? (sum === 1 ? 'Abmeldung' : 'Abmeldungen') : (sum === 1 ? 'cancellation' : 'cancellations')}</span>
                        </div>
                        {members.map(s => dayRow(s, stripGroupPrefix(s.fullTitle, g.label) || s.label))}
                      </div>
                    );
                  })}
                {quietChildren > 0 && (
                  <div className="dex-ui-muted" style={{ padding: '10px 4px 0' }}>
                    {isDe
                      ? `${quietChildren} ${isCal ? (quietChildren === 1 ? 'Tag' : 'Tage') : (quietChildren === 1 ? 'Termin' : 'Termine')} ohne Abmeldung.`
                      : `${quietChildren} ${isCal ? (quietChildren === 1 ? 'day' : 'days') : (quietChildren === 1 ? 'date' : 'dates')} without cancellations.`}
                  </div>
                )}
              </div>
            ) : (
              <div className="dex-ui-table-wrap dex-ui-table-wrap--sticky">
                <table className="dex-ui-table dex-ui-table--compact">
                  <thead>
                    {/* v31.3: Aus Vorname/Nachname/Job Title/Standort wird EINE
                        Personen-Zelle (Leitfaden 5b) — fünf Spalten weniger,
                        keine Angabe weniger. Die E-Mail bleibt eigene Spalte:
                        sie ist der Schlüssel, über den diese Übersicht baut. */}
                    <tr>
                      <th className={cx('is-sortable', (cancelledSortColumn === 'nachname' || cancelledSortColumn === 'vorname') && 'is-sorted')}
                        onClick={() => toggleSort('nachname')} title={isDe ? 'Nach Nachname sortieren' : 'Sort by last name'}>
                        Person{sortMark('nachname')}
                        {miniSortBtn('vorname', isDe ? 'Vorname' : 'First name', isDe ? 'Nach Vorname sortieren' : 'Sort by first name')}
                      </th>
                      {thSort('email', 'E-Mail')}
                      <th>{isDe ? 'Abgemeldet bei' : 'Cancelled from'}</th>
                      {thSort('date', isDe ? 'Letzte Abmeldung' : 'Last cancellation')}
                      {canDelete && (
                        <th style={{ textAlign: 'right' }}>{isDe ? 'Löschen' : 'Delete'}</th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {people.map(p => {
                      const nm = `${p.firstName} ${p.lastName}`.trim() || p.email;
                      const sub = [p.jobTitle, stripLocPrefix(p.location || '')].filter(Boolean).join(' • ');
                      return (
                        <tr key={p.email} className="is-muted">
                          <td>
                            <div className="dex-ui-person">
                              <PersonContactHover email={p.email} name={nm} size={30} subline={sub} isDe={isDe} />
                              <div style={{ minWidth: 0 }}>
                                <div className="dex-ui-person-name">{nm}</div>
                                {/* v31.3 (Nachzug): Position/Standort hatten eigene Spalten — die Zweitzeile schneidet ab, voller Text im title. */}
                                {sub && <div className="dex-ui-person-sub" title={sub}>{sub}</div>}
                              </div>
                            </div>
                          </td>
                          <td>{p.email}</td>
                          <td>
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>{chipsFor(p)}</div>
                          </td>
                          <td style={{ whiteSpace: 'nowrap' }}>{p.latest ? formatDate(new Date(p.latest).toISOString()) : dash}</td>
                          {canDelete && (
                            <td className="is-actions">
                              <button
                                type="button"
                                className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                                title={isDe ? 'Diese Person überall endgültig löschen (nur im Entwurf möglich)' : 'Permanently delete this person everywhere (drafts only)'}
                                aria-label={isDe ? 'Person überall löschen' : 'Delete person everywhere'}
                                onClick={() => { deletePerson(p).catch(() => { /* */ }); }}
                              >
                                <Trash2 size={14} />
                              </button>
                            </td>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            ));
          }
          // v24.82: Abmeldungen im selben Zeilen-Layout wie die aktiven
          // Anmeldungen (Foto + Name + „Position • Standort • Firma"), aber
          // alle Texte in hellem Grau, damit sie klar von den aktiven
          // Anmeldungen zu unterscheiden sind. Ein „Löschen"-Button erscheint
          // NUR bei Entwurf-Events (canDelete) — sonst bleiben Abmeldungen
          // wegen der einjährigen Aufbewahrungsfrist erhalten.
          // v31.3: Das Grau kommt jetzt aus `is-muted` (Zeile wird beim
          // Überfahren wieder lesbar) statt aus einer Farbe je Zelle. Neue
          // Spaltenfolge nach der Frage, in der ein Organizer liest: wer →
          // was → wann → wer hat abgemeldet → was ist auffällig.
          return shell(cancelledRegs.length, declineCount, null, (
            /* v30.87: eigener Scroll-Container (70vh) wie bei den Anmeldungen —
               bei 48 Abmeldungen lief die Tabelle sonst über die ganze Seite.
               v31.3: das macht `dex-ui-table-wrap--sticky` samt stehendem Kopf. */
            <div className="dex-ui-table-wrap dex-ui-table-wrap--sticky">
              <table className="dex-ui-table dex-ui-table--compact">
                <thead>
                  <tr>
                    <th style={{ width: 36 }}>#</th>
                    {thSort('nachname', isDe ? 'Teilnehmer' : 'Attendee')}
                    {thSort('type', isDe ? 'Art' : 'Type')}
                    {thSort('date', isDe ? 'Abgemeldet am' : 'Cancelled on')}
                    <th>{isDe ? 'Abgemeldet von' : 'Cancelled by'}</th>
                    {isConsolidatedMode && (
                      <th>{isDe ? 'Sub-Event' : 'Sub-event'}</th>
                    )}
                    <th>{isDe ? 'Nach Frist' : 'After deadline'}</th>
                    {/* v19.4: „Wurde ersetzt durch" — die nachgerückte Person, die
                        den frei gewordenen Platz übernommen hat (vom Flow gesetzt).
                        v19.11: nur bei Events mit echter Warteliste-/Nachrück-
                        Aktivität (sonst durchgehend leer). */}
                    {hasWaitlistActivity && (
                      <th>{isDe ? 'Wurde ersetzt durch' : 'Replaced by'}</th>
                    )}
                    {canDelete && (
                      <th style={{ textAlign: 'right' }}>{isDe ? 'Löschen' : 'Delete'}</th>
                    )}
                  </tr>
                </thead>
                <tbody>
                  {sorted.map((reg, rowIdx) => {
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
                      <tr key={reg.Id} className="is-muted">
                        {/* v30.87: laufende Nummer + Personen-Zelle exakt wie in der
                            Anmeldungen-Tabelle (Name fett, Zweitzeile grau) — die
                            durchgehend ausgegraute Zeile las sich wie deaktiviert. */}
                        <td style={{ color: 'var(--dex-gray-400)' }}>{rowIdx + 1}</td>
                        <td>
                          <div className="dex-ui-person">
                            <PersonContactHover email={reg.ParticipantEmail || ''} name={fullName} size={30} subline={sub} isDe={isDe} />
                            <div style={{ minWidth: 0 }}>
                              <div className="dex-ui-person-name">{fullName}</div>
                              {/* v31.3 (Nachzug): Zweitzeile schneidet ab — voller Text im title. */}
                              {sub && <div className="dex-ui-person-sub" title={sub}>{sub}</div>}
                            </div>
                          </div>
                        </td>
                        <td><span className={pillCls(declined)}>{artLabel}</span></td>
                        <td style={{ whiteSpace: 'nowrap' }}>{formatDate(reg.CancellationDate)}</td>
                        {/* v24.88: „Abgemeldet von" — selbst abgemeldet vs. durch
                            jemand anderen (Audit CancelledBy*), analog „Angemeldet von". */}
                        <td>
                          {(() => {
                            const cby = (anyReg.CancelledByEmail || '').toLowerCase();
                            const pe = (reg.ParticipantEmail || '').toLowerCase();
                            if (!cby) return dash;
                            if (cby === pe) return <span>{isDe ? 'Selbst abgemeldet' : 'Cancelled themselves'}</span>;
                            return <span title={anyReg.CancelledByEmail}>{anyReg.CancelledByName || anyReg.CancelledByEmail}</span>;
                          })()}
                        </td>
                        {isConsolidatedMode && (
                          <td>{(reg as SPRegistration & { _sectionTitle?: string })._sectionTitle || '-'}</td>
                        )}
                        {/* v24.88: Markierung, wenn die Abmeldung NACH der
                            kommunizierten Abmeldefrist (lastDeregisterDate) erfolgte. */}
                        <td>
                          {(() => {
                            if (declined) return dash;
                            let ev = selectedEvent;
                            if (isConsolidatedMode) {
                              const sid = (reg as SPRegistration & { _sectionId?: string })._sectionId;
                              if (sid && sid !== '__parent') { const ch = consolidatedChildren.find(c => c.id === sid); if (ch) ev = ch; }
                            }
                            const dlRaw = ev?.lastDeregisterDate;
                            if (!dlRaw || !reg.CancellationDate) return dash;
                            const isLate = new Date(reg.CancellationDate).getTime() > new Date(dlRaw).getTime();
                            return isLate
                              ? <span className="dex-ui-pill dex-ui-pill--orange" title={`${isDe ? 'Abmeldefrist war' : 'Deadline was'}: ${formatDate(dlRaw)}`}>{isDe ? 'Nach Frist' : 'After deadline'}</span>
                              : dash;
                          })()}
                        </td>
                        {hasWaitlistActivity && (
                          <td>
                            {(() => {
                              const email = (anyReg.ReplacedByParticipantEmail as string | undefined) || '';
                              if (!email) return dash;
                              const other = registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === email.toLowerCase());
                              const label = other ? (((other.Vorname || '') + ' ' + (other.Nachname || '')).trim() || other.ParticipantName || email) : email;
                              return <span title={email}>{label}</span>;
                            })()}
                          </td>
                        )}
                        {canDelete && (
                          <td className="is-actions">
                            {/* v31.3: Symbol-Knopf statt Rahmen-Knopf — er erscheint
                                beim Überfahren der Zeile kräftiger (is-actions) und
                                steht damit dort, wo Löschen hingehört: rechts außen. */}
                            <button
                              type="button"
                              className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                              title={isDe ? 'Registrierung endgültig löschen (nur im Entwurf möglich)' : 'Permanently delete registration (drafts only)'}
                              aria-label={isDe ? 'Registrierung endgültig löschen' : 'Permanently delete registration'}
                              onClick={() => { deleteCancelled(reg).catch(() => { /* */ }); }}
                            >
                              <Trash2 size={14} />
                            </button>
                          </td>
                        )}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          ));
};

