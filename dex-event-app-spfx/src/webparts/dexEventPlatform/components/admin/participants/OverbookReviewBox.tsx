/* OverbookReviewBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10810-11029 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich übernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3 (docs/ui-leitfaden.md 5a Punkt 2, 5b): der Kasten ist ein
 * `dex-ui-callout--danger` (die Sache kostet jemandem den Platz), die Zahlen
 * stehen als Pillen, und der Sammel-Knopf steht unter seiner Erklärung statt
 * per `margin-left:auto` rechts außen (2a′). Die Rechenlogik ist unverändert.
 */
import * as React from 'react';
import { SPRegistration } from '../../../services/EventService';
import { InfoTooltip } from '../../InfoTooltip';
import { formatDate } from '../../../utils/eventStatus';
import { AlertCircle } from '../../Icons';
import { ensureDexUiStyles } from '../../dexUi';
import { DeloitteEvent } from '../../../types';

export interface OverbookReviewBoxProps {
  isDe: boolean;
  isSplitCapacity: boolean;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setObKeepVariant: React.Dispatch<React.SetStateAction<"active" | "firstWaitlist">>;
  setObRemoveCalendar: React.Dispatch<React.SetStateAction<boolean>>;
  setObWithMail: React.Dispatch<React.SetStateAction<boolean>>;
  setOverbookModal: React.Dispatch<React.SetStateAction<{ mode: "confirm" | "keep"; targets: SPRegistration[]; }>>;
}

export const OverbookReviewBox: React.FC<OverbookReviewBoxProps> = (p) => {
  const { isDe, isSplitCapacity, registrations, selectedEvent, setObKeepVariant, setObRemoveCalendar, setObWithMail, setOverbookModal } = p;
          // v31.3: Idempotent — der Kasten steht ohne Modal/WizardFormShell auf
          // der Seite, sonst fehlten die dex-ui-Klassen (inkl. aller :hover).
          ensureDexUiStyles();
          // v11.36: Überbuchungs-Review-Box. Zeigt alle per „Überbuchung
          // prüfen" markierten Personen (OverbookReview='Pending') mit
          // Fairness-Kontext + Aktions-Buttons. Erst durch eine Aktion
          // ändert sich der Status.
          const flaggedRaw = registrations.filter(r => r.OverbookReview === 'Pending');
          if (flaggedRaw.length === 0 || !selectedEvent) return null;
          const groupOf = (r: SPRegistration): string => r.StarterType || r.PreferredStarterType || '';
          const ACTIVE_ST = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
          // Gruppen-Key: bei Split die Gruppe, sonst ein gemeinsamer Topf.
          const keyOf = (r: SPRegistration): string => isSplitCapacity ? (groupOf(r) || '?') : 'all';
          const capOf = (key: string): number => {
            if (!isSplitCapacity) return selectedEvent.maxParticipants || 0;
            if (key === 'Durchstarter') return selectedEvent.durchstarterCapacity || 0;
            if (key === 'Funstarter') return selectedEvent.funstarterCapacity || 0;
            return 0;
          };
          // Pro Gruppe: aktive Anmeldungen in Anmeldereihenfolge (Id asc =
          // Reihenfolge der Registrierung — identisch zur Detect-Logik).
          const activeByGroup: Record<string, SPRegistration[]> = {};
          registrations
            .filter(r => ACTIVE_ST.indexOf(r.Status) >= 0)
            .slice()
            .sort((a, b) => a.Id - b.Id)
            .forEach(r => { const k = keyOf(r); (activeByGroup[k] = activeByGroup[k] || []).push(r); });
          // v22.40: Nur Personen anzeigen, die WIRKLICH noch über Kapazität
          // sind. Hat sich zwischenzeitlich jemand abgemeldet, passt eine
          // markierte Person ggf. wieder regulär in die Liste (oder ist selbst
          // nicht mehr aktiv) — solche stale Marker hier ausblenden (der
          // Auto-Heal-Effekt entfernt sie zusätzlich dauerhaft).
          const flagged = flaggedRaw.filter(r => {
            const k = keyOf(r); const cap = capOf(k); const bucket = activeByGroup[k] || [];
            const idx = bucket.findIndex(x => x.Id === r.Id);
            if (idx < 0) return false;
            return cap > 0 && (idx + 1) > cap;
          });
          if (flagged.length === 0) return null;
          // Faire Wartelisten-Reihenfolge je Gruppe: die über Kapazität
          // Aktiven + bereits vorhandene Warteliste, nach RegistrationDate.
          const fairWaitByGroup: Record<string, SPRegistration[]> = {};
          Object.keys(activeByGroup).forEach(k => {
            const cap = capOf(k);
            const overCap = cap > 0 ? activeByGroup[k].slice(cap) : [];
            const existingWl = registrations.filter(r => r.Status === 'Warteliste' && keyOf(r) === k);
            fairWaitByGroup[k] = [...overCap, ...existingWl].sort((a, b) =>
              new Date(a.RegistrationDate).getTime() - new Date(b.RegistrationDate).getTime());
          });
          // Faire Aktiv-Gesamtzahl (bei sauberer Liste) für die faire ID.
          let totalFairActive = 0;
          Object.keys(activeByGroup).forEach(k => {
            const cap = capOf(k);
            totalFairActive += cap > 0 ? Math.min(activeByGroup[k].length, cap) : activeByGroup[k].length;
          });
          const fmtGap = (ms: number): string => {
            if (!isFinite(ms) || ms < 0) return '—';
            const s = Math.round(ms / 1000);
            if (s < 90) return `${s} ${isDe ? 'Sek' : 'sec'}`;
            const m = Math.round(s / 60);
            if (m < 90) return `${m} ${isDe ? 'Min' : 'min'}`;
            const h = Math.round(m / 60);
            if (h < 48) return `${h} ${isDe ? 'Std' : 'h'}`;
            return `${Math.round(h / 24)} ${isDe ? 'Tage' : 'days'}`;
          };
          // v31.3: Zweitzeile in Zellen — einmal definiert statt viermal inline.
          const subLine: React.CSSProperties = { fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 };
          return (
            <div className="dex-ui-callout dex-ui-callout--danger" style={{ flexDirection: 'column', alignItems: 'stretch', gap: 12, marginBottom: 20 }}>
              <div style={{ display: 'flex', gap: 10, alignItems: 'flex-start' }}>
                <span className="dex-ui-callout-icon" aria-hidden="true"><AlertCircle size={16} /></span>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                    {isDe ? 'Überbuchung: mehr Anmeldungen als Plätze' : 'Overbooking: more registrations than seats'}
                    <span className="dex-ui-pill dex-ui-pill--red">
                      {isDe ? `${flagged.length} zu prüfen` : `${flagged.length} to review`}
                    </span>
                  </div>
                  <div style={{ marginTop: 2 }}>
                    {isDe
                      ? 'Diese Personen haben einen Platz über der Kapazität bekommen. Entscheide je Person — danach vergibt die App die TeilnehmerIDs automatisch neu.'
                      : 'These people got a seat beyond capacity. Decide per person — afterwards the app reassigns the participant IDs automatically.'}
                  </div>
                  {/* v31.3: Der Sammel-Knopf steht unter seiner Erklärung, nicht
                      rechts außen (Leitfaden 2a′) — und sagt jetzt, was er tut. */}
                  <div className="dex-ui-inline" style={{ marginTop: 8 }}>
                    <button
                      type="button"
                      className="btn btn-secondary dex-ui-btn-sm"
                      style={{ color: 'var(--dex-red, #c00)' }}
                      onClick={() => { setOverbookModal({ mode: 'confirm', targets: flagged }); setObWithMail(true); setObRemoveCalendar(true); }}
                    >
                      {isDe ? `Alle auf die Warteliste (${flagged.length})` : `All to waitlist (${flagged.length})`}
                    </button>
                    <InfoTooltip placement="left" text={isDe
                    ? (
                    <>
                      <strong>Sammel-Aktion:</strong> setzt <strong>alle</strong> markierten Personen auf die <strong>Warteliste</strong> (gruppentreu).<br /><br />
                      Die Optionen <strong>mit/ohne Mail</strong>, <strong>Kalender-Abmeldung</strong> und <strong>Sprache</strong> gelten <strong>für alle gleich</strong> — eine gemeinsame Entscheidung.<br /><br />
                      Der Mailtext ist trotzdem <strong>pro Person personalisiert</strong> (Name + individuelle neue Warteliste-Position).<br /><br />
                      Sollen einzelne Personen <strong>anders</strong> behandelt werden (z.B. &bdquo;Platz behalten&ldquo;), nutze stattdessen die <strong>Einzel-Buttons</strong> pro Zeile.
                    </>
                    ) : (
                    <>
                      <strong>Bulk action:</strong> moves <strong>all</strong> marked people to the <strong>waitlist</strong> (group-faithful).<br /><br />
                      The options <strong>with/without email</strong>, <strong>calendar removal</strong> and <strong>language</strong> apply <strong>to all alike</strong> — a single shared decision.<br /><br />
                      The email text is still <strong>personalized per person</strong> (name + individual new waitlist position).<br /><br />
                      If individual people should be treated <strong>differently</strong> (e.g. &bdquo;keep seat&ldquo;), use the <strong>per-row buttons</strong> instead.
                    </>
                    )
                    } />
                  </div>
                </div>
              </div>
              {/* v31.3: Weiße Tabelle im roten Kasten (Leitfaden 5b). Die Spalte
                  „Gruppe" erscheint nur bei geteilter Kapazität — ohne Gruppen
                  stand dort in JEDER Zeile nur „—". thead und tbody hängen an
                  derselben Bedingung, sonst verschiebt sich die Zeile um eine
                  Spalte (CLAUDE.md v29.2). */}
              <div className="dex-ui-table-wrap" style={{ background: '#fff' }}>
                <table className="dex-ui-table dex-ui-table--compact">
                  <thead>
                    <tr>
                      <th>{isDe ? 'Teilnehmer-ID' : 'Participant ID'}</th>
                      <th>Person</th>
                      {isSplitCapacity && <th>{isDe ? 'Gruppe' : 'Group'}</th>}
                      <th>{isDe ? 'Angemeldet am' : 'Registered'}</th>
                      <th>{isDe ? 'Über Kapazität' : 'Over capacity'}</th>
                      <th>{isDe ? 'Abstand zum fairen Platz' : 'Gap to fair seat'}</th>
                      <th>{isDe ? 'Fairer Platz' : 'Fair seat'}</th>
                      <th style={{ textAlign: 'right' }}>{isDe ? 'Aktion' : 'Action'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {flagged.map(reg => {
                      const nm = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                      const k = keyOf(reg);
                      const grpLabel = isSplitCapacity ? (groupOf(reg) || '—') : '—';
                      const cap = capOf(k);
                      const bucket = activeByGroup[k] || [];
                      const idx = bucket.findIndex(x => x.Id === reg.Id); // 0-basiert
                      const position = idx >= 0 ? idx + 1 : null;
                      const overBy = (position !== null && cap > 0) ? position - cap : null;
                      const cutoff = (cap > 0 && cap - 1 < bucket.length) ? bucket[cap - 1] : null;
                      const cutoffNm = cutoff ? ((cutoff.Vorname && cutoff.Nachname) ? `${cutoff.Vorname} ${cutoff.Nachname}` : cutoff.ParticipantName) : '';
                      const gapMs = cutoff ? (new Date(reg.RegistrationDate).getTime() - new Date(cutoff.RegistrationDate).getTime()) : NaN;
                      const wl = fairWaitByGroup[k] || [];
                      const wlRank = wl.findIndex(x => x.Id === reg.Id) + 1; // 1-basiert; 0 = nicht gefunden
                      const fairId = totalFairActive + (wlRank > 0 ? wlRank : (overBy || 0));
                      // v31.3: Initialen statt Foto — ein Foto-Abruf wäre hier
                      // neuer Netzverkehr, den es vorher nicht gab.
                      const initials = ((reg.Vorname || '').charAt(0) + (reg.Nachname || '').charAt(0)).toUpperCase()
                        || (nm || reg.ParticipantEmail || '?').charAt(0).toUpperCase();
                      return (
                        <tr key={reg.Id}>
                          <td><span className="dex-ui-pill dex-ui-pill--gray">#{reg.TeilnehmerID ?? '—'}</span></td>
                          <td>
                            <div className="dex-ui-person">
                              <span className="dex-ui-avatar" aria-hidden="true">{initials}</span>
                              <div style={{ minWidth: 0 }}>
                                <div className="dex-ui-person-name">{nm}</div>
                                <div className="dex-ui-person-sub">{reg.ParticipantEmail}</div>
                              </div>
                            </div>
                          </td>
                          {isSplitCapacity && (
                            <td>{groupOf(reg) ? <span className="dex-ui-pill dex-ui-pill--gray">{grpLabel}</span> : '—'}</td>
                          )}
                          <td style={{ color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                          <td>
                            {position !== null && cap > 0
                              ? <>
                                <span className="dex-ui-pill dex-ui-pill--red">+{overBy}</span>
                                <div style={subLine}>{isDe ? `Platz ${position} bei Kapazität ${cap}` : `Seat ${position} at capacity ${cap}`}</div>
                              </>
                              : '—'}
                          </td>
                          <td>
                            {cutoff
                              ? <><strong>+{fmtGap(gapMs)}</strong><div style={subLine}>{isDe ? 'nach' : 'after'} {cutoffNm} ({formatDate(cutoff.RegistrationDate)})</div></>
                              : '—'}
                          </td>
                          <td>
                            {wlRank > 0
                              ? <>
                                <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? `Warteliste-Platz ${wlRank}` : `Waitlist position ${wlRank}`}{isSplitCapacity ? ` · ${grpLabel}` : ''}</span>
                                <div style={subLine}>{isDe ? `= TeilnehmerID ~#${fairId} bei sauberer Liste` : `= participant ID ~#${fairId} with a clean list`}</div>
                              </>
                              : '—'}
                          </td>
                          <td className="is-actions">
                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, marginRight: 10 }}>
                              <button
                                type="button"
                                className="btn btn-secondary dex-ui-btn-sm"
                                style={{ color: 'var(--dex-red, #c00)' }}
                                onClick={() => { setOverbookModal({ mode: 'confirm', targets: [reg] }); setObWithMail(true); setObRemoveCalendar(true); }}
                              >
                                {isDe ? 'Auf Warteliste' : 'To waitlist'}
                              </button>
                              <InfoTooltip placement="left" text={isDe
                                ? (
                                <>
                                  <strong>&bdquo;Auf Warteliste&ldquo;</strong> — die Person wird (gruppentreu) auf die <strong>Warteliste</strong> gesetzt; sie hatte fälschlich einen Platz.<br /><br />
                                  Im nächsten Dialog wählst du: <strong>mit oder ohne Entschuldigungs-Mail</strong> (Deloitte-Layout, geht in die Mail-Queue — nicht direkt versendet) und ob sie <strong>vom Kalendereintrag abgemeldet</strong> wird.<br /><br />
                                  Es wird ein <strong>Audit-Eintrag</strong> geschrieben (war fälschlich angemeldet, Original-Registrierung). Danach werden die <strong>TeilnehmerIDs automatisch neu vergeben</strong>.
                                </>
                                ) : (
                                <>
                                  <strong>&bdquo;To waitlist&ldquo;</strong> — the person is moved (group-faithful) to the <strong>waitlist</strong>; they had a seat by mistake.<br /><br />
                                  In the next dialog you choose: <strong>with or without an apology email</strong> (Deloitte layout, goes into the mail queue — not sent directly) and whether they are <strong>removed from the calendar entry</strong>.<br /><br />
                                  An <strong>audit entry</strong> is written (was registered by mistake, original registration). Afterwards the <strong>participant IDs are reassigned automatically</strong>.
                                </>
                                )
                              } />
                            </span>
                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                              <button
                                type="button"
                                className="btn btn-secondary dex-ui-btn-sm"
                                onClick={() => { setOverbookModal({ mode: 'keep', targets: [reg] }); setObKeepVariant('firstWaitlist'); }}
                              >
                                {isDe ? 'Platz behalten' : 'Keep seat'}
                              </button>
                              <InfoTooltip placement="left" text={isDe
                                ? (
                                <>
                                  <strong>&bdquo;Platz behalten&ldquo;</strong> — die Person verliert den Platz <strong>nicht</strong>. Im nächsten Dialog wählst du:<br /><br />
                                  <strong>(a) Erste(r) auf der Warteliste</strong> der Gruppe — rückt beim nächsten frei werdenden Platz garantiert als Erste(r) nach.<br /><br />
                                  <strong>(b) Bleibt angemeldet</strong> — die Gruppe ist dann <strong>+1</strong> über Kapazität; der nächste frei werdende Platz wird <strong>einmal nicht</strong> nachgerückt, bis die Überzahl absorbiert ist.<br /><br />
                                  Beide Varianten mit <strong>Audit-Eintrag</strong>, danach IDs neu.
                                </>
                                ) : (
                                <>
                                  <strong>&bdquo;Keep seat&ldquo;</strong> — the person does <strong>not</strong> lose the seat. In the next dialog you choose:<br /><br />
                                  <strong>(a) First on the waitlist</strong> of the group — guaranteed to move up first when the next seat becomes free.<br /><br />
                                  <strong>(b) Stays registered</strong> — the group is then <strong>+1</strong> over capacity; the next freed seat is <strong>skipped once</strong> until the surplus is absorbed.<br /><br />
                                  Both variants with an <strong>audit entry</strong>, then IDs reassigned.
                                </>
                                )
                              } />
                            </span>
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

