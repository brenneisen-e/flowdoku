/* OverbookReviewBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10810-11029 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { SPRegistration } from '../../../services/EventService';
import { InfoTooltip } from '../../InfoTooltip';
import { formatDate } from '../../../utils/eventStatus';
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
          return (
            <div style={{ marginBottom: 20, padding: 16, borderRadius: 12, border: '1px solid var(--dex-orange, #ed8b00)', background: 'rgba(237,139,0,0.07)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 12 }}>
                <strong style={{ color: 'var(--dex-orange-dark, #b35a00)', fontSize: '0.95rem' }}>
                  {isDe ? `Überbuchung – zu prüfen (${flagged.length})` : `Overbooking – to review (${flagged.length})`}
                </strong>
                <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Über Kapazität angemeldet. Pro Person entscheiden — danach werden IDs automatisch neu vergeben.'
                    : 'Registered over capacity. Decide per person — afterwards the IDs are reassigned automatically.'}
                </span>
                <div style={{ marginLeft: 'auto', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  <button
                    className="btn btn-secondary"
                    style={{ fontSize: '0.78rem', padding: '5px 12px', color: 'var(--dex-red, #c00)' }}
                    onClick={() => { setOverbookModal({ mode: 'confirm', targets: flagged }); setObWithMail(true); setObRemoveCalendar(true); }}
                  >
                    {isDe ? `Alle bestätigen (${flagged.length})` : `Confirm all (${flagged.length})`}
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
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '1px solid rgba(237,139,0,0.4)', textAlign: 'left', color: 'var(--dex-gray-600)' }}>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Aktuell' : 'Current'}</th>
                      <th style={{ padding: '4px 8px' }}>Name</th>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Gruppe' : 'Group'}</th>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Angemeldet' : 'Registered'}</th>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Über Kapazität' : 'Over capacity'}</th>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Abstand zum letzten fairen Platz' : 'Gap to last fair seat'}</th>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Fairer Platz' : 'Fair seat'}</th>
                      <th style={{ padding: '4px 8px', textAlign: 'right' }}>{isDe ? 'Aktion' : 'Action'}</th>
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
                      return (
                        <tr key={reg.Id} style={{ borderBottom: '1px solid rgba(237,139,0,0.25)' }}>
                          <td style={{ padding: '6px 8px', fontWeight: 600 }}>#{reg.TeilnehmerID ?? '—'}</td>
                          <td style={{ padding: '6px 8px' }}>
                            {nm}
                            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{reg.ParticipantEmail}</div>
                          </td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>{grpLabel}</td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>
                            {position !== null && cap > 0
                              ? (isDe
                                ? <>Platz <strong>{position}</strong> bei Kap. {cap} <span style={{ color: 'var(--dex-red, #c00)' }}>(+{overBy})</span></>
                                : <>Seat <strong>{position}</strong> at cap. {cap} <span style={{ color: 'var(--dex-red, #c00)' }}>(+{overBy})</span></>)
                              : '—'}
                          </td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>
                            {cutoff
                              ? <><strong>+{fmtGap(gapMs)}</strong><div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{isDe ? 'nach' : 'after'} {cutoffNm} ({formatDate(cutoff.RegistrationDate)})</div></>
                              : '—'}
                          </td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>
                            {wlRank > 0
                              ? (isDe
                                ? <>Warteliste-Platz <strong>{wlRank}</strong>{isSplitCapacity ? ` (${grpLabel})` : ''}<div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>= TeilnehmerID ~#{fairId} bei sauberer Liste</div></>
                                : <>Waitlist position <strong>{wlRank}</strong>{isSplitCapacity ? ` (${grpLabel})` : ''}<div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>= participant ID ~#{fairId} with a clean list</div></>)
                              : '—'}
                          </td>
                          <td style={{ padding: '6px 8px', textAlign: 'right', whiteSpace: 'nowrap' }}>
                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, marginRight: 10 }}>
                              <button
                                className="btn btn-secondary"
                                style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #c00)' }}
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
                                className="btn btn-secondary"
                                style={{ fontSize: '0.75rem', padding: '4px 10px' }}
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

