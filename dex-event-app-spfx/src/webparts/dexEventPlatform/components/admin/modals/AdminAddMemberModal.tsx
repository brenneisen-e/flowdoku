/* AdminAddMemberModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16641-17204 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { SPRegistration } from '../../../services/EventService';
import Modal from '../../Modal';
import InternationalSearchToggle from '../../InternationalSearchToggle';
import { DeloitteEvent } from '../../../types';

export interface AdminAddMemberModalProps {
  addTeamMember: (eventId: string, teamId: string, teamName: string, member: { email: string; displayName: string; }, customData?: Record<string, string>, opts?: { suppressMemberMail?: boolean; suppressOthersMail?: boolean; ccEmail?: string; }) => Promise<{ ok: boolean; status?: "Angemeldet" | "Warteliste"; reason?: string; }>;
  adminAddCcOrganizer: boolean;
  adminAddLeadRegId: number;
  adminAddMemberBusy: boolean;
  adminAddMemberConsent: boolean;
  adminAddMemberDialog: { teamId: string; teamName: string; freeSlots: number; isNewTeam?: boolean; };
  adminAddMemberError: string;
  adminAddMemberIncludeIntl: boolean;
  adminAddMemberPick: { email: string; displayName: string; };
  adminAddMemberQuery: string;
  adminAddMemberQueryTimer: React.MutableRefObject<NodeJS.Timeout>;
  adminAddMemberResults: { email: string; displayName: string; }[];
  adminAddMemberSearching: boolean;
  adminAddNewPersonMail: boolean;
  adminAddNotifyOthers: boolean;
  adminAddNotifyScope: "all" | "lead";
  adminAddSendMail: boolean;
  adminAddTeamlessPicks: Set<number>;
  assignTeamlessToTeam: (eventId: string, teamId: string, teamName: string, existingRegId: number, isLead?: boolean, opts?: { sendMail?: boolean; recipientEmail?: string; recipientFirstName?: string; recipientLastName?: string; ccEmail?: string; }) => Promise<boolean>;
  currentUser: import("../../../types/index").User;
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
  isDe: boolean;
  notifyExistingTeamMembers: (eventId: string, teamId: string, teamName: string, newMemberNames: string[], excludeEmails: string[], scope?: "all" | "lead") => Promise<void>;
  registrations: SPRegistration[];
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  selectedEvent: DeloitteEvent;
  setAdminAddCcOrganizer: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddLeadRegId: React.Dispatch<React.SetStateAction<number>>;
  setAdminAddMemberBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddMemberConsent: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddMemberDialog: React.Dispatch<React.SetStateAction<{ teamId: string; teamName: string; freeSlots: number; isNewTeam?: boolean; }>>;
  setAdminAddMemberError: React.Dispatch<React.SetStateAction<string>>;
  setAdminAddMemberIncludeIntl: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddMemberPick: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; }>>;
  setAdminAddMemberQuery: React.Dispatch<React.SetStateAction<string>>;
  setAdminAddMemberResults: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; }[]>>;
  setAdminAddMemberSearching: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddNewPersonMail: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddNotifyOthers: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddNotifyScope: React.Dispatch<React.SetStateAction<"all" | "lead">>;
  setAdminAddSendMail: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddTeamlessPicks: React.Dispatch<React.SetStateAction<Set<number>>>;
  setRegistrations: React.Dispatch<React.SetStateAction<SPRegistration[]>>;
  setTeamsToast: React.Dispatch<React.SetStateAction<string>>;
}

export const AdminAddMemberModal: React.FC<AdminAddMemberModalProps> = (p) => {
  const { addTeamMember, adminAddCcOrganizer, adminAddLeadRegId, adminAddMemberBusy, adminAddMemberConsent, adminAddMemberDialog, adminAddMemberError, adminAddMemberIncludeIntl, adminAddMemberPick, adminAddMemberQuery, adminAddMemberQueryTimer, adminAddMemberResults, adminAddMemberSearching, adminAddNewPersonMail, adminAddNotifyOthers, adminAddNotifyScope, adminAddSendMail, adminAddTeamlessPicks, assignTeamlessToTeam, currentUser, getAllRegistrations, isDe, notifyExistingTeamMembers, registrations, searchUsers, selectedEvent, setAdminAddCcOrganizer, setAdminAddLeadRegId, setAdminAddMemberBusy, setAdminAddMemberConsent, setAdminAddMemberDialog, setAdminAddMemberError, setAdminAddMemberIncludeIntl, setAdminAddMemberPick, setAdminAddMemberQuery, setAdminAddMemberResults, setAdminAddMemberSearching, setAdminAddNewPersonMail, setAdminAddNotifyOthers, setAdminAddNotifyScope, setAdminAddSendMail, setAdminAddTeamlessPicks, setRegistrations, setTeamsToast } = p;
        // v17.2: Quick-Pick aus bereits registrierten Personen ohne Team —
        // damit der Organizer nicht via Graph-Suche jeden neu picken muss,
        // wenn die Person ohnehin schon angemeldet ist.
        const teamlessActiveLocal = registrations.filter(r =>
          (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt')
          && !r.TeamId);
        const closeDlg = (): void => {
          setAdminAddMemberDialog(null);
          setAdminAddMemberPick(null);
          setAdminAddMemberQuery('');
          setAdminAddMemberResults([]);
          setAdminAddMemberConsent(false);
          setAdminAddMemberError('');
          setAdminAddMemberBusy(false);
          setAdminAddTeamlessPicks(new Set());
          setAdminAddLeadRegId(null);
          setAdminAddSendMail(false);
          setAdminAddCcOrganizer(false);
          setAdminAddNotifyOthers(false);
          setAdminAddNotifyScope('all');
          setAdminAddNewPersonMail(true);
        };
        // v17.4: Logik zur Auswertung der Multi-Pick + (optionalem) Graph-Pick.
        const hasMultiPicks = adminAddTeamlessPicks.size > 0;
        const hasGraphPick = !!adminAddMemberPick;
        // Wenn ausschliesslich teamlose Picks: keine Consent-Box (Person hat
        // bei der eigenen Anmeldung bereits zugestimmt). Sobald aber eine
        // NEUE Person via Graph-Suche dabei ist, bleibt die Consent-Pflicht.
        const onlyTeamlessPicks = hasMultiPicks && !hasGraphPick;
        const consentRequired = !onlyTeamlessPicks && hasGraphPick;
        // v22.40: Auswahl-/Kapazitäts-Zählung für Belegung-Anzeige,
        // Über-Kapazitäts-Sperre und Button-Aktivierung.
        const totalPicks = adminAddTeamlessPicks.size + (hasGraphPick ? 1 : 0);
        const freeSlots = adminAddMemberDialog.freeSlots;
        const atCap = freeSlots > 0 && totalPicks >= freeSlots;
        const overCap = freeSlots > 0 && totalPicks > freeSlots;
        const submit = async (): Promise<void> => {
          if (!adminAddMemberDialog || adminAddMemberBusy) return;
          if (!hasMultiPicks && !hasGraphPick) return;
          if (consentRequired && !adminAddMemberConsent) return;
          setAdminAddMemberBusy(true);
          setAdminAddMemberError('');
          try {
            const tid = adminAddMemberDialog.teamId;
            const tName = adminAddMemberDialog.teamName || undefined;
            let assignedCount = 0;
            // 1) Teamlose Picks zuordnen (PATCH only). v22.40: Wenn die
            // „Info-Mail"-Checkbox an ist, bekommt jede zugeordnete Person
            // die Mail direkt — die Empfänger-Daten kommen aus der bereits
            // geladenen Registrierungs-Zeile (kein erneutes Eingeben nötig).
            for (const regId of Array.from(adminAddTeamlessPicks)) {
              const isLead = adminAddLeadRegId === regId;
              const reg = teamlessActiveLocal.find(p => p.Id === regId);
              try {
                const ok = await assignTeamlessToTeam(selectedEvent.id, tid, tName, regId, isLead, {
                  sendMail: adminAddSendMail,
                  recipientEmail: reg?.ParticipantEmail,
                  recipientFirstName: reg?.Vorname,
                  recipientLastName: reg?.Nachname,
                  ccEmail: (adminAddSendMail && adminAddCcOrganizer) ? currentUser.email : undefined,
                });
                if (ok) assignedCount++;
              } catch (err) { console.warn('[DEX] assignTeamlessToTeam failed for', regId, err); }
            }
            // 2) Falls noch ein Graph-Pick dabei: addTeamMember (neuer Insert).
            // v22.49: Kommunikation an die neue Person optional
            // (adminAddNewPersonMail); die „übrige Mitglieder"-Info wird hier
            // unterdrückt und unten zentral (mit Reichweite alle/Lead) gesteuert.
            if (hasGraphPick && adminAddMemberPick) {
              const res = await addTeamMember(selectedEvent.id, tid, tName, adminAddMemberPick, undefined, {
                suppressMemberMail: !adminAddNewPersonMail,
                suppressOthersMail: true,
                ccEmail: (adminAddNewPersonMail && adminAddCcOrganizer) ? currentUser.email : undefined,
              });
              if (!res.ok) {
                if (res.reason && res.reason.startsWith('already-registered')) {
                  setAdminAddMemberError('Person bereits beim Event angemeldet — Picker aus „Bereits angemeldet"-Liste benutzen.');
                } else if (res.reason === 'team-full') {
                  setAdminAddMemberError('Das Team ist bereits voll.');
                } else {
                  setAdminAddMemberError('Hinzufügen fehlgeschlagen.');
                }
                setAdminAddMemberBusy(false);
                return;
              }
              assignedCount++;
            }
            // v22.49: „Neues Mitglied"-Info an die übrigen Team-Mitglieder
            // (Reichweite alle / nur Lead), sofern gewählt. excludeEmails =
            // die gerade neu hinzugefügten Personen (nicht sich selbst melden).
            if (adminAddNotifyOthers) {
              const assignedRegs = Array.from(adminAddTeamlessPicks)
                .map(id => teamlessActiveLocal.find(p => p.Id === id))
                .filter(Boolean) as SPRegistration[];
              const newNames = assignedRegs.map(r => `${r.Vorname || ''} ${r.Nachname || ''}`.trim() || r.ParticipantName || r.ParticipantEmail);
              const excludeEmails = assignedRegs.map(r => r.ParticipantEmail || '').filter(Boolean);
              if (hasGraphPick && adminAddMemberPick) {
                newNames.push(adminAddMemberPick.displayName || adminAddMemberPick.email);
                excludeEmails.push(adminAddMemberPick.email);
              }
              if (newNames.length > 0) {
                try { await notifyExistingTeamMembers(selectedEvent.id, tid, tName, newNames, excludeEmails, adminAddNotifyScope); }
                catch (err) { console.warn('[DEX] notifyExistingTeamMembers failed:', err); }
              }
            }
            const teamLabel = tName ? `„${tName}"` : 'das Team';
            const toastMsg = adminAddSendMail
              ? `${assignedCount} ${assignedCount === 1 ? 'Person' : 'Personen'} ${teamLabel} zugeordnet — Info-Mail wird versendet.`
              : `${assignedCount} ${assignedCount === 1 ? 'Person' : 'Personen'} ${teamLabel} zugeordnet (ohne Mail-Versand).`;
            setTeamsToast(toastMsg);
            // TODO v17.5: Wenn adminAddSendMail=true UND assignTeamlessToTeam-
            // Pfad genutzt wurde, hier explizit eine „Du bist jetzt im Team
            // <Name>"-Mail queuen. Aktuell läuft die Mail nur über den
            // addTeamMember-Pfad (Graph-Pick) automatisch.
            window.setTimeout(() => setTeamsToast(''), 4500);
            const regs = await getAllRegistrations(selectedEvent.id);
            setRegistrations(regs);
            closeDlg();
          } catch {
            setAdminAddMemberError('Hinzufügen fehlgeschlagen.');
            setAdminAddMemberBusy(false);
          }
        };
        return (
          <Modal
            open={true}
            onClose={closeDlg}
            dismissable={!adminAddMemberBusy}
            maxWidth={540}
            ariaLabel="Person zum Team hinzufügen"
          >
              <h3 style={{ margin: 0, fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
                {adminAddMemberDialog.isNewTeam
                  ? 'Neues Team anlegen — Mitglieder zuordnen'
                  : adminAddMemberDialog.teamName
                    ? `Mitglieder zum Team „${adminAddMemberDialog.teamName}" hinzufügen`
                    : 'Mitglieder zum Team hinzufügen'}
              </h3>
              <div style={{ fontSize: '0.85rem', color: overCap ? 'var(--dex-red, #c00)' : 'var(--dex-gray-600)' }}>
                {/* v22.40: Belegung berücksichtigt die aktuelle Auswahl live
                    (bisherige Belegung + ausgewählte Personen). */}
                {(selectedEvent.teamSize || 0) > 0
                  ? `Team-Belegung: ${((selectedEvent.teamSize || 0) - freeSlots) + totalPicks}/${selectedEvent.teamSize}${totalPicks > 0 ? ' (inkl. Auswahl)' : ''}${overCap ? ' — zu viele ausgewählt!' : ''}`
                  : 'Belegung wird nach dem Hinzufügen aktualisiert.'}
              </div>
              {/* v17.1: Team-Name-Eingabe nur im „Neues Team anlegen"-Flow.
                  Optional — wenn leer, bekommt das Team beim Insert keinen
                  Namen, der Lead kann ihn aber später nicht mehr setzen,
                  daher direkt hier abfragen. */}
              {adminAddMemberDialog.isNewTeam && (
                <div style={{ marginTop: 4 }}>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                    Team-Name {selectedEvent.askTeamName ? <span style={{ color: 'var(--dex-red, #c00)' }}>*</span> : <span style={{ color: 'var(--dex-gray-400)', fontWeight: 400 }}>(optional)</span>}
                  </label>
                  <input
                    type="text"
                    className="form-input"
                    placeholder="z.B. „Borntowin"
                    value={adminAddMemberDialog.teamName}
                    onChange={e => setAdminAddMemberDialog(d => d ? { ...d, teamName: e.target.value } : d)}
                    style={{ width: '100%' }}
                  />
                </div>
              )}
              {/* v17.4: Consent-Box nur, wenn eine wirklich NEUE Person
                  via Graph hinzugefügt wird. Bei reiner Team-Zuordnung
                  schon-angemeldeter Personen brauchen wir keine zusätzliche
                  Zustimmung — die haben sie bei der eigenen Anmeldung
                  bereits gegeben. */}
              {consentRequired ? (
                <div style={{
                  padding: '14px 16px',
                  background: 'rgba(237,139,0,0.10)',
                  border: '2px solid var(--dex-orange, #ed8b00)',
                  borderRadius: 8,
                  color: '#7a4a00',
                  fontSize: '0.88rem',
                  lineHeight: 1.5,
                }}>
                  <div style={{ fontWeight: 700, marginBottom: 6 }}>
                    Vorab die Zustimmung des Mitglieds einholen
                  </div>
                  <div>
                    {'Mit dem Hinzufügen meldest du diese Person an. Sie erhält automatisch '}
                    {'eine Anmeldebestätigung per Mail, einen Outlook-Termin und sieht das '}
                    {'Event in „Meine Events". Bitte stelle sicher, dass die Person ihrer '}
                    {'Anmeldung '}<strong>vorher zugestimmt</strong>{' hat.'}
                  </div>
                </div>
              ) : (onlyTeamlessPicks && (
                <div style={{
                  padding: '10px 14px',
                  background: 'rgba(33,150,243,0.06)',
                  border: '1px solid var(--dex-info, #2196f3)',
                  borderRadius: 8,
                  color: 'var(--dex-gray-700)',
                  fontSize: '0.82rem',
                  lineHeight: 1.5,
                }}>
                  Du ordnest bereits-angemeldete Teilnehmer einem Team zu — keine neue Anmeldung, keine Bestätigungsmail an die Personen (es sei denn du hakst &bdquo;Info-Mail an die zugeordneten&hellip;&ldquo; unten an).
                </div>
              ))}
              <div>
                {/* v22.45: Drei klare Abschnitte — 1. Bestehende Teilnehmer
                    (bereits angemeldet, nur zuordnen), 2. Neue Teilnehmer (per
                    Suche stellvertretend anmelden), 3. Kommunikation ans Team. */}
                {teamlessActiveLocal.length > 0 && (
                  <div style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-green-dark, #4a7c1f)', marginBottom: 4 }}>
                    1 · Bestehende Teilnehmer
                  </div>
                )}
                {/* v17.4: Multi-Select aus bereits registrierten Personen
                    ohne Team. Checkbox-Liste; bei Mehrfach-Auswahl
                    erscheint zusätzlich die Lead-Radio-Auswahl. */}
                {teamlessActiveLocal.length > 0 && (
                  <div style={{ marginBottom: 12, padding: 10, border: '1px dashed var(--dex-orange, #ed8b00)', borderRadius: 6, background: 'rgba(237,139,0,0.04)' }}>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 600, marginBottom: 6 }}>
                      Bereits angemeldet ohne Team ({teamlessActiveLocal.length}) — mehrere auswählbar:
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4, maxHeight: 280, overflowY: 'auto' }}>
                      {teamlessActiveLocal.map(p => {
                        const nm = `${p.Vorname || ''} ${p.Nachname || ''}`.trim() || p.ParticipantName || p.ParticipantEmail;
                        const isPicked = adminAddTeamlessPicks.has(p.Id);
                        const isLead = adminAddLeadRegId === p.Id;
                        return (
                          <div
                            key={p.Id}
                            style={{
                              display: 'flex', alignItems: 'center', gap: 8,
                              padding: '6px 10px',
                              border: `1px solid ${isPicked ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-200)'}`,
                              borderRadius: 6,
                              background: isPicked ? 'rgba(237,139,0,0.08)' : '#fff',
                            }}
                          >
                            <input
                              type="checkbox"
                              checked={isPicked}
                              // v22.40: Über-Kapazitäts-Sperre — nicht mehr als
                              // freie Plätze auswählbar; bereits Gewählte bleiben
                              // abwählbar.
                              disabled={!isPicked && atCap}
                              onChange={e => {
                                setAdminAddTeamlessPicks(prev => {
                                  const next = new Set(prev);
                                  if (e.target.checked) next.add(p.Id);
                                  else next.delete(p.Id);
                                  return next;
                                });
                                // Wenn Lead deselektiert wurde: Lead zurücksetzen.
                                if (!e.target.checked && adminAddLeadRegId === p.Id) setAdminAddLeadRegId(null);
                              }}
                              style={{ flexShrink: 0, cursor: (!isPicked && atCap) ? 'not-allowed' : 'pointer' }}
                            />
                            <img
                              src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(p.ParticipantEmail)}&size=S`}
                              alt={nm}
                              onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                              style={{ width: 24, height: 24, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                            />
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontSize: '0.82rem', fontWeight: 500 }}>{nm}</div>
                              <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{p.ParticipantEmail}</div>
                            </div>
                            {isPicked && (
                              <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: '0.72rem', color: 'var(--dex-gray-700)', cursor: 'pointer' }}>
                                <input
                                  type="radio"
                                  name="lead-pick"
                                  checked={isLead}
                                  onChange={() => setAdminAddLeadRegId(p.Id)}
                                  style={{ margin: 0 }}
                                />
                                Lead
                              </label>
                            )}
                          </div>
                        );
                      })}
                    </div>
                    {adminAddTeamlessPicks.size > 0 && (
                      <div style={{ marginTop: 8, display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.78rem', color: 'var(--dex-gray-700)', flexWrap: 'wrap' }}>
                        <strong>{adminAddTeamlessPicks.size}</strong> ausgewählt
                        {!adminAddLeadRegId && adminAddTeamlessPicks.size > 0 && (
                          <span style={{ color: 'var(--dex-gray-500)' }}>
                            — bitte einen Lead markieren (oder leer = kein Lead).
                          </span>
                        )}
                        {atCap && (
                          <span style={{ color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 600 }}>
                            — Team voll ({freeSlots} {freeSlots === 1 ? 'Platz' : 'Plätze'}).
                          </span>
                        )}
                      </div>
                    )}
                    <div style={{ marginTop: 8, fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                      Oder weiter unten via Suche eine zusätzliche neue Person hinzufügen.
                    </div>
                  </div>
                )}
                {/* v22.45: 2 · Neue Teilnehmer — Person, die noch NICHT beim
                    Event angemeldet ist, per Suche stellvertretend hinzufügen.
                    (Kommunikation/Info-Mail folgt als Abschnitt 3 weiter unten.) */}
                <div style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-green-dark, #4a7c1f)', marginTop: teamlessActiveLocal.length > 0 ? 14 : 0, marginBottom: 2 }}>
                  {teamlessActiveLocal.length > 0 ? '2 · Neue Teilnehmer' : 'Neue Teilnehmer'}
                </div>
                <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginBottom: 6 }}>
                  Jemand, der noch nicht beim Event angemeldet ist — per Suche hinzufügen (wird stellvertretend angemeldet).
                </div>
                {adminAddMemberPick ? (
                  <div style={{
                    display: 'inline-flex', alignItems: 'center', gap: 10,
                    padding: '6px 10px 6px 6px',
                    border: '1px solid var(--dex-gray-200)',
                    borderRadius: 'var(--dex-radius)',
                    background: 'var(--dex-gray-50, #f7f7f7)',
                    maxWidth: '100%',
                  }}>
                    <img
                      src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(adminAddMemberPick.email)}&size=S`}
                      alt={adminAddMemberPick.displayName}
                      onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                      style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                    />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: '0.9rem' }}>{adminAddMemberPick.displayName}</div>
                      <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem' }}>{adminAddMemberPick.email}</div>
                    </div>
                    <button
                      type="button"
                      onClick={() => { setAdminAddMemberPick(null); setAdminAddMemberQuery(''); setAdminAddMemberResults([]); }}
                      title={isDe ? 'Auswahl entfernen' : 'Remove selection'}
                      style={{
                        background: 'var(--dex-gray-200)', border: 'none', color: 'var(--dex-gray-700)',
                        width: 22, height: 22, borderRadius: '50%', cursor: 'pointer',
                        fontSize: '0.9rem', lineHeight: 1,
                      }}
                    >×</button>
                  </div>
                ) : (
                  <div style={{ position: 'relative' }}>
                    <input
                      className="form-input"
                      value={adminAddMemberQuery}
                      // v22.40: Suche sperren, wenn das Team durch die Auswahl
                      // bereits voll ist (keine zusätzliche neue Person mehr).
                      disabled={atCap}
                      placeholder={atCap ? 'Team voll — keine weitere Person' : 'Name oder E-Mail eingeben…'}
                      onChange={e => {
                        const val = e.target.value;
                        setAdminAddMemberQuery(val);
                        if (adminAddMemberQueryTimer.current) clearTimeout(adminAddMemberQueryTimer.current);
                        if (val.length >= 2) {
                          adminAddMemberQueryTimer.current = setTimeout(async () => {
                            setAdminAddMemberSearching(true);
                            try {
                              const res = await searchUsers(val, adminAddMemberIncludeIntl);
                              setAdminAddMemberResults(res.map(r => ({ email: r.email, displayName: r.displayName })));
                            } catch { setAdminAddMemberResults([]); }
                            setAdminAddMemberSearching(false);
                          }, 300);
                        } else {
                          setAdminAddMemberResults([]);
                        }
                      }}
                    />
                    <InternationalSearchToggle
                      query={adminAddMemberQuery}
                      checked={adminAddMemberIncludeIntl}
                      onChange={setAdminAddMemberIncludeIntl}
                      isDe={isDe}
                    />
                    {(adminAddMemberResults.length > 0 || adminAddMemberSearching) && (
                      <div style={{
                        position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 10,
                        background: '#fff', border: '1px solid var(--dex-gray-200)',
                        borderRadius: 6, marginTop: 4, maxHeight: 220, overflowY: 'auto',
                        boxShadow: '0 4px 12px rgba(0,0,0,0.12)',
                      }}>
                        {adminAddMemberSearching && (
                          <div style={{ padding: 10, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
                            Suche…
                          </div>
                        )}
                        {adminAddMemberResults.map(r => (
                          <button
                            key={r.email}
                            type="button"
                            onClick={() => { setAdminAddMemberPick(r); setAdminAddMemberResults([]); setAdminAddMemberQuery(''); }}
                            style={{
                              display: 'flex', alignItems: 'center', gap: 8,
                              width: '100%', padding: '6px 10px', border: 'none',
                              background: '#fff', cursor: 'pointer', textAlign: 'left',
                            }}
                          >
                            <img
                              src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(r.email)}&size=S`}
                              alt={r.displayName}
                              onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                              style={{ width: 28, height: 28, borderRadius: '50%' }}
                            />
                            <div>
                              <div style={{ fontSize: '0.85rem', fontWeight: 600 }}>{r.displayName}</div>
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{r.email}</div>
                            </div>
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                )}
                {/* v22.45/v22.49: 3 · Kommunikation an das Team. */}
                {(adminAddTeamlessPicks.size > 0 || hasGraphPick) && (
                  <>
                    <div style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-green-dark, #4a7c1f)', marginTop: 14, marginBottom: 4 }}>
                      3 · Kommunikation an das Team
                    </div>
                    {/* a) Neue Person (Graph-Pick): Anmeldebestätigung + Outlook
                        optional (Default an — echte Neu-Anmeldung). */}
                    {hasGraphPick && (
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, fontSize: '0.82rem', cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={adminAddNewPersonMail}
                          onChange={e => setAdminAddNewPersonMail(e.target.checked)}
                          style={{ marginTop: 2 }}
                        />
                        <span>
                          Anmeldebestätigung &amp; Kalendereinladung an die neue Person senden
                          <span style={{ display: 'block', fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                            Default an — die Person wird ja neu angemeldet. Abwählen = still hinzufügen.
                          </span>
                        </span>
                      </label>
                    )}
                    {/* b) Info-Mail an die zugeordneten (bereits angemeldeten) Personen. */}
                    {adminAddTeamlessPicks.size > 0 && (
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, fontSize: '0.82rem', cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={adminAddSendMail}
                          onChange={e => setAdminAddSendMail(e.target.checked)}
                          style={{ marginTop: 2 }}
                        />
                        <span>
                          Info-Mail an die zugeordneten Team-Mitglieder versenden
                          <span style={{ display: 'block', fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                            Default aus — die Person ist ja bereits beim Event angemeldet.
                          </span>
                        </span>
                      </label>
                    )}
                    {(adminAddSendMail || (hasGraphPick && adminAddNewPersonMail)) && (
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, marginLeft: 24, fontSize: '0.82rem', cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={adminAddCcOrganizer}
                          onChange={e => setAdminAddCcOrganizer(e.target.checked)}
                          style={{ marginTop: 2 }}
                        />
                        <span>
                          Bestätigungsmail als Kopie (CC) an mich
                          <span style={{ display: 'block', fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                            {currentUser.email} bekommt diese Mail(s) in Kopie.
                          </span>
                        </span>
                      </label>
                    )}
                    {/* c) Übrige Team-Mitglieder informieren — Reichweite alle / nur Lead. */}
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 4, fontSize: '0.82rem', cursor: 'pointer' }}>
                      <input
                        type="checkbox"
                        checked={adminAddNotifyOthers}
                        onChange={e => setAdminAddNotifyOthers(e.target.checked)}
                        style={{ marginTop: 2 }}
                      />
                      <span>
                        Auch die übrigen Team-Mitglieder informieren
                        <span style={{ display: 'block', fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                          Schickt den bisherigen Mitgliedern eine „neues Mitglied“-Info.
                        </span>
                      </span>
                    </label>
                    {adminAddNotifyOthers && (
                      <div style={{ display: 'flex', gap: 16, marginLeft: 24, marginBottom: 8, fontSize: '0.82rem' }}>
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: 'pointer' }}>
                          <input type="radio" name="notifyScope" checked={adminAddNotifyScope === 'all'} onChange={() => setAdminAddNotifyScope('all')} style={{ margin: 0 }} />
                          Alle Mitglieder
                        </label>
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: 'pointer' }}>
                          <input type="radio" name="notifyScope" checked={adminAddNotifyScope === 'lead'} onChange={() => setAdminAddNotifyScope('lead')} style={{ margin: 0 }} />
                          Nur den Team-Lead
                        </label>
                      </div>
                    )}
                  </>
                )}
              </div>
              {/* v17.4: Consent-Checkbox nur bei wirklich neuer Person via Graph. */}
              {consentRequired && (
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
                  <input
                    type="checkbox"
                    checked={adminAddMemberConsent}
                    onChange={e => setAdminAddMemberConsent(e.target.checked)}
                    style={{ marginTop: 3 }}
                  />
                  <span>
                    <span style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                    Ich bestätige, dass die Person ihrer Anmeldung zugestimmt hat.
                  </span>
                </label>
              )}
              {adminAddMemberError && (
                <div style={{ padding: 10, borderRadius: 6, background: 'rgba(220,38,38,0.10)', color: '#b91c1c', fontSize: '0.85rem' }}>
                  {adminAddMemberError}
                </div>
              )}
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={closeDlg}
                  disabled={adminAddMemberBusy}
                >
                  Abbrechen
                </button>
                {(() => {
                  // v17.1: Bei „Neues Team anlegen" + askTeamName=true ist
                  // der Team-Name Pflicht (analog Self-Registration-Flow).
                  const needName = !!adminAddMemberDialog.isNewTeam && !!selectedEvent.askTeamName;
                  const nameOk = !needName || (adminAddMemberDialog.teamName.trim().length > 0);
                  // v22.40-Bugfix: Vorher verlangte `disabled` zwingend einen
                  // Graph-Pick + Consent — dadurch war der Button bei reiner
                  // Zuordnung bereits-angemeldeter Personen NIE klickbar. Jetzt:
                  // mindestens eine Auswahl (teamlos ODER Graph), Consent nur bei
                  // echtem Graph-Neu-Pick, nicht über Kapazität, Name ok.
                  const consentOk = !consentRequired || adminAddMemberConsent;
                  const disabled = totalPicks === 0 || !consentOk || adminAddMemberBusy || !nameOk || overCap;
                  const title = !nameOk ? 'Bitte einen Team-Namen eingeben.'
                    : overCap ? `Zu viele ausgewählt — nur noch ${freeSlots} Platz/Plätze frei.`
                    : totalPicks === 0 ? 'Bitte mindestens eine Person auswählen.'
                    : (!consentOk ? 'Bitte die Zustimmung bestätigen.' : '');
                  return (
                    <button
                      type="button"
                      className="btn btn-primary"
                      onClick={() => { submit().catch(() => { /* */ }); }}
                      disabled={disabled}
                      title={title}
                    >
                      {adminAddMemberBusy
                        ? 'Wird gespeichert…'
                        : (adminAddMemberDialog.isNewTeam
                          ? `Team anlegen${totalPicks > 0 ? ` (${totalPicks})` : ''}`
                          : `Hinzufügen${totalPicks > 0 ? ` (${totalPicks})` : ''}`)}
                    </button>
                  );
                })()}
              </div>
          </Modal>
        );
};

