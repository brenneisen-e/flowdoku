/* AdminAddMemberModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16641-17204 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { SPRegistration } from '../../../services/EventService';
import Modal from '../../Modal';
import InternationalSearchToggle from '../../InternationalSearchToggle';
import { DeloitteEvent } from '../../../types';
// v31.2: Gemeinsame Klassen statt Inline-Styles — das Modal hängt im Portal an
// document.body, dort greift nur dieses Stylesheet; und Inline-Styles können
// kein :hover (Leitfaden 1.3).
import { cx } from '../../dexUi';
import { AlertCircle, Check, Search, Users, X } from '../../Icons';

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
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
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
  setTeamsToast: React.Dispatch<React.SetStateAction<string>>;
}

export const AdminAddMemberModal: React.FC<AdminAddMemberModalProps> = (p) => {
  const { addTeamMember, adminAddCcOrganizer, adminAddLeadRegId, adminAddMemberBusy, adminAddMemberConsent, adminAddMemberDialog, adminAddMemberError, adminAddMemberIncludeIntl, adminAddMemberPick, adminAddMemberQuery, adminAddMemberQueryTimer, adminAddMemberResults, adminAddMemberSearching, adminAddNewPersonMail, adminAddNotifyOthers, adminAddNotifyScope, adminAddSendMail, adminAddTeamlessPicks, assignTeamlessToTeam, currentUser, isDe, notifyExistingTeamMembers, registrations, reloadRegistrations, searchUsers, selectedEvent, setAdminAddCcOrganizer, setAdminAddLeadRegId, setAdminAddMemberBusy, setAdminAddMemberConsent, setAdminAddMemberDialog, setAdminAddMemberError, setAdminAddMemberIncludeIntl, setAdminAddMemberPick, setAdminAddMemberQuery, setAdminAddMemberResults, setAdminAddMemberSearching, setAdminAddNewPersonMail, setAdminAddNotifyOthers, setAdminAddNotifyScope, setAdminAddSendMail, setAdminAddTeamlessPicks, setTeamsToast } = p;
        // v31.2: Der Dialog war rein deutsch; beide Sprachen laufen jetzt über
        // denselben Schalter wie der Rest des Organizer Centers.
        const t = (de: string, en: string): string => (isDe ? de : en);
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
                  // v31.2: Der Text nennt die Liste so, wie sie oben im Dialog heißt.
                  setAdminAddMemberError(t('Diese Person ist schon beim Event angemeldet — wähle sie oben unter „Bereits angemeldet, noch ohne Team“ aus.', 'This person is already registered for the event — pick them above under “Already registered, no team yet”.'));
                } else if (res.reason === 'team-full') {
                  setAdminAddMemberError(t('Das Team ist bereits voll.', 'The team is already full.'));
                } else {
                  setAdminAddMemberError(t('Hinzufügen fehlgeschlagen.', 'Adding failed.'));
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
            // v31.2: „dem Team zugeordnet" statt „das Team zugeordnet", und beide Sprachen.
            const teamLabel = tName ? `„${tName}“` : t('dem Team', 'the team');
            const who = `${assignedCount} ${assignedCount === 1 ? t('Person', 'person') : t('Personen', 'people')}`;
            const toastMsg = isDe
              ? `${who} ${teamLabel} zugeordnet${adminAddSendMail ? ' — Info-Mail wird versendet.' : ' (ohne Mail-Versand).'}`
              : `${who} assigned to ${teamLabel}${adminAddSendMail ? ' — info mail is being sent.' : ' (no mail sent).'}`;
            setTeamsToast(toastMsg);
            // TODO v17.5: Wenn adminAddSendMail=true UND assignTeamlessToTeam-
            // Pfad genutzt wurde, hier explizit eine „Du bist jetzt im Team
            // <Name>"-Mail queuen. Aktuell läuft die Mail nur über den
            // addTeamMember-Pfad (Graph-Pick) automatisch.
            window.setTimeout(() => setTeamsToast(''), 4500);
            // v30.67 (Review): gemeinsamer Nachlade-Pfad statt `[]` bei 429.
            await reloadRegistrations();
            closeDlg();
          } catch {
            setAdminAddMemberError(t('Hinzufügen fehlgeschlagen.', 'Adding failed.'));
            setAdminAddMemberBusy(false);
          }
        };
        // v17.1: Bei „Neues Team anlegen" + askTeamName=true ist
        // der Team-Name Pflicht (analog Self-Registration-Flow).
        const needName = !!adminAddMemberDialog.isNewTeam && !!selectedEvent.askTeamName;
        const nameOk = !needName || (adminAddMemberDialog.teamName.trim().length > 0);
        // v22.40-Bugfix: Vorher verlangte `disabled` zwingend einen
        // Graph-Pick + Consent — dadurch war der Button bei reiner
        // Zuordnung bereits-angemeldeter Personen NIE klickbar. Jetzt:
        // mindestens eine Auswahl (teamlos ODER Graph), Consent nur bei
        // echtem Graph-Neu-Pick, nicht über Kapazität, Name ok.
        // v31.2: aus der IIFE im Fuß nach oben gezogen — der Fuß ist jetzt eine
        // Modal-Prop, und der Hinweis am Knopf nennt weiter den ersten Grund.
        const consentOk = !consentRequired || adminAddMemberConsent;
        const disabled = totalPicks === 0 || !consentOk || adminAddMemberBusy || !nameOk || overCap;
        const submitTitle = !nameOk ? t('Bitte einen Team-Namen eingeben.', 'Please enter a team name.')
          : overCap ? t(`Zu viele ausgewählt — nur noch ${freeSlots} Platz/Plätze frei.`, `Too many selected — only ${freeSlots} seat(s) left.`)
          : totalPicks === 0 ? t('Bitte mindestens eine Person auswählen.', 'Please pick at least one person.')
          : (!consentOk ? t('Bitte die Zustimmung bestätigen.', 'Please confirm consent.') : '');
        // v22.40: Belegung berücksichtigt die aktuelle Auswahl live
        // (bisherige Belegung + ausgewählte Personen).
        const teamSize = selectedEvent.teamSize || 0;
        const occupancy = teamSize > 0
          ? `${t('Belegung', 'Seats')} ${(teamSize - freeSlots) + totalPicks}/${teamSize}${totalPicks > 0 ? t(' inkl. Auswahl', ' incl. selection') : ''}${overCap ? t(' — zu viele ausgewählt', ' — too many selected') : ''}`
          : t('Belegung wird nach dem Hinzufügen aktualisiert', 'Seat count updates after adding');
        const photoOf = (email: string): string => `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=S`;
        const hidePhoto = (e: React.SyntheticEvent<HTMLImageElement>): void => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; };
        return (
          <Modal
            open={true}
            onClose={closeDlg}
            dismissable={!adminAddMemberBusy}
            maxWidth={560}
            ariaLabel="Person zum Team hinzufügen"
            icon={<Users size={20} />}
            title={adminAddMemberDialog.isNewTeam
              ? t('Neues Team anlegen', 'Create a new team')
              : adminAddMemberDialog.teamName
                ? t(`Mitglieder für „${adminAddMemberDialog.teamName}“`, `Members for “${adminAddMemberDialog.teamName}”`)
                : t('Mitglieder zum Team hinzufügen', 'Add members to the team')}
            subtitle={t(
              'Bereits Angemeldete ordnest du nur zu. Wer noch nicht angemeldet ist, wird von dir stellvertretend angemeldet.',
              'People who are already registered are only assigned. Anyone not yet registered is registered by you on their behalf.')}
            footer={<>
              <button type="button" className="btn btn-secondary" onClick={closeDlg} disabled={adminAddMemberBusy}>
                {t('Abbrechen', 'Cancel')}
              </button>
              <button type="button" className="btn btn-primary" onClick={() => { submit().catch(() => { /* */ }); }} disabled={disabled} title={submitTitle}>
                {adminAddMemberBusy
                  ? t('Wird gespeichert…', 'Saving…')
                  : (adminAddMemberDialog.isNewTeam
                    ? `${t('Team anlegen', 'Create team')}${totalPicks > 0 ? ` (${totalPicks})` : ''}`
                    : `${t('Hinzufügen', 'Add')}${totalPicks > 0 ? ` (${totalPicks})` : ''}`)}
              </button>
            </>}
          >
            <div className="dex-ui-modal-body">
              {/* v17.1: Team-Name-Eingabe nur im „Neues Team anlegen"-Flow.
                  Optional — wenn leer, bekommt das Team beim Insert keinen
                  Namen, der Lead kann ihn aber später nicht mehr setzen,
                  daher direkt hier abfragen. v31.2: Pflicht zuerst — deshalb
                  ganz oben, als Frage statt als Feldname. */}
              {adminAddMemberDialog.isNewTeam && (
                <div className="dex-ui-field">
                  <label className="dex-ui-label" htmlFor="dex-addmember-teamname">
                    {t('Wie soll das Team heißen?', 'What should the team be called?')}
                    {selectedEvent.askTeamName
                      ? <span style={{ color: 'var(--dex-red, #c00)' }}>*</span>
                      : <span className="dex-ui-label-optional">{t('(optional)', '(optional)')}</span>}
                  </label>
                  <input
                    id="dex-addmember-teamname"
                    type="text"
                    className="dex-ui-input"
                    placeholder={t('z.B. „Borntowin“', 'e.g. “Borntowin”')}
                    value={adminAddMemberDialog.teamName}
                    onChange={e => setAdminAddMemberDialog(d => d ? { ...d, teamName: e.target.value } : d)}
                  />
                  <div className="dex-ui-help">
                    {t('Der Lead kann den Namen später nicht mehr selbst nachtragen — deshalb gleich hier.', 'The lead cannot add the name later — so set it right here.')}
                  </div>
                </div>
              )}
              <div>
                <div className="dex-ui-section-title">
                  {t('Wer kommt ins Team?', 'Who joins the team?')}
                  <span className={cx('dex-ui-pill', overCap ? 'dex-ui-pill--red' : atCap ? 'dex-ui-pill--orange' : teamSize > 0 ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')} style={{ textTransform: 'none', letterSpacing: 0 }}>
                    {occupancy}
                  </span>
                </div>
                {/* v17.4: Multi-Select aus bereits registrierten Personen
                    ohne Team. Checkbox-Liste; bei Auswahl erscheint zusätzlich
                    der Lead-Chip (v31.2: Chip statt Radio, gleiche Bindung). */}
                {teamlessActiveLocal.length > 0 && (
                  <div className="dex-ui-card dex-ui-card--soft" style={{ padding: '12px 14px', marginBottom: 12 }}>
                    <div className="dex-ui-label" style={{ marginBottom: 2 }}>
                      {t('Bereits angemeldet, noch ohne Team', 'Already registered, no team yet')}
                      <span className="dex-ui-pill dex-ui-pill--gray">{teamlessActiveLocal.length}</span>
                    </div>
                    <div className="dex-ui-help" style={{ margin: '0 0 8px' }}>
                      {t('Mehrere auswählbar — sie werden nur zugeordnet, nicht neu angemeldet.', 'Pick several — they are only assigned, not registered again.')}
                    </div>
                    <div className="dex-ui-stack" style={{ gap: 6, maxHeight: 280, overflowY: 'auto' }}>
                      {teamlessActiveLocal.map(p => {
                        const nm = `${p.Vorname || ''} ${p.Nachname || ''}`.trim() || p.ParticipantName || p.ParticipantEmail;
                        const isPicked = adminAddTeamlessPicks.has(p.Id);
                        const isLead = adminAddLeadRegId === p.Id;
                        // v22.40: Über-Kapazitäts-Sperre — nicht mehr als
                        // freie Plätze auswählbar; bereits Gewählte bleiben
                        // abwählbar.
                        const locked = !isPicked && atCap;
                        const cbId = `dex-addmember-teamless-${p.Id}`;
                        return (
                          <div key={p.Id} className={cx('dex-ui-toggle-row', isPicked && 'is-active', locked && 'is-disabled')} style={{ alignItems: 'center', padding: '8px 12px' }}>
                            <input
                              id={cbId}
                              type="checkbox"
                              checked={isPicked}
                              disabled={locked}
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
                              style={{ marginTop: 0 }}
                            />
                            <label htmlFor={cbId} className="dex-ui-toggle-row-body" style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: locked ? 'not-allowed' : 'pointer' }}>
                              <img className="dex-ui-avatar" src={photoOf(p.ParticipantEmail)} alt={nm} onError={hidePhoto} />
                              <span className="dex-ui-row-main">
                                <span className="dex-ui-row-title" style={{ display: 'block' }}>{nm}</span>
                                <span className="dex-ui-row-sub" style={{ display: 'block' }}>{p.ParticipantEmail}</span>
                              </span>
                            </label>
                            {isPicked && (
                              <label className={cx('dex-ui-chip', isLead && 'is-active')} title={t('Diese Person wird Team-Lead', 'This person becomes team lead')}>
                                <input type="radio" name="lead-pick" className="dex-ui-sr-only" checked={isLead} onChange={() => setAdminAddLeadRegId(p.Id)} />
                                {isLead && <Check size={12} />}Lead
                              </label>
                            )}
                          </div>
                        );
                      })}
                    </div>
                    {adminAddTeamlessPicks.size > 0 && (
                      <div className="dex-ui-help" style={{ marginTop: 8 }}>
                        <strong>{adminAddTeamlessPicks.size}</strong> {t('ausgewählt', 'selected')}
                        {!adminAddLeadRegId && ` — ${t('markiere eine Person als Lead (oder lass es leer = kein Lead).', 'mark one person as lead (or leave it empty = no lead).')}`}
                        {atCap && (
                          <span style={{ color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 600 }}>
                            {' — '}{t(`Team voll (${freeSlots} ${freeSlots === 1 ? 'Platz' : 'Plätze'}).`, `Team full (${freeSlots} ${freeSlots === 1 ? 'seat' : 'seats'}).`)}
                          </span>
                        )}
                        {/* v17.4: Nur Zuordnung — die Personen haben bei der eigenen
                            Anmeldung schon zugestimmt, also keine Consent-Box. */}
                        {onlyTeamlessPicks && (
                          <span style={{ display: 'block', marginTop: 4 }}>
                            {t('Keine neue Anmeldung, keine Bestätigungsmail — außer du schaltest unten die Info-Mail ein.', 'No new registration, no confirmation mail — unless you switch on the info mail below.')}
                          </span>
                        )}
                      </div>
                    )}
                  </div>
                )}
                {/* v22.45: Neue Teilnehmer — Person, die noch NICHT beim Event
                    angemeldet ist, per Suche stellvertretend hinzufügen. */}
                <div className="dex-ui-field">
                  <div className="dex-ui-label" style={{ marginBottom: 2 }}>
                    {teamlessActiveLocal.length > 0 ? t('Oder eine neue Person hinzufügen', 'Or add a new person') : t('Neue Person hinzufügen', 'Add a new person')}
                  </div>
                  <div className="dex-ui-help" style={{ margin: '0 0 8px' }}>
                    {t('Jemand, der noch nicht beim Event angemeldet ist — du meldest die Person stellvertretend an.', 'Someone not yet registered for the event — you register them on their behalf.')}
                  </div>
                  {adminAddMemberPick ? (
                    <div className="dex-ui-card dex-ui-card--accent" style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '10px 14px' }}>
                      <img className="dex-ui-avatar dex-ui-avatar--lg" src={photoOf(adminAddMemberPick.email)} alt={adminAddMemberPick.displayName} onError={hidePhoto} />
                      <div className="dex-ui-row-main">
                        <div className="dex-ui-row-title">{adminAddMemberPick.displayName}</div>
                        <div className="dex-ui-row-sub">{adminAddMemberPick.email}</div>
                      </div>
                      <span className="dex-ui-pill dex-ui-pill--green">{t('wird neu angemeldet', 'new registration')}</span>
                      <button
                        type="button"
                        className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                        onClick={() => { setAdminAddMemberPick(null); setAdminAddMemberQuery(''); setAdminAddMemberResults([]); }}
                        title={t('Auswahl entfernen', 'Remove selection')}
                        aria-label={t('Auswahl entfernen', 'Remove selection')}
                      ><X size={16} /></button>
                    </div>
                  ) : (
                    <div style={{ position: 'relative' }}>
                      <span style={{ position: 'absolute', left: 12, top: 12, color: 'var(--dex-gray-400)', display: 'inline-flex', pointerEvents: 'none' }}><Search size={16} /></span>
                      <input
                        className="dex-ui-input"
                        style={{ paddingLeft: 36 }}
                        value={adminAddMemberQuery}
                        // v22.40: Suche sperren, wenn das Team durch die Auswahl
                        // bereits voll ist (keine zusätzliche neue Person mehr).
                        disabled={atCap}
                        placeholder={atCap ? t('Team voll — keine weitere Person', 'Team full — no further person') : t('Name oder E-Mail eingeben…', 'Type a name or e-mail…')}
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
                        <div className="dex-ui-card" style={{ position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 10, marginTop: 4, padding: 4, maxHeight: 220, overflowY: 'auto', boxShadow: '0 6px 20px rgba(0,0,0,0.12)' }}>
                          {adminAddMemberSearching && (
                            <div className="dex-ui-muted" style={{ padding: 10 }}>{t('Suche…', 'Searching…')}</div>
                          )}
                          {/* v31.2: `dex-ui-textbtn` setzt den Button-Rahmen zurück, `dex-ui-row`
                              (später im Stylesheet) liefert Layout und Hover — kein Inline-
                              Background, der den Hover überdecken würde. */}
                          {adminAddMemberResults.map(r => (
                            <button
                              key={r.email}
                              type="button"
                              className="dex-ui-row dex-ui-textbtn"
                              onClick={() => { setAdminAddMemberPick(r); setAdminAddMemberResults([]); setAdminAddMemberQuery(''); }}
                              style={{ width: '100%', textAlign: 'left' }}
                            >
                              <img className="dex-ui-avatar" src={photoOf(r.email)} alt={r.displayName} onError={hidePhoto} />
                              <span className="dex-ui-row-main">
                                <span className="dex-ui-row-title" style={{ display: 'block' }}>{r.displayName}</span>
                                <span className="dex-ui-row-sub" style={{ display: 'block' }}>{r.email}</span>
                              </span>
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                  {/* v17.4: Consent nur, wenn eine wirklich NEUE Person via Graph
                      hinzugefügt wird. v31.2: Hinweis und Bestätigung stehen jetzt
                      direkt unter der gewählten Person — vorher lag der Kasten oben
                      und die Checkbox ganz unten. */}
                  {consentRequired && (
                    <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 10, flexDirection: 'column', gap: 8 }}>
                      <div>
                        <strong>{t('Vorab die Zustimmung einholen.', 'Get their consent first.')}</strong>{' '}
                        {t('Mit dem Hinzufügen meldest du diese Person an: Sie bekommt eine Anmeldebestätigung per Mail, einen Outlook-Termin und sieht das Event unter „Meine Events“.', 'Adding registers this person: they get a confirmation mail, an Outlook appointment and see the event under “My events”.')}
                      </div>
                      <label className={cx('dex-ui-toggle-row', adminAddMemberConsent && 'is-active')} style={{ padding: '8px 12px' }}>
                        <input type="checkbox" checked={adminAddMemberConsent} onChange={e => setAdminAddMemberConsent(e.target.checked)} />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">
                            <span style={{ color: 'var(--dex-red)' }}>*</span>
                            {t('Ich bestätige: Die Person hat ihrer Anmeldung zugestimmt.', 'I confirm: this person agreed to being registered.')}
                          </span>
                        </span>
                      </label>
                    </div>
                  )}
                </div>
              </div>
              {/* v22.45/v22.49: Kommunikation an das Team — erst sichtbar, wenn
                  jemand ausgewählt ist. */}
              {(adminAddTeamlessPicks.size > 0 || hasGraphPick) && (
                <div>
                  <div className="dex-ui-section-title">{t('Wer bekommt eine Mail?', 'Who gets a mail?')}</div>
                  <div className="dex-ui-stack" style={{ gap: 8 }}>
                    {/* a) Neue Person (Graph-Pick): Anmeldebestätigung + Outlook
                        optional (Default an — echte Neu-Anmeldung). */}
                    {hasGraphPick && (
                      <label className={cx('dex-ui-toggle-row', adminAddNewPersonMail && 'is-active')}>
                        <input type="checkbox" checked={adminAddNewPersonMail} onChange={e => setAdminAddNewPersonMail(e.target.checked)} />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">{t('Anmeldebestätigung und Kalendereinladung an die neue Person', 'Confirmation and calendar invite to the new person')}</span>
                          <span className="dex-ui-toggle-row-desc">{t('Vorgabe an, weil die Person neu angemeldet wird. Ausgeschaltet heißt: still hinzufügen, ohne Mail.', 'On by default because the person is newly registered. Off means: add silently, no mail.')}</span>
                        </span>
                      </label>
                    )}
                    {/* b) Info-Mail an die zugeordneten (bereits angemeldeten) Personen. */}
                    {adminAddTeamlessPicks.size > 0 && (
                      <label className={cx('dex-ui-toggle-row', adminAddSendMail && 'is-active')}>
                        <input type="checkbox" checked={adminAddSendMail} onChange={e => setAdminAddSendMail(e.target.checked)} />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">{t('Info-Mail an die zugeordneten Personen', 'Info mail to the assigned people')}</span>
                          <span className="dex-ui-toggle-row-desc">{t('Vorgabe aus, weil diese Personen schon beim Event angemeldet sind.', 'Off by default because these people are already registered for the event.')}</span>
                        </span>
                      </label>
                    )}
                    {(adminAddSendMail || (hasGraphPick && adminAddNewPersonMail)) && (
                      <label className={cx('dex-ui-toggle-row', adminAddCcOrganizer && 'is-active')} style={{ marginLeft: 24 }}>
                        <input type="checkbox" checked={adminAddCcOrganizer} onChange={e => setAdminAddCcOrganizer(e.target.checked)} />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">{t('Kopie (CC) an mich', 'Copy (CC) to me')}</span>
                          <span className="dex-ui-toggle-row-desc">{currentUser.email} {t('bekommt diese Mail(s) in Kopie.', 'receives a copy of these mails.')}</span>
                        </span>
                      </label>
                    )}
                    {/* c) Übrige Team-Mitglieder informieren — Reichweite alle / nur Lead
                        (v31.2: Chips statt Radio-Paar, gleiche Bindung). */}
                    <label className={cx('dex-ui-toggle-row', adminAddNotifyOthers && 'is-active')}>
                      <input type="checkbox" checked={adminAddNotifyOthers} onChange={e => setAdminAddNotifyOthers(e.target.checked)} />
                      <span className="dex-ui-toggle-row-body">
                        <span className="dex-ui-toggle-row-title">{t('Die übrigen Team-Mitglieder informieren', 'Tell the other team members')}</span>
                        <span className="dex-ui-toggle-row-desc">{t('Die bisherigen Mitglieder bekommen eine kurze „Neues Mitglied“-Info.', 'Existing members get a short “new member” notice.')}</span>
                      </span>
                    </label>
                    {adminAddNotifyOthers && (
                      <div className="dex-ui-inline" style={{ marginLeft: 24 }}>
                        <span className="dex-ui-muted">{t('Wen genau?', 'Whom exactly?')}</span>
                        <button type="button" className={cx('dex-ui-chip', adminAddNotifyScope === 'all' && 'is-active')} aria-pressed={adminAddNotifyScope === 'all'} onClick={() => setAdminAddNotifyScope('all')}>
                          {t('Alle Mitglieder', 'All members')}
                        </button>
                        <button type="button" className={cx('dex-ui-chip', adminAddNotifyScope === 'lead' && 'is-active')} aria-pressed={adminAddNotifyScope === 'lead'} onClick={() => setAdminAddNotifyScope('lead')}>
                          {t('Nur den Team-Lead', 'Only the team lead')}
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              )}
              {adminAddMemberError && (
                <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
                  <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                  <span>{adminAddMemberError}</span>
                </div>
              )}
            </div>
          </Modal>
        );
};
