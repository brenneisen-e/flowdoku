/* TeamsSection — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 11031-11568 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { EventService, SPRegistration } from '../../../services/EventService';
import { ChevronDown, ChevronUp, Pencil, Plus, RefreshCw, Users } from '../../Icons';
import { Icon } from '@fluentui/react/lib/Icon';
import { DeloitteEvent } from '../../../types';

export interface TeamsSectionProps {
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  currentUser: import("../../../types/index").User;
  dragOverTid: string;
  dragRegId: number;
  eventServiceRef: EventService;
  getActiveTeams: () => Array<{    tid: string;    teamName: string;    members: SPRegistration[];}>;
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
  isAdmin: boolean;
  isDe: boolean;
  isLoadingRegs: boolean;
  isMobile: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  leadTransferBusy: boolean;
  leadTransferOpenFor: string;
  moveRegToTeam: (reg: SPRegistration, targetTid: string, targetTeamName: string | undefined) => Promise<void>;
  onTeamDrop: (targetTid: string, targetTeamName: string | undefined) => void;
  openTeamMailDialog: () => void;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setAdminAddCcOrganizer: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddLeadRegId: React.Dispatch<React.SetStateAction<number>>;
  setAdminAddMemberConsent: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddMemberDialog: React.Dispatch<React.SetStateAction<{ teamId: string; teamName: string; freeSlots: number; isNewTeam?: boolean; }>>;
  setAdminAddMemberError: React.Dispatch<React.SetStateAction<string>>;
  setAdminAddMemberPick: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; }>>;
  setAdminAddMemberQuery: React.Dispatch<React.SetStateAction<string>>;
  setAdminAddMemberResults: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; }[]>>;
  setAdminAddNewPersonMail: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddNotifyOthers: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddNotifyScope: React.Dispatch<React.SetStateAction<"all" | "lead">>;
  setAdminAddSendMail: React.Dispatch<React.SetStateAction<boolean>>;
  setAdminAddTeamlessPicks: React.Dispatch<React.SetStateAction<Set<number>>>;
  setDragOverTid: React.Dispatch<React.SetStateAction<string>>;
  setDragRegId: React.Dispatch<React.SetStateAction<number>>;
  setLeadTransferBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setLeadTransferOpenFor: React.Dispatch<React.SetStateAction<string>>;
  setRegistrations: React.Dispatch<React.SetStateAction<SPRegistration[]>>;
  setTeamEditOpenFor: React.Dispatch<React.SetStateAction<string>>;
  setTeamsCollapsed: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamsToast: React.Dispatch<React.SetStateAction<string>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  teamEditOpenFor: string;
  teamsCollapsed: boolean;
  transferTeamLead: (eventId: string, teamId: string, newLeadEmail: string) => Promise<{ ok: boolean; reason?: string; }>;
}

export const TeamsSection: React.FC<TeamsSectionProps> = (p) => {
  const { confirmDialog, currentUser, dragOverTid, dragRegId, eventServiceRef, getActiveTeams, getAllRegistrations, isAdmin, isDe, isLoadingRegs, isMobile, isOrganizerFor, leadTransferBusy, leadTransferOpenFor, moveRegToTeam, onTeamDrop, openTeamMailDialog, registrations, selectedEvent, setAdminAddCcOrganizer, setAdminAddLeadRegId, setAdminAddMemberConsent, setAdminAddMemberDialog, setAdminAddMemberError, setAdminAddMemberPick, setAdminAddMemberQuery, setAdminAddMemberResults, setAdminAddNewPersonMail, setAdminAddNotifyOthers, setAdminAddNotifyScope, setAdminAddSendMail, setAdminAddTeamlessPicks, setDragOverTid, setDragRegId, setLeadTransferBusy, setLeadTransferOpenFor, setRegistrations, setTeamEditOpenFor, setTeamsCollapsed, setTeamsToast, showAlert, teamEditOpenFor, teamsCollapsed, transferTeamLead } = p;
          // v11.84: Teams-Section — Admin-Center-Team-Management.
          // Sichtbar nur für Events mit aktivierter Team-Anmeldung. Listet
          // alle Teams (gruppiert per TeamId, abgemeldete Mitglieder
          // ausgeblendet), mit Lead-Badge und Buttons für „+ Person
          // hinzufügen" und „Lead-Rolle übergeben". Reagiert live auf
          // `registrations` — kein zusätzlicher Roundtrip.
          if (!selectedEvent || !selectedEvent.teamRegistrationEnabled) return null;
          if (isLoadingRegs) return null;

          // groupBy TeamId, abgemeldete Personen NICHT eingehen lassen.
          const teamsByid: Record<string, SPRegistration[]> = {};
          // v16.2: Teilnehmer ohne Team in eine eigene Liste — werden
          // unten als „Teilnehmer ohne Team"-Sektion gerendert, damit
          // der Organizer sie sieht und ggf. einem (neuen) Team zuordnen
          // kann.
          const teamlessActive: SPRegistration[] = [];
          for (const r of registrations) {
            if (r.Status === 'Abgemeldet') continue;
            const tid = r.TeamId || '';
            if (!tid) {
              teamlessActive.push(r);
              continue;
            }
            (teamsByid[tid] = teamsByid[tid] || []).push(r);
          }
          // Sortierung: aelteste Lead-RegistrationDate zuerst.
          const teamEntries = Object.entries(teamsByid)
            .map(([tid, members]) => {
              // Lead oben, dann TeilnehmerID aufsteigend.
              members.sort((a, b) => {
                if (!!a.TeamLead !== !!b.TeamLead) return a.TeamLead ? -1 : 1;
                const aT = (a.TeilnehmerID ?? 9_999_999) as number;
                const bT = (b.TeilnehmerID ?? 9_999_999) as number;
                return aT - bT;
              });
              const lead = members.find(m => !!m.TeamLead) || members[0];
              const leadDate = lead?.RegistrationDate ? new Date(lead.RegistrationDate).getTime() : Number.MAX_SAFE_INTEGER;
              return { tid, members, lead, leadDate };
            })
            .sort((a, b) => a.leadDate - b.leadDate);

          const teamSizeCfg = selectedEvent.teamSize || 0;
          const count = teamEntries.length;
          const canManage = isAdmin || isOrganizerFor(selectedEvent);

          // v26.x (Mobile): HTML5-Drag&Drop feuert auf Touch-Geräten nicht.
          // Deshalb auf dem Handy pro Person ein simples Auswahlmenü zum
          // Umsortieren anbieten — reine Zusatz-UI, die Drag-Logik bleibt
          // unangetastet. Nutzt denselben Pfad (moveRegToTeam) wie der Drop.
          const teamSelectOptions = teamEntries.map(te => ({
            tid: te.tid,
            label: te.members.find(mm => !!mm.TeamName)?.TeamName || `${selectedEvent.teamTermSingular || 'Team'} ${te.tid}`,
          }));
          const MobileTeamSelect = (reg: SPRegistration): React.ReactElement => {
            const curTid = reg.TeamId || '';
            return (
              <select
                value={curTid}
                aria-label={isDe ? 'Team ändern' : 'Change team'}
                onChange={e => {
                  const target = e.target.value;
                  const opt = teamSelectOptions.find(o => o.tid === target);
                  moveRegToTeam(reg, target, opt?.label).catch(() => { /* */ });
                }}
                style={{
                  marginTop: 6, width: '100%', fontSize: '0.82rem',
                  padding: '6px 8px', borderRadius: 8,
                  border: '1px solid var(--dex-gray-300)', background: '#fff',
                }}
              >
                <option value="">{isDe ? `Ohne ${selectedEvent.teamTermSingular || 'Team'}` : `No ${selectedEvent.teamTermSingular || 'team'}`}</option>
                {teamSelectOptions.map(o => (
                  <option key={o.tid} value={o.tid}>{o.label}</option>
                ))}
              </select>
            );
          };

          const statusBadge = (st: string): React.ReactElement | null => {
            if (!st || st === 'Angemeldet') return null;
            const colorMap: Record<string, string> = {
              'Warteliste': '#b35a00',
              'QR versendet': '#3a7dbf',
              'Eingecheckt': '#4a7c1f',
            };
            const color = colorMap[st] || 'var(--dex-gray-500)';
            return (
              <span style={{
                display: 'inline-block', padding: '1px 8px', borderRadius: 10,
                background: `${color}15`, color, fontSize: '0.7rem', fontWeight: 600, marginLeft: 6,
              }}>{st}</span>
            );
          };

          return (
            <div style={{ marginBottom: 20, padding: 16, borderRadius: 12, border: '1px solid var(--dex-gray-200)', background: '#fff' }}>
              <div
                role="button"
                tabIndex={0}
                onClick={() => setTeamsCollapsed(v => !v)}
                onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setTeamsCollapsed(v => !v); } }}
                style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer', userSelect: 'none' }}
              >
                <Users size={20} />
                <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1rem' }}>
                  {(selectedEvent?.teamTermPlural || 'Teams')} ({count})
                </strong>
                <span style={{ marginLeft: 'auto', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  {teamsCollapsed ? <ChevronDown size={18} /> : <ChevronUp size={18} />}
                </span>
              </div>
              {!teamsCollapsed && (
                <div style={{ marginTop: 14, display: 'flex', flexDirection: 'column', gap: 12 }}>
                  {/* v23.0: Drag&Drop-Hinweis. */}
                  {canManage && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.8rem', color: 'var(--dex-gray-600)', background: 'rgba(134,188,37,0.08)', border: '1px solid var(--dex-green, #86bc25)', borderRadius: 8, padding: '7px 12px' }}>
                      <Icon iconName="DragObject" style={{ fontSize: 15, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                      {isDe
                        ? `Tipp: Personen per Drag & Drop zwischen ${(selectedEvent?.teamTermPlural || 'Teams')} und „ohne ${(selectedEvent?.teamTermSingular || 'Team')}" verschieben.`
                        : `Tip: drag & drop people between ${(selectedEvent?.teamTermPlural || 'teams')} and “no ${(selectedEvent?.teamTermSingular || 'team')}”.`}
                    </div>
                  )}
                  {teamEntries.length === 0 && (
                    <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.88rem', fontStyle: 'italic' }}>
                      Keine Team-Anmeldungen bisher.
                    </div>
                  )}
                  {/* v16.2: „Neues Team anlegen"-Button + Teamless-Sektion.
                      v23.0: zusätzlich „Mail an <Teams>"-Button (Per-Team-Info-Mail). */}
                  {canManage && (
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.85rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                        onClick={() => {
                          // Neue lokale TeamID generieren und Add-Member-Dialog
                          // direkt damit öffnen. Sobald die erste Person hinzu-
                          // gefügt wird, wird die TeamId im SP-Item gespeichert.
                          const newTid = (typeof crypto !== 'undefined' && crypto.randomUUID)
                            ? crypto.randomUUID()
                            : `team-${Date.now()}-${Math.floor(Math.random() * 1e9)}`;
                          setAdminAddMemberDialog({ teamId: newTid, teamName: '', freeSlots: teamSizeCfg || 99, isNewTeam: true });
                          setAdminAddMemberPick(null);
                          setAdminAddMemberQuery('');
                          setAdminAddMemberResults([]);
                          setAdminAddMemberConsent(false);
                          setAdminAddMemberError('');
                          setAdminAddTeamlessPicks(new Set());
                          setAdminAddLeadRegId(null);
                          setAdminAddSendMail(false);
                          setAdminAddCcOrganizer(false);
                          setAdminAddNotifyOthers(false);
                          setAdminAddNotifyScope('all');
                          setAdminAddNewPersonMail(true);
                        }}
                      >
                        <Plus size={14} /> {isDe ? `Neue ${selectedEvent?.teamTermSingular || 'Team'} anlegen` : `Create new ${selectedEvent?.teamTermSingular || 'team'}`}
                      </button>
                      {getActiveTeams().length > 0 && (
                        <button
                          type="button"
                          className="btn btn-secondary"
                          style={{ fontSize: '0.85rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                          onClick={openTeamMailDialog}
                          title={isDe ? 'Jedem Mitglied eine eigene Mail mit team-spezifischer Info senden (z.B. Teams-Einwahllink).' : 'Send each member an individual mail with team-specific info (e.g. a Teams join link).'}
                        >
                          <Icon iconName="Mail" style={{ fontSize: 14 }} /> {isDe ? `Mail an ${selectedEvent?.teamTermPlural || 'Teams'}` : `Mail to ${selectedEvent?.teamTermPlural || 'teams'}`}
                        </button>
                      )}
                    </div>
                  )}
                  {/* v23.5: „ohne Team"-Box ist jetzt IMMER ein Drop-Ziel (für
                      canManage), auch wenn gerade niemand teamlos ist — sonst
                      konnte man eine Person per Drag&Drop nicht aus ihrem Team
                      nehmen (die Box war nur bei vorhandenen teamlosen Personen
                      da). Leerer Zustand zeigt einen Hinweis als Drop-Fläche. */}
                  {(canManage || teamlessActive.length > 0) && (
                    <div
                      onDragOver={canManage ? (e => { e.preventDefault(); setDragOverTid(''); }) : undefined}
                      onDragLeave={canManage ? (() => setDragOverTid(prev => (prev === '' ? null : prev))) : undefined}
                      onDrop={canManage ? (() => onTeamDrop('', undefined)) : undefined}
                      style={{ padding: 14, border: dragOverTid === '' ? '2px dashed var(--dex-green, #86bc25)' : '1px dashed var(--dex-orange, #ed8b00)', borderRadius: 10, background: dragOverTid === '' ? 'rgba(134,188,37,0.10)' : 'rgba(237,139,0,0.04)' }}
                    >
                      <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, flexWrap: 'wrap', marginBottom: 10 }}>
                        <strong style={{ fontSize: '0.95rem', color: 'var(--dex-orange-dark, #b35a00)' }}>
                          {isDe ? `Teilnehmer ohne ${selectedEvent?.teamTermSingular || 'Team'}` : `Attendees without ${selectedEvent?.teamTermSingular || 'team'}`} ({teamlessActive.length})
                        </strong>
                        <span style={{ color: 'var(--dex-gray-600)', fontSize: '0.82rem' }}>
                          — Einzel-Anmeldungen ohne Team-Zuordnung
                        </span>
                      </div>
                      {teamlessActive.length === 0 && (
                        <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', fontStyle: 'italic', padding: '6px 2px' }}>
                          {isDe
                            ? `Aktuell ist niemand ohne ${selectedEvent?.teamTermSingular || 'Team'}. Zieh eine Person aus einem ${selectedEvent?.teamTermSingular || 'Team'} hierher, um die Zuordnung zu lösen.`
                            : `Nobody is currently without a ${selectedEvent?.teamTermSingular || 'team'}. Drag a person from a ${selectedEvent?.teamTermSingular || 'team'} here to remove their assignment.`}
                        </div>
                      )}
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                        {teamlessActive.map(m => {
                          const name = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const dept = (m as any).Department || '';
                          return (
                            <div
                              key={m.Id}
                              draggable={canManage}
                              onDragStart={canManage ? (() => setDragRegId(m.Id)) : undefined}
                              onDragEnd={canManage ? (() => { setDragRegId(null); setDragOverTid(null); }) : undefined}
                              title={canManage ? (isDe ? 'Ziehen, um zuzuordnen' : 'Drag to assign') : undefined}
                              style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '4px 6px', borderRadius: 6, cursor: canManage ? 'grab' : 'default', opacity: dragRegId === m.Id ? 0.4 : 1, background: dragRegId === m.Id ? 'var(--dex-gray-100)' : 'transparent' }}
                            >
                              <img
                                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=L`}
                                alt={name}
                                onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                              />
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontSize: '0.88rem', fontWeight: 500 }}>{name}{statusBadge(m.Status)}</div>
                                <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)' }}>{m.ParticipantEmail}</div>
                                {dept && <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', marginTop: 1 }}>{dept}</div>}
                                {isMobile && canManage && teamSelectOptions.length > 0 && MobileTeamSelect(m)}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                  {/* v19.0: Teams in einem responsiven 3-Spalten-Raster +
                      durchnummeriert — spart vertikalen Platz im Organizer-Center. */}
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))', gap: 12, alignItems: 'stretch' }}>
                  {teamEntries.map(({ tid, members, lead }, teamIdx) => {
                    const teamName = members.find(m => !!m.TeamName)?.TeamName || '';
                    const total = members.length;
                    const free = teamSizeCfg > 0 ? Math.max(0, teamSizeCfg - total) : 0;
                    const canAdd = canManage && (teamSizeCfg === 0 || total < teamSizeCfg);
                    const leadEmail = lead?.ParticipantEmail || '';
                    const otherMembers = members.filter(m => m.Id !== lead?.Id);
                    // v19.19: Teams mit freien Plätzen farblich (orange) hervorheben,
                    // damit der Organizer auf einen Blick sieht, welche Teams noch
                    // nicht voll belegt sind.
                    const hasFreeSlots = free > 0;
                    const isDropTarget = dragOverTid === tid;
                    return (
                      <div
                        key={tid}
                        onDragOver={canManage ? (e => { e.preventDefault(); setDragOverTid(tid); }) : undefined}
                        onDragLeave={canManage ? (() => setDragOverTid(prev => (prev === tid ? null : prev))) : undefined}
                        onDrop={canManage ? (() => onTeamDrop(tid, teamName || undefined)) : undefined}
                        style={{
                          padding: 14,
                          border: isDropTarget ? '2px solid var(--dex-green, #86bc25)' : (hasFreeSlots ? '1px solid var(--dex-orange, #ed8b00)' : '1px solid var(--dex-gray-200)'),
                          borderRadius: 10,
                          background: isDropTarget ? 'rgba(134,188,37,0.12)' : (hasFreeSlots ? 'rgba(237,139,0,0.06)' : 'var(--dex-gray-50, #f7f7f7)'),
                          // v19.19: Flex-Spalte, damit der Aktions-Block (u.a.
                          // „Lead-Rolle übergeben") per marginTop:auto immer am
                          // unteren Kartenrand sitzt → alle Karten gleich hoch.
                          display: 'flex',
                          flexDirection: 'column',
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, flexWrap: 'wrap', marginBottom: 10 }}>
                          <strong style={{ fontSize: '0.95rem', color: 'var(--dex-gray-800)' }}>
                            <span style={{ color: 'var(--dex-gray-400)', marginRight: 4 }}>{teamIdx + 1}.</span>{teamName ? `Team „${teamName}"` : 'Team (ohne Namen)'}
                          </strong>
                          <span style={{ color: hasFreeSlots ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-600)', fontSize: '0.85rem', fontWeight: hasFreeSlots ? 600 : 400 }}>
                            {teamSizeCfg > 0 ? `${total}/${teamSizeCfg} belegt` : `${total} Mitglieder`}
                          </span>
                          {hasFreeSlots && (
                            <span style={{
                              display: 'inline-block', padding: '1px 8px', borderRadius: 10,
                              background: 'var(--dex-orange, #ed8b00)', color: '#fff',
                              fontSize: '0.7rem', fontWeight: 700,
                            }}>{free} frei</span>
                          )}
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                          {members.map(m => {
                            const name = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
                            const isLead = !!m.TeamLead;
                            return (
                              <div
                                key={m.Id}
                                draggable={canManage}
                                onDragStart={canManage ? (() => setDragRegId(m.Id)) : undefined}
                                onDragEnd={canManage ? (() => { setDragRegId(null); setDragOverTid(null); }) : undefined}
                                title={canManage ? (isDe ? 'Ziehen, um in ein anderes Team / „ohne Team" zu verschieben' : 'Drag to move to another team / “no team”') : undefined}
                                style={{
                                  display: 'flex', alignItems: 'center', gap: 10,
                                  padding: '4px 6px', borderRadius: 6,
                                  cursor: canManage ? 'grab' : 'default',
                                  opacity: dragRegId === m.Id ? 0.4 : 1,
                                  background: dragRegId === m.Id ? 'var(--dex-gray-100)' : 'transparent',
                                }}
                              >
                                <div style={{ position: 'relative', width: 32, height: 32, flexShrink: 0 }}>
                                  <img
                                    src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=L`}
                                    alt={name}
                                    onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                    style={{
                                      width: 32, height: 32, borderRadius: '50%', objectFit: 'cover',
                                      background: 'var(--dex-gray-100)',
                                      transition: 'transform 0.18s ease',
                                      transformOrigin: 'left center',
                                      cursor: 'pointer',
                                    }}
                                    onMouseEnter={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(2.4)'; (e.currentTarget as HTMLImageElement).style.zIndex = '10'; (e.currentTarget as HTMLImageElement).style.position = 'relative'; (e.currentTarget as HTMLImageElement).style.boxShadow = '0 8px 24px rgba(0,0,0,0.25)'; }}
                                    onMouseLeave={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(1)'; (e.currentTarget as HTMLImageElement).style.zIndex = ''; (e.currentTarget as HTMLImageElement).style.position = ''; (e.currentTarget as HTMLImageElement).style.boxShadow = ''; }}
                                  />
                                </div>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontSize: '0.88rem', fontWeight: 500 }}>
                                    {name}
                                    {statusBadge(m.Status)}
                                  </div>
                                  <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)' }}>{m.ParticipantEmail}</div>
                                  {/* v16.1: Business Area / Department aus
                                      der SP-Registrierung mit anzeigen,
                                      damit der Organizer auf einen Blick
                                      sieht, aus welcher Practice die
                                      Mitglieder kommen. */}
                                  {(() => {
                                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                    const dept = (m as any).Department || '';
                                    if (!dept) return null;
                                    return (
                                      <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', marginTop: 1 }}>{dept}</div>
                                    );
                                  })()}
                                  {isMobile && canManage && MobileTeamSelect(m)}
                                </div>
                                {isLead && (
                                  <span style={{
                                    display: 'inline-block', padding: '2px 10px', borderRadius: 12,
                                    background: 'var(--dex-green, #86bc25)', color: '#fff',
                                    fontSize: '0.72rem', fontWeight: 700,
                                  }}>Lead</span>
                                )}
                                {/* v22.41/v22.45: „Aus Team entfernen" — löst NUR
                                    die Team-Zuordnung (TeamId/Lead/Name leeren),
                                    die Anmeldung inkl. Status (z.B. Warteliste)
                                    bleibt bestehen. Erscheint erst im „Anpassen"-
                                    Modus des Teams (nicht dauerhaft an jedem Namen). */}
                                {canManage && teamEditOpenFor === tid && eventServiceRef && selectedEvent.subsiteUrl && (
                                  <button
                                    type="button"
                                    title="Aus dem Team entfernen (Anmeldung bleibt bestehen)"
                                    onClick={async () => {
                                      const sub = selectedEvent.subsiteUrl;
                                      if (!sub) return;
                                      const stHint = m.Status && m.Status !== 'Angemeldet' ? ` (Status: ${m.Status})` : '';
                                      const ok = await confirmDialog(
                                        `${name} aus dem Team „${teamName || ''}" entfernen?\n\nDie Anmeldung${stHint} bleibt bestehen — die Person steht danach ohne Team da und kann einem anderen Team zugeordnet werden.`,
                                        { danger: true, confirmLabel: 'Aus Team entfernen' }
                                      );
                                      if (!ok) return;
                                      try {
                                        await eventServiceRef.assignRegistrationToTeam(sub, m.Id, '', '', false);
                                        // Lead entfernt + andere bleiben → frühestes Mitglied nachziehen.
                                        if (isLead) {
                                          const rest = members
                                            .filter(x => x.Id !== m.Id && x.Status !== 'Abgemeldet')
                                            .sort((a, b) => ((a.TeilnehmerID ?? 9_999_999) as number) - ((b.TeilnehmerID ?? 9_999_999) as number));
                                          if (rest.length > 0) {
                                            await eventServiceRef.assignRegistrationToTeam(sub, rest[0].Id, tid, teamName || undefined, true);
                                          }
                                        }
                                        await eventServiceRef.writeChangeLog({
                                          action: 'TeamMemberRemoved',
                                          targetType: 'Participant',
                                          targetId: m.ParticipantEmail,
                                          targetName: name,
                                          eventId: selectedEvent.id,
                                          eventTitle: selectedEvent.title,
                                          details: { teamId: tid, removedBy: currentUser.email, keptStatus: m.Status },
                                        }).catch(() => { /* */ });
                                        setTeamsToast(`${name} wurde aus dem Team entfernt — Anmeldung bleibt bestehen.`);
                                        window.setTimeout(() => setTeamsToast(''), 4500);
                                        const regs = await getAllRegistrations(selectedEvent.id);
                                        setRegistrations(regs);
                                      } catch (err) {
                                        console.warn('[DEX] removeFromTeam failed:', err);
                                        showAlert('Entfernen aus dem Team fehlgeschlagen.', { variant: 'error' });
                                      }
                                    }}
                                    style={{
                                      background: 'none', border: 'none', cursor: 'pointer',
                                      color: 'var(--dex-red, #c00)', fontSize: '0.72rem',
                                      textDecoration: 'underline', padding: '2px 4px', flexShrink: 0,
                                    }}
                                  >
                                    Entfernen
                                  </button>
                                )}
                              </div>
                            );
                          })}
                        </div>
                        {canManage && (
                          <div style={{ marginTop: 'auto', paddingTop: 12, display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 8, position: 'relative' }}>
                            {canAdd && (
                              <button
                                type="button"
                                className="btn btn-secondary"
                                style={{ fontSize: '0.8rem', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                                onClick={() => {
                                  setAdminAddMemberDialog({ teamId: tid, teamName, freeSlots: free });
                                  setAdminAddMemberPick(null);
                                  setAdminAddMemberQuery('');
                                  setAdminAddMemberResults([]);
                                  setAdminAddMemberConsent(false);
                                  setAdminAddMemberError('');
                                  setAdminAddTeamlessPicks(new Set());
                                  setAdminAddLeadRegId(null);
                                  setAdminAddSendMail(false);
                                  setAdminAddCcOrganizer(false);
                                  setAdminAddNotifyOthers(false);
                                  setAdminAddNotifyScope('all');
                                  setAdminAddNewPersonMail(true);
                                }}
                              >
                                <Plus size={14} /> Person hinzufügen
                                {teamSizeCfg > 0 && ` (${free} Slot${free === 1 ? '' : 's'} frei)`}
                              </button>
                            )}
                            {/* v22.45: „Anpassen" schaltet den Bearbeiten-Modus
                                des Teams ein/aus — erst dann erscheinen die
                                „Entfernen"-Buttons pro Mitglied. */}
                            {canManage && (
                              <button
                                type="button"
                                className={teamEditOpenFor === tid ? 'btn btn-primary' : 'btn btn-secondary'}
                                style={{ fontSize: '0.8rem', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                                onClick={() => setTeamEditOpenFor(teamEditOpenFor === tid ? null : tid)}
                              >
                                <Pencil size={14} /> {teamEditOpenFor === tid ? 'Fertig' : 'Anpassen'}
                              </button>
                            )}
                            {otherMembers.length > 0 && (
                              <>
                                <button
                                  type="button"
                                  className="btn btn-secondary"
                                  style={{ fontSize: '0.8rem', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                                  onClick={() => setLeadTransferOpenFor(leadTransferOpenFor === tid ? null : tid)}
                                >
                                  <RefreshCw size={14} /> Lead-Rolle übergeben
                                </button>
                                {leadTransferOpenFor === tid && (
                                  <div style={{
                                    position: 'absolute', top: '100%', left: 0, marginTop: 6,
                                    background: '#fff', border: '1px solid var(--dex-gray-300)',
                                    borderRadius: 8, boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
                                    zIndex: 20, minWidth: 280, maxWidth: 360, padding: 6,
                                  }}>
                                    <div style={{ padding: '6px 10px', fontSize: '0.78rem', color: 'var(--dex-gray-600)', borderBottom: '1px solid var(--dex-gray-100)' }}>
                                      Neue Lead-Rolle übertragen an:
                                    </div>
                                    {otherMembers.map(m => {
                                      const nm = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
                                      return (
                                        <button
                                          key={m.Id}
                                          type="button"
                                          disabled={leadTransferBusy}
                                          onClick={async () => {
                                            if (leadTransferBusy) return;
                                            setLeadTransferBusy(true);
                                            try {
                                              const res = await transferTeamLead(selectedEvent.id, tid, m.ParticipantEmail);
                                              if (res.ok) {
                                                setTeamsToast(`Lead-Rolle wurde an ${nm} übergeben.`);
                                                const regs = await getAllRegistrations(selectedEvent.id);
                                                setRegistrations(regs);
                                                window.setTimeout(() => setTeamsToast(''), 4500);
                                              } else {
                                                setTeamsToast(`Lead-Übergabe fehlgeschlagen: ${res.reason || 'Unbekannter Fehler'}.`);
                                                window.setTimeout(() => setTeamsToast(''), 4500);
                                              }
                                            } finally {
                                              setLeadTransferBusy(false);
                                              setLeadTransferOpenFor(null);
                                            }
                                          }}
                                          style={{
                                            display: 'flex', alignItems: 'center', gap: 8,
                                            width: '100%', padding: '8px 10px', border: 'none',
                                            background: 'transparent', cursor: 'pointer',
                                            textAlign: 'left', borderRadius: 6,
                                          }}
                                          onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.background = 'var(--dex-gray-100)'; }}
                                          onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.background = 'transparent'; }}
                                        >
                                          <img
                                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=S`}
                                            alt={nm}
                                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                            style={{ width: 24, height: 24, borderRadius: '50%' }}
                                          />
                                          <div style={{ minWidth: 0 }}>
                                            <div style={{ fontSize: '0.85rem', fontWeight: 500 }}>{nm}</div>
                                            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{m.ParticipantEmail}</div>
                                          </div>
                                        </button>
                                      );
                                    })}
                                    <button
                                      type="button"
                                      onClick={() => setLeadTransferOpenFor(null)}
                                      style={{
                                        width: '100%', padding: '6px 10px',
                                        border: 'none', borderTop: '1px solid var(--dex-gray-100)',
                                        background: 'transparent', cursor: 'pointer',
                                        fontSize: '0.78rem', color: 'var(--dex-gray-500)',
                                      }}
                                    >Abbrechen</button>
                                  </div>
                                )}
                              </>
                            )}
                            {/* leadEmail nur als Referenz für den Lead-Lookup behalten — nicht für's TS-Linting wegwerfen. */}
                            <span style={{ display: 'none' }}>{leadEmail}</span>
                          </div>
                        )}
                      </div>
                    );
                  })}
                  </div>
                </div>
              )}
            </div>
          );
};

