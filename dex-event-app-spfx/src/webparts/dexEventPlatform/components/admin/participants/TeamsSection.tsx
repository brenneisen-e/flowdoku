/* TeamsSection — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 11031-11568 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich übernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { EventService, SPRegistration } from '../../../services/EventService';
import { ChevronDown, ChevronUp, Pencil, Plus, RefreshCw, Trash2, Users } from '../../Icons';
import { Icon } from '@fluentui/react/lib/Icon';
import { DeloitteEvent } from '../../../types';
// v31.3: Gemeinsame Klassen des Organizer Centers (Karten, Zeilen, Pillen,
// Aufklapper) — siehe docs/ui-leitfaden.md.
import { cx, ensureDexUiStyles } from '../../dexUi';

export interface TeamsSectionProps {
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  currentUser: import("../../../types/index").User;
  dragOverTid: string;
  dragRegId: number;
  eventServiceRef: EventService;
  getActiveTeams: () => Array<{    tid: string;    teamName: string;    members: SPRegistration[];}>;
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
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
  setTeamEditOpenFor: React.Dispatch<React.SetStateAction<string>>;
  setTeamsCollapsed: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamsToast: React.Dispatch<React.SetStateAction<string>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  teamEditOpenFor: string;
  teamsCollapsed: boolean;
  transferTeamLead: (eventId: string, teamId: string, newLeadEmail: string) => Promise<{ ok: boolean; reason?: string; }>;
}

export const TeamsSection: React.FC<TeamsSectionProps> = (p) => {
  const { confirmDialog, currentUser, dragOverTid, dragRegId, eventServiceRef, getActiveTeams, isAdmin, isDe, isLoadingRegs, isMobile, isOrganizerFor, leadTransferBusy, leadTransferOpenFor, moveRegToTeam, onTeamDrop, openTeamMailDialog, registrations, reloadRegistrations, selectedEvent, setAdminAddCcOrganizer, setAdminAddLeadRegId, setAdminAddMemberConsent, setAdminAddMemberDialog, setAdminAddMemberError, setAdminAddMemberPick, setAdminAddMemberQuery, setAdminAddMemberResults, setAdminAddNewPersonMail, setAdminAddNotifyOthers, setAdminAddNotifyScope, setAdminAddSendMail, setAdminAddTeamlessPicks, setDragOverTid, setDragRegId, setLeadTransferBusy, setLeadTransferOpenFor, setTeamEditOpenFor, setTeamsCollapsed, setTeamsToast, showAlert, teamEditOpenFor, teamsCollapsed, transferTeamLead } = p;
          // v31.3: Idempotent — die Sektion rendert auch ohne offenes Modal,
          // und ohne das Stylesheet hätten die dex-ui-Klassen keine Wirkung.
          ensureDexUiStyles();
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
          // v31.3: Die Bezeichnung des Events (z.B. „Gruppe" statt „Team") stand
          // an fünfzehn Stellen als `selectedEvent?.teamTermSingular || 'Team'`
          // im Text — einmal ausgerechnet ist sie auch einmal zu ändern.
          const termOne = selectedEvent?.teamTermSingular || 'Team';
          const termMany = selectedEvent?.teamTermPlural || 'Teams';
          const termOneEn = selectedEvent?.teamTermSingular || 'team';
          const termManyEn = selectedEvent?.teamTermPlural || 'teams';

          // v26.x (Mobile): HTML5-Drag&Drop feuert auf Touch-Geräten nicht.
          // Deshalb auf dem Handy pro Person ein simples Auswahlmenü zum
          // Umsortieren anbieten — reine Zusatz-UI, die Drag-Logik bleibt
          // unangetastet. Nutzt denselben Pfad (moveRegToTeam) wie der Drop.
          const teamSelectOptions = teamEntries.map(te => ({
            tid: te.tid,
            label: te.members.find(mm => !!mm.TeamName)?.TeamName || `${termOne} ${te.tid}`,
          }));
          const MobileTeamSelect = (reg: SPRegistration): React.ReactElement => {
            const curTid = reg.TeamId || '';
            return (
              <select
                value={curTid}
                aria-label={isDe ? `${termOne} wechseln` : `Change ${termOneEn}`}
                onChange={e => {
                  const target = e.target.value;
                  const opt = teamSelectOptions.find(o => o.tid === target);
                  moveRegToTeam(reg, target, opt?.label).catch(() => { /* */ });
                }}
                className="dex-ui-select dex-ui-select--sm"
                style={{ marginTop: 6 }}
              >
                <option value="">{isDe ? `Ohne ${termOne}` : `No ${termOneEn}`}</option>
                {teamSelectOptions.map(o => (
                  <option key={o.tid} value={o.tid}>{o.label}</option>
                ))}
              </select>
            );
          };

          // v31.3: Status als Pille nach Leitfaden 5b — dieselben Farben wie in
          // der Teilnehmertabelle (blau eingecheckt, orange Warteliste), damit
          // dieselbe Person in beiden Ansichten gleich aussieht.
          const statusBadge = (st: string): React.ReactElement | null => {
            if (!st || st === 'Angemeldet') return null;
            const tone = st === 'Warteliste' ? 'dex-ui-pill--orange'
              : st === 'Eingecheckt' ? 'dex-ui-pill--blue'
              : st === 'QR versendet' ? 'dex-ui-pill--green'
              : 'dex-ui-pill--gray';
            return <span className={cx('dex-ui-pill', tone)}>{st}</span>;
          };

          // v31.3: Personen-Zelle (Foto, Name + Status, Mail, Practice) — bis
          // v31.2 stand dieselbe Optik zweimal inline im Code, einmal je Liste;
          // die Teams-Liste hatte eine Lupe beim Foto, die Teamlosen nicht.
          // `zoom` hält diesen Unterschied, ohne die Darstellung zu doppeln.
          const personCell = (m: SPRegistration, name: string, zoom: boolean): React.ReactElement => {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const dept = (m as any).Department || '';
            const img = (
              <img
                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=L`}
                alt={name}
                onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                className="dex-ui-avatar"
                // v31.3: `zoom-in` statt `pointer` — das Foto vergrößert sich beim
                // Überfahren, ein Klick tut aber nichts (Leitfaden 1.3).
                style={zoom ? { transition: 'transform 0.18s ease', transformOrigin: 'left center', cursor: 'zoom-in' } : undefined}
                onMouseEnter={zoom ? (e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(2.4)'; (e.currentTarget as HTMLImageElement).style.zIndex = '10'; (e.currentTarget as HTMLImageElement).style.position = 'relative'; (e.currentTarget as HTMLImageElement).style.boxShadow = '0 8px 24px rgba(0,0,0,0.25)'; }) : undefined}
                onMouseLeave={zoom ? (e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(1)'; (e.currentTarget as HTMLImageElement).style.zIndex = ''; (e.currentTarget as HTMLImageElement).style.position = ''; (e.currentTarget as HTMLImageElement).style.boxShadow = ''; }) : undefined}
              />
            );
            return (
              <>
                {zoom ? <div style={{ position: 'relative', width: 32, height: 32, flexShrink: 0 }}>{img}</div> : img}
                <div className="dex-ui-row-main">
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6, minWidth: 0 }}>
                    {/* minWidth:0 — ohne das schrumpft ein langer Name nicht und
                        läuft aus der Karte heraus, statt mit Ellipse zu enden. */}
                    <span className="dex-ui-person-name" style={{ minWidth: 0 }}>{name}</span>
                    {statusBadge(m.Status)}
                  </div>
                  <div className="dex-ui-person-sub">{m.ParticipantEmail}</div>
                  {/* v16.1: Business Area / Department aus der SP-Registrierung
                      mit anzeigen, damit der Organizer auf einen Blick sieht,
                      aus welcher Practice die Mitglieder kommen. */}
                  {dept && <div className="dex-ui-person-sub" style={{ fontSize: '0.72rem' }}>{dept}</div>}
                  {/* Innerhalb eines Teams ist `teamSelectOptions` nie leer — die
                      Prüfung stammt aus der Teamlosen-Liste und schadet hier nicht. */}
                  {isMobile && canManage && teamSelectOptions.length > 0 && MobileTeamSelect(m)}
                </div>
              </>
            );
          };

          return (
            <div className={cx('dex-ui-card', 'dex-ui-card--list')} style={{ marginBottom: 20 }}>
              {/* v31.3: Die Kopfzeile IST der Aufklapper — deshalb Hover und
                  Zeiger (Leitfaden 1.3). Zähler als Pille daneben; wer noch ohne
                  Zuordnung ist, steht auch im eingeklappten Zustand da, denn
                  genau das ist die Arbeit, die hier wartet. */}
              <div
                role="button"
                tabIndex={0}
                onClick={() => setTeamsCollapsed(v => !v)}
                onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setTeamsCollapsed(v => !v); } }}
                className="dex-ui-row"
                style={{ cursor: 'pointer', userSelect: 'none' }}
              >
                <span className="dex-ui-card-head-title"><Users size={18} /> {termMany}</span>
                <span className="dex-ui-pill dex-ui-pill--gray">{count}</span>
                {teamlessActive.length > 0 && (
                  <span className="dex-ui-pill dex-ui-pill--orange">
                    {isDe ? `${teamlessActive.length} ohne ${termOne}` : `${teamlessActive.length} without ${termOneEn}`}
                  </span>
                )}
                <span style={{ marginLeft: 'auto', display: 'inline-flex', alignItems: 'center', color: 'var(--dex-gray-600)' }}>
                  {teamsCollapsed ? <ChevronDown size={18} /> : <ChevronUp size={18} />}
                </span>
              </div>
              {!teamsCollapsed && (
                <div style={{ padding: '4px 6px 6px', display: 'flex', flexDirection: 'column', gap: 12 }}>
                  {/* v16.2: „Neues Team anlegen"-Button + Teamless-Sektion.
                      v23.0: zusätzlich „Mail an <Teams>"-Button (Per-Team-Info-Mail).
                      v31.3: Beide Knöpfe stehen jetzt ZUERST und links unter dem
                      Titel (Leitfaden 2a′) — vorher lagen sie hinter Hinweis und
                      Teamlosen-Kasten, also hinter dem, was sie erzeugen. */}
                  {canManage && (
                    <div className="dex-ui-inline">
                      <button
                        type="button"
                        className="btn btn-secondary dex-ui-btn-sm"
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
                        {/* v31.3: „Team anlegen" statt „Neue Team anlegen" — der
                            Artikel passte nur zu weiblichen Bezeichnungen. */}
                        <Plus size={14} /> {isDe ? `${termOne} anlegen` : `Create new ${termOneEn}`}
                      </button>
                      {getActiveTeams().length > 0 && (
                        <button
                          type="button"
                          className="btn btn-secondary dex-ui-btn-sm"
                          onClick={openTeamMailDialog}
                          title={isDe ? 'Jedem Mitglied eine eigene Mail mit team-spezifischer Info senden (z.B. Teams-Einwahllink).' : 'Send each member an individual mail with team-specific info (e.g. a Teams join link).'}
                        >
                          <Icon iconName="Mail" style={{ fontSize: 14 }} /> {isDe ? `Mail an ${termMany}` : `Mail to ${termManyEn}`}
                        </button>
                      )}
                    </div>
                  )}
                  {/* v23.0: Drag&Drop-Hinweis. v31.3: als ruhiger Hinweiskasten —
                      grün ist der Farbe der Auswahl vorbehalten (Leitfaden 1.1). */}
                  {canManage && (
                    <div className="dex-ui-callout dex-ui-callout--neutral dex-ui-callout--sm">
                      <span className="dex-ui-callout-icon"><Icon iconName="DragObject" style={{ fontSize: 14 }} /></span>
                      <span>
                        {isDe
                          ? `Tipp: Personen per Drag & Drop zwischen ${termMany} und „ohne ${termOne}“ verschieben.`
                          : `Tip: drag & drop people between ${termManyEn} and “no ${termOneEn}”.`}
                      </span>
                    </div>
                  )}
                  {/* v19.0: Teams in einem responsiven 3-Spalten-Raster +
                      durchnummeriert — spart vertikalen Platz im Organizer-Center.
                      v31.3: Die Teams stehen jetzt VOR den Teilnehmern ohne
                      Zuordnung — die Sektion heißt so, und der Kasten darunter ist
                      der Rest, nicht der Anfang. Ohne Teams ein leerer Zustand
                      statt einer leeren Fläche (Leitfaden 1.6). */}
                  {teamEntries.length === 0 ? (
                    <div className="dex-ui-empty">
                      <div className="dex-ui-empty-icon"><Users size={20} /></div>
                      <div className="dex-ui-empty-title">{isDe ? `Noch keine ${termMany}` : `No ${termManyEn} yet`}</div>
                      {isDe
                        ? `Oben über „${termOne} anlegen“ startest du — wer sich einzeln angemeldet hat, steht darunter.`
                        : `Start with “Create new ${termOneEn}” above — people who registered individually are listed below.`}
                    </div>
                  ) : (
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))', gap: 12, alignItems: 'stretch' }}>
                  {teamEntries.map(({ tid, members, lead }, teamIdx) => {
                    const teamName = members.find(m => !!m.TeamName)?.TeamName || '';
                    const total = members.length;
                    const free = teamSizeCfg > 0 ? Math.max(0, teamSizeCfg - total) : 0;
                    const canAdd = canManage && (teamSizeCfg === 0 || total < teamSizeCfg);
                    const otherMembers = members.filter(m => m.Id !== lead?.Id);
                    // v19.19: Teams mit freien Plätzen bleiben erkennbar — seit
                    // v31.3 über Zähler-Pille und Auslastungsbalken statt über eine
                    // komplett orange Karte: Farbe nur dort, wo sie etwas bedeutet
                    // (Leitfaden 1.1), sonst leuchtet die halbe Sektion.
                    const hasFreeSlots = free > 0;
                    const isDropTarget = dragOverTid === tid;
                    return (
                      <div
                        key={tid}
                        onDragOver={canManage ? (e => { e.preventDefault(); setDragOverTid(tid); }) : undefined}
                        onDragLeave={canManage ? (() => setDragOverTid(prev => (prev === tid ? null : prev))) : undefined}
                        onDrop={canManage ? (() => onTeamDrop(tid, teamName || undefined)) : undefined}
                        className={cx('dex-ui-card', 'dex-ui-card--list')}
                        style={{
                          display: 'flex',
                          flexDirection: 'column',
                          // v31.3: Beim Ziehen zeigt die Zielkarte grüne Kante und
                          // grünen Grund — vorher unterschied sie sich nur um einen
                          // Pixel Rahmenstärke von den anderen.
                          ...(isDropTarget ? { borderColor: 'var(--dex-green, #86bc25)', background: 'rgba(134,188,37,0.10)' } : null),
                        }}
                      >
                        <div className="dex-ui-card-head" style={{ padding: '6px 8px 4px' }}>
                          <h4 className="dex-ui-card-head-title">
                            <span style={{ color: 'var(--dex-gray-400)' }}>{teamIdx + 1}.</span>
                            {teamName ? `${termOne} „${teamName}“` : (isDe ? `${termOne} (ohne Namen)` : `${termOneEn} (unnamed)`)}
                          </h4>
                          <span className={cx('dex-ui-pill', hasFreeSlots ? 'dex-ui-pill--orange' : 'dex-ui-pill--gray')}>
                            {teamSizeCfg > 0 ? `${total}/${teamSizeCfg}` : (isDe ? `${total} Mitglieder` : `${total} members`)}
                          </span>
                          {hasFreeSlots && (
                            <span className="dex-ui-card-head-meta">
                              {isDe ? `${free} ${free === 1 ? 'Platz' : 'Plätze'} frei` : `${free} ${free === 1 ? 'slot' : 'slots'} free`}
                            </span>
                          )}
                        </div>
                        {/* v31.3: Auslastung als Balken — „3/4" liest man, ein
                            halbvoller Balken sieht man. */}
                        {teamSizeCfg > 0 && (
                          <div className="dex-ui-progress" style={{ margin: '0 8px 6px' }} aria-hidden="true">
                            <div
                              className={cx('dex-ui-progress-bar', total > teamSizeCfg && 'dex-ui-progress-bar--red')}
                              style={{ width: `${Math.min(100, Math.round((total / teamSizeCfg) * 100))}%` }}
                            />
                          </div>
                        )}
                        {/* v31.3: Die Aktionen des Teams stehen jetzt links unter
                            seinem Namen statt unten am Kartenrand (Leitfaden 2a′) —
                            wer den Namen liest, sieht sofort, was er damit tun kann.
                            `position: relative` trägt das Lead-Menü. */}
                        {canManage && (
                          <div className="dex-ui-inline" style={{ padding: '0 8px 8px', position: 'relative' }}>
                            {canAdd && (
                              <button
                                type="button"
                                className="btn btn-secondary dex-ui-btn-sm"
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
                                {/* v31.3: Die freien Plätze stehen jetzt oben in der
                                    Kopfzeile — zweimal dieselbe Zahl braucht niemand. */}
                                <Plus size={14} /> {isDe ? 'Person hinzufügen' : 'Add person'}
                              </button>
                            )}
                            {/* v22.45: „Anpassen" schaltet den Bearbeiten-Modus
                                des Teams ein/aus — erst dann erscheinen die
                                „Entfernen"-Knöpfe pro Mitglied. v31.3: als Chip mit
                                is-active, damit der eingeschaltete Zustand sichtbar
                                bleibt, ohne einen zweiten Primär-Knopf zu setzen. */}
                            <button
                              type="button"
                              className={cx('dex-ui-chip', teamEditOpenFor === tid && 'is-active')}
                              onClick={() => setTeamEditOpenFor(teamEditOpenFor === tid ? null : tid)}
                            >
                              <Pencil size={13} /> {teamEditOpenFor === tid ? (isDe ? 'Fertig' : 'Done') : (isDe ? 'Anpassen' : 'Edit')}
                            </button>
                            {otherMembers.length > 0 && (
                              <>
                                <button
                                  type="button"
                                  className={cx('dex-ui-chip', leadTransferOpenFor === tid && 'is-active')}
                                  onClick={() => setLeadTransferOpenFor(leadTransferOpenFor === tid ? null : tid)}
                                >
                                  <RefreshCw size={13} /> {isDe ? 'Lead übergeben' : 'Transfer lead'}
                                </button>
                                {leadTransferOpenFor === tid && (
                                  <div style={{
                                    position: 'absolute', top: '100%', left: 0, marginTop: 6,
                                    background: '#fff', border: '1px solid var(--dex-gray-200)',
                                    borderRadius: 12, boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
                                    zIndex: 20, minWidth: 280, maxWidth: 360, padding: 6,
                                  }}>
                                    <div className="dex-ui-muted" style={{ padding: '4px 8px 6px' }}>
                                      {isDe ? 'Wer übernimmt die Lead-Rolle?' : 'Who takes over as lead?'}
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
                                                setTeamsToast(isDe ? `Lead-Rolle wurde an ${nm} übergeben.` : `Lead role handed over to ${nm}.`);
                                                await reloadRegistrations();
                                                window.setTimeout(() => setTeamsToast(''), 4500);
                                              } else {
                                                setTeamsToast(isDe
                                                  ? `Lead-Übergabe fehlgeschlagen: ${res.reason || 'Unbekannter Fehler'}.`
                                                  : `Lead handover failed: ${res.reason || 'unknown error'}.`);
                                                window.setTimeout(() => setTeamsToast(''), 4500);
                                              }
                                            } finally {
                                              setLeadTransferBusy(false);
                                              setLeadTransferOpenFor(null);
                                            }
                                          }}
                                          className={cx('dex-ui-row', 'dex-ui-rowbtn')}
                                          style={{ padding: '6px 8px' }}
                                        >
                                          <img
                                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=S`}
                                            alt={nm}
                                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                            className="dex-ui-avatar"
                                            style={{ width: 24, height: 24 }}
                                          />
                                          <div className="dex-ui-row-main">
                                            <div className="dex-ui-row-title">{nm}</div>
                                            <div className="dex-ui-row-sub">{m.ParticipantEmail}</div>
                                          </div>
                                        </button>
                                      );
                                    })}
                                    <button
                                      type="button"
                                      onClick={() => setLeadTransferOpenFor(null)}
                                      className="dex-ui-textbtn dex-ui-textbtn--muted"
                                      style={{ width: '100%', justifyContent: 'center', marginTop: 2, borderTop: '1px solid var(--dex-gray-100)', borderRadius: 0 }}
                                    >{isDe ? 'Abbrechen' : 'Cancel'}</button>
                                  </div>
                                )}
                              </>
                            )}
                          </div>
                        )}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                          {members.map(m => {
                            const name = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
                            const isLead = !!m.TeamLead;
                            return (
                              <div
                                key={m.Id}
                                draggable={canManage}
                                onDragStart={canManage ? (() => setDragRegId(m.Id)) : undefined}
                                onDragEnd={canManage ? (() => { setDragRegId(null); setDragOverTid(null); }) : undefined}
                                title={canManage ? (isDe ? `Ziehen, um in ein anderes ${termOne} / „ohne ${termOne}“ zu verschieben` : `Drag to move to another ${termOneEn} / “no ${termOneEn}”`) : undefined}
                                className="dex-ui-row"
                                style={{
                                  padding: '6px 8px',
                                  cursor: canManage ? 'grab' : 'default',
                                  opacity: dragRegId === m.Id ? 0.4 : 1,
                                  background: dragRegId === m.Id ? 'var(--dex-gray-100)' : undefined,
                                }}
                              >
                                {/* v31.3: Griff — die ganze Zeile ist weiterhin
                                    ziehbar, aber ohne sichtbaren Griff sah man ihr
                                    das nicht an. */}
                                {canManage && <span className="dex-ui-drag-handle" aria-hidden="true">≡</span>}
                                {personCell(m, name, true)}
                                {isLead && <span className="dex-ui-pill dex-ui-pill--green">Lead</span>}
                                {/* v22.41/v22.45: „Aus Team entfernen" — löst NUR
                                    die Team-Zuordnung (TeamId/Lead/Name leeren),
                                    die Anmeldung inkl. Status (z.B. Warteliste)
                                    bleibt bestehen. Erscheint erst im „Anpassen"-
                                    Modus des Teams (nicht dauerhaft an jedem Namen). */}
                                {canManage && teamEditOpenFor === tid && eventServiceRef && selectedEvent.subsiteUrl && (
                                  <span className="dex-ui-row-actions">
                                  <button
                                    type="button"
                                    className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                                    title={isDe ? `Aus dem ${termOne} entfernen (Anmeldung bleibt bestehen)` : `Remove from ${termOneEn} (registration stays)`}
                                    aria-label={isDe ? `${name} aus dem ${termOne} entfernen` : `Remove ${name} from ${termOneEn}`}
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
                                        setTeamsToast(isDe
                                          ? `${name} wurde aus dem ${termOne} entfernt — Anmeldung bleibt bestehen.`
                                          : `${name} was removed from the ${termOneEn} — the registration stays.`);
                                        window.setTimeout(() => setTeamsToast(''), 4500);
                                        await reloadRegistrations();
                                      } catch (err) {
                                        console.warn('[DEX] removeFromTeam failed:', err);
                                        showAlert(isDe ? `Entfernen aus dem ${termOne} fehlgeschlagen.` : `Removing from the ${termOneEn} failed.`, { variant: 'error' });
                                      }
                                    }}
                                  >
                                    <Trash2 size={15} />
                                  </button>
                                  </span>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                  </div>
                  )}
                  {/* v23.5: „ohne Team"-Box ist IMMER ein Drop-Ziel (für
                      canManage), auch wenn gerade niemand teamlos ist — sonst
                      konnte man eine Person per Drag&Drop nicht aus ihrem Team
                      nehmen (die Box war nur bei vorhandenen teamlosen Personen
                      da). Leerer Zustand zeigt einen Hinweis als Drop-Fläche.
                      v31.3: gestrichelte Karte statt orangem Kasten — die Zahl
                      derer ohne Zuordnung steht als Pille im Kopf und oben in
                      der Sektionszeile. */}
                  {(canManage || teamlessActive.length > 0) && (
                    <div
                      onDragOver={canManage ? (e => { e.preventDefault(); setDragOverTid(''); }) : undefined}
                      onDragLeave={canManage ? (() => setDragOverTid(prev => (prev === '' ? null : prev))) : undefined}
                      onDrop={canManage ? (() => onTeamDrop('', undefined)) : undefined}
                      className={cx('dex-ui-card', 'dex-ui-card--list')}
                      style={{
                        borderStyle: 'dashed',
                        ...(dragOverTid === '' ? { borderColor: 'var(--dex-green, #86bc25)', background: 'rgba(134,188,37,0.10)' } : null),
                      }}
                    >
                      <div className="dex-ui-card-head" style={{ padding: '6px 8px 4px' }}>
                        <h4 className="dex-ui-card-head-title">
                          {isDe ? `Teilnehmer ohne ${termOne}` : `Attendees without ${termOneEn}`}
                        </h4>
                        <span className={cx('dex-ui-pill', teamlessActive.length > 0 ? 'dex-ui-pill--orange' : 'dex-ui-pill--gray')}>
                          {teamlessActive.length}
                        </span>
                        <span className="dex-ui-card-head-meta">
                          {isDe ? 'Einzel-Anmeldungen ohne Zuordnung' : 'Individual registrations without an assignment'}
                        </span>
                      </div>
                      {teamlessActive.length === 0 && (
                        <div className="dex-ui-muted" style={{ padding: '2px 8px 8px' }}>
                          {isDe
                            ? `Aktuell ist niemand ohne ${termOne}. Zieh eine Person aus einem ${termOne} hierher, um die Zuordnung zu lösen.`
                            : `Nobody is currently without a ${termOneEn}. Drag a person from a ${termOneEn} here to remove their assignment.`}
                        </div>
                      )}
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                        {teamlessActive.map(m => {
                          const name = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
                          return (
                            <div
                              key={m.Id}
                              draggable={canManage}
                              onDragStart={canManage ? (() => setDragRegId(m.Id)) : undefined}
                              onDragEnd={canManage ? (() => { setDragRegId(null); setDragOverTid(null); }) : undefined}
                              title={canManage ? (isDe ? 'Ziehen, um zuzuordnen' : 'Drag to assign') : undefined}
                              className="dex-ui-row"
                              style={{
                                padding: '6px 8px',
                                cursor: canManage ? 'grab' : 'default',
                                opacity: dragRegId === m.Id ? 0.4 : 1,
                                background: dragRegId === m.Id ? 'var(--dex-gray-100)' : undefined,
                              }}
                            >
                              {canManage && <span className="dex-ui-drag-handle" aria-hidden="true">≡</span>}
                              {personCell(m, name, false)}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>
          );
};

