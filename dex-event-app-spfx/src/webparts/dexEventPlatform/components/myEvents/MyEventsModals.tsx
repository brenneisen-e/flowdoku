/* MyEventsModals — aus MyEventsPage.tsx ausgelagert (Zeilen 2975-3589 des
 * urspruenglichen Stands, v30.65). Die fuenf Dialoge der Seite: Mitglied zum
 * Team hinzufuegen, Team verwalten, Cascade-Abmeldung, eigener QR-Code und das
 * Kommunikations-Log. Das JSX ist zeichengleich uebernommen; die
 * Anzeige-Bedingung (`addMemberDialog && …`) bleibt beim Aufrufer, die
 * Komponente rendert ihren Inhalt unbedingt.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import Modal from '../Modal';
import InternationalSearchToggle from '../InternationalSearchToggle';
import { SPRegistration, EventCommRow } from '../../services/EventService';
import { formatDate } from './myEventsHelpers';


export interface AddMemberModalProps {
  addMemberBusy: boolean;
  addMemberConsent: boolean;
  addMemberError: string;
  addMemberIncludeIntl: boolean;
  addMemberPick: { email: string; displayName: string; };
  addMemberQuery: string;
  addMemberQueryTimer: React.MutableRefObject<NodeJS.Timeout>;
  addMemberResults: { email: string; displayName: string; }[];
  addMemberSearching: boolean;
  closeAddMemberDialog: () => void;
  isDe: boolean;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  setAddMemberConsent: React.Dispatch<React.SetStateAction<boolean>>;
  setAddMemberIncludeIntl: React.Dispatch<React.SetStateAction<boolean>>;
  setAddMemberPick: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; }>>;
  setAddMemberQuery: React.Dispatch<React.SetStateAction<string>>;
  setAddMemberResults: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; }[]>>;
  setAddMemberSearching: React.Dispatch<React.SetStateAction<boolean>>;
  submitAddMember: () => Promise<void>;
}

export function AddMemberModal(props: AddMemberModalProps): React.ReactElement {
  const { addMemberBusy, addMemberConsent, addMemberError, addMemberIncludeIntl, addMemberPick, addMemberQuery, addMemberQueryTimer, addMemberResults, addMemberSearching, closeAddMemberDialog, isDe, searchUsers, setAddMemberConsent, setAddMemberIncludeIntl, setAddMemberPick, setAddMemberQuery, setAddMemberResults, setAddMemberSearching, submitAddMember } = props;
  return (
        <Modal
          open={true}
          onClose={closeAddMemberDialog}
          dismissable={!addMemberBusy}
          maxWidth={540}
          ariaLabel={isDe ? 'Mitglied zum Team hinzufügen' : 'Add member to team'}
        >
            <h3 style={{ margin: 0, fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Mitglied zum Team hinzufügen' : 'Add member to team'}
            </h3>
            {/* Pflicht-Hinweisbox — orange, analog zur Initial-Team-Anmeldung. */}
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
                {isDe ? 'Vorab Zustimmung einholen' : 'Get consent up front'}
              </div>
              <div style={{ marginBottom: 6 }}>
                {isDe
                  ? 'Die ausgewählte Person wird sofort und ohne weitere Rückfrage zum Team hinzugefügt. Sie bekommt automatisch:'
                  : 'The selected person is added to the team immediately, without further confirmation. They automatically receive:'}
              </div>
              <ul style={{ margin: '0 0 4px 18px', padding: 0 }}>
                <li>{isDe ? 'eine Anmeldebestätigung per Mail' : 'a confirmation email'}</li>
                <li>{isDe ? 'einen Outlook-Termin im Kalender' : 'an Outlook calendar invite'}</li>
                <li>{isDe ? 'den Event in „Meine Events"' : 'the event in „My Events"'}</li>
              </ul>
              <div style={{ marginTop: 4 }}>
                {isDe
                  ? <>Bitte stelle sicher, dass die Person ihrer Anmeldung <strong>vorher zugestimmt</strong> hat.</>
                  : <>Make sure the person has <strong>consented up front</strong>.</>}
              </div>
            </div>
            {/* People-Picker — simple Inline-Variante mit der searchUsers-API. */}
            <div>
              <label style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)', display: 'block', marginBottom: 4 }}>
                <span style={{ color: 'var(--dex-red)' }}>*</span> {isDe ? 'Person auswählen' : 'Pick a person'}
              </label>
              {addMemberPick ? (
                <div style={{
                  display: 'inline-flex', alignItems: 'center', gap: 10,
                  padding: '6px 10px 6px 6px',
                  border: '1px solid var(--dex-gray-200)',
                  borderRadius: 'var(--dex-radius)',
                  background: 'var(--dex-gray-50, #f7f7f7)',
                  maxWidth: '100%',
                }}>
                  <img
                    src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(addMemberPick.email)}&size=S`}
                    alt={addMemberPick.displayName}
                    onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                    style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                  />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: '0.9rem' }}>{addMemberPick.displayName}</div>
                    <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem' }}>{addMemberPick.email}</div>
                  </div>
                  <button
                    type="button"
                    onClick={() => { setAddMemberPick(null); setAddMemberQuery(''); setAddMemberResults([]); }}
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
                    value={addMemberQuery}
                    placeholder={isDe ? 'Name oder E-Mail eingeben…' : 'Type a name or email…'}
                    onChange={e => {
                      const val = e.target.value;
                      setAddMemberQuery(val);
                      if (addMemberQueryTimer.current) clearTimeout(addMemberQueryTimer.current);
                      if (val.length >= 2) {
                        addMemberQueryTimer.current = setTimeout(async () => {
                          setAddMemberSearching(true);
                          try {
                            const res = await searchUsers(val, addMemberIncludeIntl);
                            setAddMemberResults(res.map(r => ({ email: r.email, displayName: r.displayName })));
                          } catch { setAddMemberResults([]); }
                          setAddMemberSearching(false);
                        }, 300);
                      } else {
                        setAddMemberResults([]);
                      }
                    }}
                  />
                  <div style={{ marginTop: 2 }}>
                    <InternationalSearchToggle query={addMemberQuery} checked={addMemberIncludeIntl} onChange={setAddMemberIncludeIntl} isDe={isDe} />
                  </div>
                  {(addMemberResults.length > 0 || addMemberSearching) && (
                    <div style={{
                      position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 10,
                      background: '#fff', border: '1px solid var(--dex-gray-200)',
                      borderRadius: 6, marginTop: 4, maxHeight: 220, overflowY: 'auto',
                      boxShadow: '0 4px 12px rgba(0,0,0,0.12)',
                    }}>
                      {addMemberSearching && (
                        <div style={{ padding: 10, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
                          {isDe ? 'Suche…' : 'Searching…'}
                        </div>
                      )}
                      {addMemberResults.map(r => (
                        <button
                          key={r.email}
                          type="button"
                          onClick={() => { setAddMemberPick(r); setAddMemberResults([]); setAddMemberQuery(''); }}
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
            </div>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
              <input
                type="checkbox"
                checked={addMemberConsent}
                onChange={e => setAddMemberConsent(e.target.checked)}
                style={{ marginTop: 3 }}
              />
              <span>
                <span style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                {isDe
                  ? 'Ich bestätige, dass die ausgewählte Person ihrer Anmeldung zugestimmt hat.'
                  : 'I confirm that the selected person has consented to this registration.'}
              </span>
            </label>
            {addMemberError && (
              <div style={{ padding: 10, borderRadius: 6, background: 'rgba(220,38,38,0.10)', color: '#b91c1c', fontSize: '0.85rem' }}>
                {addMemberError}
              </div>
            )}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={closeAddMemberDialog}
                disabled={addMemberBusy}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={submitAddMember}
                disabled={!addMemberPick || !addMemberConsent || addMemberBusy}
              >
                {addMemberBusy
                  ? (isDe ? 'Wird hinzugefügt…' : 'Adding…')
                  : (isDe ? 'Hinzufügen' : 'Add')}
              </button>
            </div>
        </Modal>
  );
}


export interface ManageTeamModalProps {
  closeManageTeamDialog: () => void;
  currentUserEmail: string;
  isDe: boolean;
  manageTeamBusyId: number;
  manageTeamConfirm: SPRegistration;
  manageTeamDialog: { eventId: string; teamId: string; teamName: string; teamSize: number; };
  manageTeamMembers: SPRegistration[];
  performManageTeamCancel: (member: SPRegistration) => Promise<void>;
  setManageTeamConfirm: React.Dispatch<React.SetStateAction<SPRegistration>>;
}

export function ManageTeamModal(props: ManageTeamModalProps): React.ReactElement {
  const { closeManageTeamDialog, currentUserEmail, isDe, manageTeamBusyId, manageTeamConfirm, manageTeamDialog, manageTeamMembers, performManageTeamCancel, setManageTeamConfirm } = props;
        const activeMembers = manageTeamMembers.filter(m => m.Status !== 'Abgemeldet');
        const sortedAll = [...manageTeamMembers].sort((a, b) => {
          const aLead = a.TeamLead ? 0 : 1;
          const bLead = b.TeamLead ? 0 : 1;
          if (aLead !== bLead) return aLead - bLead;
          const aTid = typeof a.TeilnehmerID === 'number' ? a.TeilnehmerID : Number.MAX_SAFE_INTEGER;
          const bTid = typeof b.TeilnehmerID === 'number' ? b.TeilnehmerID : Number.MAX_SAFE_INTEGER;
          if (aTid !== bTid) return aTid - bTid;
          return a.Id - b.Id;
        });
        return (
          <Modal
            open={true}
            onClose={closeManageTeamDialog}
            dismissable={manageTeamBusyId === null}
            maxWidth={620}
            ariaLabel={isDe ? 'Team verwalten' : 'Manage team'}
          >
              <div>
                <h3 style={{ margin: 0, fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
                  {isDe
                    ? `Team „${manageTeamDialog.teamName || 'Unbenannt'}" verwalten`
                    : `Manage team „${manageTeamDialog.teamName || 'Unnamed'}"`}
                </h3>
                <div style={{ marginTop: 4, fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? `Belegung: ${activeMembers.length}/${manageTeamDialog.teamSize}`
                    : `Occupancy: ${activeMembers.length}/${manageTeamDialog.teamSize}`}
                </div>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {sortedAll.map(m => {
                  const isCancelled = m.Status === 'Abgemeldet';
                  const isMemberLead = !!m.TeamLead && !isCancelled;
                  const isSelf = (m.ParticipantEmail || '').toLowerCase() === (currentUserEmail || '').toLowerCase();
                  const fullName = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantEmail;
                  const loc = (m.Location || '').trim();
                  const busy = manageTeamBusyId === m.Id;
                  return (
                    <div
                      key={m.Id}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 12,
                        padding: '8px 10px', borderRadius: 6,
                        background: isCancelled ? 'rgba(0,0,0,0.04)' : '#fff',
                        border: '1px solid var(--dex-gray-200)',
                        opacity: isCancelled ? 0.55 : 1,
                      }}
                    >
                      <div style={{ width: 40, height: 40, borderRadius: '50%', flexShrink: 0, position: 'relative' }}>
                        <img
                          src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=L`}
                          alt={fullName}
                          onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                          style={{
                            width: 40, height: 40, borderRadius: '50%',
                            objectFit: 'cover', background: 'var(--dex-gray-100)',
                            transition: 'transform 160ms ease, box-shadow 160ms ease',
                            transformOrigin: 'left center',
                            /* v11.94: kein zoom-in-Cursor */
                          }}
                          onMouseEnter={e => {
                            if (isCancelled) return;
                            const img = e.currentTarget as HTMLImageElement;
                            img.style.transform = 'scale(2.4)';
                            img.style.boxShadow = '0 8px 24px rgba(0,0,0,0.35)';
                            img.style.zIndex = '50';
                            img.style.position = 'relative';
                          }}
                          onMouseLeave={e => {
                            const img = e.currentTarget as HTMLImageElement;
                            img.style.transform = 'scale(1)';
                            img.style.boxShadow = 'none';
                            img.style.zIndex = '';
                            img.style.position = '';
                          }}
                        />
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)', fontSize: '0.9rem' }}>
                          {fullName}
                        </div>
                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>
                          {m.ParticipantEmail}{loc ? ` · ${loc}` : ''}
                        </div>
                      </div>
                      {isCancelled ? (
                        <span style={{
                          padding: '2px 8px', borderRadius: 999,
                          background: 'var(--dex-gray-300, #c8c8c8)', color: '#fff',
                          fontSize: '0.68rem', fontWeight: 600, flexShrink: 0,
                        }}>
                          {isDe ? 'abgemeldet' : 'cancelled'}
                        </span>
                      ) : isMemberLead ? (
                        <span style={{
                          padding: '2px 8px', borderRadius: 999,
                          background: 'var(--dex-green, #86bc25)', color: '#fff',
                          fontSize: '0.68rem', fontWeight: 600, flexShrink: 0,
                        }}>
                          Lead
                        </span>
                      ) : null}
                      {/* Trash-Button — nur für aktive Nicht-Lead-Nicht-Self-Mitglieder. */}
                      {!isCancelled && !isMemberLead && !isSelf && (
                        <button
                          type="button"
                          onClick={() => setManageTeamConfirm(m)}
                          disabled={busy || manageTeamBusyId !== null}
                          title={isDe ? 'Diese Person aus dem Team abmelden' : 'Remove this person from the team'}
                          style={{
                            background: 'var(--dex-red, #d62828)', color: '#fff', border: 'none',
                            width: 32, height: 32, borderRadius: 6, cursor: 'pointer',
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            flexShrink: 0, opacity: busy ? 0.6 : 1,
                          }}
                        >
                          <Icon iconName="Delete" style={{ fontSize: 14 }} />
                        </button>
                      )}
                    </div>
                  );
                })}
              </div>
              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', lineHeight: 1.45 }}>
                {isDe
                  ? <>{'Du als Team-Lead kannst dich nicht selbst über diesen Pfad abmelden — nutze dafür den normalen „Abmelden"-Button auf deiner Event-Karte. Dort übernimmt die App automatisch den Lead-Wechsel an das früheste verbleibende Team-Mitglied.'}</>
                  : <>{'As team lead you cannot cancel yourself via this dialog — use the normal „Cancel" button on your event card instead. The app then automatically transfers the lead role to the earliest remaining team member.'}</>}
              </div>
              <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={closeManageTeamDialog}
                  disabled={manageTeamBusyId !== null}
                >
                  {isDe ? 'Schließen' : 'Close'}
                </button>
              </div>
            {/* Confirm-Modal (zweite Ebene) — sitzt auf demselben z-index-Layer. */}
            {manageTeamConfirm && (() => {
              const cm = manageTeamConfirm;
              const fullName = `${cm.Vorname || ''} ${cm.Nachname || ''}`.trim() || cm.ParticipantEmail;
              const busy = manageTeamBusyId === cm.Id;
              return (
                <Modal
                  open={true}
                  onClose={() => setManageTeamConfirm(null)}
                  dismissable={!busy}
                  maxWidth={480}
                  ariaLabel={isDe ? 'Person aus dem Team abmelden?' : 'Cancel this team member?'}
                >
                    <h3 style={{ margin: 0, fontSize: '1.05rem', color: 'var(--dex-gray-800)' }}>
                      {isDe ? 'Person aus dem Team abmelden?' : 'Cancel this team member?'}
                    </h3>
                    <div style={{ fontSize: '0.9rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                      {isDe ? (
                        <>
                          <p style={{ marginTop: 0 }}>
                            <strong>{fullName}</strong> wird vom Event abgemeldet. Die Person bekommt eine Abmelde-Bestätigungs-Mail und (falls Outlook aktiv) eine Termin-Absage.
                          </p>
                          <p style={{ marginBottom: 0, color: 'var(--dex-gray-600)' }}>
                            {'Hinweis: die Person könnte sich auch selbst über „Meine Events" abmelden — du machst das nur stellvertretend.'}
                          </p>
                        </>
                      ) : (
                        <>
                          <p style={{ marginTop: 0 }}>
                            <strong>{fullName}</strong> will be cancelled from the event. They will receive a cancellation confirmation email and (if Outlook is active) a calendar removal.
                          </p>
                          <p style={{ marginBottom: 0, color: 'var(--dex-gray-600)' }}>
                            {'Note: this person could also cancel themselves via „My Events" — you are doing it on their behalf.'}
                          </p>
                        </>
                      )}
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        onClick={() => setManageTeamConfirm(null)}
                        disabled={busy}
                      >
                        {isDe ? 'Abbrechen' : 'Cancel'}
                      </button>
                      <button
                        type="button"
                        className="btn btn-primary"
                        onClick={() => performManageTeamCancel(cm)}
                        disabled={busy}
                        style={{ background: 'var(--dex-red, #d62828)', borderColor: 'var(--dex-red, #d62828)' }}
                      >
                        {busy
                          ? (isDe ? 'Wird abgemeldet…' : 'Cancelling…')
                          : (isDe ? 'Person abmelden' : 'Cancel member')}
                      </button>
                    </div>
                </Modal>
              );
            })()}
          </Modal>
        );
}


export interface CascadeCancelModalProps {
  cascadeDialog: { parentTitle: string; subEvents: { id: string; title: string; startDate?: string; location?: string; }[]; resolve: (_choice: "cascade" | "parent-only" | "abort") => void; isSectionedEvent?: boolean; };
  isDe: boolean;
}

export function CascadeCancelModal(props: CascadeCancelModalProps): React.ReactElement {
  const { cascadeDialog, isDe } = props;
        const dlg = cascadeDialog;
        const choose = (c: 'cascade' | 'parent-only' | 'abort'): void => dlg.resolve(c);
        const isSec = !!dlg.isSectionedEvent;
        // v14.7: Bei „Sectioned"-Events (requireSubEventSelection an) sprechen
        // wir nicht von „Sub-Event" sondern von „Event-Section". Außerdem
        // wird die Option „nur Hauptevent abmelden, Sub-Events behalten"
        // weggelassen — der Teilnehmer war ja gar nicht „nur Hauptevent"
        // angemeldet (die Pflichtwahl hat das verhindert). Stattdessen
        // Hinweis aufs Anmelde-Bearbeiten für einzelne Sections.
        const title = isSec
          ? (isDe ? 'Komplett vom Event abmelden?' : 'Cancel registration entirely?')
          : (isDe ? 'Auch von Sub-Events abmelden?' : 'Cancel sub-events too?');
        return (
          <Modal
            open={true}
            onClose={() => choose('abort')}
            maxWidth={520}
            ariaLabel={title}
          >
              <h3 style={{ margin: 0, fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
                {title}
              </h3>
              <p style={{ margin: 0, fontSize: '0.9rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                {isSec
                  ? (isDe
                      ? <>Du bist für <strong>{dlg.subEvents.length}</strong> Event-Section{dlg.subEvents.length === 1 ? '' : 's'} von <strong>&bdquo;{dlg.parentTitle}&ldquo;</strong> angemeldet:</>
                      : <>You are registered for <strong>{dlg.subEvents.length}</strong> event section{dlg.subEvents.length === 1 ? '' : 's'} of <strong>&bdquo;{dlg.parentTitle}&ldquo;</strong>:</>)
                  : (isDe
                      ? <>Du bist für <strong>{dlg.subEvents.length}</strong> Sub-Event{dlg.subEvents.length === 1 ? '' : 's'} von <strong>&bdquo;{dlg.parentTitle}&ldquo;</strong> angemeldet:</>
                      : <>You are registered for <strong>{dlg.subEvents.length}</strong> sub-event{dlg.subEvents.length === 1 ? '' : 's'} of <strong>&bdquo;{dlg.parentTitle}&ldquo;</strong>:</>)
                }
              </p>
              <ul style={{ margin: 0, paddingLeft: 20, fontSize: '0.85rem', color: 'var(--dex-gray-700)', maxHeight: 200, overflowY: 'auto' }}>
                {/* v15.8: pro Sub-Event Titel + Datum + Ort listen, damit
                    der User auf einen Blick sieht was er da abmeldet. */}
                {dlg.subEvents.map(s => {
                  const dateStr = s.startDate
                    ? new Date(s.startDate).toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })
                    : '';
                  const subParts = [dateStr, s.location].filter(Boolean).join(' · ');
                  return (
                    <li key={s.id} style={{ marginBottom: 4 }}>
                      <div style={{ fontWeight: 600 }}>{s.title}</div>
                      {subParts && (
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>{subParts}</div>
                      )}
                    </li>
                  );
                })}
              </ul>
              {isSec && (
                <div style={{
                  padding: '10px 12px', borderRadius: 8,
                  background: 'rgba(0,118,168,0.08)',
                  border: '1px solid var(--dex-blue, #0076a8)',
                  fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.45,
                }}>
                  {isDe
                    ? <>Du willst nur <strong>einzelne Sections</strong> abwählen (z.B. &bdquo;Dinner absagen, aber Meeting bleibt&ldquo;)? Dann breche hier ab und nutze stattdessen <strong>&bdquo;Anmeldung bearbeiten&ldquo;</strong> auf der Event-Karte — dort kannst du gezielt einzelne Sections an- oder abwählen, ohne die ganze Teilnahme zu stornieren.</>
                    : <>Want to drop only <strong>individual sections</strong> (e.g. &bdquo;cancel dinner, keep meeting&ldquo;)? Then abort here and use <strong>&bdquo;Edit registration&ldquo;</strong> on the event card instead — there you can pick / deselect individual sections without cancelling the whole registration.</>}
                </div>
              )}
              <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.4 }}>
                {isDe
                  ? 'Wähle wie weiter:'
                  : 'How do you want to proceed?'}
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                <button
                  className="btn btn-primary"
                  onClick={() => choose('cascade')}
                  style={{ fontSize: '0.9rem', padding: '10px 16px' }}
                >
                  {isSec
                    ? (isDe
                        ? `Ja, komplett abmelden (Event + ${dlg.subEvents.length} Section${dlg.subEvents.length === 1 ? '' : 's'})`
                        : `Yes, cancel entirely (event + ${dlg.subEvents.length} section${dlg.subEvents.length === 1 ? '' : 's'})`)
                    : (isDe
                        ? `Alles abmelden (Hauptevent + ${dlg.subEvents.length} Sub-Event${dlg.subEvents.length === 1 ? '' : 's'})`
                        : `Cancel everything (main event + ${dlg.subEvents.length} sub-event${dlg.subEvents.length === 1 ? '' : 's'})`)}
                </button>
                {!isSec && (
                  <button
                    className="btn btn-secondary"
                    onClick={() => choose('parent-only')}
                    style={{ fontSize: '0.9rem', padding: '10px 16px' }}
                  >
                    {isDe
                      ? 'Nur Hauptevent abmelden, Sub-Events behalten'
                      : 'Cancel main event only, keep sub-events'}
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => choose('abort')}
                  style={{
                    background: 'transparent', border: 'none',
                    color: 'var(--dex-gray-500)', cursor: 'pointer',
                    fontSize: '0.85rem', textDecoration: 'underline',
                    padding: '6px 16px',
                  }}
                >
                  {isDe ? 'Abbrechen — nichts abmelden' : 'Cancel — keep everything'}
                </button>
              </div>
          </Modal>
        );
}


export interface MyQrModalProps {
  isDe: boolean;
  myQrModal: { dataUrl: string; name: string; tid?: number; eventTitle: string; };
  setMyQrModal: React.Dispatch<React.SetStateAction<{ dataUrl: string; name: string; tid?: number; eventTitle: string; }>>;
}

export function MyQrModal(props: MyQrModalProps): React.ReactElement {
  const { isDe, myQrModal, setMyQrModal } = props;
  return (
        <Modal
          open={true}
          onClose={() => setMyQrModal(null)}
          maxWidth={420}
          ariaLabel={isDe ? 'Mein QR-Code' : 'My QR code'}
        >
          <h3 style={{ margin: 0, fontSize: '1.05rem', textAlign: 'center' }}>
            {isDe ? 'Mein Check-in-QR-Code' : 'My check-in QR code'}
          </h3>
          <p style={{ margin: 0, fontSize: '0.82rem', color: 'var(--dex-gray-500)', textAlign: 'center' }}>
            {myQrModal.eventTitle}
          </p>
          <div style={{ textAlign: 'center' }}>
            <img
              src={myQrModal.dataUrl}
              alt="QR-Code"
              style={{ width: 280, maxWidth: '90%', height: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: 10, background: '#fff' }}
            />
          </div>
          <p style={{ margin: 0, textAlign: 'center', fontWeight: 700, fontSize: '0.95rem' }}>
            {myQrModal.name}
            {myQrModal.tid ? <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}> · Nr. {myQrModal.tid}</span> : null}
          </p>
          <p style={{ margin: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)', textAlign: 'center', lineHeight: 1.5 }}>
            {isDe
              ? 'Zeig diesen Code am Event-Tag dem Check-in-Team — er ist derselbe wie in deiner QR-Mail. Falls der Scan nicht klappt, reicht dein Name.'
              : 'Show this code to the check-in team on event day — it is the same as in your QR email. If the scan fails, your name is enough.'}
          </p>
          <div style={{ textAlign: 'right' }}>
            <button className="btn btn-secondary" onClick={() => setMyQrModal(null)} style={{ fontSize: '0.85rem' }}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          </div>
        </Modal>
  );
}


export interface EventCommsModalProps {
  commsLoading: boolean;
  commsModal: { eventId: string; eventTitle: string; };
  commsOpenId: number;
  commsRows: EventCommRow[];
  isDe: boolean;
  setCommsModal: React.Dispatch<React.SetStateAction<{ eventId: string; eventTitle: string; }>>;
  setCommsOpenId: React.Dispatch<React.SetStateAction<number>>;
}

export function EventCommsModal(props: EventCommsModalProps): React.ReactElement {
  const { commsLoading, commsModal, commsOpenId, commsRows, isDe, setCommsModal, setCommsOpenId } = props;
  return (
        <Modal
          open={true}
          onClose={() => setCommsModal(null)}
          maxWidth={640}
          ariaLabel={isDe ? 'Nachrichten zum Event' : 'Event messages'}
        >
          <h3 style={{ margin: 0, fontSize: '1.05rem' }}>
            {isDe ? 'Nachrichten zum Event' : 'Event messages'}
          </h3>
          <p style={{ margin: 0, fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>
            {commsModal.eventTitle}
          </p>
          {commsLoading ? (
            <p style={{ margin: '8px 0', fontSize: '0.9rem', color: 'var(--dex-gray-600)' }}>
              {isDe ? 'Wird geladen…' : 'Loading…'}
            </p>
          ) : commsRows.length === 0 ? (
            <p style={{ margin: '8px 0', fontSize: '0.9rem', color: 'var(--dex-gray-600)' }}>
              {isDe ? 'Zu diesem Event gibt es noch keine Nachrichten.' : 'There are no messages for this event yet.'}
            </p>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8, maxHeight: '60vh', overflowY: 'auto' }}>
              {commsRows.map(row => {
                const isOpen = commsOpenId === row.id;
                return (
                  <div
                    key={row.id}
                    style={{
                      border: '1px solid var(--dex-gray-200)', borderRadius: 10,
                      background: 'var(--dex-gray-50, #fafafa)', overflow: 'hidden',
                    }}
                  >
                    <button
                      type="button"
                      onClick={() => setCommsOpenId(prev => (prev === row.id ? null : row.id))}
                      aria-expanded={isOpen}
                      style={{
                        display: 'block', width: '100%', textAlign: 'left',
                        background: 'none', border: 'none', cursor: 'pointer', padding: '10px 14px',
                      }}
                    >
                      <div style={{ fontWeight: 700, fontSize: '0.9rem', color: 'var(--dex-gray-800)', wordBreak: 'break-word' }}>
                        {row.subject || (isDe ? '(ohne Betreff)' : '(no subject)')}
                      </div>
                      <div style={{ marginTop: 2, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                        {formatDate(row.created)} · {isDe ? 'von' : 'from'} {row.sentByName || row.sentByEmail || '—'}
                      </div>
                    </button>
                    {isOpen && (
                      <div style={{ borderTop: '1px solid var(--dex-gray-200)', background: '#fff' }}>
                        <iframe
                          title={row.subject || 'message'}
                          srcDoc={row.bodyHtml || ''}
                          sandbox=""
                          style={{ width: '100%', height: 360, border: 'none', display: 'block' }}
                        />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
          <div style={{ textAlign: 'right' }}>
            <button className="btn btn-secondary" onClick={() => setCommsModal(null)} style={{ fontSize: '0.85rem' }}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          </div>
        </Modal>
  );
}
