/* MyEventsModals — aus MyEventsPage.tsx ausgelagert (Zeilen 2975-3589 des
 * urspruenglichen Stands, v30.65). Die fuenf Dialoge der Seite: Mitglied zum
 * Team hinzufuegen, Team verwalten, Cascade-Abmeldung, eigener QR-Code und das
 * Kommunikations-Log. Das JSX ist zeichengleich uebernommen; die
 * Anzeige-Bedingung (`addMemberDialog && …`) bleibt beim Aufrufer, die
 * Komponente rendert ihren Inhalt unbedingt.
 */
import * as React from 'react';
import Modal from '../Modal';
import InternationalSearchToggle from '../InternationalSearchToggle';
import { SPRegistration, EventCommRow } from '../../services/EventService';
import { formatDate, formatDateTimeRange } from './myEventsHelpers';
import { groupSubEventTabs, stripGroupPrefix } from '../../utils/subEventGroups';
// v31.2: Gemeinsame UI-Klassen (Zeilen, Pills, Kacheln, Hinweiskästen) und die
// Inline-SVG-Symbole — Kopf/Fuß der Dialoge kommen aus den Modal-Props.
import { cx } from '../dexUi';
import { AlertCircle, ChevronDown, Info, Mail, QrCode, Trash2, Users, X } from '../Icons';


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
  // v31.2: Reihenfolge Pflicht → Warnung → Bestätigung. Vorher stand ein
  // achtzeiliger oranger Kasten ÜBER dem Suchfeld, und der eigentliche
  // Pflichtschritt lag unter dem Falz. Was die Person sofort bekommt, steht
  // jetzt als eine Zeile im Untertitel; der Kasten trägt nur noch die Warnung
  // und sitzt direkt über dem Haken, zu dem er gehört.
  return (
        <Modal
          open={true}
          onClose={closeAddMemberDialog}
          dismissable={!addMemberBusy}
          maxWidth={540}
          ariaLabel={isDe ? 'Mitglied zum Team hinzufügen' : 'Add member to team'}
          title={isDe ? 'Mitglied zum Team hinzufügen' : 'Add member to team'}
          subtitle={isDe
            ? 'Die Person wird sofort und ohne Rückfrage angemeldet — sie bekommt eine Bestätigungs-Mail, einen Outlook-Termin und den Event in „Meine Events".'
            : 'The person is registered immediately, without further confirmation — they get a confirmation email, an Outlook invite and the event in „My Events".'}
          icon={<Users size={20} />}
          footer={<>
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
          </>}
        >
            {/* People-Picker — simple Inline-Variante mit der searchUsers-API. */}
            <div className="dex-ui-field">
              <label className="dex-ui-label">
                <span style={{ color: 'var(--dex-red)' }}>*</span>
                {isDe ? 'Wen möchtest du ins Team holen?' : 'Who do you want to add to the team?'}
              </label>
              {addMemberPick ? (
                <div className="dex-ui-card dex-ui-card--soft" style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px' }}>
                  <img
                    className="dex-ui-avatar"
                    src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(addMemberPick.email)}&size=S`}
                    alt={addMemberPick.displayName}
                    onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                  />
                  <div className="dex-ui-row-main">
                    <div className="dex-ui-row-title">{addMemberPick.displayName}</div>
                    <div className="dex-ui-row-sub">{addMemberPick.email}</div>
                  </div>
                  <button
                    type="button"
                    className="dex-ui-iconbtn"
                    onClick={() => { setAddMemberPick(null); setAddMemberQuery(''); setAddMemberResults([]); }}
                    title={isDe ? 'Auswahl entfernen' : 'Remove selection'}
                    aria-label={isDe ? 'Auswahl entfernen' : 'Remove selection'}
                  >
                    <X size={16} />
                  </button>
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
                    <div className="dex-ui-card" style={{ position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 10, marginTop: 4, padding: 4, maxHeight: 220, overflowY: 'auto', boxShadow: '0 6px 20px rgba(0,0,0,0.10)' }}>
                      {addMemberSearching && (
                        <div className="dex-ui-muted" style={{ padding: '8px 10px' }}>
                          {isDe ? 'Suche…' : 'Searching…'}
                        </div>
                      )}
                      {addMemberResults.map(r => (
                        <button
                          key={r.email}
                          type="button"
                          className="dex-ui-row"
                          onClick={() => { setAddMemberPick(r); setAddMemberResults([]); setAddMemberQuery(''); }}
                          style={{ width: '100%', border: 'none', background: 'transparent', cursor: 'pointer', textAlign: 'left', font: 'inherit', padding: '6px 10px' }}
                        >
                          <img
                            className="dex-ui-avatar"
                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(r.email)}&size=S`}
                            alt={r.displayName}
                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                          />
                          <div className="dex-ui-row-main">
                            <div className="dex-ui-row-title">{r.displayName}</div>
                            <div className="dex-ui-row-sub">{r.email}</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>
            {/* Pflicht-Hinweis — orange, analog zur Initial-Team-Anmeldung. */}
            <div className="dex-ui-callout dex-ui-callout--warn">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span>
                {isDe
                  ? <><strong>Vorab Zustimmung einholen.</strong> Stell sicher, dass die Person ihrer Anmeldung <strong>vorher zugestimmt</strong> hat — es gibt keine weitere Rückfrage.</>
                  : <><strong>Get consent up front.</strong> Make sure the person has <strong>consented</strong> to this registration — there is no further confirmation.</>}
              </span>
            </div>
            <label className={cx('dex-ui-toggle-row', addMemberConsent && 'is-active')}>
              <input
                type="checkbox"
                checked={addMemberConsent}
                onChange={e => setAddMemberConsent(e.target.checked)}
              />
              <span className="dex-ui-toggle-row-body">
                <span className="dex-ui-toggle-row-title">
                  <span style={{ color: 'var(--dex-red)' }}>*</span>
                  {isDe
                    ? 'Ich bestätige, dass die ausgewählte Person ihrer Anmeldung zugestimmt hat.'
                    : 'I confirm that the selected person has consented to this registration.'}
                </span>
                <span className="dex-ui-toggle-row-desc">
                  {isDe ? <>Ohne diesen Haken bleibt &bdquo;Hinzufügen&ldquo; gesperrt.</> : <>Without this, &bdquo;Add&ldquo; stays disabled.</>}
                </span>
              </span>
            </label>
            {addMemberError && (
              <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>{addMemberError}</span>
              </div>
            )}
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
        // v31.2: Kopf/Fuß über die Modal-Props, Mitglieder als Zeilen mit Hover,
        // Status als Pill, Entfernen als runder Symbol-Knopf (rot erst beim
        // Überfahren) — der rote Vollflächen-Knopf je Zeile las sich vorher wie
        // eine Warnung an jedem Mitglied.
        return (
          <Modal
            open={true}
            onClose={closeManageTeamDialog}
            dismissable={manageTeamBusyId === null}
            maxWidth={620}
            ariaLabel={isDe ? 'Team verwalten' : 'Manage team'}
            title={isDe
              ? <>Team &bdquo;{manageTeamDialog.teamName || 'Unbenannt'}&ldquo; verwalten</>
              : <>Manage team &bdquo;{manageTeamDialog.teamName || 'Unnamed'}&ldquo;</>}
            subtitle={isDe
              ? `${activeMembers.length} von ${manageTeamDialog.teamSize} Plätzen belegt`
              : `${activeMembers.length} of ${manageTeamDialog.teamSize} seats taken`}
            icon={<Users size={20} />}
            footer={
              <button
                type="button"
                className="btn btn-secondary"
                onClick={closeManageTeamDialog}
                disabled={manageTeamBusyId !== null}
              >
                {isDe ? 'Schließen' : 'Close'}
              </button>
            }
          >
              <div className="dex-ui-card" style={{ padding: '4px 6px' }}>
                {sortedAll.map(m => {
                  const isCancelled = m.Status === 'Abgemeldet';
                  const isMemberLead = !!m.TeamLead && !isCancelled;
                  const isSelf = (m.ParticipantEmail || '').toLowerCase() === (currentUserEmail || '').toLowerCase();
                  const fullName = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantEmail;
                  const loc = (m.Location || '').trim();
                  const busy = manageTeamBusyId === m.Id;
                  return (
                    <div key={m.Id} className="dex-ui-row dex-ui-row--bordered" style={{ opacity: isCancelled ? 0.55 : 1 }}>
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
                      <div className="dex-ui-row-main">
                        <div className="dex-ui-row-title">{fullName}</div>
                        <div className="dex-ui-row-sub">{m.ParticipantEmail}{loc ? ` · ${loc}` : ''}</div>
                      </div>
                      {isCancelled ? (
                        <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'abgemeldet' : 'cancelled'}</span>
                      ) : isMemberLead ? (
                        <span className="dex-ui-pill dex-ui-pill--green">Lead</span>
                      ) : null}
                      {/* Trash-Button — nur für aktive Nicht-Lead-Nicht-Self-Mitglieder. */}
                      {!isCancelled && !isMemberLead && !isSelf && (
                        <div className="dex-ui-row-actions">
                          <button
                            type="button"
                            className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                            onClick={() => setManageTeamConfirm(m)}
                            disabled={busy || manageTeamBusyId !== null}
                            title={isDe ? 'Diese Person aus dem Team abmelden' : 'Remove this person from the team'}
                            aria-label={isDe ? 'Diese Person aus dem Team abmelden' : 'Remove this person from the team'}
                          >
                            <Trash2 size={16} />
                          </button>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
              <div className="dex-ui-callout dex-ui-callout--neutral">
                <span className="dex-ui-callout-icon"><Info size={16} /></span>
                <span>
                  {isDe
                    ? <>Dich selbst meldest du als Team-Lead über <strong>&bdquo;Abmelden&ldquo;</strong> auf deiner Event-Karte ab — die App übergibt den Lead dann automatisch an das früheste verbleibende Team-Mitglied.</>
                    : <>As team lead you cancel yourself via <strong>&bdquo;Cancel&ldquo;</strong> on your event card — the app then hands the lead role to the earliest remaining team member.</>}
                </span>
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
                  title={isDe ? 'Person aus dem Team abmelden?' : 'Cancel this team member?'}
                  subtitle={isDe
                    ? <><strong>{fullName}</strong> wird vom Event abgemeldet und bekommt eine Abmelde-Bestätigung per Mail — bei aktivem Outlook auch eine Termin-Absage.</>
                    : <><strong>{fullName}</strong> will be cancelled from the event and receives a cancellation email — plus a calendar removal if Outlook is active.</>}
                  icon={<Trash2 size={18} />}
                  footer={<>
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
                      className="btn btn-danger"
                      onClick={() => performManageTeamCancel(cm)}
                      disabled={busy}
                    >
                      {busy
                        ? (isDe ? 'Wird abgemeldet…' : 'Cancelling…')
                        : (isDe ? 'Person abmelden' : 'Cancel member')}
                    </button>
                  </>}
                >
                    <div className="dex-ui-callout dex-ui-callout--neutral">
                      <span className="dex-ui-callout-icon"><Info size={16} /></span>
                      <span>
                        {isDe
                          ? <>Die Person könnte sich auch selbst über &bdquo;Meine Events&ldquo; abmelden — du machst das nur stellvertretend.</>
                          : <>This person could also cancel themselves via &bdquo;My Events&ldquo; — you are doing it on their behalf.</>}
                      </span>
                    </div>
                </Modal>
              );
            })()}
          </Modal>
        );
}


export interface CascadeCancelModalProps {
  cascadeDialog: { parentTitle: string; subEvents: { id: string; title: string; startDate?: string; endDate?: string; location?: string; }[]; resolve: (_choice: "cascade" | "parent-only" | "abort") => void; isSectionedEvent?: boolean; };
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
        // v31.2: Die Bezeichnung einmal bestimmen — sie steht im Untertitel und
        // in beiden Kacheln. „Sub-Event" bleibt hier fest, weil die Seite kein
        // childTermSingular/Plural an diesen Dialog durchreicht (Props-Vertrag).
        const n = dlg.subEvents.length;
        const unit = isSec
          ? (isDe ? (n === 1 ? 'Event-Section' : 'Event-Sections') : (n === 1 ? 'event section' : 'event sections'))
          : (isDe ? (n === 1 ? 'Sub-Event' : 'Sub-Events') : (n === 1 ? 'sub-event' : 'sub-events'));
        return (
          <Modal
            open={true}
            onClose={() => choose('abort')}
            maxWidth={520}
            ariaLabel={title}
            title={title}
            subtitle={isDe
              ? <>Du bist für <strong>{n}</strong> {unit} von <strong>&bdquo;{dlg.parentTitle}&ldquo;</strong> angemeldet.</>
              : <>You are registered for <strong>{n}</strong> {unit} of <strong>&bdquo;{dlg.parentTitle}&ldquo;</strong>.</>}
            icon={<AlertCircle size={20} />}
            footer={
              <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={() => choose('abort')}>
                {isDe ? 'Abbrechen — nichts abmelden' : 'Cancel — keep everything'}
              </button>
            }
          >
              <div className="dex-ui-card dex-ui-card--soft" style={{ padding: '10px 14px', fontSize: '0.85rem', color: 'var(--dex-gray-700)', maxHeight: 220, overflowY: 'auto' }}>
                {/* v15.8: pro Sub-Event Titel + Datum + Ort listen, damit
                    der User auf einen Blick sieht was er da abmeldet.
                    v30.79: von–bis statt nur Start, und bei vielen Terminen
                    dieselbe Präfix-Gruppierung wie in der Termin-Liste
                    („Day 1", „Day 2" …) als Zwischenüberschriften. */}
                {(() => {
                  const renderItem = (s: { id: string; title: string; startDate?: string; endDate?: string; location?: string }, shownTitle: string): React.ReactElement => {
                    const subParts = [formatDateTimeRange(s.startDate, s.endDate, isDe), s.location].filter(Boolean).join(' · ');
                    return (
                      <li key={s.id} style={{ marginBottom: 4 }}>
                        <div style={{ fontWeight: 600 }}>{shownTitle}</div>
                        {subParts && (
                          <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>{subParts}</div>
                        )}
                      </li>
                    );
                  };
                  const grouping = groupSubEventTabs(dlg.subEvents.map(s => s.title));
                  if (!grouping.grouped) {
                    return (
                      <ul style={{ margin: 0, paddingLeft: 20 }}>
                        {dlg.subEvents.map(s => renderItem(s, s.title))}
                      </ul>
                    );
                  }
                  return grouping.groups.map(g => {
                    const members = g.idxs.map(i => dlg.subEvents[i]).filter(Boolean);
                    const label = g.label === 'Weitere' ? (isDe ? 'Weitere' : 'Other') : g.label;
                    return (
                      <div key={g.label} style={{ marginBottom: 6 }}>
                        <div style={{ display: 'flex', alignItems: 'baseline', gap: 8, padding: '4px 0 2px', borderBottom: '1px solid var(--dex-gray-200)', marginBottom: 4 }}>
                          <span style={{ fontWeight: 800, fontSize: '0.86rem', color: 'var(--dex-gray-800)' }}>{label}</span>
                          <span style={{ fontSize: '0.74rem', color: 'var(--dex-gray-600)' }}>{members.length} {isDe ? (members.length === 1 ? 'Termin' : 'Termine') : (members.length === 1 ? 'date' : 'dates')}</span>
                        </div>
                        <ul style={{ margin: 0, paddingLeft: 20 }}>
                          {members.map(s => renderItem(s, stripGroupPrefix(s.title, g.label)))}
                        </ul>
                      </div>
                    );
                  });
                })()}
              </div>
              {isSec && (
                <div className="dex-ui-callout dex-ui-callout--info">
                  <span className="dex-ui-callout-icon"><Info size={16} /></span>
                  <span>
                    {isDe
                      ? <>Nur <strong>einzelne Sections</strong> abwählen (z.B. &bdquo;Dinner absagen, Meeting behalten&ldquo;)? Dann brich hier ab und nutze <strong>&bdquo;Anmeldung bearbeiten&ldquo;</strong> auf der Event-Karte — dort wählst du Sections gezielt an oder ab, ohne die ganze Teilnahme zu stornieren.</>
                      : <>Want to drop only <strong>individual sections</strong> (e.g. &bdquo;cancel dinner, keep meeting&ldquo;)? Then abort here and use <strong>&bdquo;Edit registration&ldquo;</strong> on the event card — there you pick or deselect sections without cancelling the whole registration.</>}
                  </span>
                </div>
              )}
              {/* v31.2: Die Alternativen als Kacheln mit je einer Zeile Folge statt
                  eines Knopfstapels — der Teilnehmer sieht vor dem Klick, was
                  danach noch steht und was nicht. Die Klammer-Abmeldung selbst
                  entscheidet weiter der Aufrufer über `resolve`. */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Was möchtest du abmelden?' : 'What do you want to cancel?'}</div>
                <div className="dex-ui-stack" style={{ gap: 8 }}>
                  <button type="button" className="dex-ui-choice" onClick={() => choose('cascade')}>
                    <div className="dex-ui-choice-body">
                      <div className="dex-ui-choice-title">
                        {isSec
                          ? (isDe ? 'Ja, komplett abmelden' : 'Yes, cancel entirely')
                          : (isDe ? 'Alles abmelden' : 'Cancel everything')}
                      </div>
                      <div className="dex-ui-choice-desc">
                        {isSec
                          ? (isDe ? `Event + ${n} ${unit} — du bist danach nirgends mehr angemeldet.` : `Event + ${n} ${unit} — you will no longer be registered anywhere.`)
                          : (isDe ? `Hauptevent + ${n} ${unit} — du bist danach nirgends mehr angemeldet.` : `Main event + ${n} ${unit} — you will no longer be registered anywhere.`)}
                      </div>
                    </div>
                  </button>
                  {!isSec && (
                    <button type="button" className="dex-ui-choice" onClick={() => choose('parent-only')}>
                      <div className="dex-ui-choice-body">
                        <div className="dex-ui-choice-title">{isDe ? 'Nur Hauptevent abmelden' : 'Cancel main event only'}</div>
                        <div className="dex-ui-choice-desc">
                          {isDe ? `Deine ${n} ${unit} bleiben bestehen.` : `Your ${n} ${unit} stay registered.`}
                        </div>
                      </div>
                    </button>
                  )}
                </div>
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
  // v31.2: Kopf/Fuß über die Modal-Props; die Teilnehmer-Nummer als Pill,
  // weil sie seit v30.33 der Weg ist, wenn der Scan nicht klappt — sie muss
  // am Einlass auf einen Blick ablesbar sein.
  return (
        <Modal
          open={true}
          onClose={() => setMyQrModal(null)}
          maxWidth={420}
          ariaLabel={isDe ? 'Mein QR-Code' : 'My QR code'}
          title={isDe ? 'Mein Check-in-QR-Code' : 'My check-in QR code'}
          subtitle={myQrModal.eventTitle}
          icon={<QrCode size={20} />}
          footer={
            <button type="button" className="btn btn-secondary" onClick={() => setMyQrModal(null)}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          }
        >
          <div className="dex-ui-card dex-ui-card--soft" style={{ textAlign: 'center', padding: 16 }}>
            <img
              src={myQrModal.dataUrl}
              alt="QR-Code"
              style={{ width: 260, maxWidth: '90%', height: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: 10, background: '#fff', display: 'inline-block' }}
            />
            <div style={{ marginTop: 12, fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-gray-800)' }}>
              {myQrModal.name}
            </div>
            {myQrModal.tid ? (
              <div style={{ marginTop: 6 }}>
                <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'Nr.' : 'No.'} {myQrModal.tid}</span>
              </div>
            ) : null}
          </div>
          <p className="dex-ui-muted" style={{ margin: 0, textAlign: 'center', lineHeight: 1.5 }}>
            {isDe
              ? 'Zeig diesen Code am Event-Tag dem Check-in-Team — er ist derselbe wie in deiner QR-Mail. Klappt der Scan nicht, reichen dein Name oder deine Nummer.'
              : 'Show this code to the check-in team on event day — it is the same as in your QR email. If the scan fails, your name or number is enough.'}
          </p>
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
  // v31.2: Kopf und Fuß über die Modal-Props; jede Nachricht als Zeile mit
  // Hover und Pfeil — vorher hatte der Betreff-Knopf keinen Hover und las sich
  // als Überschrift, nicht als etwas, das man aufklappen kann.
  return (
        <Modal
          open={true}
          onClose={() => setCommsModal(null)}
          maxWidth={640}
          ariaLabel={isDe ? 'Nachrichten zum Event' : 'Event messages'}
          title={isDe ? 'Nachrichten zum Event' : 'Event messages'}
          subtitle={commsModal.eventTitle}
          icon={<Mail size={20} />}
          footer={
            <button type="button" className="btn btn-secondary" onClick={() => setCommsModal(null)}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          }
        >
          {commsLoading ? (
            <p className="dex-ui-muted" style={{ margin: 0 }}>
              {isDe ? 'Wird geladen…' : 'Loading…'}
            </p>
          ) : commsRows.length === 0 ? (
            <div className="dex-ui-empty">
              <span className="dex-ui-empty-icon"><Mail size={20} /></span>
              <div className="dex-ui-empty-title">{isDe ? 'Noch keine Nachrichten' : 'No messages yet'}</div>
              {isDe ? 'Zu diesem Event gibt es noch keine Nachrichten.' : 'There are no messages for this event yet.'}
            </div>
          ) : (
            <div className="dex-ui-stack" style={{ gap: 8, maxHeight: '60vh', overflowY: 'auto' }}>
              {commsRows.map(row => {
                const isOpen = commsOpenId === row.id;
                return (
                  <div key={row.id} className="dex-ui-card" style={{ padding: 0, overflow: 'hidden' }}>
                    <button
                      type="button"
                      className="dex-ui-row"
                      onClick={() => setCommsOpenId(prev => (prev === row.id ? null : row.id))}
                      aria-expanded={isOpen}
                      style={{ width: '100%', border: 'none', background: 'transparent', cursor: 'pointer', textAlign: 'left', font: 'inherit', padding: '10px 14px', borderRadius: 0 }}
                    >
                      <div className="dex-ui-row-main">
                        <div className="dex-ui-row-title" style={{ whiteSpace: 'normal', wordBreak: 'break-word' }}>
                          {row.subject || (isDe ? '(ohne Betreff)' : '(no subject)')}
                        </div>
                        <div className="dex-ui-row-sub">
                          {formatDate(row.created)} · {isDe ? 'von' : 'from'} {row.sentByName || row.sentByEmail || '—'}
                        </div>
                      </div>
                      <span className="dex-ui-disclosure-chevron" aria-hidden="true" style={{ transform: isOpen ? 'rotate(180deg)' : 'none' }}>
                        <ChevronDown size={16} />
                      </span>
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
        </Modal>
  );
}
