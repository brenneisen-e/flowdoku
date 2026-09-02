/* MyEventCard — aus MyEventsPage.tsx ausgelagert (Zeilen 1737-2832 des
 * urspruenglichen Stands, v30.65; mit dem Kopf der frueheren map-Callback-
 * Funktion `renderMyEventCard`, Zeilen 1668-1736). Eine Karte in „Meine
 * Events": Kopfzeile, Antworten, Dokumente, Quiz, Team, Sub-Events und die
 * Abmelde-Wege. Der Rumpf ist zeichengleich uebernommen; die 49 Werte, die er
 * aus dem Seiten-Scope gelesen hat, kommen jetzt als Props herein.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import OrganizerList from '../OrganizerList';
import { CachedImg } from '../CachedImage';
import { UserFieldPicker } from '../UserFieldPicker';
import { isEventVisibleForUser } from '../EventListPage';
import { DeloitteEvent, EventSpecificField, AgendaItem, TransferTime } from '../../types';
import { SPRegistration } from '../../services/EventService';
import { isEventOver, formatAllDayPeriod } from '../../utils/eventFormat';
import { selfCancelLocked, selfCancelLockReason } from '../../utils/cancelPolicy';
import { X, Pencil, QrCode, Mail } from '../Icons';
import { TeamsJoinButton } from '../TeamsJoinButton';
import { eventTeamsLink, locationWithoutTeamsUrl } from '../../utils/teamsLink';
import DocumentsViewer from './DocumentsViewer';
import QuizPlayer from './QuizPlayer';
import MyEventSubEvents from './MyEventSubEvents';
import MyEventUpload from './MyEventUpload';
import MyEventDocField from './MyEventDocField';
import { FieldAnswerTag, MyEventEntry, formatDate, formatDateRange, getStatusBadgeClass, getStatusLabel } from './myEventsHelpers';

export interface MyEventCardProps {
  /** Der Eintrag, den die Karte zeigt — frueher das destrukturierte Argument
   *  von `renderMyEventCard`. */
  entry: MyEventEntry;
  cancellingId: string;
  cancelRegistration: (eventId: string, opts?: { suppressNotifications?: boolean; skipReload?: boolean; }) => Promise<boolean>;
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  deleteFieldDocument: (eventId: string, fileName: string, participantEmail?: string) => Promise<boolean>;
  deleteMyEventAttachment: (eventId: string, fileName: string) => Promise<boolean>;
  descExpanded: Record<string, boolean>;
  editData: Record<string, string>;
  editingId: string;
  enqueueJoinReqFetch: (eventId: string, teamId: string) => void;
  enqueueTeamFetch: (eventId: string, teamId: string) => void;
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
  getMyRegistration: (eventId: string) => Promise<SPRegistration>;
  handleCancel: (eventId: string) => Promise<void>;
  handleDecideJoinRequest: (eventId: string, teamId: string, requestId: number, decision: 'Approved' | 'Rejected') => Promise<void>;
  isCancelling: boolean;
  isDe: boolean;
  isSaving: boolean;
  joinReqBusyId: number;
  joinRequestsCache: Record<string, { Id: number; RequesterEmail: string; RequesterDisplayName: string; Status: string; Created: string; }[]>;
  listFieldDocuments: (eventId: string, fieldId: string, participantEmail?: string) => Promise<{ fileName: string; serverRelativeUrl: string; displayName: string; }[]>;
  listMyEventAttachments: (eventId: string) => Promise<{ fileName: string; serverRelativeUrl: string; }[]>;
  loadMyRegistrations: (silent?: boolean) => Promise<void>;
  locale: import("../../context/LanguageContext").Locale;
  openComms: (ev: DeloitteEvent) => void;
  openManageTeamDialog: (eventId: string, teamId: string, teamName: string, teamSize: number, members: SPRegistration[]) => void;
  openMyQr: (ev: DeloitteEvent, reg: SPRegistration) => Promise<void>;
  registerForEvent: (eventId: string, customData: Record<string, string>, participantFirstName?: string, participantLastName?: string, participantEmail?: string, preferredStarterType?: string, opts?: { skipShadowParent?: boolean; suppressMail?: boolean; suppressOutlook?: boolean; extraCc?: string; proxyConsentConfirmed?: boolean; actorAllowedAsAssistant?: boolean; skipReload?: boolean; bundledItems?: import("../../utils/bundledComm").BundledItem[]; }) => Promise<{ ok: boolean; status: "Angemeldet" | "Warteliste"; reason?: string; }>;
  searchUser: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string; company?: string; }>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  setAddMemberConsent: React.Dispatch<React.SetStateAction<boolean>>;
  setAddMemberDialog: React.Dispatch<React.SetStateAction<{ eventId: string; teamId: string; teamName: string; freeSlots: number; }>>;
  setAddMemberError: React.Dispatch<React.SetStateAction<string>>;
  setAddMemberPick: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; }>>;
  setAddMemberQuery: React.Dispatch<React.SetStateAction<string>>;
  setAddMemberResults: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; }[]>>;
  setCancellingId: React.Dispatch<React.SetStateAction<string>>;
  setDescExpanded: React.Dispatch<React.SetStateAction<Record<string, boolean>>>;
  setEditData: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  setEditingId: React.Dispatch<React.SetStateAction<string>>;
  setIsSaving: React.Dispatch<React.SetStateAction<boolean>>;
  setMyEvents: React.Dispatch<React.SetStateAction<MyEventEntry[]>>;
  showAlert: (message: React.ReactNode, opts?: import("../../context/DialogContext").AlertOptions) => void;
  switchSplitGroup: (eventId: string, newType: "Durchstarter" | "Funstarter") => Promise<{ ok: boolean; status: "Angemeldet" | "Warteliste" | "Failed"; full: boolean; }>;
  t: (key: string) => string;
  teamMembersCache: Record<string, SPRegistration[]>;
  updateMyRegistration: (eventId: string, customData: Record<string, string>) => Promise<boolean>;
  uploadFieldDocument: (eventId: string, fieldId: string, file: File, participantEmail?: string) => Promise<boolean>;
  uploadMyEventAttachment: (eventId: string, file: File) => Promise<boolean>;
}

export default function MyEventCard(props: MyEventCardProps): React.ReactElement | null {
  const { cancellingId, cancelRegistration, childEventsOf, confirmDialog, deleteFieldDocument, deleteMyEventAttachment, descExpanded, editData, editingId, enqueueJoinReqFetch, enqueueTeamFetch, getAllRegistrations, getMyRegistration, handleCancel, handleDecideJoinRequest, isCancelling, isDe, isSaving, joinReqBusyId, joinRequestsCache, listFieldDocuments, listMyEventAttachments, loadMyRegistrations, locale, openComms, openManageTeamDialog, openMyQr, registerForEvent, searchUser, searchUsers, setAddMemberConsent, setAddMemberDialog, setAddMemberError, setAddMemberPick, setAddMemberQuery, setAddMemberResults, setCancellingId, setDescExpanded, setEditData, setEditingId, setIsSaving, setMyEvents, showAlert, switchSplitGroup, t, teamMembersCache, updateMyRegistration, uploadFieldDocument, uploadMyEventAttachment } = props;
  const { event, registration, sessionsOnly, subEventTitles, hiddenRow } = props.entry;
            // Custom Data parsen und IDs zu Labels mappen
            let customData: Record<string, string> = {};
            try {
              if (registration.CustomData) customData = JSON.parse(registration.CustomData);
            } catch { /* */ }

            // Feld-ID zu Label-Map aus den Event-Feldern erstellen.
            // v17.22: im Bilingual-Modus + EN-Locale die EN-Variante des
            // Labels nehmen — vorher rein DE, obwohl der Teilnehmer sich
            // bei der Anmeldung die EN-Labels angesehen hatte.
            const useEnDisplay = locale === 'en' && !!event.bilingualFields;
            const fieldLabelMap: Record<string, string> = {};
            // v17.22: Wert-Übersetzung für Select-Optionen (DE-Wert →
            // EN-Anzeige) pro Feld, damit auch die Antwort selbst zweisprachig
            // erscheint. Map: fieldId → (DE-Option → EN-Option).
            const fieldOptionEnMap: Record<string, Record<string, string>> = {};
            // v19.34: Feldtyp pro ID merken, damit People-Picker-Antworten
            // (`user`/`roommate`) als Foto-Tag gerendert werden können.
            const fieldTypeMap: Record<string, string> = {};
            for (const field of event.eventSpecificFields) {
              fieldTypeMap[field.id] = field.type;
              fieldLabelMap[field.id] = (useEnDisplay && field.labelEn && field.labelEn.trim())
                ? field.labelEn
                : field.label;
              if (useEnDisplay && field.options && field.optionsEn) {
                const m: Record<string, string> = {};
                field.options.forEach((opt, i) => {
                  const en = (field.optionsEn || [])[i];
                  if (en && en.trim()) m[opt] = en;
                });
                fieldOptionEnMap[field.id] = m;
              }
            }
            // Fallback-Labels für ad-hoc Keys, die NICHT als EventSpecificField
            // registriert sind (z.B. Leistungsnachweis, der direkt aus der Starter-
            // Typ-Sektion kommt). Sonst würde der technische Key angezeigt.
            const adHocLabels: Record<string, string> = {
              b2run_leistungsnachweis: t('reg.starter.proof') || 'Leistungsnachweis vorhanden',
            };

            // "salutation" überspringen (wird schon im Namen angezeigt)
            // Boolean-Werte ('true'/'false') zu lesbarem Ja/Nein konvertieren,
            // damit das UI nicht "true" als technischen String zeigt.
            const yesLabel = t('general.yes') || 'Ja';
            const noLabel = t('general.no') || 'Nein';
            const displayData = Object.keys(customData)
              .filter(key => key !== 'salutation' && customData[key])
              .map(key => {
                const raw = customData[key];
                let value: string = raw;
                if (raw === 'true') value = yesLabel;
                else if (raw === 'false') value = noLabel;
                else if (fieldOptionEnMap[key]) {
                  // v17.22: Select-Antworten im EN-Modus übersetzen. Multi-
                  // Select-Werte sind " | "-getrennt — jeden Teil einzeln mappen.
                  value = raw.split(' | ').map(part => {
                    const trimmed = part.trim();
                    return fieldOptionEnMap[key][trimmed] || trimmed;
                  }).join(' | ');
                }
                return {
                  label: fieldLabelMap[key] || adHocLabels[key] || key,
                  value,
                  type: fieldTypeMap[key],
                };
              });

            return (
              <div key={event.id} id={`dex-myevent-${event.id}`} className="card my-event-card">
                {/* Header-Zeile: Thumbnail links + Titel/Details rechts */}
                <div style={{ display: 'flex', gap: 16, alignItems: 'flex-start' }}>
                  {event.imageUrl && (
                    <div
                      className="my-event-card__thumb"
                      style={{
                        flexShrink: 0,
                        width: 140,
                        height: 100,
                        borderRadius: 'var(--dex-radius, 12px)',
                        background: 'var(--dex-gray-50, #fafafa)',
                        border: '1px solid var(--dex-gray-200)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        overflow: 'hidden',
                      }}
                    >
                      <CachedImg
                        src={event.imageUrl}
                        alt={event.title}
                        loading="lazy"
                        decoding="async"
                        style={{
                          maxWidth: '100%',
                          maxHeight: '100%',
                          objectFit: 'contain',
                          display: 'block',
                        }}
                      />
                    </div>
                  )}
                  <div style={{ flex: 1, minWidth: 0 }}>
                    {/* Header: Titel + Status Badge */}
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                      <h3 style={{ margin: 0, fontSize: '1.1rem' }}>{event.title}</h3>
                      {sessionsOnly ? (() => {
                        // v15.15: Im subEventsOnlyMode komplett ausblenden —
                        // Badge UND Hinweisbox sind dort redundant, weil
                        // „nur Sub-Events"-Anmeldung der einzig mögliche
                        // Zustand ist (kein Hauptevent zum Anmelden).
                        if (event.subEventsOnlyMode) return null;
                        const term = event.childEventTermPlural || '';
                        const badgeText = term
                          ? (isDe ? `Nur ${term}` : `${term} only`)
                          : (t('myevents.sessionsonly.badge') || 'Nur Sub-Events');
                        return (
                          <span
                            className="badge"
                            title={t('myevents.sessionsonly.hint')}
                            style={{
                              flexShrink: 0, marginLeft: 12,
                              background: 'var(--dex-orange, #ed8b00)', color: '#fff',
                            }}
                          >
                            {badgeText}
                          </span>
                        );
                      })() : (
                        <span className={`badge ${getStatusBadgeClass(registration.Status)}`} style={{ flexShrink: 0, marginLeft: 12 }}>
                          {registration.Status === 'Warteliste' && registration.TeilnehmerID && event.maxParticipants > 0
                            ? `${getStatusLabel(registration.Status, t)} #${registration.TeilnehmerID - event.maxParticipants}`
                            : getStatusLabel(registration.Status, t)}
                        </span>
                      )}
                    </div>
                    {/* v28.23: Stellvertretend angelegte Anmeldung, deren Zeile
                        für die Person (noch) nicht lesbar ist. Sie SIEHT die
                        Anmeldung jetzt — inkl. Status —, kann sie aber nicht
                        selbst bearbeiten oder stornieren. */}
                    {hiddenRow && (
                      <div style={{
                        marginTop: 6, padding: '8px 10px', borderRadius: 6,
                        background: 'rgba(0,118,168,0.07)', border: '1px solid rgba(0,118,168,0.35)',
                        color: 'var(--dex-gray-700)', fontSize: '0.78rem', lineHeight: 1.5,
                      }}>
                        {isDe
                          ? <>Diese Anmeldung wurde <strong>für dich angelegt</strong> (z.B. durch deine Assistenz oder die Organizer). Deine Anmeldung ist gültig — die Detailangaben und das Abmelden liegen aber bei der Person, die dich angemeldet hat. Bitte wende dich für Änderungen an sie oder an die Organizer. <strong>Melde dich nicht erneut an</strong>, sonst entsteht eine doppelte Anmeldung.</>
                          : <>This registration was <strong>created for you</strong> (e.g. by your assistant or the organizers). Your registration is valid — the details and cancellation stay with whoever registered you. Please contact them or the organizers for changes. <strong>Do not register again</strong>, that would create a duplicate.</>}
                      </div>
                    )}
                    {/* v28.39: Hotel-Zuordnung — nur wenn der Organizer die
                        Anzeige im Hotel-Bereich freigegeben hat UND für diese
                        Person ein Hotel hinterlegt ist. Rein lesend; Änderungen
                        laufen über die Organizer. */}
                    {event.hotelVisibleToAttendees && (registration.Hotel || '').trim() && (() => {
                      const hotelName = (registration.Hotel || '').trim();
                      const h = (event.hotels || []).filter(x => x.name === hotelName)[0];
                      const day = (iso?: string): string => (iso ? String(iso).substring(0, 10) : '');
                      const from = day(registration.HotelFrom);
                      const to = day(registration.HotelTo);
                      const nights = (from && to)
                        ? Math.max(0, Math.round((Date.parse(`${to}T00:00:00Z`) - Date.parse(`${from}T00:00:00Z`)) / 86400000))
                        : 0;
                      const fmt = (d: string): string => {
                        if (!d) return '—';
                        const t = Date.parse(`${d}T00:00:00Z`);
                        return isNaN(t) ? d : new Date(t).toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric' });
                      };
                      return (
                        <div style={{
                          marginTop: 6, padding: '10px 12px', borderRadius: 8,
                          background: 'rgba(134,188,37,0.07)', border: '1px solid var(--dex-green, #86bc25)',
                          color: 'var(--dex-gray-700)', fontSize: '0.82rem', lineHeight: 1.55,
                        }}>
                          <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>
                            {isDe ? 'Deine Unterkunft' : 'Your accommodation'}
                          </strong>
                          <div style={{ marginTop: 2 }}>
                            <strong>{hotelName}</strong>
                            {h && h.address && <span style={{ color: 'var(--dex-gray-600)' }}> · {h.address}</span>}
                          </div>
                          {(from || to) && (
                            <div style={{ color: 'var(--dex-gray-600)' }}>
                              {isDe ? 'Anreise' : 'Arrival'} {fmt(from)} · {isDe ? 'Abreise' : 'Departure'} {fmt(to)}
                              {nights > 0 && <> · {nights} {isDe ? (nights === 1 ? 'Nacht' : 'Nächte') : (nights === 1 ? 'night' : 'nights')}</>}
                            </div>
                          )}
                          <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                            {isDe ? 'Für Änderungen wende dich bitte an die Organizer.' : 'For changes please contact the organizers.'}
                          </div>
                        </div>
                      );
                    })()}
                    {/* v15.15: Hinweisbox „nur für Sub-Events angemeldet"
                        nur außerhalb des subEventsOnlyMode anzeigen — dort
                        ist sie redundant, weil es gar keine andere Option
                        gibt. */}
                    {sessionsOnly && !event.subEventsOnlyMode && (() => {
                      const term = event.childEventTermPlural || '';
                      const subList = (subEventTitles && subEventTitles.length > 0)
                        ? ` (${subEventTitles.join(', ')})`
                        : '';
                      const hintText = term
                        ? (isDe
                            ? `Du bist für ${term} dieses Events angemeldet, aber NICHT für das Haupt-Event selbst`
                            : `You are registered for ${term} of this event but NOT for the main event itself`)
                        : t('myevents.sessionsonly.hint');
                      return (
                        <div style={{
                          marginTop: 6, padding: '6px 10px', borderRadius: 6,
                          background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                          color: 'var(--dex-orange)', fontSize: '0.78rem',
                        }}>
                          {hintText}{subList}.
                        </div>
                      );
                    })()}

                    {/* v11.82: Team-Badge — sichtbar wenn die eigene Anmeldung
                        eine TeamId hat. Lazy-Load der anderen Mitglieder via
                        getTeamMembers. Zeigt: Team-Name (falls vorhanden),
                        Belegungs-Anzahl, Liste der Mitglieder (Lead zuerst). */}
                    {registration.TeamId && (() => {
                      const cacheKey = `${event.id}|${registration.TeamId}`;
                      const cached = teamMembersCache[cacheKey];
                      if (!cached) {
                        // Async laden (nur einmal pro Karte, idempotent über ref-Set).
                        enqueueTeamFetch(event.id, registration.TeamId);
                        return null;
                      }
                      const activeMembers = cached.filter(m => m.Status !== 'Abgemeldet');
                      const total = activeMembers.length;
                      const teamSizeCfg = event.teamSize || total;
                      const tn = registration.TeamName || (cached.find(m => m.TeamName)?.TeamName) || '';
                      const teamTermS = event.teamTermSingular || 'Team';
                      const isLead = !!registration.TeamLead;
                      // v11.86: Sortierung — Lead zuerst, dann nach TeilnehmerID,
                      // dann nach Id. Abgemeldete Mitglieder werden ausgegraut
                      // mit eigenem Badge weiter unten gerendert.
                      const sortedAll = [...cached].sort((a, b) => {
                        const aLead = a.TeamLead ? 0 : 1;
                        const bLead = b.TeamLead ? 0 : 1;
                        if (aLead !== bLead) return aLead - bLead;
                        const aTid = typeof a.TeilnehmerID === 'number' ? a.TeilnehmerID : Number.MAX_SAFE_INTEGER;
                        const bTid = typeof b.TeilnehmerID === 'number' ? b.TeilnehmerID : Number.MAX_SAFE_INTEGER;
                        if (aTid !== bTid) return aTid - bTid;
                        return a.Id - b.Id;
                      });
                      return (
                        <div style={{
                          marginTop: 8, padding: '10px 14px', borderRadius: 6,
                          background: 'rgba(134,188,37,0.08)', border: '1px solid var(--dex-green, #86bc25)',
                          color: 'var(--dex-green-dark, #3f5f10)', fontSize: '0.82rem', lineHeight: 1.45,
                        }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                            <Icon iconName="People" style={{ fontSize: 14 }} />
                            <strong>
                              {isDe
                                ? `${teamTermS} „${tn || 'Unbenannt'}" — ${total}/${teamSizeCfg} belegt`
                                : `${teamTermS} „${tn || 'Unnamed'}" — ${total}/${teamSizeCfg} taken`}
                            </strong>
                            {isLead && (
                              <span style={{
                                padding: '1px 8px', borderRadius: 999,
                                background: 'var(--dex-green, #86bc25)', color: '#fff',
                                fontSize: '0.7rem', fontWeight: 600,
                              }}>
                                {isDe ? `du bist ${teamTermS}-Lead` : `you are ${teamTermS} lead`}
                              </span>
                            )}
                          </div>
                          {/* v11.86: Mitglieder-Karten — pro Person Foto + Name + Email + Standort. */}
                          {sortedAll.length > 0 && (
                            <div style={{ marginTop: 10, display: 'flex', flexDirection: 'column', gap: 8 }}>
                              {sortedAll.map(m => {
                                const isCancelled = m.Status === 'Abgemeldet';
                                const isMemberLead = !!m.TeamLead && !isCancelled;
                                const fullName = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantEmail;
                                const loc = (m.Location || '').trim();
                                return (
                                  <div
                                    key={m.Id}
                                    style={{
                                      display: 'flex', alignItems: 'center', gap: 12,
                                      padding: '6px 8px', borderRadius: 6,
                                      background: isCancelled ? 'rgba(0,0,0,0.04)' : '#fff',
                                      border: '1px solid var(--dex-gray-200)',
                                      opacity: isCancelled ? 0.55 : 1,
                                      position: 'relative',
                                    }}
                                  >
                                    <div
                                      style={{
                                        width: 40, height: 40, borderRadius: '50%',
                                        overflow: 'visible', flexShrink: 0,
                                        position: 'relative',
                                      }}
                                    >
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
                                      <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)', fontSize: '0.85rem' }}>
                                        {fullName}
                                      </div>
                                      <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)' }}>
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
                                    ) : isMemberLead && (
                                      <span style={{
                                        padding: '2px 8px', borderRadius: 999,
                                        background: 'var(--dex-green, #86bc25)', color: '#fff',
                                        fontSize: '0.68rem', fontWeight: 600, flexShrink: 0,
                                      }}>
                                        Lead
                                      </span>
                                    )}
                                  </div>
                                );
                              })}
                            </div>
                          )}
                          {/* v11.83/v11.86: Aktion-Buttons — Edit (alle Leads) +
                              Add (nur bei freien Slots). */}
                          {isLead && (
                            <div style={{ marginTop: 10, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                              <button
                                type="button"
                                className="btn btn-secondary"
                                onClick={() => openManageTeamDialog(
                                  event.id,
                                  registration.TeamId!,
                                  tn || '',
                                  teamSizeCfg,
                                  cached
                                )}
                                style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                              >
                                <Icon iconName="Edit" style={{ fontSize: 12, marginRight: 6 }} />
                                {isDe ? 'Team bearbeiten' : 'Edit team'}
                              </button>
                              {total < teamSizeCfg && (
                                <button
                                  type="button"
                                  className="btn btn-secondary"
                                  onClick={() => {
                                    setAddMemberDialog({
                                      eventId: event.id,
                                      teamId: registration.TeamId!,
                                      teamName: tn || '',
                                      freeSlots: teamSizeCfg - total,
                                    });
                                    setAddMemberPick(null);
                                    setAddMemberQuery('');
                                    setAddMemberResults([]);
                                    setAddMemberConsent(false);
                                    setAddMemberError('');
                                  }}
                                  style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                                >
                                  <Icon iconName="AddFriend" style={{ fontSize: 12, marginRight: 6 }} />
                                  {isDe
                                    ? `Mitglied hinzufügen (${teamSizeCfg - total} Slot${(teamSizeCfg - total) === 1 ? '' : 's'} frei)`
                                    : `Add member (${teamSizeCfg - total} slot${(teamSizeCfg - total) === 1 ? '' : 's'} free)`}
                                </button>
                              )}
                            </div>
                          )}
                          {/* v11.83: Beitritts-Anfragen-Block — nur für Leads, wenn das
                              Event Approval aktiviert hat UND es Pending-Anfragen gibt. */}
                          {isLead && event.teamJoinRequiresApproval && (() => {
                            const jKey = `${event.id}|${registration.TeamId}`;
                            const jReqs = joinRequestsCache[jKey];
                            if (jReqs === undefined) {
                              enqueueJoinReqFetch(event.id, registration.TeamId!);
                              return null;
                            }
                            if (jReqs.length === 0) return null;
                            return (
                              <div style={{
                                marginTop: 10,
                                padding: '8px 12px',
                                borderRadius: 6,
                                background: 'rgba(237,139,0,0.10)',
                                border: '1px solid var(--dex-orange, #ed8b00)',
                                color: '#7a4a00',
                              }}>
                                <div style={{ fontWeight: 700, marginBottom: 6 }}>
                                  {isDe ? 'Beitritts-Anfragen' : 'Join requests'} ({jReqs.length})
                                </div>
                                {jReqs.map(r => {
                                  const busy = joinReqBusyId === r.Id;
                                  return (
                                    <div key={r.Id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', borderTop: '1px solid rgba(237,139,0,0.25)' }}>
                                      <div style={{ flex: 1, fontSize: '0.82rem' }}>
                                        <div style={{ fontWeight: 600 }}>{r.RequesterDisplayName || r.RequesterEmail}</div>
                                        <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)' }}>{r.RequesterEmail}</div>
                                      </div>
                                      <button
                                        type="button"
                                        className="btn btn-primary"
                                        disabled={busy}
                                        onClick={() => handleDecideJoinRequest(event.id, registration.TeamId!, r.Id, 'Approved')}
                                        style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                      >
                                        {isDe ? 'Bestätigen' : 'Approve'}
                                      </button>
                                      <button
                                        type="button"
                                        className="btn btn-secondary"
                                        disabled={busy}
                                        onClick={() => handleDecideJoinRequest(event.id, registration.TeamId!, r.Id, 'Rejected')}
                                        style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                      >
                                        {isDe ? 'Ablehnen' : 'Reject'}
                                      </button>
                                    </div>
                                  );
                                })}
                              </div>
                            );
                          })()}
                        </div>
                      );
                    })()}

                    {/* Kompakte Info-Zeile: Location + Datum inline, umbricht auf schmalen Bildschirmen */}
                    <div style={{ marginTop: 8, display: 'flex', flexWrap: 'wrap', alignItems: 'flex-start', gap: '6px 24px', fontSize: '0.88rem', color: 'var(--dex-gray-700)' }}>
                      {/* v27.8: Ort einzeilig als „Name, Stadt" (vorher zweizeilig
                          mit voller Adresse). */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <Icon iconName="MapPin" style={{ fontSize: 14, color: 'var(--dex-gray-500)' }} />
                        <span style={{ fontWeight: 700, color: 'var(--dex-gray-800)' }}>
                          {(() => {
                            // v29.39: Eine Teams-URL im Ort gehört nicht in die
                            // Ort-Zeile — sie steht daneben als Knopf.
                            const loc = eventTeamsLink(event) ? locationWithoutTeamsUrl(event.location) : (event.location || '');
                            return [loc, event.locationAddress && event.locationAddress.city].filter(Boolean).join(', ')
                              || (eventTeamsLink(event) ? (isDe ? 'Online' : 'Online') : '-');
                          })()}
                        </span>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}><Icon iconName="Calendar" style={{ fontSize: 14, color: 'var(--dex-gray-500)' }} /> {event.allDay ? formatAllDayPeriod(event.startDate, event.endDate, isDe) : formatDateRange(event.startDate, event.endDate)}</div>
                      {/* v29.39: Teilnahme-Knopf direkt bei den Eckdaten — wer in
                          „Meine Events" nachsieht, will von dort in die
                          Besprechung, nicht erst den Kalender suchen. */}
                      {eventTeamsLink(event) && (
                        <TeamsJoinButton url={eventTeamsLink(event)} isDe={isDe} variant="link" />
                      )}
                    </div>

                    {/* v27.7: Gruppe (Durchstarter/Funstarter bzw. eigene
                        Beschriftung) anzeigen, in der die Person angemeldet ist.
                        StarterType = effektive Gruppe; auf der Warteliste ist er
                        leer, dann steht der Wunsch in PreferredStarterType. */}
                    {(() => {
                      const grp = (registration.StarterType || registration.PreferredStarterType || '').trim();
                      if (!grp) return null;
                      const grpLabel = grp === 'Durchstarter'
                        ? ((event.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter')
                        : grp === 'Funstarter'
                          ? ((event.splitLabelB && event.splitLabelB.trim()) || 'Funstarter')
                          : grp;
                      return (
                        <div style={{ marginTop: 8 }}>
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.78rem', fontWeight: 600, padding: '4px 10px', borderRadius: 999, background: 'rgba(134,188,37,0.14)', color: 'var(--dex-green-dark, #4a7c1f)', border: '1px solid rgba(134,188,37,0.30)' }}>
                            <Icon iconName="Group" style={{ fontSize: 13 }} />
                            {(isDe ? 'Gruppe: ' : 'Group: ')}{grpLabel}
                          </span>
                        </div>
                      );
                    })()}

                    {/* Organizer mit Foto (Hover vergrößert). v24.12: einzelne ausblendbar. */}
                    {event.organizers.length > 0 && !(event.hideOrganizer && !event.hideOrganizerIndividualOnly) && (
                      <div style={{ marginTop: 10 }}>
                        <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4 }}>Organizer</div>
                        <OrganizerList
                          names={event.organizers.reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map(o => {
                            const trimmed = o.trim();
                            const parts = trimmed.split(',').map(s => s.trim());
                            return parts.length === 2 ? `${parts[1]} ${parts[0]}` : trimmed;
                          }).filter(Boolean)}
                          emails={event.organizerEmails}
                          hiddenEmails={(event.hideOrganizer && event.hideOrganizerIndividualOnly) ? event.hiddenOrganizerEmails : []}
                          size="sm"
                        />
                      </div>
                    )}
                    {/* v10.26: Optionaler Ansprechpartner — frei eingegebene Person
                        außerhalb des App-User-Pools. Reines Anzeige-Feld; Mailto-Link
                        wenn Email gesetzt. Wird nur gerendert wenn mindestens Name
                        oder Email gepflegt sind. Spiegelt das Verhalten der
                        Registration-Page in My Events wider. */}
                    {(event.contactName || event.contactEmail || event.contactInfo) && (
                      <div style={{ marginTop: 10 }}>
                        {/* v28.7: Überschrift AUSSERHALB der Box — gleiches
                            Muster wie „Organizer" darüber (und wie auf der
                            Anmelde-Seite seit v28.6). */}
                        <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4 }}>
                          {(event.emailLanguage || 'EN').toUpperCase() === 'DE' ? 'Ansprechpartner' : 'Contact'}
                        </div>
                        <div style={{ padding: '8px 10px', background: 'var(--dex-gray-50, #f7f7f7)', borderRadius: 6, border: '1px solid var(--dex-gray-200)' }}>
                        {event.contactName && (
                          <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>{event.contactName}</div>
                        )}
                        {event.contactEmail && (
                          <div style={{ fontSize: '0.78rem', marginTop: 2 }}>
                            <a href={`mailto:${event.contactEmail}`} style={{ color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'none' }}>{event.contactEmail}</a>
                          </div>
                        )}
                        {event.contactInfo && (
                          <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-700)', marginTop: 4, whiteSpace: 'pre-wrap', lineHeight: 1.45 }}>{event.contactInfo}</div>
                        )}
                        </div>
                      </div>
                    )}
                  </div>
                </div>

                {/* v17.22: Event-Beschreibung auch unter „Meine Events"
                    anzeigen (vorher nur auf der Anmeldeseite). RichText-HTML
                    aus dem eigenen Tenant — gleiche Render-Logik wie auf der
                    RegistrationPage (HTML erlaubt, sonst \n→<br>).
                    v17.23: standardmäßig eingeklappt, per Button aufklappbar. */}
                {event.description && (!editingId || editingId !== event.id) && (() => {
                  const isOpen = !!descExpanded[event.id];
                  return (
                    <div style={{ marginTop: 10 }}>
                      <button
                        type="button"
                        onClick={() => setDescExpanded(prev => ({ ...prev, [event.id]: !prev[event.id] }))}
                        aria-expanded={isOpen}
                        style={{
                          display: 'inline-flex', alignItems: 'center', gap: 6,
                          background: 'none', border: 'none', cursor: 'pointer', padding: 0,
                          fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-green-dark, #4a7c1f)',
                        }}
                      >
                        <span style={{
                          display: 'inline-block', transition: 'transform 0.15s',
                          transform: isOpen ? 'rotate(90deg)' : 'rotate(0deg)', fontSize: '0.7rem',
                        }}>▶</span>
                        {isDe ? 'Beschreibung' : 'Description'}
                      </button>
                      {isOpen && (
                        <div
                          className="dex-event-desc"
                          style={{
                            marginTop: 6, padding: '10px 14px',
                            color: 'var(--dex-gray-700)', background: 'var(--dex-gray-50, #fafafa)',
                            borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-gray-200)',
                            wordBreak: 'break-word',
                          }}
                          dangerouslySetInnerHTML={{
                            __html: (() => {
                              const raw = event.description || '';
                              const isHtml = /<[a-z][\s\S]*>/i.test(raw);
                              return isHtml
                                ? raw
                                : raw.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
                            })(),
                          }}
                        />
                      )}
                    </div>
                  );
                })()}

                {/* v10.26: Custom-Field-Antworten als rechteckige
                    pastellgrüne Tags — visuell eindeutig von den
                    abgerundeten grauen Organizer-Chips getrennt, damit der
                    User sofort sieht: das ist „was ich angegeben habe", nicht
                    „wer organisiert das". */}
                {!editingId || editingId !== event.id ? (
                  // v18.38: Edit-Bereich anzeigen, sobald das Event überhaupt
                  // bearbeitbare Felder hat — NICHT mehr nur wenn schon Werte
                  // ausgefüllt sind. Sonst kann ein Teilnehmer, der ein
                  // optionales Feld leer gelassen hat (z.B. „zusätzliche
                  // Nacht"), es später nicht mehr nachtragen.
                  // v20.9 BUG-FIX: Der Block rendert jetzt AUCH, wenn das Event
                  // keine Custom-Felder hat, aber die Anmeldung QR-fähig ist —
                  // sonst fehlte der „Mein QR-Code"-Button bei Events ohne
                  // Abfragefelder (z.B. einfaches Sommerfest).
                  // Die Aktions-Zeile rendert immer, da sie jetzt mindestens den
                  // „Nachrichten zum Event"-Button enthält (zusätzlich zu den
                  // optionalen Angaben-Tags und dem QR-Button).
                  (
                    <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 6, alignItems: 'center' }}>
                      {displayData.map(({ label, value, type }) => (
                        <FieldAnswerTag key={label} label={label} value={value} type={type} />
                      ))}
                      {/* v11.30: Edit-Button direkt neben den Angaben-Tags
                          (statt unten in der Aktions-Zeile). Näher am Inhalt
                          den er bearbeitet.
                          v18.38: zeigt jetzt „Angaben ergänzen", wenn noch
                          nichts ausgefüllt wurde — sonst „Angaben bearbeiten". */}
                      {!hiddenRow && (event.eventSpecificFields || []).filter((f: EventSpecificField) => f.label).length > 0 && (
                      <button
                        type="button"
                        className="btn btn-outline"
                        onClick={() => { setEditData(customData); setEditingId(event.id); }}
                        style={{
                          fontSize: '0.78rem', padding: '5px 12px', borderRadius: 6,
                          width: 'auto', cursor: 'pointer',
                          display: 'inline-flex', alignItems: 'center', gap: 4,
                        }}
                        title={displayData.length > 0 ? t('myevents.edit') : (isDe ? 'Angaben ergänzen' : 'Add details')}
                      >
                        <Pencil size={12} /> {displayData.length > 0 ? t('myevents.edit') : (isDe ? 'Angaben ergänzen' : 'Add details')}
                      </button>
                      )}
                      {/* v20.7: Persönlicher Check-in-QR — gleicher Code wie
                          in der QR-Mail. v28.7: erst sichtbar, NACHDEM die
                          QR-Codes fürs Event versendet wurden (Status
                          'QR versendet'/'Eingecheckt') — vorher wirkte der
                          Button, als gäbe es schon einen gültigen Check-in. */}
                      {!sessionsOnly && !hiddenRow && !!event.eventNumber && ['QR versendet', 'Eingecheckt'].indexOf(registration.Status) >= 0 && (
                        <button
                          type="button"
                          className="btn btn-outline"
                          onClick={() => { openMyQr(event, registration).catch(() => { /* */ }); }}
                          style={{
                            fontSize: '0.78rem', padding: '5px 12px', borderRadius: 6,
                            width: 'auto', cursor: 'pointer',
                            display: 'inline-flex', alignItems: 'center', gap: 4,
                          }}
                          title={isDe ? 'Deinen persönlichen Check-in-QR-Code anzeigen' : 'Show your personal check-in QR code'}
                        >
                          <QrCode size={12} /> {isDe ? 'Mein QR-Code' : 'My QR code'}
                        </button>
                      )}
                      {/* Nachrichten zum Event: Broadcast-Mails (Einladung,
                          Ankündigungen) aus dem Kommunikations-Log lesen. */}
                      <button
                        type="button"
                        className="btn btn-outline"
                        onClick={() => openComms(event)}
                        style={{
                          fontSize: '0.78rem', padding: '5px 12px', borderRadius: 6,
                          width: 'auto', cursor: 'pointer',
                          display: 'inline-flex', alignItems: 'center', gap: 4,
                        }}
                        title={isDe ? 'Nachrichten zu diesem Event ansehen' : 'View messages for this event'}
                      >
                        <Mail size={12} /> {isDe ? 'Nachrichten zum Event' : 'Event messages'}
                      </button>
                    </div>
                  )
                ) : (
                  <div style={{ marginTop: 12 }}>
                    {/* v17.22: EN-Varianten auch im „Meine Events"-Edit-Formular
                        berücksichtigen — vorher rein DE, obwohl der Teilnehmer
                        sich auf der Anmeldeseite die EN-Labels angesehen hatte. */}
                    {(() => {
                      const useEnEdit = locale === 'en' && !!event.bilingualFields;
                      const eLabel = (f: EventSpecificField): string =>
                        (useEnEdit && f.labelEn && f.labelEn.trim()) ? f.labelEn : f.label;
                      const eOpt = (f: EventSpecificField, opt: string, idx: number): string =>
                        (useEnEdit && f.optionsEn && f.optionsEn[idx] && f.optionsEn[idx].trim()) ? f.optionsEn[idx] : opt;
                      return event.eventSpecificFields.map((field: EventSpecificField) => (
                        <div className="form-group" key={field.id} style={{ marginBottom: 10 }}>
                          <label className="form-label" style={{ fontSize: '0.82rem', marginBottom: 2 }}>
                            {field.required && <span className="required">*</span>}
                            {eLabel(field)}
                          </label>
                          {field.type === 'select' ? (
                            <select className="form-select" value={editData[field.id] || ''} onChange={e => setEditData({ ...editData, [field.id]: e.target.value })}>
                              <option value="">—</option>
                              {field.options && field.options.map((opt, optIdx) => <option key={opt} value={opt}>{eOpt(field, opt, optIdx)}</option>)}
                            </select>
                          ) : (field.type === 'user' || field.type === 'roommate') ? (
                            /* v18.61: People-Picker-Felder beim Bearbeiten wieder als
                               echter People-Picker mit Profilfoto (vorher Rohtext). */
                            <UserFieldPicker
                              value={editData[field.id] || ''}
                              onChange={v => setEditData({ ...editData, [field.id]: v })}
                              // v29.40: Nachträglich ergänzte Angaben dürfen die
                              // Verteiler-Begrenzung des Feldes nicht umgehen —
                              // sonst wäre der Umweg über „Angaben ergänzen"
                              // genau das Schlupfloch, das die Option schließt.
                              searchUsers={field.audienceOnly
                                ? (async (q: string, intl?: boolean) => {
                                  const res = await searchUsers(q, intl);
                                  return res.filter(u => isEventVisibleForUser(event, u.email, u.location || '', [], u.jobTitle || ''));
                                })
                                : searchUsers}
                              searchUserByEmail={searchUser}
                              placeholder={locale === 'de' ? 'Name oder E-Mail eingeben…' : 'Type a name or email…'}
                              errorStyle={{}}
                            />
                          ) : field.type === 'checkbox' ? (
                            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.85rem' }}>
                              <input
                                type="checkbox"
                                checked={editData[field.id] === 'true'}
                                onChange={e => setEditData({ ...editData, [field.id]: e.target.checked ? 'true' : 'false' })}
                                style={{ marginTop: 2 }}
                              />
                              <span>{(useEnEdit && field.confirmLabelEn && field.confirmLabelEn.trim() ? field.confirmLabelEn : field.confirmLabel) || eLabel(field)}</span>
                            </label>
                          ) : (
                            <input className="form-input" value={editData[field.id] || ''} onChange={e => setEditData({ ...editData, [field.id]: e.target.value })} placeholder={eLabel(field)} type={field.type === 'number' ? 'number' : 'text'} />
                          )}
                        </div>
                      ));
                    })()}
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button className="btn btn-primary" style={{ fontSize: '0.82rem' }} disabled={isSaving} onClick={async () => { setIsSaving(true); await updateMyRegistration(event.id, editData); await loadMyRegistrations(); setEditingId(null); setIsSaving(false); }}>
                        {isSaving ? t('myevents.saving') : t('myevents.save')}
                      </button>
                      <button className="btn btn-secondary" style={{ fontSize: '0.82rem' }} onClick={() => setEditingId(null)}>{t('general.cancel')}</button>
                    </div>
                  </div>
                )}

                {/* Agenda / Timeline - mehrspaltig bei mehreren Tagen, horizontal scrollbar auf Mobile */}
                {event.agenda && event.agenda.length > 0 && (() => {
                  const grouped = Object.entries(
                    event.agenda.reduce((groups: Record<string, AgendaItem[]>, item: AgendaItem) => {
                      const key = item.date || 'TBD';
                      if (!groups[key]) groups[key] = [];
                      groups[key].push(item);
                      return groups;
                    }, {} as Record<string, AgendaItem[]>)
                  ).sort(([a], [b]) => a.localeCompare(b));
                  const dayCount = grouped.length;
                  // v22.36: Durchlaufende Nummerierung der Agenda-Schritte
                  // (über alle Tage, sortiert nach Datum + Uhrzeit) — die
                  // Einzel-Schritte zeigen keine Icons mehr, nur die Sektion.
                  const agendaOrderIds = event.agenda
                    .slice()
                    .sort((a: AgendaItem, b: AgendaItem) => ((a.date || '') + (a.time || '')).localeCompare((b.date || '') + (b.time || '')))
                    .map((x: AgendaItem) => x.id);

                  return (
                    <div style={{ marginTop: 12 }}>
                      <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 8, display: 'flex', alignItems: 'center', gap: 6 }}>
                        <Icon iconName="Calendar" style={{ fontSize: 14, color: 'var(--dex-green-dark, #6b9a1e)' }} />
                        {t('myevents.agenda')} {dayCount > 1 && <span style={{ fontWeight: 400, fontSize: '0.72rem', color: 'var(--dex-gray-400)' }}>· {dayCount} {t('myevents.agenda') === 'Programm' ? 'Tage (seitwärts scrollen)' : 'days (swipe)'}</span>}
                      </div>
                      {/* Horizontal scrollbarer Container - funktioniert auf Desktop und Mobile */}
                      <div
                        style={{
                          display: 'flex',
                          flexWrap: 'nowrap',
                          gap: 16,
                          overflowX: 'auto',
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          WebkitOverflowScrolling: 'touch' as any,
                          scrollSnapType: 'x mandatory',
                          paddingBottom: 4,
                        }}
                      >
                        {grouped.map(([date, items]) => (
                          <div key={date} style={{
                            flex: `0 0 ${dayCount === 1 ? '100%' : dayCount === 2 ? 'calc(50% - 8px)' : 'min(280px, 85%)'}`,
                            scrollSnapAlign: 'start',
                            background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, padding: 12,
                            border: '1px solid var(--dex-gray-200)',
                            minWidth: 0,
                          }}>
                            <div style={{
                              fontSize: '0.78rem', fontWeight: 700, color: '#fff', marginBottom: 8,
                              background: 'var(--dex-green-dark, #6b9a1e)', borderRadius: 8, padding: '6px 12px',
                              textAlign: 'center',
                            }}>
                              {date !== 'TBD' ? new Date(date + 'T00:00').toLocaleDateString('de-DE', { weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric' }) : 'TBD'}
                            </div>
                            {items.sort((a: AgendaItem, b: AgendaItem) => (a.time || '').localeCompare(b.time || '')).map((item: AgendaItem) => (
                              <div key={item.id} style={{
                                display: 'flex', alignItems: 'flex-start', gap: 8, padding: '6px 0',
                                borderLeft: '2px solid var(--dex-green)', marginLeft: 4, paddingLeft: 10,
                              }}>
                                <span style={{ width: 26, height: 26, borderRadius: '50%', background: 'var(--dex-green-dark, #6b9a1e)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, color: '#fff', fontWeight: 700, fontSize: '0.72rem', lineHeight: 1 }}>
                                  {agendaOrderIds.indexOf(item.id) + 1}
                                </span>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontSize: '0.8rem', fontWeight: 600 }}>
                                    {item.time}{item.endTime ? ` – ${item.endTime}` : ''}
                                  </div>
                                  <div style={{ fontSize: '0.8rem', wordBreak: 'break-word' }}>{item.title}</div>
                                  {item.description && (
                                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 1, wordBreak: 'break-word' }}>{item.description}</div>
                                  )}
                                </div>
                              </div>
                            ))}
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })()}

                {/* Transferzeiten */}
                {event.transferTimes && event.transferTimes.length > 0 && (
                  <div style={{ marginTop: 12 }}>
                    <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 6 }}>
                      {t('myevents.transfers')}
                    </div>
                    {event.transferTimes.sort((a: TransferTime, b: TransferTime) => (a.date + a.departureTime).localeCompare(b.date + b.departureTime)).map((tr: TransferTime) => (
                      <div key={tr.id} style={{
                        display: 'flex', gap: 10, padding: '8px 12px', marginBottom: 6, fontSize: '0.82rem',
                        background: 'var(--dex-gray-50, #fafafa)', borderRadius: 10,
                        borderLeft: '3px solid var(--dex-orange)',
                      }}>
                        <span style={{ width: 28, height: 28, borderRadius: '50%', background: 'var(--dex-orange)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                          <Icon iconName="Bus" style={{ fontSize: 13, color: '#fff' }} />
                        </span>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontWeight: 600 }}>
                            {tr.location}{tr.meetingPoint ? ` – ${tr.meetingPoint}` : ''}
                          </div>
                          {tr.address && (
                            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
                              <Icon iconName="MapPin" style={{ fontSize: 11, marginRight: 4 }} />{tr.address}
                            </div>
                          )}
                          <div style={{ marginTop: 2 }}>
                            <Icon iconName="Calendar" style={{ fontSize: 11, color: 'var(--dex-gray-400)', marginRight: 4 }} />
                            {new Date(tr.date + 'T00:00').toLocaleDateString('de-DE', {weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric'})},
                            {' '}{tr.departureTime}{tr.arrivalTime ? ` → ${tr.arrivalTime}` : ''} Uhr
                          </div>
                          {tr.description && <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>{tr.description}</div>}
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                {/* Dokumente mit Viewer */}
                {event.documents && event.documents.length > 0 && (
                  <DocumentsViewer documents={event.documents} t={t} />
                )}

                {/* Fun-Zone Quiz */}
                {!sessionsOnly && !hiddenRow && event.quiz && event.quiz.length > 0 && (
                  <QuizPlayer
                    quiz={event.quiz}
                    t={t}
                    clusterSize={event.quizClusterSize}
                    initialAnswers={(() => {
                      // Zuvor gespeicherte Antworten aus der Teilnehmer-Registrierung laden
                      // (QuizAnswers ist JSON-String eines number[][]).
                      try {
                        const raw = registration.QuizAnswers;
                        if (!raw) return undefined;
                        const parsed = JSON.parse(raw);
                        return Array.isArray(parsed) ? parsed : undefined;
                      } catch { return undefined; }
                    })()}
                    onProgress={async (score: number, answers: number[][], isComplete: boolean) => {
                      // Nach jedem Cluster-Wechsel in die Subsite-Teilnehmerliste schreiben.
                      // isComplete=true setzt zusätzlich QuizCompletedAt (für Statistik-Filter).
                      if (!event.subsiteUrl) {
                        console.warn('[DEX] saveQuizProgress: event.subsiteUrl leer - Save übersprungen', { eventId: event.id });
                        return;
                      }
                      try {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const ctx = (window as any).__dexSpfxContext;
                        if (!ctx) {
                          console.warn('[DEX] saveQuizProgress: __dexSpfxContext fehlt - Save übersprungen');
                          return;
                        }
                        const { EventService } = await import('../../services/EventService');
                        const svc = new EventService(ctx);
                        const ok = await svc.saveQuizProgress(event.subsiteUrl, registration.Id, score, answers, isComplete);
                        const answeredNow = answers.filter(a => Array.isArray(a) && a.length > 0).length;
                        console.warn(`[DEX] saveQuizProgress ok=${ok} regId=${registration.Id} score=${score} answered=${answeredNow}/${answers.length} complete=${isComplete} subsite=${event.subsiteUrl}`);
                        // VERIFY: re-read das Item und checke ob QuizAnswers tatsächlich in SP landete
                        try {
                          const verifyResp = await ctx.spHttpClient.get(
                            `${event.subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items(${registration.Id})?$select=QuizScore,QuizAnswers,QuizCompletedAt`,
                            (await import('@microsoft/sp-http')).SPHttpClient.configurations.v1
                          );
                          if (verifyResp.ok) {
                            const data = await verifyResp.json();
                            console.warn(`[DEX] saveQuizProgress VERIFY: QuizScore=${data.QuizScore} QuizAnswersLen=${(data.QuizAnswers || '').length} QuizCompletedAt=${data.QuizCompletedAt} QuizAnswersSample=${String(data.QuizAnswers || '').substring(0, 120)}`);
                          } else {
                            console.warn(`[DEX] saveQuizProgress VERIFY read failed: ${verifyResp.status}`);
                          }
                        } catch (vErr) {
                          console.warn('[DEX] saveQuizProgress VERIFY error:', vErr);
                        }
                        // Lokale myEvents-Liste nach erfolgreichem Save aktualisieren, damit die
                        // registration im Parent-State die frischen QuizScore/QuizAnswers hat —
                        // sonst sieht der User beim Wiedereintritt noch das alte (leere) Feld.
                        if (ok) {
                          setMyEvents(prev => prev.map(entry => {
                            if (entry.event.id !== event.id) return entry;
                            return {
                              ...entry,
                              registration: {
                                ...entry.registration,
                                QuizScore: score,
                                QuizAnswers: JSON.stringify(answers),
                                ...(isComplete ? { QuizCompletedAt: new Date().toISOString() } : {}),
                              },
                            };
                          }));
                        }
                      } catch (err) { console.warn('[DEX] saveQuizProgress failed:', err); }
                    }}
                  />
                )}

                {/* Sub-Events (Trainingssessions etc.) — nur wenn Event welche hat.
                    Seit v6.4: Sub-Events sind eigene DEX_Events-Items, werden über
                    childEventsOf(parentId) aus dem Context gezogen. */}
                {childEventsOf(event.id).length > 0 && (
                  <MyEventSubEvents
                    parentEvent={event}
                    childEvents={childEventsOf(event.id)}
                    registerForEvent={registerForEvent}
                    cancelRegistration={cancelRegistration}
                    getMyRegistration={getMyRegistration}
                    getAllRegistrations={getAllRegistrations}
                    updateMyRegistration={updateMyRegistration}
                    onMutated={loadMyRegistrations}
                  />
                )}

                {/* v11.0: Datei-Upload-Block — wird nur gerendert, wenn der
                    Organizer beim Event den Upload erlaubt hat und die
                    Anmeldung aktiv ist (nicht sessionsOnly oder abgemeldet). */}
                {event.allowAttendeeUpload && !sessionsOnly && !hiddenRow && (
                  <MyEventUpload
                    event={event}
                    list={listMyEventAttachments}
                    upload={uploadMyEventAttachment}
                    remove={deleteMyEventAttachment}
                  />
                )}

                {/* v19.0: Dokument-Custom-Felder — pro Feld ein Upload-Block,
                    damit der User die Datei auch nachträglich ergänzen/ersetzen
                    kann. */}
                {!sessionsOnly && !hiddenRow && (event.eventSpecificFields || []).filter(f => f.type === 'document').map(df => (
                  <MyEventDocField
                    key={df.id}
                    event={event}
                    field={df}
                    list={listFieldDocuments}
                    upload={uploadFieldDocument}
                    remove={deleteFieldDocument}
                  />
                ))}
                {/* Registriert am + Aktionen — im Sessions-Only-Modus ausblenden,
                    weil es keine echte Parent-Registrierung gibt. Sessions werden
                    über die Sub-Event-Sektion oben gemanagt. */}
                {!sessionsOnly && !hiddenRow && (
                  <div style={{ marginTop: 12, paddingTop: 10, borderTop: '1px solid var(--dex-gray-200)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
                    <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)' }}>
                      {t('myevents.registeredon')}: {formatDate(registration.RegistrationDate)}
                    </span>
                    <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                      {cancellingId === event.id && !isCancelling && event.lastDeregisterDate && new Date(event.lastDeregisterDate) < new Date() && (
                        // v9.17: prominenter Late-Cancel-Hinweis — der User soll
                        // klar sehen, dass der Organizer durch die Abmeldung
                        // automatisch informiert wird (entscheidend wenn z.B.
                        // Hotel/Catering noch reagieren muss).
                        <span style={{
                          fontSize: '0.82rem', color: 'var(--dex-orange-dark, #b35a00)',
                          background: 'var(--dex-orange-light, #fff3e0)',
                          border: '1px solid var(--dex-orange, #ed8b00)',
                          padding: '6px 10px', borderRadius: 6,
                          display: 'block', marginBottom: 6, width: '100%',
                          fontWeight: 500,
                        }}>
                          {t('myevents.latecancel')}
                        </span>
                      )}
                      {/* v11.30: „Angaben bearbeiten"-Button wandert nach
                          oben neben die Angaben-Tags. Hier in der Aktions-
                          Zeile nur noch der Cancel-Edit-Button während
                          aktiver Bearbeitung — sonst leer. */}
                      {editingId === event.id && (
                        <button className="btn btn-secondary" style={{ fontSize: '0.85rem', padding: '8px 16px' }} onClick={() => setEditingId(null)}>
                          {t('general.cancel')}
                        </button>
                      )}
                      {/* v10.27: Gruppe wechseln bei Split-Capacity-Events.
                          Sichtbar nur wenn beide Kapazitäten > 0 sind und der
                          User aktiv angemeldet (nicht abgemeldet) ist. Ein
                          Klick öffnet einen window.confirm-Dialog mit klarem
                          Hinweis, dass der Wechsel evtl. auf die Warteliste
                          der Ziel-Gruppe führt, falls diese voll ist. */}
                      {(() => {
                        const dCap = event.durchstarterCapacity || 0;
                        const fCap = event.funstarterCapacity || 0;
                        if (dCap <= 0 || fCap <= 0) return null;
                        const labelA = (event.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
                        const labelB = (event.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
                        const currentType = registration.StarterType || registration.PreferredStarterType || '';
                        const currentLabel = currentType === 'Durchstarter' ? labelA : currentType === 'Funstarter' ? labelB : '?';
                        const targetType: 'Durchstarter' | 'Funstarter' = currentType === 'Durchstarter' ? 'Funstarter' : 'Durchstarter';
                        const targetLabel = targetType === 'Durchstarter' ? labelA : labelB;
                        return (
                          <button
                            className="btn btn-secondary"
                            style={{ fontSize: '0.85rem', padding: '8px 16px' }}
                            title={t('myevents.switchgroup.title') || `Aktuell in: ${currentLabel}`}
                            onClick={async () => {
                              const msg = `${t('myevents.switchgroup.confirm') || 'Gruppe wechseln zu'} „${targetLabel}"?\n\n` +
                                ((t('myevents.switchgroup.hint') || 'Falls die Ziel-Gruppe bereits voll ist, kommst du auf deren Warteliste und rückst nach, sobald ein Platz frei wird.'));
                              if (!(await confirmDialog(msg, { confirmLabel: isDe ? 'Wechseln' : 'Switch' }))) return;
                              const r = await switchSplitGroup(event.id, targetType);
                              if (!r.ok) {
                                showAlert(t('myevents.switchgroup.failed') || 'Gruppen-Wechsel fehlgeschlagen.', { variant: 'error' });
                                return;
                              }
                              const okMsg = r.full
                                ? `${t('myevents.switchgroup.waitlist') || 'Wechsel registriert — du stehst auf der Warteliste der Gruppe'} „${targetLabel}".`
                                : `${t('myevents.switchgroup.success') || 'Wechsel erfolgreich — du bist jetzt in Gruppe'} „${targetLabel}".`;
                              showAlert(okMsg, { variant: 'success' });
                              await loadMyRegistrations();
                            }}
                          >
                            {(t('myevents.switchgroup.btn') || 'Gruppe wechseln')} → {targetLabel}
                          </button>
                        );
                      })()}
                      {/* Abmelden-Button: prominent ausgelegt damit er auf der Karte
                          sofort gefunden wird (User-Feedback v9.8). 2-Klick-Confirm
                          bleibt — der erste Klick färbt rot und blendet den
                          "Doch behalten"-Button daneben ein.
                          v22.22: Bei bereits vergangenen Events entfällt der Button —
                          stattdessen ein grauer Hinweis (performCancel blockt
                          zusätzlich, auch für den Auto-Cancel-Deep-Link). */}
                      {/* v28.23: Fremd angelegte, für die Person nicht lesbare
                          Zeile — die Selbst-Abmeldung würde an denselben
                          Zeilen-Rechten scheitern. Statt eines Buttons, der
                          nicht funktioniert, ein klarer Hinweis. */}
                      {hiddenRow ? (
                        <span style={{
                          fontSize: '0.8rem', color: 'var(--dex-gray-500)',
                          padding: '8px 12px', borderRadius: 8,
                          background: 'var(--dex-gray-50, #fafafa)',
                          border: '1px solid var(--dex-gray-200)',
                          lineHeight: 1.4,
                        }}>
                          {isDe
                            ? 'Abmelden über die Person, die dich angemeldet hat, oder über die Organizer.'
                            : 'To cancel, contact whoever registered you, or the organizers.'}
                        </span>
                      ) : isEventOver(event) ? (
                        <span style={{
                          fontSize: '0.8rem', color: 'var(--dex-gray-500)',
                          padding: '8px 12px', borderRadius: 8,
                          background: 'var(--dex-gray-50, #fafafa)',
                          border: '1px solid var(--dex-gray-200)',
                          lineHeight: 1.4,
                        }}>
                          {isDe
                            ? 'Dieses Event liegt in der Vergangenheit — eine Abmeldung ist nicht mehr möglich.'
                            : 'This event is in the past — cancelling is no longer possible.'}
                        </span>
                      ) : selfCancelLocked(event) ? (
                        /* v29.25: Selbst-Abmeldung gesperrt (komplett oder
                           nach der Frist) — statt eines Knopfs, der nur eine
                           Fehlermeldung produziert, steht hier der Grund und
                           der Weg (performCancel blockt zusätzlich, auch für
                           den Auto-Cancel-Deep-Link aus der Mail). */
                        <span style={{
                          fontSize: '0.8rem', color: 'var(--dex-orange-dark, #b35a00)',
                          padding: '8px 12px', borderRadius: 8,
                          background: 'var(--dex-orange-light, #fff3e0)',
                          border: '1px solid var(--dex-orange, #ed8b00)',
                          lineHeight: 1.4,
                        }}>
                          {selfCancelLockReason(event) === 'always'
                            ? (isDe
                              ? 'Bei diesem Event ist die Selbst-Abmeldung deaktiviert. Bitte wende dich zum Abmelden an die Organizer.'
                              : 'Self-cancellation is disabled for this event. Please contact the organizers to cancel.')
                            : (isDe
                              ? 'Die Abmeldefrist ist abgelaufen — bei diesem Event ist eine Selbst-Abmeldung danach nicht mehr möglich. Bitte wende dich an die Organizer.'
                              : 'The cancellation deadline has passed — for this event self-cancellation is no longer possible after the deadline. Please contact the organizers.')}
                        </span>
                      ) : (
                        <>
                      <button
                        className={`btn dex-cancel-btn${cancellingId === event.id ? ' dex-cancel-btn--armed' : ''}`}
                        onClick={() => handleCancel(event.id)}
                        disabled={isCancelling}
                        style={{
                          fontSize: '0.95rem',
                          fontWeight: 600,
                          padding: '10px 20px',
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: 8,
                          background: cancellingId === event.id ? 'var(--dex-red)' : '#fff',
                          color: cancellingId === event.id ? '#fff' : 'var(--dex-red)',
                          border: `2px solid var(--dex-red)`,
                          borderRadius: 8,
                          boxShadow: cancellingId === event.id ? '0 2px 8px rgba(218,41,28,0.3)' : 'none',
                          cursor: isCancelling ? 'not-allowed' : 'pointer',
                        }}
                      >
                        <X size={16} />
                        {/* v30.20: Bei Kalender-Events sagt der Knopf, WAS er
                            abmeldet — er kappt die GANZE Buchung (alle Tage).
                            Nutzer-Befund: „man versteht den großen Abmelde-
                            Button nicht" — er wirkte wie der Weg, EINEN Tag
                            abzumelden. Einzelne Tage laufen über den Kalender
                            (Klick auf grünen Tag + Bestätigung, s.u.). */}
                        {cancellingId === event.id
                          ? (isCancelling ? '...' : t('myevents.confirmcancel'))
                          : (event.subEventCalendar
                            ? (isDe ? 'Alle Termine abmelden' : 'Cancel all dates')
                            : t('myevents.cancel'))}
                      </button>
                      {cancellingId === event.id && !isCancelling && (
                        <button className="btn btn-secondary" onClick={() => setCancellingId(null)} style={{ fontSize: '0.85rem', padding: '8px 16px' }}>{t('myevents.keepreg')}</button>
                      )}
                      {event.subEventCalendar && cancellingId !== event.id && (
                        <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', lineHeight: 1.4 }}>
                          {isDe
                            ? 'Einzelne Tage meldest du unten im Kalender ab: auf den grünen Tag klicken und bestätigen.'
                            : 'To cancel a single day, use the calendar below: click the green day and confirm.'}
                        </span>
                      )}
                        </>
                      )}
                    </div>
                  </div>
                )}
              </div>
            );
}
