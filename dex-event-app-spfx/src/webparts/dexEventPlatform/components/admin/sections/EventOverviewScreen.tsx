/* EventOverviewScreen — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 4730-5021 des Stands
 * vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { Calendar, Pin, Plus, Trash2 } from '../../Icons';
import { DeloitteEvent } from '../../../types';
import { formatDate } from '../../../utils/eventStatus';
import OrganizerList from '../../OrganizerList';

export interface EventOverviewScreenProps {
  adminEvents: DeloitteEvent[];
  archiveBusyId: string;
  archivedCount: number;
  archivedEventIds: Set<string>;
  changeLogModal: React.ReactElement<any, string | React.JSXElementConstructor<any>>;
  currentEvents: DeloitteEvent[];
  dangerZoneModal: React.ReactElement<any, string | React.JSXElementConstructor<any>>;
  deletingId: string;
  draftCount: number;
  eventSortMode: "alpha" | "date";
  handleArchiveEvent: (event: DeloitteEvent) => Promise<void>;
  handleSelectEvent: (event: DeloitteEvent) => Promise<void>;
  handleUnarchiveEvent: (event: DeloitteEvent) => Promise<void>;
  hideDrafts: boolean;
  isAdmin: boolean;
  isDe: boolean;
  isDeleting: boolean;
  isEventsLoading: boolean;
  isPastEvent: (e: DeloitteEvent) => boolean;
  locale: import("../../../context/LanguageContext").Locale;
  navigate: (page: import("../../../context/NavigationContext").Page, eventId?: string, intent?: import("../../../context/NavigationContext").NavIntent) => void;
  pastEvents: DeloitteEvent[];
  setConfirmDeleteEvent: React.Dispatch<React.SetStateAction<DeloitteEvent>>;
  setConfirmDeleteText: React.Dispatch<React.SetStateAction<string>>;
  setEventSortMode: React.Dispatch<React.SetStateAction<"alpha" | "date">>;
  setHideDrafts: React.Dispatch<React.SetStateAction<boolean>>;
  setShowArchivedEvents: React.Dispatch<React.SetStateAction<boolean>>;
  setShowPastEvents: React.Dispatch<React.SetStateAction<boolean>>;
  showArchivedEvents: boolean;
  showPastEvents: boolean;
  t: (key: string) => string;
}

export const EventOverviewScreen: React.FC<EventOverviewScreenProps> = (p) => {
  const { adminEvents, archiveBusyId, archivedCount, archivedEventIds, changeLogModal, currentEvents, dangerZoneModal, deletingId, draftCount, eventSortMode, handleArchiveEvent, handleSelectEvent, handleUnarchiveEvent, hideDrafts, isAdmin, isDe, isDeleting, isEventsLoading, isPastEvent, locale, navigate, pastEvents, setConfirmDeleteEvent, setConfirmDeleteText, setEventSortMode, setHideDrafts, setShowArchivedEvents, setShowPastEvents, showArchivedEvents, showPastEvents, t } = p;
  return (
      <div className="page-container" role="main" style={{ maxWidth: 1200, marginLeft: 'auto', marginRight: 'auto' }}>
        <style>{`@keyframes dex-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
        <h2 style={{ margin: '0 0 16px' }}>{locale === 'de' ? 'Organizer – Eventübersicht' : 'Organizer – Event overview'}</h2>
        {/* v23.44: Nur noch „Neues Event erstellen" — Teilnehmer, Prozesse,
            Audit-Log und SharePoint-Liste sind Admin-Funktionen und leben jetzt
            zentral in der Admin-Kachel (admin-hub). */}
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 16 }}>
          <button className="btn btn-primary" onClick={() => navigate('create-event')} style={{ fontSize: '0.85rem' }}>
            <Plus size={16} /> {t('admin.newevent')}
          </button>
        </div>

        {isEventsLoading ? (
          <div className="card text-center" style={{ padding: 48 }}>
            <div style={{ display: 'flex', justifyContent: 'center', marginBottom: 16 }}>
              <div style={{
                width: 40, height: 40, border: '4px solid var(--dex-gray-200)',
                borderTop: '4px solid var(--dex-green)', borderRadius: '50%',
                animation: 'dexOrbSpin 1s linear infinite',
              }} />
            </div>
            <p style={{ color: 'var(--dex-gray-400)' }}>Events werden geladen...</p>
          </div>
        ) : adminEvents.length === 0 ? (
          <div className="card text-center" style={{ padding: 48 }}>
            <p style={{ color: 'var(--dex-gray-400)' }}>{t('admin.noevents')}</p>
            <button className="btn btn-primary mt-24" onClick={() => navigate('create-event')}>
              {t('create.submit')}
            </button>
          </div>
        ) : (
          <>
          {(() => {
            const renderEventCard = (event: DeloitteEvent, opts?: { muted?: boolean }): React.ReactElement => {
              // v24.7 (Q/R): moderne Status-Leiste links über die volle Karten-
              // höhe (statt Eck-Winkel). Farbe: blau = abgeschlossen/vergangen,
              // orange = Entwurf, grün = aktiv. Bedeutung erklärt die Legende.
              const past = isPastEvent(event);
              const statusColor = past ? 'var(--dex-blue, #0076a8)' : event.isFictive ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)';
              const statusLabel = past ? (isDe ? 'Abgeschlossen' : 'Completed') : event.isFictive ? (isDe ? 'Entwurf' : 'Draft') : (isDe ? 'Aktiv' : 'Active');
              return (
              <div
                key={event.id}
                className="card card-clickable"
                style={{ position: 'relative', padding: '26px 24px 22px 28px', cursor: 'pointer', opacity: opts?.muted ? 0.85 : 1, overflow: 'hidden' }}
              >
                <span
                  title={statusLabel}
                  style={{ position: 'absolute', top: 0, bottom: 0, left: 0, width: 6, borderTopLeftRadius: 'var(--dex-radius)', borderBottomLeftRadius: 'var(--dex-radius)', background: statusColor }}
                />
                <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: 12 }}>
                  <div onClick={() => handleSelectEvent(event)} style={{ flex: '1 1 260px', display: 'flex', alignItems: 'center', gap: 16, cursor: 'pointer' }}>
                    {/* v23.42: größeres Thumbnail. */}
                    <div style={{
                      width: 84, height: 60, borderRadius: 'var(--dex-radius)', flexShrink: 0,
                      background: event.imageUrl
                        ? `url(${event.imageUrl}) center/cover no-repeat`
                        : 'linear-gradient(135deg, var(--dex-gray-200), var(--dex-gray-100))',
                      filter: opts?.muted ? 'grayscale(0.4)' : 'none',
                      display: 'flex', alignItems: 'center', justifyContent: 'center',
                      color: 'var(--dex-gray-400)', fontSize: '0.7rem',
                    }}>
                      {!event.imageUrl && '—'}
                    </div>
                    <div style={{ minWidth: 0 }}>
                      <h3 style={{ margin: '0 0 6px' }}>{event.title}</h3>
                      {/* v23.42: Datum + Ort mit Icon (wie auf der Anmeldeseite). */}
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
                        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                          <Calendar size={14} />
                          {formatDate(event.startDate)} {isDe ? 'bis' : 'until'} {formatDate(event.endDate)}
                        </span>
                        {event.location && (
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                            <Pin size={14} />
                            {event.location}
                          </span>
                        )}
                        {/* v23.42: Organizer mit Foto + Hover (wie Anmeldeseite). */}
                        <div onClick={e => e.stopPropagation()} style={{ marginTop: 4 }}>
                          <OrganizerList
                            names={event.organizers.reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map(o => o.trim()).filter(Boolean)}
                            emails={event.organizerEmails}
                            size="sm"
                            compact
                          />
                        </div>
                      </div>
                    </div>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                    {/* v23.42: Teilnehmerzahl + Warteliste UNTEREINANDER. */}
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 2 }}>
                      <span style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                        {(() => {
                          const split = (event.durchstarterCapacity || 0) + (event.funstarterCapacity || 0);
                          const isSplitEv = (event.durchstarterCapacity || 0) > 0 && (event.funstarterCapacity || 0) > 0;
                          const eff = event.maxParticipants && event.maxParticipants > 0 ? event.maxParticipants : split;
                          const shown = (isSplitEv && eff > 0) ? Math.min(event.currentParticipants, eff) : event.currentParticipants;
                          return `${shown}/${eff || '∞'} ${isDe ? 'Teilnehmer' : 'attendees'}`;
                        })()}
                      </span>
                      {event.waitlistCount > 0 && (
                        <span style={{ fontSize: '0.8rem', color: 'var(--dex-orange, #ed8b00)', fontWeight: 600 }}>
                          {event.waitlistCount} {isDe ? 'auf Warteliste' : 'on waitlist'}
                        </span>
                      )}
                    </div>
                    {/* v24.6: abgelaufene Events aus der EIGENEN Übersicht aus-/einblenden. */}
                    {isPastEvent(event) && (
                      archivedEventIds.has(event.id) ? (
                        <button
                          type="button"
                          className="btn btn-secondary"
                          style={{ fontSize: '0.74rem', padding: '4px 10px' }}
                          disabled={archiveBusyId === event.id}
                          onClick={e => { e.stopPropagation(); void handleUnarchiveEvent(event); }}
                        >{archiveBusyId === event.id ? '…' : (isDe ? 'Einblenden' : 'Unhide')}</button>
                      ) : (
                        <button
                          type="button"
                          className="btn btn-secondary"
                          style={{ fontSize: '0.74rem', padding: '4px 10px' }}
                          title={isDe ? 'Aus meiner Übersicht ausblenden — das Event bleibt erhalten und für andere sichtbar.' : 'Hide from my overview — the event is kept and stays visible to others.'}
                          disabled={archiveBusyId === event.id}
                          onClick={e => { e.stopPropagation(); void handleArchiveEvent(event); }}
                        >{archiveBusyId === event.id ? '…' : (isDe ? 'Archivieren' : 'Archive')}</button>
                      )
                    )}
                    {/* v18.3: Demo-Event hat keinen Löschen-Button (kein Backend). */}
                    {!event.isDemoShowcase && (
                      <button
                        className="btn btn-secondary"
                        style={{ fontSize: '0.8rem', padding: '6px 12px', color: 'var(--dex-red, #c00)' }}
                        onClick={(e) => {
                          e.stopPropagation();
                          setConfirmDeleteEvent(event);
                          setConfirmDeleteText('');
                        }}
                        disabled={isDeleting}
                      >
                        <Trash2 size={14} /> {isDeleting && deletingId === event.id ? (isDe ? 'Wird gelöscht...' : 'Deleting...') : (isDe ? 'Löschen' : 'Delete')}
                      </button>
                    )}
                  </div>
                </div>
              </div>
              );
            };
            return (
              <>
                {/* v18.2: Sortier- + Entwurf-Filter-Leiste über der Event-Liste. */}
                <div style={{
                  display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap',
                  marginBottom: 16, padding: '10px 14px',
                  background: 'var(--dex-gray-50, #fafafa)', borderRadius: 'var(--dex-radius, 12px)',
                  border: '1px solid var(--dex-gray-200)',
                }}>
                  <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                    <span style={{ fontWeight: 600 }}>{isDe ? 'Sortierung:' : 'Sort:'}</span>
                    <select
                      className="form-select"
                      value={eventSortMode}
                      onChange={e => setEventSortMode(e.target.value as 'alpha' | 'date')}
                      style={{ fontSize: '0.85rem', padding: '4px 34px 4px 10px', minWidth: 210, width: 'auto' }}
                    >
                      <option value="alpha">{isDe ? 'Alphabetisch (A–Z)' : 'Alphabetical (A–Z)'}</option>
                      <option value="date">{isDe ? 'Datum aufsteigend' : 'Date ascending'}</option>
                    </select>
                  </label>
                  {draftCount > 0 && (
                    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', color: 'var(--dex-gray-700)', cursor: 'pointer' }}>
                      <input
                        type="checkbox"
                        checked={hideDrafts}
                        onChange={e => setHideDrafts(e.target.checked)}
                        style={{ accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
                      />
                      {isDe ? `Entwürfe ausblenden (${draftCount})` : `Hide drafts (${draftCount})`}
                    </label>
                  )}
                  {/* v24.6: Archivierte (für mich ausgeblendete) Events einblenden. */}
                  {archivedCount > 0 && (
                    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', color: 'var(--dex-gray-700)', cursor: 'pointer' }}>
                      <input
                        type="checkbox"
                        checked={showArchivedEvents}
                        onChange={e => setShowArchivedEvents(e.target.checked)}
                        style={{ accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
                      />
                      {isDe ? `Archivierte anzeigen (${archivedCount})` : `Show archived (${archivedCount})`}
                    </label>
                  )}
                </div>
                {/* v23.44/v24.7: Farb-Legende für die Status-Leiste links an den Karten. */}
                <div style={{ display: 'flex', alignItems: 'center', gap: 18, margin: '0 4px 10px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', flexWrap: 'wrap' }}>
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 12, height: 12, borderRadius: 3, background: 'var(--dex-green, #86bc25)', display: 'inline-block' }} />
                    {isDe ? 'Aktiv (für Teilnehmer sichtbar)' : 'Active (visible to attendees)'}
                  </span>
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 12, height: 12, borderRadius: 3, background: 'var(--dex-orange, #ed8b00)', display: 'inline-block' }} />
                    {isDe ? 'Entwurf (noch nicht sichtbar)' : 'Draft (not yet visible)'}
                  </span>
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 12, height: 12, borderRadius: 3, background: 'var(--dex-blue, #0076a8)', display: 'inline-block' }} />
                    {isDe ? 'Abgeschlossen (vorbei)' : 'Completed (past)'}
                  </span>
                </div>
                {/* v24.11: Hinweis auf abgelaufene Entwürfe (wie auf der Startseite),
                    jetzt auch hier im Organizer Center. Löschen über die übliche
                    Sicherheitsabfrage (Entwürfe = einfaches Ja). */}
                {(() => {
                  const dayMs = 24 * 60 * 60 * 1000;
                  const nowTs = Date.now();
                  const stale = adminEvents.filter(e => {
                    if (e.parentEventId || !e.isFictive) return false;
                    const endRaw = e.endDate || e.startDate;
                    const endTs = endRaw ? new Date(endRaw).getTime() : 0;
                    return endTs > 0 && endTs < nowTs - dayMs;
                  });
                  if (stale.length === 0) return null;
                  return (
                    <div style={{ margin: '0 4px 14px', padding: '12px 16px', background: '#fff', border: '1px solid rgba(237,139,0,0.5)', borderRadius: 12 }}>
                      <div style={{ fontWeight: 800, fontSize: '0.88rem', color: 'var(--dex-gray-800)', marginBottom: 4 }}>
                        {isDe ? 'Abgelaufene Entwürfe aufräumen' : 'Clean up expired drafts'}
                      </div>
                      <p style={{ margin: '0 0 8px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                        {isDe
                          ? <><strong>{stale.length}</strong> {stale.length === 1 ? 'Entwurf ist' : 'Entwürfe sind'} abgelaufen (Datum vorbei, nie aktiviert) und {stale.length === 1 ? 'kann' : 'können'} gelöscht werden.</>
                          : <><strong>{stale.length}</strong> draft(s) expired (date passed, never activated) and can be deleted.</>}
                      </p>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                        {stale.map(ev => (
                          <div key={ev.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
                            <span style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-800)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{ev.title || (isDe ? 'Ohne Titel' : 'Untitled')}</span>
                            <button
                              className="btn btn-secondary"
                              style={{ fontSize: '0.74rem', padding: '4px 10px', color: 'var(--dex-red, #c00)', flexShrink: 0 }}
                              onClick={() => { setConfirmDeleteEvent(ev); setConfirmDeleteText(''); }}
                            >
                              {isDe ? 'Löschen' : 'Delete'}
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })()}
                <div className="my-events-list">
                  {currentEvents.map(ev => renderEventCard(ev))}
                </div>
                {isAdmin && pastEvents.length > 0 && (
                  <div style={{ marginTop: 24 }}>
                    <button
                      type="button"
                      onClick={() => setShowPastEvents(!showPastEvents)}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        width: '100%', padding: '12px 20px',
                        background: 'var(--dex-gray-50, #f8f9fa)',
                        border: '1px dashed var(--dex-gray-300)',
                        borderRadius: 'var(--dex-radius, 12px)',
                        cursor: 'pointer',
                        fontSize: '0.9rem', fontWeight: 600,
                        color: 'var(--dex-gray-700)',
                        textAlign: 'left',
                      }}
                    >
                      <span style={{ fontSize: '1rem' }}>{showPastEvents ? '▾' : '▸'}</span>
                      <span style={{ flex: 1 }}>
                        Vergangene Events ({pastEvents.length})
                      </span>
                      <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                        Klicken zum {showPastEvents ? 'Einklappen' : 'Ausklappen'}
                      </span>
                    </button>
                    {showPastEvents && (
                      <div className="my-events-list" style={{ marginTop: 12 }}>
                        {pastEvents.map(ev => renderEventCard(ev, { muted: true }))}
                      </div>
                    )}
                  </div>
                )}
              </>
            );
          })()}
          </>
        )}
        {dangerZoneModal}
        {changeLogModal}
      </div>
  );
};

