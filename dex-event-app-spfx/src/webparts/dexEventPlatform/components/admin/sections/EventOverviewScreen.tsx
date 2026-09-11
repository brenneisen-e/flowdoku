/* EventOverviewScreen — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 4730-5021 des Stands
 * vor dem Schnitt). Der Inhalt war bis v31.2 zeichengleich übernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3: Nach docs/ui-leitfaden.md 5a umgebaut. Warum: Die ganze Karte trug
 * `cursor: pointer`, geöffnet hat aber nur ihre linke Hälfte; der Status hing
 * als farbige Leiste dran, seine Bedeutung in einer Legende darüber
 * (nachschlagen statt lesen). Jetzt: Kopf mit dem einen Primär-Knopf,
 * Werkzeugleiste aus Chips, jede Karte EINE klickbare Fläche mit Status-Pille,
 * Vergangenes hinter einem Aufklapper.
 */
import * as React from 'react';
import { AlertCircle, Calendar, ChevronDown, Pin, Plus, Search, Trash2, X } from '../../Icons';
import { cx, ensureDexUiStyles } from '../../dexUi';
import { DeloitteEvent } from '../../../types';
import { formatDate } from '../../../utils/eventStatus';
import OrganizerList from '../../OrganizerList';

export interface EventOverviewScreenProps {
  adminEvents: DeloitteEvent[];
  archiveBusyId: string;
  archivedCount: number;
  archivedEventIds: Set<string>;
  changeLogModal: React.ReactElement | null;
  currentEvents: DeloitteEvent[];
  dangerZoneModal: React.ReactElement | null;
  deletingId: string;
  draftCount: number;
  /** v31.27: Suchtext der Eventuebersicht — grenzt die Liste ein, statt zu springen. */
  eventQuery: string;
  setEventQuery: React.Dispatch<React.SetStateAction<string>>;
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
  const { adminEvents, archiveBusyId, archivedCount, archivedEventIds, changeLogModal, currentEvents, dangerZoneModal, deletingId, draftCount, eventQuery, setEventQuery, eventSortMode, handleArchiveEvent, handleSelectEvent, handleUnarchiveEvent, hideDrafts, isAdmin, isDe, isDeleting, isEventsLoading, isPastEvent, locale, navigate, pastEvents, setConfirmDeleteEvent, setConfirmDeleteText, setEventSortMode, setHideDrafts, setShowArchivedEvents, setShowPastEvents, showArchivedEvents, showPastEvents, t } = p;
  // v31.3: Idempotent — die Übersicht steht ohne Modal und ohne WizardFormShell
  // auf der Seite; ohne diesen Aufruf fehlten ihr die dex-ui-Klassen.
  ensureDexUiStyles();
  return (
      <div className="page-container" role="main" style={{ maxWidth: 1200, marginLeft: 'auto', marginRight: 'auto' }}>
        <style>{`@keyframes dex-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
        <div className="dex-ui-page-head">
          <div style={{ minWidth: 0 }}>
            <h2 className="dex-ui-page-head-title">{locale === 'de' ? 'Organizer – Eventübersicht' : 'Organizer – Event overview'}</h2>
            {/* v31.3: Was ein Klick auf ein Event tut. `dex-ui-muted` statt
                `-page-head-meta` — das ist ein Erklärsatz, keine Meta-Zeile. */}
            <div className="dex-ui-muted" style={{ marginTop: 4 }}>
              {isDe
                ? 'Klick ein Event an, um Teilnehmer, Mails und Check-in zu verwalten.'
                : 'Pick an event to manage attendees, emails and check-in.'}
            </div>
          </div>
          {/* v23.44: Nur noch „Neues Event erstellen" — Teilnehmer, Prozesse,
              Audit-Log und SharePoint-Liste sind Admin-Funktionen und leben jetzt
              zentral in der Admin-Kachel (admin-hub). */}
          <div className="dex-ui-page-head-actions">
            <button className="btn btn-primary" onClick={() => navigate('create-event')} style={{ fontSize: '0.85rem' }}>
              <Plus size={16} /> {t('admin.newevent')}
            </button>
          </div>
        </div>

        {isEventsLoading ? (
          <div className="dex-ui-card" style={{ padding: 40, textAlign: 'center' }}>
            <div style={{ display: 'flex', justifyContent: 'center', marginBottom: 14 }}>
              <div style={{
                width: 40, height: 40, border: '4px solid var(--dex-gray-200)',
                borderTop: '4px solid var(--dex-green)', borderRadius: '50%',
                animation: 'dexOrbSpin 1s linear infinite',
              }} />
            </div>
            {/* v31.3: Der Ladetext stand nur auf Deutsch da. */}
            <p style={{ margin: 0, color: 'var(--dex-gray-500)' }}>{isDe ? 'Events werden geladen …' : 'Loading events …'}</p>
          </div>
        ) : adminEvents.length === 0 ? (
          <div className="dex-ui-empty">
            <div className="dex-ui-empty-icon"><Calendar size={20} /></div>
            <div className="dex-ui-empty-title">{t('admin.noevents')}</div>
            {/* v31.3: Kein zweiter Primär-Knopf — der eine steht oben im Kopf. */}
            <p style={{ margin: '0 0 14px' }}>
              {isDe
                ? 'Leg dein erstes Event an — danach verwaltest du hier Teilnehmer, Mails und Check-in.'
                : 'Create your first event — then you manage attendees, emails and check-in here.'}
            </p>
            <button className="btn btn-outline" onClick={() => navigate('create-event')}>
              {t('create.submit')}
            </button>
          </div>
        ) : (
          <>
          {(() => {
            const renderEventCard = (event: DeloitteEvent, opts?: { muted?: boolean }): React.ReactElement => {
              // v24.7 (Q/R) → v31.3: Aus der farbigen Leiste (blau vergangen,
              // orange Entwurf, grün aktiv) mit Legende über der Liste ist eine
              // beschriftete Pille neben dem Titel geworden — lesbar statt
              // nachschlagbar; die Legenden-Erklärung hängt als Tooltip daran.
              const past = isPastEvent(event);
              const archived = archivedEventIds.has(event.id);
              const busy = archiveBusyId === event.id;
              const statusTone = past ? 'blue' : event.isFictive ? 'orange' : 'green';
              const statusLabel = past ? (isDe ? 'Abgeschlossen' : 'Completed') : event.isFictive ? (isDe ? 'Entwurf' : 'Draft') : (isDe ? 'Aktiv' : 'Active');
              const statusHint = past
                ? (isDe ? 'Abgeschlossen (vorbei)' : 'Completed (past)')
                : event.isFictive
                  ? (isDe ? 'Entwurf — für Teilnehmer noch nicht sichtbar' : 'Draft — not yet visible to attendees')
                  : (isDe ? 'Aktiv — für Teilnehmer sichtbar' : 'Active — visible to attendees');
              return (
              // v31.3: EINE klickbare Fläche — vorher öffnete nur die linke
              // Kartenhälfte. Die Knöpfe darin stoppen die Weitergabe (Klick UND
              // Taste), damit Archivieren und Löschen nicht mit öffnen. Kein
              // aria-label: `role="button"` benennt sich aus dem Inhalt (wie die
              // Sub-Event-Karten in Schritt 1); ein Label würde Status, Zahlen
              // und Knöpfe verschlucken.
              <div
                key={event.id}
                role="button"
                tabIndex={0}
                className={cx('dex-ui-card', 'dex-ui-card--hover', opts?.muted && 'dex-ui-card--muted')}
                onClick={() => { void handleSelectEvent(event); }}
                onKeyDown={e => {
                  // Nur mit Fokus auf der Karte selbst — sonst öffnet Enter auf
                  // „Archivieren"/„Löschen" das Event und preventDefault stoppt
                  // obendrein den Klick des Knopfes (Wächter wie in BasicsStep).
                  if (e.target !== e.currentTarget) return;
                  if (e.key !== 'Enter' && e.key !== ' ') return;
                  e.preventDefault();
                  void handleSelectEvent(event);
                }}
                style={{ cursor: 'pointer', display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 14 }}
              >
                {/* v23.42: größeres Thumbnail. */}
                <div style={{
                  width: 84, height: 60, borderRadius: 10, flexShrink: 0,
                  background: event.imageUrl
                    ? `url(${event.imageUrl}) center/cover no-repeat`
                    : 'linear-gradient(135deg, var(--dex-gray-200), var(--dex-gray-100))',
                  filter: opts?.muted ? 'grayscale(0.4)' : 'none',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  color: 'var(--dex-gray-400)', fontSize: '0.7rem',
                }}>
                  {!event.imageUrl && '—'}
                </div>
                <div style={{ flex: '1 1 280px', minWidth: 0 }}>
                  <div className="dex-ui-inline">
                    <h3 style={{ margin: 0, fontSize: '1.05rem' }}>{event.title}</h3>
                    <span className={`dex-ui-pill dex-ui-pill--${statusTone}`} title={statusHint}>{statusLabel}</span>
                    {/* v31.3: Ohne Marke wirken Archivierte wie normale Events. */}
                    {archived && (
                      <span
                        className="dex-ui-pill dex-ui-pill--gray"
                        title={isDe ? 'Nur für dich ausgeblendet — für alle anderen unverändert sichtbar.' : 'Hidden for you only — unchanged for everyone else.'}
                      >{isDe ? 'Archiviert' : 'Archived'}</span>
                    )}
                  </div>
                  {/* v23.42/v31.3: Datum, Ort und Zahlen in EINER Zeile beim Titel. */}
                  <div className="dex-ui-inline" style={{ gap: '4px 14px', marginTop: 5, fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                      <Calendar size={14} />
                      {formatDate(event.startDate)} {isDe ? 'bis' : 'until'} {formatDate(event.endDate)}
                    </span>
                    {event.location && (
                      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                        <Pin size={14} />
                        {event.location}
                      </span>
                    )}
                    <span className="dex-ui-pill dex-ui-pill--gray">
                      {(() => {
                        const split = (event.durchstarterCapacity || 0) + (event.funstarterCapacity || 0);
                        const isSplitEv = (event.durchstarterCapacity || 0) > 0 && (event.funstarterCapacity || 0) > 0;
                        const eff = event.maxParticipants && event.maxParticipants > 0 ? event.maxParticipants : split;
                        const shown = (isSplitEv && eff > 0) ? Math.min(event.currentParticipants, eff) : event.currentParticipants;
                        return `${shown}/${eff || '∞'} ${isDe ? 'Teilnehmer' : 'attendees'}`;
                      })()}
                    </span>
                    {event.waitlistCount > 0 && (
                      <span className="dex-ui-pill dex-ui-pill--orange">
                        {event.waitlistCount} {isDe ? 'auf Warteliste' : 'on waitlist'}
                      </span>
                    )}
                  </div>
                  {/* v23.42: Organizer mit Foto + Hover (wie Anmeldeseite). */}
                  <div onClick={e => e.stopPropagation()} style={{ marginTop: 6 }}>
                    <OrganizerList
                      names={event.organizers.reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map(o => o.trim()).filter(Boolean)}
                      emails={event.organizerEmails}
                      size="sm"
                      compact
                    />
                  </div>
                </div>
                {/* v31.3: Rechts außen nur, was WEG vom Inhalt führt (Leitfaden 2a′). */}
                <div className="dex-ui-inline" style={{ marginLeft: 'auto', flexShrink: 0, gap: 2 }}>
                  {/* v24.6: abgelaufene Events aus der EIGENEN Übersicht aus-/einblenden. */}
                  {past && (
                    archived ? (
                      <button
                        type="button"
                        className="dex-ui-textbtn dex-ui-textbtn--muted"
                        title={isDe ? 'Wieder in meiner Übersicht anzeigen.' : 'Show in my overview again.'}
                        disabled={busy}
                        onClick={e => { e.stopPropagation(); void handleUnarchiveEvent(event); }}
                      >{busy ? '…' : (isDe ? 'Einblenden' : 'Unhide')}</button>
                    ) : (
                      <button
                        type="button"
                        className="dex-ui-textbtn dex-ui-textbtn--muted"
                        title={isDe ? 'Aus meiner Übersicht ausblenden — das Event bleibt erhalten und für andere sichtbar.' : 'Hide from my overview — the event is kept and stays visible to others.'}
                        disabled={busy}
                        onClick={e => { e.stopPropagation(); void handleArchiveEvent(event); }}
                      >{busy ? '…' : (isDe ? 'Archivieren' : 'Archive')}</button>
                    )
                  )}
                  {/* v18.3: Demo-Event hat keinen Löschen-Button (kein Backend). */}
                  {!event.isDemoShowcase && (
                    <button
                      type="button"
                      className="dex-ui-textbtn dex-ui-textbtn--danger"
                      title={isDe ? 'Event löschen — du wirst vorher gefragt.' : 'Delete the event — you are asked first.'}
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
              );
            };
            return (
              <>
                {/* v24.11/v31.3: Abgelaufene Entwürfe stehen jetzt GANZ OBEN — was
                    Handeln verlangt, kommt vor Sortierung und Liste (Leitfaden 5a);
                    Löschen weiter über die übliche Rückfrage. */}
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
                    <div className="dex-ui-callout dex-ui-callout--warn" style={{ flexDirection: 'column', alignItems: 'stretch', gap: 10, marginBottom: 14 }}>
                      <div style={{ display: 'flex', gap: 10, alignItems: 'flex-start' }}>
                        <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                        <div style={{ minWidth: 0 }}>
                          <div style={{ fontWeight: 700 }}>{isDe ? 'Abgelaufene Entwürfe aufräumen' : 'Clean up expired drafts'}</div>
                          <div style={{ marginTop: 2 }}>
                            {isDe
                              ? <><strong>{stale.length}</strong> {stale.length === 1 ? 'Entwurf ist' : 'Entwürfe sind'} abgelaufen (Datum vorbei, nie aktiviert) und {stale.length === 1 ? 'kann' : 'können'} gelöscht werden.</>
                              : <><strong>{stale.length}</strong> draft(s) expired (date passed, never activated) and can be deleted.</>}
                          </div>
                        </div>
                      </div>
                      {/* Kein Zeilen-Hover: die Zeile selbst tut nichts, nur ihr Knopf. */}
                      <div className="dex-ui-stack" style={{ gap: 4, paddingLeft: 26 }}>
                        {stale.map(ev => (
                          <div key={ev.id} style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                            <span style={{ flex: 1, minWidth: 0, fontSize: '0.8rem', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{ev.title || (isDe ? 'Ohne Titel' : 'Untitled')}</span>
                            <button
                              type="button"
                              className="dex-ui-textbtn dex-ui-textbtn--danger"
                              style={{ flexShrink: 0 }}
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
                {/* v18.2/v31.3: Sortierung und Filter als Werkzeugleiste statt
                    Dropdown mit zwei Einträgen und Checkboxen (Leitfaden 2b). */}
                {/* v31.27: Suchzeile ueber Sortierung und Filter — sie grenzt
                    die Liste ein, waehrend die globale Suche im Kopf direkt in
                    ein Event springt. Zwei verschiedene Fragen, deshalb zwei
                    Felder (Leitfaden 2b: Filter vor der Liste, nicht dahinter,
                    wo er unauffindbar waere). */}
                <div className="dex-ui-searchbar" style={{ marginBottom: 10, maxWidth: 420 }}>
                  <span className="dex-ui-searchbar-icon" aria-hidden="true"><Search size={15} /></span>
                  <input
                    type="text"
                    className="dex-ui-input"
                    value={eventQuery}
                    onChange={e => setEventQuery(e.target.value)}
                    placeholder={isDe ? 'Event, Ort, Nummer oder Organizer suchen' : 'Search event, location, number or organizer'}
                    aria-label={isDe ? 'Events durchsuchen' : 'Search events'}
                    style={{ paddingRight: 34 }}
                  />
                  {eventQuery && (
                    <button
                      type="button"
                      className="dex-ui-iconbtn"
                      onClick={() => setEventQuery('')}
                      aria-label={isDe ? 'Suche leeren' : 'Clear search'}
                      style={{ position: 'absolute', right: 3, top: '50%', transform: 'translateY(-50%)', width: 32, height: 32 }}
                    ><X size={14} /></button>
                  )}
                </div>
                <div className="dex-ui-toolbar">
                  <span className="dex-ui-muted" style={{ fontWeight: 600 }}>{isDe ? 'Sortierung' : 'Sort'}</span>
                  <button
                    type="button"
                    className={cx('dex-ui-chip', eventSortMode === 'alpha' && 'is-active')}
                    aria-pressed={eventSortMode === 'alpha'}
                    onClick={() => setEventSortMode('alpha')}
                  >{isDe ? 'Alphabetisch (A–Z)' : 'Alphabetical (A–Z)'}</button>
                  <button
                    type="button"
                    className={cx('dex-ui-chip', eventSortMode === 'date' && 'is-active')}
                    aria-pressed={eventSortMode === 'date'}
                    onClick={() => setEventSortMode('date')}
                  >{isDe ? 'Datum aufsteigend' : 'Date ascending'}</button>
                  {(draftCount > 0 || archivedCount > 0) && (
                    <span className="dex-ui-muted" style={{ fontWeight: 600, marginLeft: 6 }}>{isDe ? 'Filter' : 'Filters'}</span>
                  )}
                  {draftCount > 0 && (
                    <button
                      type="button"
                      className={cx('dex-ui-chip', hideDrafts && 'is-active')}
                      aria-pressed={hideDrafts}
                      title={isDe ? 'Entwürfe sind für Teilnehmer noch nicht sichtbar.' : 'Drafts are not yet visible to attendees.'}
                      onClick={() => setHideDrafts(!hideDrafts)}
                    >{isDe ? `Entwürfe ausblenden (${draftCount})` : `Hide drafts (${draftCount})`}</button>
                  )}
                  {/* v24.6: Archivierte (für mich ausgeblendete) Events einblenden. */}
                  {archivedCount > 0 && (
                    <button
                      type="button"
                      className={cx('dex-ui-chip', showArchivedEvents && 'is-active')}
                      aria-pressed={showArchivedEvents}
                      title={isDe ? 'Archivierte Events sind nur für dich ausgeblendet.' : 'Archived events are hidden for you only.'}
                      onClick={() => setShowArchivedEvents(!showArchivedEvents)}
                    >{isDe ? `Archivierte anzeigen (${archivedCount})` : `Show archived (${archivedCount})`}</button>
                  )}
                  <span className="dex-ui-toolbar-spacer" />
                  <span className="dex-ui-muted">{isDe ? `${currentEvents.length} angezeigt` : `${currentEvents.length} shown`}</span>
                </div>
                {/* v31.3: Farb-Legende entfallen — die Pille je Karte sagt es selbst. */}
                {currentEvents.length === 0 ? (
                  <div className="dex-ui-empty">
                    <div className="dex-ui-empty-icon"><Calendar size={20} /></div>
                    <div className="dex-ui-empty-title">{isDe ? 'Keine Events in dieser Ansicht' : 'No events in this view'}</div>
                    {isDe
                      ? 'Vielleicht blendet ein Filter oben sie aus — oder du legst ein neues Event an.'
                      : 'A filter above may be hiding them — or create a new event.'}
                  </div>
                ) : (
                  /* v31.3: `card` ist hier kein Aussehen, sondern der Griff der
                     Organizer-Tour (Selektor `.page-container .card`) — ohne die
                     Klasse pollt ihr Schritt „Deine Event-Liste" vier Sekunden ins
                     Leere. Die Inline-Werte nehmen der Klasse ihre Optik (Polster
                     32, Radius 16); `data-tour` liegt für einen zentralen Umstieg
                     des Selektors bereit. */
                  <div className="card dex-ui-stack" data-tour="admin-event-list" style={{ gap: 12, padding: 0, background: 'transparent', border: 'none', borderRadius: 0, boxShadow: 'none', overflow: 'visible' }}>
                    {currentEvents.map(ev => renderEventCard(ev))}
                  </div>
                )}
                {/* v31.3: Aufklapper statt gestricheltem Kasten — und in beiden
                    Sprachen: Beschriftung, Zähler und „Klicken zum Ausklappen"
                    standen nur auf Deutsch da; den Satz ersetzt der Chevron. */}
                {isAdmin && pastEvents.length > 0 && (
                  <div style={{ marginTop: 22 }}>
                    <button
                      type="button"
                      className={cx('dex-ui-disclosure', showPastEvents && 'is-open')}
                      aria-expanded={showPastEvents}
                      onClick={() => setShowPastEvents(!showPastEvents)}
                    >
                      <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                      {isDe ? 'Vergangene Events' : 'Past events'}
                      <span className="dex-ui-disclosure-count">{pastEvents.length}</span>
                    </button>
                    {showPastEvents && (
                      <div className="dex-ui-disclosure-body dex-ui-stack" style={{ gap: 12, marginTop: 4 }}>
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

