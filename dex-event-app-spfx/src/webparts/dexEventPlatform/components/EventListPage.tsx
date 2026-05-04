/**
 * Event-Uebersicht
 *
 * Zeigt Events als Karten an.
 * Standort + Zielgruppen-Filter: User sehen nur passende Events.
 * Admin/Organizer sehen alle Events.
 */

import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useNavigation } from '../context/NavigationContext';
import { DeloitteEvent } from '../types';
import { useLanguage } from '../context/LanguageContext';
import { RefreshCw } from './Icons';
import { Icon } from '@fluentui/react/lib/Icon';
import EventCard from './EventCard';

/**
 * Prüft ob ein User-Standort zu einem LocationFilter passt.
 * "DE - Koeln" matcht "Köln" oder "Koeln".
 */
function matchesLocation(userLocation: string, locationFilters: string[]): boolean {
  if (locationFilters.length === 0) return true;
  const filters = locationFilters.map(s => s.trim().toLowerCase());
  if (filters.indexOf('all') >= 0) return true;
  if (!userLocation) return false;
  const loc = userLocation.toLowerCase();
  return filters.some(f => {
    const normalized = f.replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ä/g, 'ae').replace(/ß/g, 'ss');
    return loc.indexOf(f) >= 0 || loc.indexOf(normalized) >= 0;
  });
}

/**
 * Prüft ob ein User zur Zielgruppe passt.
 * Zielgruppe kann Verteilergruppen (DEALL, SAPALL) oder Einzelmails sein.
 * Einzelmails: exakter Match auf User-Email.
 * Gruppen: werden gegen User-Email-Prefix gematcht (z.B. DEALL matcht alle @deloitte.de).
 */
function matchesAudience(userEmail: string, userLocation: string, audienceFilters: string[]): boolean {
  if (audienceFilters.length === 0) return true;
  const email = userEmail.toLowerCase();
  const loc = (userLocation || '').toLowerCase();

  return audienceFilters.some(filter => {
    const f = filter.trim().toLowerCase();
    if (!f) return false;

    // Einzelne E-Mail-Adresse
    if (f.indexOf('@') >= 0) {
      return email === f;
    }

    // Gruppen-Patterns
    if (f === 'all' || f === 'deall') return true; // Alle Mitarbeiter
    // Standort-Gruppen: DEKOELN, DEHAMBURG, etc.
    if (f.startsWith('de')) {
      const city = f.substring(2);
      const normalized = city.replace(/oe/g, 'ö').replace(/ue/g, 'ü').replace(/ae/g, 'ä');
      return loc.indexOf(city) >= 0 || loc.indexOf(normalized) >= 0;
    }
    // Abteilungs-Gruppen: SAPALL, TECHALL, etc. (Zukunft: Graph API Integration)
    return false;
  });
}

/**
 * Audience-Liste normalisieren: 'All'/'DEALL' bedeuten "kein Audience-Filter"
 * und werden weggefiltert. So wird "München + All" korrekt als "nur München"
 * interpretiert (statt 'All' als Override fuer alle User).
 */
function normalizeAudience(audience: string[]): string[] {
  return audience
    .map(s => s.trim())
    .filter(s => s && s.toLowerCase() !== 'all' && s.toLowerCase() !== 'deall');
}

/**
 * Prüft ob ein Event fuer den User sichtbar ist.
 *
 * Default: AND (Schnittmenge). Wenn Standort UND Zielgruppe gesetzt sind,
 * muss BEIDES passen (z.B. "München + SAP" = nur Munich-SAP-Mitarbeiter).
 * Mit filterMode='OR' wird stattdessen die Vereinigung verwendet (z.B.
 * "Köln-Mitarbeiter ODER explizit eingeladene Gäste").
 */
function isEventVisibleForUser(
  event: DeloitteEvent,
  userEmail: string,
  userLocation: string
): boolean {
  // v8.6: Exclude-Liste hat Vorrang. Wer hier drin ist, sieht das Event NIE
  // — egal ob er ueber Standortfilter oder Verteiler-Mitgliedschaft sonst
  // sichtbar waere.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const excluded = ((event as any).excludedUsers || []) as string[];
  if (excluded.length > 0 && userEmail) {
    const emailLc = userEmail.toLowerCase().trim();
    if (excluded.some(e => (e || '').toLowerCase().trim() === emailLc)) return false;
  }

  const hasLocationFilter = event.locationAudience.length > 0;
  const normalizedAud = normalizeAudience(event.audienceFilter);
  const hasAudienceFilter = normalizedAud.length > 0;

  if (!hasLocationFilter && !hasAudienceFilter) return true;

  const locMatch = matchesLocation(userLocation, event.locationAudience);
  const audMatch = matchesAudience(userEmail, userLocation, normalizedAud);

  // Default = AND. Nur wenn explizit OR konfiguriert wird Union genutzt.
  if (event.filterMode === 'OR') {
    if (hasLocationFilter && hasAudienceFilter) return locMatch || audMatch;
    if (hasLocationFilter) return locMatch;
    return audMatch;
  }
  // AND (Default): beide Filter muessen passen, falls beide gesetzt
  if (hasLocationFilter && hasAudienceFilter) return locMatch && audMatch;
  if (hasLocationFilter) return locMatch;
  return audMatch;
}

export default function EventListPage(): React.ReactElement {
  // Seit v6.4: nur Top-Level-Events anzeigen. Sub-Events (parentEventId gesetzt)
  // erscheinen im Details-View des Parents (RegistrationPage), nicht eigenständig.
  const { topLevelEvents: events, isEventsLoading, getMyEventNumbers, refreshEvents, testTeamEmails } = useEvents();
  const [isRefreshing, setIsRefreshing] = React.useState(false);
  const handleRefresh = async (): Promise<void> => {
    if (isRefreshing) return;
    setIsRefreshing(true);
    try { await refreshEvents(); } finally { setIsRefreshing(false); }
  };
  const { currentUser } = useCurrentUser();
  const { isAdmin } = useRoles();
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  const [onlyActive, setOnlyActive] = React.useState(true);
  // View-Mode (Cards | List) - persistiert in localStorage
  const [viewMode, setViewMode] = React.useState<'cards' | 'list'>(() => {
    try { return (localStorage.getItem('dex-eventlist-view') as 'cards' | 'list') || 'cards'; }
    catch { return 'cards'; }
  });
  const switchView = (m: 'cards' | 'list'): void => {
    setViewMode(m);
    try { localStorage.setItem('dex-eventlist-view', m); } catch { /* */ }
  };
  // v9.18: Debug-Button entfernt — wurde im Live-Betrieb nicht gebraucht.
  const [myNumbers, setMyNumbers] = React.useState<{ registered: number[]; waitlisted: number[] }>({ registered: [], waitlisted: [] });

  React.useEffect(() => {
    getMyEventNumbers().then(setMyNumbers).catch(err => console.warn('[DEX]', err));
  }, [events]);

  const statusFiltered = onlyActive
    ? events.filter((e) => e.status === 'Active')
    : events;

  // Admin sieht ALLE Events. Organizer sieht nur (a) Events, die zur Filterlogik
  // passen UND (b) Events, bei denen er selbst in organizerEmails steht — NICHT
  // tenant-weit alle Events. User sieht nur Filter-passende Events.
  // Fictive (Test-)Events sind sichtbar fuer:
  //   - Admins (immer)
  //   - Event-eigene Organizer (immer)
  //   - v9.16: Mitglieder des globalen Test-Teams (TestTeamEmails in _Config)
  //   - v9.18: Co-Organizer + QR-Scanner des konkreten Events
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const inTestTeam = (testTeamEmails || []).some(e => e === currentEmailLc);
  const fictiveFiltered = statusFiltered.filter((e: DeloitteEvent) => {
    if (!e.isFictive) return true;
    if (isAdmin) return true;
    if (inTestTeam) return true;
    if (e.organizerEmails.some((em: string) => (em || '').toLowerCase() === currentEmailLc)) return true;
    if ((e.coOrganizerEmails || []).some((em: string) => (em || '').toLowerCase() === currentEmailLc)) return true;
    if ((e.qrScannerEmails || []).some((em: string) => (em || '').toLowerCase() === currentEmailLc)) return true;
    return false;
  });
  const filteredEvents = (isAdmin
    ? fictiveFiltered
    : fictiveFiltered.filter((e: DeloitteEvent) =>
        isEventVisibleForUser(e, currentUser.email, currentUser.location)
        || e.organizerEmails.some((em: string) => (em || '').toLowerCase() === currentEmailLc)
      )
  ).slice().sort((a: DeloitteEvent, b: DeloitteEvent) => {
    // Chronologisch nach Startdatum (frueheste zuerst). Events ohne Datum ans Ende.
    const ta = a.startDate ? new Date(a.startDate).getTime() : Number.POSITIVE_INFINITY;
    const tb = b.startDate ? new Date(b.startDate).getTime() : Number.POSITIVE_INFINITY;
    return ta - tb;
  });

  if (isEventsLoading) {
    return (
      <div className="page-container text-center">
        <div style={{ padding: 48 }}>
          {/* SVG-Spinner mit SMIL animateTransform - laeuft unabhaengig von
              CSS-Keyframes / SPFx-Style-Hashing / React inline-style-Ordering.
              Voller hellgruener Ring + dunkelgruener Arc der rotiert. */}
          <svg width={48} height={48} viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" style={{ display: 'block', margin: '0 auto 16px' }}>
            <circle cx={24} cy={24} r={20} fill="none" stroke="rgba(134,188,37,0.20)" strokeWidth={4} />
            <path d="M 24 4 A 20 20 0 0 1 44 24" fill="none" stroke="#86bc25" strokeWidth={4} strokeLinecap="round">
              <animateTransform attributeName="transform" type="rotate" from="0 24 24" to="360 24 24" dur="1s" repeatCount="indefinite" />
            </path>
          </svg>
          <p style={{ color: 'var(--dex-gray-400)' }}>{t('myevents.loading')}</p>
        </div>
      </div>
    );
  }

  return (
    <div className="page-container">
      {/* Toolbar: Refresh-Button immer sichtbar (User koennen die Liste
          frisch laden ohne die App zu reloaden), Debug-Button nur fuer
          Organizer/Admin. */}
      <div style={{ marginBottom: 8, display: 'flex', justifyContent: 'flex-end', gap: 8, alignItems: 'center' }}>
        <button
          type="button"
          onClick={handleRefresh}
          disabled={isRefreshing || isEventsLoading}
          title={isDe ? 'Events neu laden' : 'Refresh events'}
          style={{
            // v9.18: moderner Button-Look — größeres Padding, Schatten,
            // klare Card-Box statt schmaler Outline. Hover hebt sich leicht.
            display: 'inline-flex', alignItems: 'center', gap: 8,
            fontSize: '0.85rem', padding: '8px 16px',
            background: '#fff',
            border: '1px solid var(--dex-gray-200, #e5e7eb)',
            borderRadius: 8, color: 'var(--dex-gray-800)',
            cursor: (isRefreshing || isEventsLoading) ? 'not-allowed' : 'pointer',
            opacity: (isRefreshing || isEventsLoading) ? 0.6 : 1,
            boxShadow: '0 1px 3px rgba(0,0,0,0.06)',
            fontWeight: 500,
            transition: 'box-shadow 120ms ease, transform 120ms ease',
          }}
          onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.boxShadow = '0 4px 12px rgba(0,0,0,0.08)'; }}
          onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.boxShadow = '0 1px 3px rgba(0,0,0,0.06)'; }}
        >
          <span style={{
            display: 'inline-flex',
            animation: isRefreshing ? 'dex-spin 0.8s linear infinite' : 'none',
          }}>
            <RefreshCw size={14} />
          </span>
          {isRefreshing
            ? (isDe ? 'Wird geladen…' : 'Loading…')
            : (isDe ? 'Aktualisieren' : 'Refresh')}
        </button>
      </div>
      <style>{`@keyframes dex-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
      {/* Titel + Hinweis: Events sind fuer den User personalisiert */}
      <div className="card" style={{
        padding: '16px 20px',
        marginBottom: 16,
        background: 'linear-gradient(135deg, rgba(134,188,37,0.08) 0%, rgba(134,188,37,0.02) 100%)',
        border: '1px solid var(--dex-green, #86bc25)',
      }}>
        <h2 style={{ margin: 0, marginBottom: 6, fontSize: '1.1rem', fontWeight: 700 }}>
          {t('eventlist.title')}
        </h2>
        <p style={{ margin: 0, fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
          {t('eventlist.hint')}
        </p>
      </div>
      <div className="flex-between mb-16">
        {/* View-Mode Switcher: Cards / List */}
        <div style={{ display: 'inline-flex', borderRadius: 6, overflow: 'hidden', border: '1px solid var(--dex-gray-300)' }}>
          <button
            onClick={() => switchView('cards')}
            style={{
              padding: '6px 12px', fontSize: '0.78rem', cursor: 'pointer', border: 'none',
              background: viewMode === 'cards' ? 'var(--dex-green)' : 'transparent',
              color: viewMode === 'cards' ? '#fff' : 'var(--dex-gray-600)',
              fontWeight: viewMode === 'cards' ? 600 : 400,
              display: 'inline-flex', alignItems: 'center', gap: 6,
            }}
          >
            <Icon iconName="GridViewMedium" style={{ fontSize: 14 }} /> Cards
          </button>
          <button
            onClick={() => switchView('list')}
            style={{
              padding: '6px 12px', fontSize: '0.78rem', cursor: 'pointer', border: 'none',
              background: viewMode === 'list' ? 'var(--dex-green)' : 'transparent',
              color: viewMode === 'list' ? '#fff' : 'var(--dex-gray-600)',
              fontWeight: viewMode === 'list' ? 600 : 400,
              display: 'inline-flex', alignItems: 'center', gap: 6,
            }}
          >
            <Icon iconName="GroupedList" style={{ fontSize: 14 }} /> List
          </button>
        </div>
        <div className="toggle-wrapper">
          <label className="toggle">
            <input
              type="checkbox"
              checked={onlyActive}
              onChange={(e) => setOnlyActive(e.target.checked)}
            />
            <span className="toggle-slider" />
          </label>
          <span>{t('eventlist.onlyactive')}</span>
        </div>
      </div>
      {viewMode === 'cards' ? (
        <div className="event-grid">
          {filteredEvents.map((event, i) => (
            <EventCard
              key={event.id}
              event={event}
              index={i}
              isRegistered={myNumbers.registered.includes(event.eventNumber)}
              isWaitlisted={myNumbers.waitlisted.includes(event.eventNumber)}
            />
          ))}
        </div>
      ) : (
        <EventListView
          events={filteredEvents}
          myNumbers={myNumbers}
          formatDate={(iso) => {
            if (!iso) return '';
            const d = new Date(iso);
            if (isNaN(d.getTime())) return '';
            return d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ' ' +
                   d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
          }}
        />
      )}
      {filteredEvents.length === 0 && (
        <p className="text-center mt-24" style={{ color: 'var(--dex-gray-400)' }}>
          Keine Events für dich gefunden.
        </p>
      )}
    </div>
  );
}

/**
 * Listen-Ansicht der Events - im Stil der Admin/Organizer-Seite.
 */
function EventListView({ events, myNumbers, formatDate }: {
  events: DeloitteEvent[];
  myNumbers: { registered: number[]; waitlisted: number[] };
  formatDate: (iso: string) => string;
}): React.ReactElement {
  const { navigate } = useNavigation();
  return (
    <div className="my-events-list">
      {events.map(event => {
        const isReg = myNumbers.registered.includes(event.eventNumber);
        const isWait = myNumbers.waitlisted.includes(event.eventNumber);
        const targetPage = (isReg || isWait) ? 'my-events' : 'registration';
        return (
          <div
            key={event.id}
            className="card card-clickable"
            style={{ padding: '20px 24px', cursor: 'pointer' }}
            onClick={() => navigate(targetPage, event.id)}
          >
            <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: 12 }}>
              <div style={{ flex: '1 1 200px', display: 'flex', alignItems: 'center', gap: 16 }}>
                {event.imageUrl && (
                  <div style={{
                    width: 60, height: 40, borderRadius: 'var(--dex-radius)', flexShrink: 0,
                    background: `url(${event.imageUrl}) center/cover no-repeat`,
                  }} />
                )}
                <div>
                  <h3 style={{ marginBottom: 4 }}>{event.title}</h3>
                  <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', margin: 0 }}>
                    {formatDate(event.startDate)} - {formatDate(event.endDate)}
                    {event.location ? ` · ${event.location}` : ''}
                  </p>
                  <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)', margin: '2px 0 0' }}>
                    Organizer: {event.organizers.map(o => { const p = o.split(',').map(s => s.trim()); return p.length === 2 ? `${p[1]} ${p[0]}` : o; }).join(', ')}
                  </p>
                </div>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                <span style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                  {event.currentParticipants}/{event.maxParticipants || '∞'} Teilnehmer
                </span>
                {isReg && (
                  <span style={{ padding: '2px 10px', borderRadius: 12, fontSize: '0.75rem', fontWeight: 600, background: 'rgba(134,188,37,0.18)', color: 'var(--dex-green-dark)' }}>Angemeldet</span>
                )}
                {isWait && (
                  <span style={{ padding: '2px 10px', borderRadius: 12, fontSize: '0.75rem', fontWeight: 600, background: 'rgba(237,139,0,0.18)', color: 'var(--dex-orange, #ed8b00)' }}>Warteliste</span>
                )}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}
