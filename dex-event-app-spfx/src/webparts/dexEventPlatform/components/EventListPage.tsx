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
// v11.99: RefreshCw nicht mehr benötigt (Page-Level-Refresh-Button entfernt).
import { Icon } from '@fluentui/react/lib/Icon';
import EventCard from './EventCard';
import { CachedBg } from './CachedImage';
import { prewarmImages } from '../utils/imageCache';

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
function matchesAudience(
  userEmail: string,
  userLocation: string,
  audienceFilters: string[],
  userGroupEmails: string[] = [],
  eventResolvedEmails: string[] = [],
): boolean {
  if (audienceFilters.length === 0) return true;
  const email = userEmail.toLowerCase();
  const loc = (userLocation || '').toLowerCase();
  const groupSet = new Set((userGroupEmails || []).map(g => (g || '').toLowerCase().trim()).filter(Boolean));
  const resolvedSet = new Set((eventResolvedEmails || []).map(e => (e || '').toLowerCase().trim()).filter(Boolean));

  // v16.4: Pre-compiled-Pfad — wenn die DLs beim Event-Save aufgeloest und
  // in event.audienceResolvedEmails gespeichert wurden, reicht ein direkter
  // E-Mail-Lookup. Funktioniert auch fuer verschachtelte DLs, die
  // /me/memberOf nicht zurueckliefert.
  if (resolvedSet.has(email)) return true;

  return audienceFilters.some(filter => {
    const f = filter.trim().toLowerCase();
    if (!f) return false;

    // E-Mail-Adresse: entweder direkte User-Adresse, ODER der User ist
    // Mitglied in dieser Mailverteiler-/Gruppen-Adresse (v15.27 — Fix
    // fuer Check-Visibility vs Runtime-Mismatch).
    if (f.indexOf('@') >= 0) {
      if (email === f) return true;
      if (groupSet.has(f)) return true;
      return false;
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
  userLocation: string,
  userGroupEmails: string[] = [],
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
  const audMatch = matchesAudience(userEmail, userLocation, normalizedAud, userGroupEmails, event.audienceResolvedEmails || []);

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
  const { topLevelEvents: events, isEventsLoading, getMyEventNumbers, childEventsOf, refreshParticipantCounts } = useEvents();
  const { currentUser, groupEmails } = useCurrentUser();
  const { isAdmin, canCreateEvents } = useRoles();

  // v19.17: KEIN 5-Sekunden-Polling mehr — das verursachte einen sichtbaren
  // Re-Render der ganzen Liste. Stattdessen die Teilnehmerzahlen nur EINMAL beim
  // Öffnen der Übersicht im Hintergrund frisch nachladen (fixt die veraltete
  // Boot-Snapshot-Zahl, z.B. „0/1" nach Flow-Nachrücken). refreshParticipantCounts
  // aktualisiert dabei nur die Events, deren Zahl sich tatsächlich geändert hat
  // (referenzschonend, siehe EventContext) → kein Flackern.
  React.useEffect(() => {
    refreshParticipantCounts().catch(() => { /* best-effort */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  // v19.22: Event-Bilder im Hintergrund in den Cache vorwärmen, sobald die
  // Event-Liste da ist — beim Öffnen eines Events liegt das Bild dann bereits
  // lokal vor (kein SharePoint-Roundtrip).
  React.useEffect(() => {
    prewarmImages((events || []).map(e => e.imageUrl));
  }, [events]);
  const { t } = useLanguage();
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

  // v11.89: "Nur aktive Events" blendet jetzt sowohl Entwürfe (IsFictive
  // oder ActiveFrom in der Zukunft) als auch Cancelled/Completed aus —
  // konsistent mit der Konsolidierung „Entwurf = Entwurf, kein zweites
  // Under-Construction-Konzept mehr".
  // v18.17 — BUG-FIX: vorher hat dieser Filter ALLE isFictive-Events (und
  // ActiveFrom-zukünftige) hart entfernt, BEVOR die Test-Team-/Co-Organizer-
  // /QR-Scanner-Sichtbarkeitsprüfung weiter unten greifen konnte. Folge:
  // wer ohne Organizer-Rechte ins Test-Team eines Entwurfs aufgenommen wurde,
  // konnte das Event trotzdem nicht sehen. Der Entwurfs-/ActiveFrom-Check
  // läuft jetzt ausschließlich in `fictiveFiltered` unten, wo die
  // per-User-Whitelist (Test-Team etc.) korrekt berücksichtigt wird.
  // Admin sieht ALLE Events. Organizer sieht nur (a) Events, die zur Filterlogik
  // passen UND (b) Events, bei denen er selbst in organizerEmails steht — NICHT
  // tenant-weit alle Events. User sieht nur Filter-passende Events.
  // Fictive (Test-)Events sind sichtbar fuer:
  //   - Admins (immer)
  //   - Event-eigene Organizer (immer)
  //   - Co-Organizer + QR-Scanner des konkreten Events
  //   - v9.21: Test-Team-Mitglieder DIESES Events (vorher global, jetzt per-Event)
  //
  // v9.21: ActiveFrom-Logik — wenn das Event ein Aktiv-ab-Datum hat und das
  // in der Zukunft liegt, wird es fuer regulaere User behandelt wie ein Entwurf.
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  // v11.90: Events, bei denen der eingeloggte User selbst Organizer ist
  // (Haupt- oder Co-Organizer), kommen oben links — ABER nur solange das
  // Event nicht abgelaufen ist (EndDate in der Vergangenheit). Abgelaufene
  // eigene Events sortieren mit allen anderen rein nach Startdatum.
  // v20.0 (Audit): useCallback, damit die memoizte Filterkette unten eine
  // stabile Referenz als Dependency bekommt.
  const isOwnOrganizer = React.useCallback((e: DeloitteEvent): boolean => {
    const own = (e.organizerEmails || []).some((em: string) => (em || '').toLowerCase() === currentEmailLc);
    const co = (e.coOrganizerEmails || []).some((em: string) => (em || '').toLowerCase() === currentEmailLc);
    return own || co;
  }, [currentEmailLc]);
  // v20.0 (Audit): die komplette Filter-/Sortier-Kette (Status → Entwurf →
  // Sichtbarkeit → Sortierung) lief vorher bei JEDEM Render neu über alle
  // Events. Jetzt memoizt — neu berechnet nur wenn sich Events, Filter,
  // Rolle oder User-Identität ändern.
  const filteredEvents = React.useMemo(() => {
    const nowTs = Date.now();
    const statusFiltered = onlyActive
      ? events.filter((e) => e.status === 'Active')
      : events;
    const fictiveFiltered = statusFiltered.filter((e: DeloitteEvent) => {
      // Effektiv "noch im Entwurf": entweder isFictive oder ActiveFrom in der Zukunft
      const activeFromTs = e.activeFrom ? new Date(e.activeFrom).getTime() : 0;
      const isInDraftMode = e.isFictive || (activeFromTs > 0 && activeFromTs > nowTs);
      if (!isInDraftMode) return true;
      if (isAdmin) return true;
      if (e.organizerEmails.some((em: string) => (em || '').toLowerCase() === currentEmailLc)) return true;
      if ((e.coOrganizerEmails || []).some((em: string) => (em || '').toLowerCase() === currentEmailLc)) return true;
      if ((e.qrScannerEmails || []).some((em: string) => (em || '').toLowerCase() === currentEmailLc)) return true;
      if ((e.testTeamEmails || []).some((em: string) => (em || '').toLowerCase() === currentEmailLc)) return true;
      return false;
    });
    const isExpired = (e: DeloitteEvent): boolean => {
      const t = e.endDate ? new Date(e.endDate).getTime() : 0;
      return t > 0 && t < nowTs;
    };
    return (isAdmin
      ? fictiveFiltered
      : fictiveFiltered.filter((e: DeloitteEvent) =>
          isEventVisibleForUser(e, currentUser.email, currentUser.location, groupEmails)
          || e.organizerEmails.some((em: string) => (em || '').toLowerCase() === currentEmailLc)
        )
    ).slice().sort((a: DeloitteEvent, b: DeloitteEvent) => {
      // v11.90: Priorität — eigene, noch nicht abgelaufene Events zuerst.
      const aPrio = isOwnOrganizer(a) && !isExpired(a) ? 0 : 1;
      const bPrio = isOwnOrganizer(b) && !isExpired(b) ? 0 : 1;
      if (aPrio !== bPrio) return aPrio - bPrio;
      // Innerhalb der Priorität: chronologisch nach Startdatum (frueheste zuerst).
      // Events ohne Datum ans Ende.
      const ta = a.startDate ? new Date(a.startDate).getTime() : Number.POSITIVE_INFINITY;
      const tb = b.startDate ? new Date(b.startDate).getTime() : Number.POSITIVE_INFINITY;
      return ta - tb;
    });
  }, [events, onlyActive, isAdmin, currentEmailLc, currentUser.email, currentUser.location, groupEmails, isOwnOrganizer]);

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
      {/* v11.99: Page-Level-Refresh-Button entfernt — Header oben rechts
          hat bereits einen Aktualisieren-Button, doppelt verwirrt. */}
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
      <div className="flex-between mb-16" style={{ alignItems: 'flex-end' }}>
        {/* View-Mode Switcher: Cards / List — mit „Ansicht"-Label darueber,
            damit klar ist, was die beiden Buttons umschalten (v19.6). */}
        <div>
          <div style={{
            fontSize: '0.8rem', fontWeight: 700, color: 'var(--dex-gray-600)',
            textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 6,
          }}>
            {t('eventlist.view') || 'Ansicht'}
          </div>
          <div style={{ display: 'inline-flex', borderRadius: 8, overflow: 'hidden', border: '1px solid var(--dex-gray-300)' }}>
            <button
              onClick={() => switchView('cards')}
              style={{
                padding: '9px 18px', fontSize: '0.98rem', cursor: 'pointer', border: 'none',
                background: viewMode === 'cards' ? 'var(--dex-green)' : 'transparent',
                color: viewMode === 'cards' ? '#fff' : 'var(--dex-gray-600)',
                fontWeight: viewMode === 'cards' ? 700 : 500,
                display: 'inline-flex', alignItems: 'center', gap: 8,
              }}
            >
              <Icon iconName="GridViewMedium" style={{ fontSize: 18 }} /> Cards
            </button>
            <button
              onClick={() => switchView('list')}
              style={{
                padding: '9px 18px', fontSize: '0.98rem', cursor: 'pointer', border: 'none',
                background: viewMode === 'list' ? 'var(--dex-green)' : 'transparent',
                color: viewMode === 'list' ? '#fff' : 'var(--dex-gray-600)',
                fontWeight: viewMode === 'list' ? 700 : 500,
                display: 'inline-flex', alignItems: 'center', gap: 8,
              }}
            >
              <Icon iconName="GroupedList" style={{ fontSize: 18 }} /> List
            </button>
          </div>
        </div>
        {/* v15.19: „Nur aktive Events"-Toggle nur fuer Admin/Organizer.
            Reine User sehen ohnehin nur Events ihres Standorts und brauchen
            den Switch nicht — sie sollen vergangene Events nicht
            ausversehen einblenden koennen. */}
        {(isAdmin || canCreateEvents) && (
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
        )}
      </div>
      {/* v15.21: Zwei klar getrennte Sektionen — eigene Events (Organizer)
          zuerst, danach alle weiteren Events sortiert nach Datum. */}
      {(() => {
        const ownEvents = filteredEvents.filter(e => isOwnOrganizer(e));
        const otherEvents = filteredEvents.filter(e => !isOwnOrganizer(e));
        const sectionTitle = (text: string): React.ReactElement => (
          <h3 style={{
            margin: '24px 0 12px', fontSize: '1rem',
            color: 'var(--dex-gray-800)', fontWeight: 700,
            textTransform: 'uppercase', letterSpacing: 0.6,
          }}>{text}</h3>
        );
        const formatDate = (iso: string): string => {
          if (!iso) return '';
          const d = new Date(iso);
          if (isNaN(d.getTime())) return '';
          return d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ' ' +
                 d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
        };
        // v15.22: Bei subEventsOnlyMode-Events gibt es keine Parent-
        // Registrierung — die EventNumber des Hauptevents taucht in
        // myNumbers.registered nicht auf. Stattdessen muessen wir
        // die Sub-Events durchgehen: ist mindestens eines davon
        // angemeldet/wartelistig, soll auch die Parent-Karte den
        // „Registered"-Overlay zeigen.
        const isRegisteredFor = (event: DeloitteEvent): boolean => {
          if (myNumbers.registered.includes(event.eventNumber)) return true;
          if (event.subEventsOnlyMode) {
            const kids = childEventsOf(event.id);
            return kids.some(k => myNumbers.registered.includes(k.eventNumber));
          }
          return false;
        };
        const isWaitlistedFor = (event: DeloitteEvent): boolean => {
          if (myNumbers.waitlisted.includes(event.eventNumber)) return true;
          if (event.subEventsOnlyMode) {
            const kids = childEventsOf(event.id);
            return kids.some(k => myNumbers.waitlisted.includes(k.eventNumber));
          }
          return false;
        };
        const renderSection = (list: DeloitteEvent[]): React.ReactElement => {
          if (viewMode === 'cards') {
            return (
              <div className="event-grid">
                {list.map((event, i) => (
                  <EventCard
                    key={event.id}
                    event={event}
                    index={i}
                    isRegistered={isRegisteredFor(event)}
                    isWaitlisted={isWaitlistedFor(event)}
                    isOwnOrganizer={isOwnOrganizer(event)}
                  />
                ))}
              </div>
            );
          }
          // v15.22: Fuer den Listen-View die Parent-EventNumbers von
          // subEventsOnlyMode-Parents augmentieren, wenn ein Sub-Event
          // angemeldet ist — damit der Listen-View ebenfalls
          // „angemeldet" markiert.
          const augReg: number[] = myNumbers.registered.slice();
          const augWait: number[] = myNumbers.waitlisted.slice();
          for (const ev of list) {
            if (isRegisteredFor(ev) && !augReg.includes(ev.eventNumber)) augReg.push(ev.eventNumber);
            if (isWaitlistedFor(ev) && !augWait.includes(ev.eventNumber)) augWait.push(ev.eventNumber);
          }
          return (
            <EventListView
              events={list}
              myNumbers={{ registered: augReg, waitlisted: augWait }}
              currentUserEmailLc={currentEmailLc}
              formatDate={formatDate}
            />
          );
        };
        return (
          <>
            {ownEvents.length > 0 && (
              <>
                {sectionTitle(t('eventlist.ownevents') || 'Meine Events als Organizer')}
                {renderSection(ownEvents)}
              </>
            )}
            {otherEvents.length > 0 && (
              <>
                {ownEvents.length > 0 && sectionTitle(t('eventlist.otherevents') || 'Weitere Events')}
                {renderSection(otherEvents)}
              </>
            )}
            {filteredEvents.length === 0 && (
              <p className="text-center mt-24" style={{ color: 'var(--dex-gray-400)' }}>
                Keine Events für dich gefunden.
              </p>
            )}
          </>
        );
      })()}
    </div>
  );
}

/**
 * Listen-Ansicht der Events - im Stil der Admin/Organizer-Seite.
 */
function EventListView({ events, myNumbers, formatDate, currentUserEmailLc }: {
  events: DeloitteEvent[];
  myNumbers: { registered: number[]; waitlisted: number[] };
  formatDate: (iso: string) => string;
  currentUserEmailLc: string;
}): React.ReactElement {
  const { navigate } = useNavigation();
  const { t } = useLanguage();
  const { canCreateEvents } = useRoles();
  return (
    <div className="my-events-list">
      {events.map(event => {
        const isReg = myNumbers.registered.includes(event.eventNumber);
        const isWait = myNumbers.waitlisted.includes(event.eventNumber);
        const targetPage = (isReg || isWait) ? 'my-events' : 'registration';
        // v11.90: Organizer-Badge in der List-View auch
        const isOwn = (event.organizerEmails || []).some(em => (em || '').toLowerCase() === currentUserEmailLc)
          || (event.coOrganizerEmails || []).some(em => (em || '').toLowerCase() === currentUserEmailLc);
        return (
          <div
            key={event.id}
            className="card card-clickable"
            style={{
              padding: '20px 24px', cursor: 'pointer',
              // v19.16: Angemeldet/Warteliste deutlich markieren — analog zum
              // Overlay der Kachel-Ansicht, damit man auf einen Blick sieht, dass
              // man für dieses Event angemeldet ist. Grüner Akzent + zarte Tönung
              // (bzw. orange bei Warteliste).
              ...(isReg ? {
                borderLeft: '5px solid var(--dex-green, #86bc25)',
                background: 'linear-gradient(90deg, rgba(134,188,37,0.10) 0%, rgba(134,188,37,0.015) 60%)',
              } : isWait ? {
                borderLeft: '5px solid var(--dex-orange, #ed8b00)',
                background: 'linear-gradient(90deg, rgba(237,139,0,0.10) 0%, rgba(237,139,0,0.015) 60%)',
              } : {}),
            }}
            onClick={() => navigate(targetPage, event.id)}
          >
            <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: 12, width: '100%' }}>
              <div style={{ flex: '1 1 200px', minWidth: 0, display: 'flex', alignItems: 'center', gap: 16 }}>
                <CachedBg url={event.imageUrl} style={{
                  width: 60, height: 40, borderRadius: 'var(--dex-radius)', flexShrink: 0,
                }} />
                <div style={{ minWidth: 0 }}>
                  <h3 style={{ marginBottom: 4, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                    {event.title}
                    {isOwn && (
                      <span style={{
                        padding: '2px 8px', borderRadius: 999, fontSize: '0.65rem', fontWeight: 700,
                        background: 'var(--dex-green, #86bc25)', color: '#fff', letterSpacing: 0.5,
                      }}>Organizer</span>
                    )}
                    {event.isFictive && (
                      <span style={{
                        padding: '2px 8px', borderRadius: 999, fontSize: '0.65rem', fontWeight: 700,
                        background: 'var(--dex-orange, #ed8b00)', color: '#fff', letterSpacing: 0.5,
                      }}>Entwurf</span>
                    )}
                  </h3>
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
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, padding: '3px 10px', borderRadius: 12, fontSize: '0.75rem', fontWeight: 700, background: 'rgba(134,188,37,0.22)', color: 'var(--dex-green-dark)' }}>
                    <Icon iconName="CompletedSolid" style={{ fontSize: 13 }} /> Angemeldet
                  </span>
                )}
                {isWait && (
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, padding: '3px 10px', borderRadius: 12, fontSize: '0.75rem', fontWeight: 700, background: 'rgba(237,139,0,0.22)', color: 'var(--dex-orange, #ed8b00)' }}>
                    <Icon iconName="Clock" style={{ fontSize: 13 }} /> Warteliste
                  </span>
                )}
                {/* v19.15: Aktionen auch in der Listen-Ansicht — vorher nur in den
                    Cards (Overlay). Bei bereits angemeldeten Events „Meine Events"
                    + (für Organizer/Admin) „Für andere Person registrieren", sonst
                    „Registrieren". stopPropagation, damit der Button nicht den
                    Zeilen-Klick mit auslöst. */}
                {(isReg || isWait) ? (
                  <>
                    <button
                      className="btn btn-primary"
                      style={{ fontSize: '0.78rem', padding: '6px 14px' }}
                      onClick={(e) => { e.stopPropagation(); navigate('my-events'); }}
                    >
                      {t('myevents.title')}
                    </button>
                    {canCreateEvents && (
                      <button
                        className="btn btn-secondary"
                        style={{ fontSize: '0.78rem', padding: '6px 14px' }}
                        onClick={(e) => { e.stopPropagation(); navigate('registration', event.id, 'register-other'); }}
                      >
                        {t('reg.registerother')}
                      </button>
                    )}
                  </>
                ) : (
                  <button
                    className="btn btn-primary"
                    style={{ fontSize: '0.78rem', padding: '6px 14px' }}
                    onClick={(e) => { e.stopPropagation(); navigate('registration', event.id); }}
                  >
                    {t('reg.register')}
                  </button>
                )}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}
