// Start-Seite - Navigation zu Registration, My Events und Admin

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useLanguage } from '../context/LanguageContext';
import { Calendar, Pin, Settings, QrCode, Star, Users, MessageSquare } from './Icons';
import InquiryModal from './InquiryModal';
import { useTickets } from '../context/TicketContext';

export default function StartPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { canCreateEvents, isAdmin, isPowerUser } = useRoles();
  const { events, isEventsLoading, getMyProxyRegistrations } = useEvents();
  const { powerUserQueue } = useTickets();
  const { currentUser } = useCurrentUser();
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  // v12.5: Organizer-Kachel wird jetzt IMMER gerendert. Wer keine Organizer-
  // Rechte hat sieht sie ausgegraut mit Overlay-Button „Want to become an
  // organizer?" — Klick öffnet das gleiche Inquiry-Modal wie das
  // Bubble-CTA auf der Landing-Page.
  const [showInquiry, setShowInquiry] = React.useState(false);

  // v11.38/v11.46: Co-Organizer pro Event (per-Event-Rolle, ohne globale
  // Organizer-Rolle in DEX_Roles) sehen die Organizer-Kachel ebenfalls —
  // AdminPage gewährt ihnen ohnehin Zugriff auf "ihre" Events (siehe
  // isOrganizerFor dort), aber ohne Kachel im Startmenü gab es bisher
  // keinen Einstieg.
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isOrganizerOfAnyEvent = !!currentEmailLc && (events || []).some(e => {
    const inOrg = (e.organizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
    if (inOrg) return true;
    return (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
  });
  // v23.36: Im Demo-Modus sieht der Demo-User die Start-Kacheln EXAKT wie ein
  // normaler User — d.h. die Organizer-Kachel ist ausgegraut.
  const isOrganizer = canCreateEvents || isOrganizerOfAnyEvent;
  // v13.12: Check-In-Team-Mitgliedschaft = User ist QR-Scanner/Organizer eines
  // AKTIVEN Events; dann erscheint die Check-In-Kachel.
  const isCheckInTeamOfActive = !!currentEmailLc && (events || []).some(e => {
    if (e.status !== 'Active') return false;
    if ((e.qrScannerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc)) return true;
    if ((e.organizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc)) return true;
    return (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
  });
  const showCheckInTile = isCheckInTeamOfActive || isAdmin || isOrganizer;
  // v24.24: „Admin"-Kachel (Hub) — nur echte Admins.
  const showAdminHubTile = isAdmin;

  // v26: „Tickets"-Kachel — Power-User + Admins beantworten hier die Fragen aus
  // dem Ticketsystem (Badge = Anzahl noch offener/in Bearbeitung befindlicher).
  const showTicketsTile = isAdmin || isPowerUser;
  const openTicketCount = powerUserQueue.filter(tk => tk.status !== 'Closed').length;

  // v24.36: „Assistenz"-Kachel — nur sichtbar, wenn der User STELLVERTRETEND
  // für eine andere Person angemeldet hat (oder Admin ist). Erkennung läuft im
  // Hintergrund über alle Teilnehmerlisten, pro Tag gecacht (localStorage).
  const [hasProxyRegs, setHasProxyRegs] = React.useState<boolean>(false);
  const proxyScanStartedRef = React.useRef(false);
  React.useEffect(() => {
    if (isAdmin) return; // Admin sieht die Kachel immer — kein Scan nötig.
    if (isEventsLoading || (events || []).length === 0) return;
    if (proxyScanStartedRef.current) return;
    const key = `dex_assist_${currentEmailLc}`;
    try {
      const raw = window.localStorage.getItem(key);
      if (raw) {
        const cached = JSON.parse(raw) as { ts: number; has: boolean };
        if (cached && Date.now() - cached.ts < 24 * 60 * 60 * 1000) {
          setHasProxyRegs(!!cached.has);
          return; // frischer Cache — kein erneuter Scan.
        }
      }
    } catch { /* */ }
    proxyScanStartedRef.current = true;
    getMyProxyRegistrations()
      .then(list => {
        const has = (list || []).length > 0;
        setHasProxyRegs(has);
        try { window.localStorage.setItem(key, JSON.stringify({ ts: Date.now(), has })); } catch { /* */ }
      })
      .catch(() => { /* best-effort */ });
  }, [isAdmin, isEventsLoading, events, currentEmailLc, getMyProxyRegistrations]);
  const showAssistTile = isAdmin || hasProxyRegs;

  // v26: Kacheln als Elemente — werden danach in beschriftete Cluster gruppiert
  // und als quadratische Kacheln in einem zweispaltigen Raster gerendert
  // (data-tour-Anker bleiben für das geführte Tutorial erhalten).
  const tileRegister = (
    <div className="card card-clickable start-card" data-tour="tile-register" onClick={() => navigate('register')}>
      <div className="start-card__icon"><Calendar size={64} strokeWidth={1} /></div>
      <h2>{t('start.register')}</h2>
      <p>{t('start.register.desc')}</p>
    </div>
  );
  const tileMyEvents = (
    <div className="card card-clickable start-card" data-tour="tile-myevents" onClick={() => navigate('my-events')}>
      <div className="start-card__icon"><Pin size={64} strokeWidth={1} /></div>
      <h2>{t('start.myevents')}</h2>
      <p>{t('start.myevents.desc')}</p>
    </div>
  );
  const tileOrganizer = isOrganizer ? (
    <div className="card card-clickable start-card start-card--admin" data-tour="tile-admin" onClick={() => navigate('admin')}>
      <div className="start-card__icon"><Settings size={64} strokeWidth={1} /></div>
      <h2>{t('start.admin')}</h2>
      <p>{t('start.admin.desc')}</p>
    </div>
  ) : (
    <div className="card start-card start-card--admin" style={{ position: 'relative', cursor: 'default', opacity: 0.55 }}>
      <div className="start-card__icon"><Settings size={64} strokeWidth={1} /></div>
      <h2>{t('start.admin')}</h2>
      <p>{t('start.admin.desc')}</p>
      <div style={{ position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 10 }}>
        <button
          type="button"
          onClick={(e) => { e.stopPropagation(); setShowInquiry(true); }}
          style={{
            background: 'var(--dex-green, #86bc25)', color: '#fff', border: 'none', borderRadius: 14,
            padding: '8px 12px', fontWeight: 700, cursor: 'pointer', boxShadow: '0 4px 14px rgba(0,0,0,0.18)',
            fontSize: '0.74rem', lineHeight: 1.25, textAlign: 'center',
          }}
        >
          {isDe ? 'Organizer werden?' : 'Want to become an organizer?'}
        </button>
      </div>
    </div>
  );
  const tileCheckIn = (
    <div className="card card-clickable start-card start-card--checkin" data-tour="tile-checkin" onClick={() => navigate('check-in')}>
      <div className="start-card__icon"><QrCode size={64} strokeWidth={1} /></div>
      <h2>{isDe ? 'Check-In' : 'Check-in'}</h2>
      <p>{isDe ? 'Teilnehmer einchecken' : 'Check in attendees'}</p>
    </div>
  );
  const tileTickets = (
    <div className="card card-clickable start-card" onClick={() => navigate('tickets')} style={{ position: 'relative' }}>
      <div className="start-card__icon"><MessageSquare size={64} strokeWidth={1} /></div>
      <h2>Tickets</h2>
      <p>{isDe ? 'Fragen beantworten' : 'Answer questions'}</p>
      {openTicketCount > 0 && (
        <span style={{
          position: 'absolute', top: 10, right: 10, background: '#ed8b00', color: '#fff',
          borderRadius: 12, minWidth: 22, height: 22, padding: '0 6px',
          display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700,
        }}>{openTicketCount}</span>
      )}
    </div>
  );
  const tileAssist = (
    <div className="card card-clickable start-card" onClick={() => navigate('assistant')}>
      <div className="start-card__icon"><Users size={64} /></div>
      <h2>{isDe ? 'Assistenz' : 'Assistant'}</h2>
      <p>{isDe ? 'Anmeldungen für andere' : 'Registrations for others'}</p>
    </div>
  );
  const tileAdminHub = (
    <div className="card card-clickable start-card" onClick={() => navigate('admin-hub')}>
      <div className="start-card__icon"><Star size={64} strokeWidth={1} /></div>
      <h2>Admin</h2>
      <p>{isDe ? 'Verwaltung & Prozesse' : 'Administration & processes'}</p>
    </div>
  );

  const clusters: Array<{ key: string; title: string; tiles: React.ReactNode[] }> = [
    { key: 'teilnahme', title: isDe ? 'Teilnahme' : 'Participation', tiles: [tileRegister, tileMyEvents] },
    { key: 'organisation', title: isDe ? 'Organisation' : 'Organization', tiles: [tileOrganizer, ...(showCheckInTile ? [tileCheckIn] : [])] },
    { key: 'support', title: isDe ? 'Support' : 'Support', tiles: showTicketsTile ? [tileTickets] : [] },
    { key: 'verwaltung', title: isDe ? 'Verwaltung' : 'Administration', tiles: [...(showAssistTile ? [tileAssist] : []), ...(showAdminHubTile ? [tileAdminHub] : [])] },
  ].filter(c => c.tiles.length > 0);

  return (
    <div className="page-container">
      {/* v9.36: Keyframes inline injizieren — SPFx hasht sonst die Names im
          .module.scss und die Animation findet sie nicht.
          v26: Layout in beschriftete Cluster (quadratische Kacheln, zweizeilig). */}
      <style>{`
        @keyframes dexStartIconBounce {
          0%   { transform: translateY(0) scale(1); }
          30%  { transform: translateY(-8px) scale(1.1); }
          60%  { transform: translateY(0) scale(1); }
          80%  { transform: translateY(-3px) scale(1.04); }
          100% { transform: translateY(0) scale(1); }
        }
        @keyframes dexStartIconWiggle {
          0%   { transform: rotate(0deg) scale(1); }
          20%  { transform: rotate(-10deg) scale(1.08); }
          40%  { transform: rotate(8deg) scale(1.08); }
          60%  { transform: rotate(-6deg) scale(1.05); }
          80%  { transform: rotate(4deg) scale(1.03); }
          100% { transform: rotate(0deg) scale(1); }
        }
        @keyframes dexStartIconSpin {
          0%   { transform: rotate(0deg) scale(1); }
          100% { transform: rotate(360deg) scale(1); }
        }
        @keyframes dexStartIconScanPulse {
          0%   { transform: scale(1); filter: drop-shadow(0 0 0 rgba(134,188,37,0.0)); }
          40%  { transform: scale(1.12); filter: drop-shadow(0 0 6px rgba(134,188,37,0.55)); }
          70%  { transform: scale(0.96); filter: drop-shadow(0 0 0 rgba(134,188,37,0.0)); }
          100% { transform: scale(1); filter: drop-shadow(0 0 0 rgba(134,188,37,0.0)); }
        }
        .start-card--checkin:hover .start-card__icon svg { animation: dexStartIconScanPulse 0.9s ease; }

        /* v26: Cluster-Raster — bis zu 2 Cluster nebeneinander (→ zwei Reihen),
           darin quadratische Kacheln. */
        /* v26.2: jeder Cluster in einer EIGENEN Zeile (untereinander) → größere Kacheln. */
        .dex-cluster-grid {
          display: grid; grid-template-columns: 1fr;
          gap: 24px; max-width: 660px; margin: 0 auto;
        }
        .dex-cluster-title {
          font-size: 0.82rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px;
          color: var(--dex-gray-500); margin: 0 0 12px; padding-bottom: 7px;
          border-bottom: 1px solid var(--dex-gray-200);
        }
        /* v26.6: alle Kacheln GLEICH GROSS (fix 300px, quadratisch — wie die
           Tickets-Kachel). Passen zwei nebeneinander, stehen sie in einer Zeile;
           sonst brechen sie um, statt unterschiedlich zu schrumpfen. */
        .dex-cluster-tiles { display: flex; flex-wrap: wrap; gap: 18px; align-items: stretch; }
        .dex-cluster .start-card {
          flex: 0 0 auto !important; width: 300px !important; max-width: 100% !important;
          padding: 24px 16px !important; min-height: 0 !important; aspect-ratio: 1 / 1; gap: 8px !important;
        }
        /* v26.5: echter KREIS hinter dem Icon — flex:0 0 auto + aspect-ratio
           verhindert, dass die Flex-Stauchung den Kreis zur Ellipse macht. */
        .dex-cluster .start-card__icon { width: 84px !important; height: 84px !important; flex: 0 0 auto !important; aspect-ratio: 1 / 1 !important; border-radius: 50% !important; margin-bottom: 4px !important; }
        .dex-cluster .start-card__icon svg { width: 44px !important; height: 44px !important; }
        .dex-cluster .start-card h2 { font-size: 1.12rem !important; }
        .dex-cluster .start-card p { font-size: 0.82rem !important; white-space: normal !important; }
        @media (max-width: 480px) {
          .dex-cluster .start-card__icon { width: 60px !important; height: 60px !important; aspect-ratio: 1 / 1 !important; }
          .dex-cluster .start-card__icon svg { width: 32px !important; height: 32px !important; }
          .dex-cluster .start-card h2 { font-size: 0.98rem !important; }
        }
      `}</style>
      <div className="dex-cluster-grid">
        {clusters.map(c => (
          <div key={c.key} className="dex-cluster">
            <div className="dex-cluster-title">{c.title}</div>
            <div className="dex-cluster-tiles">
              {c.tiles.map((tile, i) => <React.Fragment key={i}>{tile}</React.Fragment>)}
            </div>
          </div>
        ))}
      </div>
      {/* v24.24: „Organizer werden?" öffnet dasselbe Anfrage-Modal wie die grüne
          Box auf der Landing Page. */}
      <InquiryModal open={showInquiry} onClose={() => setShowInquiry(false)} />
    </div>
  );
}
