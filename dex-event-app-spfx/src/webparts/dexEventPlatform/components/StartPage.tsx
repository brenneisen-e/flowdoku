// Start-Seite - Navigation zu Registration, My Events und Admin

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useLanguage } from '../context/LanguageContext';
import { Calendar, Pin, Settings, QrCode, Star, Users, MessageSquare, FileText, ChevronRight } from './Icons';
import { cx, ensureDexUiStyles } from './dexUi';
import InquiryModal from './InquiryModal';
import { useTickets } from '../context/TicketContext';
import { useIsMobile } from '../utils/useIsMobile';

export default function StartPage(): React.ReactElement {
  const isMobile = useIsMobile();
  const { navigate } = useNavigation();
  const { canCreateEvents, isAdmin, isPowerUser, isFA, rolesReadStatus } = useRoles();
  const { events, isEventsLoading, eventsReadStatus, getMyProxyRegistrations } = useEvents();
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
  // v31.8: `events` ist bei einem 403 auf DEX_Events LEER — die beiden
  // Ableitungen unten sagen dann „kein Organizer, kein Check-in-Team", obwohl
  // sie in Wahrheit nichts wissen. Der Hinweiskasten unter der Kachel haengt
  // deshalb seit v31.8 auch an `eventsReadStatus`, nicht nur an den Rollen.
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
  // v30.5: „F&A Center"-Kachel — Rolle F&A und Admins (Fachkonzept Abschnitt 8).
  const showFATile = isFA || isAdmin;

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

  // v31.9: Die Zeilen-Ansicht unten nutzt `dex-ui-`-Klassen; die Seite hängt
  // weder an Modal noch an WizardFormShell, die das Stylesheet sonst einziehen.
  // Idempotent, kein Hook.
  ensureDexUiStyles();

  // v31.9: EIN Menü statt zwei.
  //
  // Bis v31.8 stand dasselbe Menü zweimal im Code — einmal als Kacheln
  // (Desktop), einmal als Zeilen (Handy) — und die beiden liefen bereits
  // auseinander: andere Untertitel bei der Organizer-Kachel, andere
  // `strokeWidth`, kein einziges `data-tour` im Mobil-Zweig (die Tour zeigte
  // dort also ins Leere) und der Hinweis „Rollen nicht lesbar" fehlte auf dem
  // Handy ganz. Zwei Quellen für dieselbe Aussage laufen immer auseinander,
  // nicht vielleicht. Deshalb: die Menüpunkte sind Daten, die beiden Zweige
  // sind nur noch zwei Darstellungen davon.
  type MenuItem = {
    key: string;
    /** Tour-Anker — bleibt an demselben Menüpunkt wie bisher (tutorialTours.ts). */
    tour?: string;
    /** Symbol in beiden Größen: Kachel 64/1, Zeile 24/1.6. */
    icon: (size: number, strokeWidth: number) => React.ReactNode;
    title: string;
    desc: string;
    onClick?: () => void;
    /** Organizer-Menüpunkt ohne Rechte: öffnet das Anfrage-Modal statt zu navigieren. */
    inquiry?: boolean;
    /** Beschriftung der Anfrage-Handlung (Kachel: Knopf, Zeile: Textzeile). */
    cta?: string;
    badge?: number;
    /** Zusatzklasse der Kachel (Farbakzent aus dem SCSS-Modul). */
    cardClass?: string;
    /** Sichtbarer Grund, warum der Punkt gesperrt ist — auf BEIDEN Wegen. */
    note?: React.ReactNode;
  };

  // v30.81: 403 auf DEX_Roles heißt „Rollenliste nicht lesbar" — wer dort als
  // Organizer steht, hat dann keine Rolle, sondern ein fehlendes Leserecht.
  // Vorher sah die Person nur „Organizer werden?" und niemand wusste, warum die
  // Kachel grau ist.
  // v31.9: Der Hinweis stand nur auf der Kachel — auf dem Handy fehlte genau
  // der Satz, der den grauen Punkt erklärt. Jetzt hängt er am Menüpunkt und
  // wird auf beiden Wegen gerendert. Außerdem standen die typografischen
  // Anführungszeichen als `&bdquo;`/`&ldquo;` in einem JS-String: React
  // maskiert das, die Person las die Entity im Klartext.
  const rightsHint = (rolesReadStatus === 'forbidden' || eventsReadStatus === 'forbidden') ? (
    isDe
      ? 'Deine Rollen oder die Event-Liste konnten nicht geladen werden (fehlendes Leserecht). Bist du bereits Organizer? Dann bitte einen Admin, in der Rollenverwaltung „Rechte prüfen“ auszuführen.'
      : 'Your roles or the event list could not be loaded (missing read access). Already an organizer? Ask an admin to run “Check rights” in role management.'
  ) : null;

  const itemRegister: MenuItem = {
    key: 'register', tour: 'tile-register',
    icon: (s, w) => <Calendar size={s} strokeWidth={w} />,
    title: t('start.register'), desc: t('start.register.desc'),
    onClick: () => navigate('register'),
  };
  const itemMyEvents: MenuItem = {
    key: 'my-events', tour: 'tile-myevents',
    icon: (s, w) => <Pin size={s} strokeWidth={w} />,
    title: t('start.myevents'), desc: t('start.myevents.desc'),
    onClick: () => navigate('my-events'),
  };
  const itemOrganizer: MenuItem = isOrganizer ? {
    key: 'admin', tour: 'tile-admin', cardClass: 'start-card--admin',
    icon: (s, w) => <Settings size={s} strokeWidth={w} />,
    title: t('start.admin'), desc: t('start.admin.desc'),
    onClick: () => navigate('admin'),
  } : {
    key: 'admin', cardClass: 'start-card--admin',
    icon: (s, w) => <Settings size={s} strokeWidth={w} />,
    title: t('start.admin'), desc: t('start.admin.desc'),
    inquiry: true,
    cta: isDe ? 'Organizer werden?' : 'Want to become an organizer?',
    note: rightsHint,
  };
  const itemCheckIn: MenuItem = {
    key: 'check-in', tour: 'tile-checkin', cardClass: 'start-card--checkin',
    icon: (s, w) => <QrCode size={s} strokeWidth={w} />,
    title: isDe ? 'Check-In' : 'Check-in',
    desc: isDe ? 'Teilnehmer einchecken' : 'Check in attendees',
    onClick: () => navigate('check-in'),
  };
  const itemTickets: MenuItem = {
    key: 'tickets',
    icon: (s, w) => <MessageSquare size={s} strokeWidth={w} />,
    title: 'Tickets', desc: isDe ? 'Fragen beantworten' : 'Answer questions',
    badge: openTicketCount > 0 ? openTicketCount : undefined,
    onClick: () => navigate('tickets'),
  };
  const itemAssist: MenuItem = {
    key: 'assistant',
    icon: (s, w) => <Users size={s} strokeWidth={w} />,
    title: isDe ? 'Assistenz' : 'Assistant',
    desc: isDe ? 'Anmeldungen für andere' : 'Registrations for others',
    onClick: () => navigate('assistant'),
  };
  const itemAdminHub: MenuItem = {
    key: 'admin-hub',
    icon: (s, w) => <Star size={s} strokeWidth={w} />,
    title: 'Admin', desc: isDe ? 'Verwaltung & Prozesse' : 'Administration & processes',
    onClick: () => navigate('admin-hub'),
  };
  const itemFA: MenuItem = {
    key: 'fa-center',
    icon: (s, w) => <FileText size={s} strokeWidth={w} />,
    title: 'F&A Center',
    desc: isDe ? 'Abrechnungsrelevante Events' : 'Billing-relevant events',
    onClick: () => navigate('fa-center'),
  };

  const clusters: Array<{ key: string; title: string; items: MenuItem[] }> = [
    { key: 'teilnahme', title: isDe ? 'Teilnahme' : 'Participation', items: [itemRegister, itemMyEvents] },
    { key: 'organisation', title: isDe ? 'Organisation' : 'Organization', items: [itemOrganizer, ...(showCheckInTile ? [itemCheckIn] : [])] },
    { key: 'support', title: isDe ? 'Support' : 'Support', items: showTicketsTile ? [itemTickets] : [] },
    { key: 'verwaltung', title: isDe ? 'Verwaltung' : 'Administration', items: [...(showAssistTile ? [itemAssist] : []), ...(showAdminHubTile ? [itemAdminHub] : []), ...(showFATile ? [itemFA] : [])] },
  ].filter(c => c.items.length > 0);

  if (isMobile) {
    return (
      <div className="page-container">
        <div className="dex-start-rows" style={{ maxWidth: 520, margin: '0 auto', display: 'flex', flexDirection: 'column', gap: 20 }}>
          {clusters.map(c => (
            <div key={c.key}>
              <div className="dex-ui-section-title" style={{ margin: '0 4px 8px' }}>{c.title}</div>
              {/* v31.9: Karte mit Zeilen aus dem gemeinsamen Klassensatz
                  (`dex-ui-card--list` + `dex-ui-row`) statt handgebauter
                  Inline-Styles — dieselbe Zeile wie in „Meine Events". */}
              <div className="dex-ui-card dex-ui-card--list">
                {c.items.map((it, i) => (
                  <button
                    key={it.key}
                    type="button"
                    data-tour={it.tour}
                    className={cx('dex-ui-rowbtn', 'dex-ui-row')}
                    // Innenabstand und Trennlinie stehen bewusst hier: Der
                    // Knopf-Reset `dex-ui-rowbtn` steht im Stylesheet HINTER
                    // `dex-ui-row` und setzt `padding: 0; border: none` — die
                    // Klassen allein ergäben eine randlose, gequetschte Zeile.
                    style={{
                      minHeight: 58, padding: '10px 12px',
                      borderTop: i === 0 ? undefined : '1px solid var(--dex-gray-100)',
                    }}
                    onClick={it.inquiry ? () => setShowInquiry(true) : it.onClick}
                  >
                    <span style={{
                      flex: '0 0 auto', width: 42, height: 42, borderRadius: '50%',
                      display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                      background: 'rgba(134,188,37,0.12)', color: 'var(--dex-green-dark, #4a7c1f)',
                    }}>{it.icon(24, 1.6)}</span>
                    <span className="dex-ui-row-main">
                      {/* v31.9: Titel und Untertitel brechen um, statt zu
                          kürzen — was hier abgeschnitten wird, steht auf dem
                          Handy nirgendwo sonst (Leitfaden 6b). */}
                      <span className="dex-ui-row-title dex-ui-row-title--wrap" style={{ display: 'block', fontSize: '0.98rem' }}>
                        {it.title}
                      </span>
                      <span className="dex-ui-row-sub" style={{ display: 'block', fontSize: '0.8rem' }}>
                        {it.desc}
                      </span>
                      {it.cta && (
                        <span className="dex-ui-row-link" style={{ display: 'block', marginTop: 4, fontWeight: 600, fontSize: '0.8rem' }}>
                          {it.cta}
                        </span>
                      )}
                      {it.note && (
                        <span className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" style={{ display: 'flex', marginTop: 6 }}>
                          {it.note}
                        </span>
                      )}
                    </span>
                    {typeof it.badge === 'number' && (
                      <span className="dex-ui-pill dex-ui-pill--orange" style={{ flex: '0 0 auto' }}>{it.badge}</span>
                    )}
                    <span style={{ flex: '0 0 auto', color: 'var(--dex-gray-400)', display: 'inline-flex', marginLeft: 2 }} aria-hidden="true">
                      <ChevronRight size={18} />
                    </span>
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
        <InquiryModal open={showInquiry} onClose={() => setShowInquiry(false)} />
      </div>
    );
  }

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
          gap: 28px; max-width: 800px; margin: 0 auto;
        }
        .dex-cluster-title {
          font-size: 0.9rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px;
          color: var(--dex-gray-500); margin: 0 0 14px; padding-bottom: 8px;
          border-bottom: 1px solid var(--dex-gray-200);
        }
        /* v26.6: alle Kacheln GLEICH GROSS (quadratisch — wie die Tickets-Kachel).
           v26.8: auf dem Desktop größer (380px) + größerer Text, damit nichts
           untergeht. Passen zwei nebeneinander, stehen sie in einer Zeile; sonst
           brechen sie um, statt unterschiedlich zu schrumpfen. */
        .dex-cluster-tiles { display: flex; flex-wrap: wrap; gap: 22px; align-items: stretch; }
        .dex-cluster .start-card {
          flex: 0 0 auto !important; width: 380px !important; max-width: 100% !important;
          padding: 32px 22px !important; min-height: 0 !important; aspect-ratio: 1 / 1; gap: 10px !important;
        }
        /* v26.5: echter KREIS hinter dem Icon — flex:0 0 auto + aspect-ratio
           verhindert, dass die Flex-Stauchung den Kreis zur Ellipse macht. */
        .dex-cluster .start-card__icon { width: 104px !important; height: 104px !important; flex: 0 0 auto !important; aspect-ratio: 1 / 1 !important; border-radius: 50% !important; margin-bottom: 6px !important; }
        .dex-cluster .start-card__icon svg { width: 54px !important; height: 54px !important; }
        .dex-cluster .start-card h2 { font-size: 1.32rem !important; }
        .dex-cluster .start-card p { font-size: 0.98rem !important; white-space: normal !important; }
        @media (max-width: 820px) {
          .dex-cluster-grid { max-width: 420px; }
        }
        @media (max-width: 480px) {
          .dex-cluster .start-card { width: 300px !important; padding: 24px 16px !important; }
          .dex-cluster .start-card__icon { width: 72px !important; height: 72px !important; aspect-ratio: 1 / 1 !important; }
          .dex-cluster .start-card__icon svg { width: 38px !important; height: 38px !important; }
          .dex-cluster .start-card h2 { font-size: 1.1rem !important; }
          .dex-cluster .start-card p { font-size: 0.9rem !important; }
        }
      `}</style>
      <div className="dex-cluster-grid">
        {clusters.map(c => (
          <div key={c.key} className="dex-cluster">
            <div className="dex-cluster-title">{c.title}</div>
            <div className="dex-cluster-tiles">
              {c.items.map(it => (
                <div
                  key={it.key}
                  className={cx('card', !it.inquiry && 'card-clickable', 'start-card', it.cardClass)}
                  data-tour={it.tour}
                  style={it.inquiry
                    ? { position: 'relative', cursor: 'default', opacity: 0.55 }
                    : (typeof it.badge === 'number' ? { position: 'relative' } : undefined)}
                  onClick={it.inquiry ? undefined : it.onClick}
                >
                  <div className="start-card__icon">{it.icon(64, 1)}</div>
                  <h2>{it.title}</h2>
                  <p>{it.desc}</p>
                  {typeof it.badge === 'number' && (
                    <span style={{
                      position: 'absolute', top: 10, right: 10, background: '#ed8b00', color: '#fff',
                      borderRadius: 12, minWidth: 22, height: 22, padding: '0 6px',
                      display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700,
                    }}>{it.badge}</span>
                  )}
                  {it.inquiry && (
                    <div style={{ position: 'absolute', inset: 0, display: 'flex', flexDirection: 'column', gap: 8, alignItems: 'center', justifyContent: 'center', padding: 10 }}>
                      <button
                        type="button"
                        onClick={(e) => { e.stopPropagation(); setShowInquiry(true); }}
                        style={{
                          background: 'var(--dex-green, #86bc25)', color: '#fff', border: 'none', borderRadius: 14,
                          padding: '8px 12px', fontWeight: 700, cursor: 'pointer', boxShadow: '0 4px 14px rgba(0,0,0,0.18)',
                          fontSize: '0.74rem', lineHeight: 1.25, textAlign: 'center',
                        }}
                      >
                        {it.cta}
                      </button>
                      {it.note && (
                        <div style={{ fontSize: '0.7rem', lineHeight: 1.3, textAlign: 'center', color: 'var(--dex-orange-dark, #b35a00)', background: 'rgba(255,255,255,0.92)', borderRadius: 8, padding: '6px 8px', maxWidth: 240 }}>
                          {it.note}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              ))}
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
