/**
 * Header-Komponente
 *
 * Sticky Header mit dynamischem Titel je nach aktueller Seite.
 * Auf der Landing Page wird das Deloitte-Logo angezeigt,
 * auf allen anderen Seiten ein Zurück-Button.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useDialog } from '../context/DialogContext';
import { DELOITTE_LOGO_HEADER } from '../data/brandLogos';
import { useEvents } from '../context/EventContext';
import { useLanguage } from '../context/LanguageContext';
import { ChevronLeft, Book, RefreshCw, Info, Users } from './Icons';
import { Icon } from '@fluentui/react/lib/Icon';
import ImpersonateModal from './ImpersonateModal';
import LandingInfoModal from './LandingInfoModal';
import GlobalSearch from './GlobalSearch';
import QuestionButton from './QuestionButton';
import { useTutorial } from './tutorial/TutorialGuide';

export default function Header(): React.ReactElement {
  const { currentPage, navigate, selectedEventId } = useNavigation();
  const { currentUser, photoUrl } = useCurrentUser();
  const { currentUserRole, originalIsAdmin } = useRoles();
  const [showImpersonate, setShowImpersonate] = React.useState(false);
  const { events } = useEvents();
  // v22.50: Das frühere Check-in-Icon im Header ist entfallen — der Zugang zur
  // Check-in-Seite läuft jetzt über die globale Such-Leiste (GlobalSearch) bzw.
  // die Aktionen im Organizer Center.
  const { t, locale, setLocale } = useLanguage();
  const { showAlert } = useDialog();
  // v24.19: Tutorial-CTA wandert von der Landing Page mittig in den Header.
  const { openTutorial } = useTutorial();
  // v24.22: „Über die App" liegt jetzt im Header (links neben Handbuch) — das
  // Info-Modal wird hier verwaltet statt auf der Landing Page.
  const [showAbout, setShowAbout] = React.useState(false);
  const [showPopup, setShowPopup] = React.useState(false);
  // v24.69: Tutorial-CTA im Header ist per X ausblendbar — Zustand bleibt in
  // localStorage erhalten (einmal weggeklickt = bleibt weg).
  const [tutorialCtaHidden, setTutorialCtaHidden] = React.useState<boolean>(() => {
    try { return window.localStorage.getItem('dex_tutorial_cta_hidden') === '1'; } catch { return false; }
  });
  const dismissTutorialCta = (): void => {
    try { window.localStorage.setItem('dex_tutorial_cta_hidden', '1'); } catch { /* */ }
    setTutorialCtaHidden(true);
  };
  const isLanding = currentPage === 'landing';
  const isStart = currentPage === 'start';

  // v18.35: Hinweis-Chip, wenn die Anmeldeseite in einer festen Sprache
  // angezeigt wird (Organizer hat sie pro Event vorgegeben). Der Text steht
  // bewusst IN der erzwungenen Sprache — passend zu dem, was der Teilnehmer
  // auf der Anmeldeseite sieht.
  const regHintEvent = events.find(e => e.id === selectedEventId);
  const forcedRegLang: 'de' | 'en' | undefined =
    (regHintEvent?.registrationLanguage === 'de' || regHintEvent?.registrationLanguage === 'en')
      ? regHintEvent.registrationLanguage : undefined;
  const showRegLangHint = currentPage === 'registration' && !!forcedRegLang;
  const regLangHintText = forcedRegLang === 'de'
    ? 'Organizer set the language to German for this registration form'
    : 'Organizer set the language to English for this registration form';
  // v24.11: Bei fest vorgegebener Formularsprache zeigt der Picker diese als
  // aktiv; der Klick auf die andere Sprache wird mit Hinweis abgelehnt.
  const pickerLang: 'de' | 'en' = forcedRegLang || locale;
  const handleLangClick = (lang: 'de' | 'en'): void => {
    if (forcedRegLang && forcedRegLang !== lang) {
      const langName = forcedRegLang === 'en'
        ? (locale === 'de' ? 'Englisch' : 'English')
        : (locale === 'de' ? 'Deutsch' : 'German');
      showAlert(
        locale === 'de'
          ? `Das ist hier nicht möglich: Der Organisator hat die Sprache dieses Anmeldeformulars aus Einheitlichkeit fest auf ${langName} gestellt.`
          : `Not possible here: the organizer set this registration form's language to ${langName} for consistency.`,
        { variant: 'info' },
      );
      return;
    }
    setLocale(lang);
  };

  const pageIdMap: Record<string, string> = {
    'landing': 'landing',
    'start': 'start',
    'register': 'event-list',
    'registration': 'register',
    'my-events': 'my-events',
    'create-event': 'event-create',
    'edit-event': 'event-edit',
    'settings': 'settings',
    'admin': 'admin-center',
    'profile': 'profile',
    'role-matrix': 'role-matrix',
    'participants': 'participants',
    'flowcharts': 'flowcharts',
    'check-in': 'check-in',
    'manual': 'manual',
    'email-templates': 'email-templates',
    'stats-archive': 'stats-archive',
  };
  // v10.19: Admin-Center hat zwei Sub-Views — die Übersichtsliste aller Events
  // ('admin-center') und die Detail-Ansicht eines konkreten Events
  // ('admin-event'). Vorher hatten beide dieselbe Page-ID, was Bug-Reports
  // ungenau gemacht hat. Detail-Modus erkennen wir an `selectedEventId`.
  const pageIdLabel = currentPage === 'admin'
    ? (selectedEventId ? 'admin-event' : 'admin-center')
    : (pageIdMap[currentPage] || currentPage);

  // v6.26: Mobile-Detection für die "Jetzt einchecken"-Sprechblase neben dem
  // QR-Icon. Wird nur auf Mobilgeräten angezeigt (Viewport <= 768px), auf
  // Desktop bleibt der Header schlank.
  // v6.28: Handbuch-Preview kann Mobile erzwingen (window.__dexForceMobile),
  // damit die Bubble im Phone-Frame-Rahmen des AppPreview sichtbar ist —
  // auch wenn der User das Handbuch am Desktop öffnet.
  const [isMobile, setIsMobile] = React.useState<boolean>(() => {
    if (typeof window === 'undefined') return false;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if ((window as any).__dexForceMobile) return true;
    return !!(window.matchMedia && window.matchMedia('(max-width: 768px)').matches);
  });
  React.useEffect(() => {
    if (typeof window === 'undefined' || !window.matchMedia) return;
    const mq = window.matchMedia('(max-width: 768px)');
    const handler = (e: MediaQueryListEvent): void => setIsMobile(e.matches);
    if (mq.addEventListener) mq.addEventListener('change', handler);
    // eslint-disable-next-line @typescript-eslint/no-deprecated
    else mq.addListener(handler);
    return () => {
      if (mq.removeEventListener) mq.removeEventListener('change', handler);
      // eslint-disable-next-line @typescript-eslint/no-deprecated
      else mq.removeListener(handler);
    };
  }, []);

  // Titel-Mapping je nach aktuellem Seitenstatus
  const getTitle = (): string => {
    switch (currentPage) {
      case 'start': return t('start.title');
      case 'register': return t('header.registration');
      case 'registration': return 'Registration Deloitte Events';
      case 'my-events': return t('myevents.title');
      case 'create-event': return t('header.createevent');
      case 'edit-event': return t('header.editevent');
      case 'settings': return t('header.settings');
      case 'profile': return t('header.profile');
      case 'admin': return 'Organizer';
      case 'role-matrix': return t('header.rolematrix');
      case 'participants': return t('header.participants');
      case 'flowcharts': return t('header.flowcharts');
      case 'check-in': return t('header.checkin');
      case 'manual': return t('header.manual');
      case 'email-templates': return locale === 'de' ? 'Mail-Vorlagen' : 'Mail templates';
      case 'stats-archive': return locale === 'de' ? 'Statistik-Archiv' : 'Statistics archive';
      default: return '';
    }
  };

  // Rollen-Farbe
  const roleColors: Record<string, { bg: string; color: string }> = {
    'Admin': { bg: '#e8f5e9', color: '#2e7d32' },
    'Organizer': { bg: '#e3f2fd', color: '#1565c0' },
    'User': { bg: '#f5f5f5', color: '#666' },
  };
  const rc = roleColors[currentUserRole] || roleColors['User'];

  // v22.28: Header beim Scrollen oben festpinnen. Das CSS-`position: sticky`
  // auf .header wird im SP-Canvas durch Overflow-Vorfahren ausgehebelt
  // (gleiche Falle wie bei PageId/fixed) — deshalb JS-Pin: ein Platzhalter-
  // Div hält die Höhe, der Header wechselt auf `position: fixed`, sobald der
  // Platzhalter aus dem sichtbaren Bereich scrollt. Der Top-Offset wird
  // dynamisch unter der SP-Chrome-Leiste gemessen (Suite-Bar ist fixed).
  const headerPlaceholderRef = React.useRef<HTMLDivElement | null>(null);
  const [headerPin, setHeaderPin] = React.useState<null | { top: number; left: number; width: number; height: number }>(null);
  React.useEffect(() => {
    const chromeTop = (): number => {
      const candidates = ['[data-automation-id="contentScrollRegion"]', '.SPPageChromeAppDiv', '#spPageCanvasContent'];
      for (const sel of candidates) {
        const el = document.querySelector(sel);
        if (el) return Math.max(0, el.getBoundingClientRect().top);
      }
      return 0;
    };
    const update = (): void => {
      const ph = headerPlaceholderRef.current;
      if (!ph) return;
      const r = ph.getBoundingClientRect();
      const top = chromeTop();
      if (r.top < top) {
        const height = ph.offsetHeight > 0 ? ph.offsetHeight : 64;
        setHeaderPin(prev => {
          const next = { top, left: r.left, width: r.width, height: prev ? prev.height : height };
          if (prev && prev.top === next.top && prev.left === next.left && prev.width === next.width) return prev;
          return next;
        });
      } else {
        setHeaderPin(null);
      }
    };
    update();
    window.addEventListener('scroll', update, true);
    window.addEventListener('resize', update);
    return () => {
      window.removeEventListener('scroll', update, true);
      window.removeEventListener('resize', update);
    };
  }, []);

  return (
    <div ref={headerPlaceholderRef} style={headerPin ? { height: headerPin.height } : undefined}>
    <header
      className="header"
      style={headerPin ? {
        position: 'fixed', top: headerPin.top, left: headerPin.left, width: headerPin.width,
        zIndex: 1000, boxShadow: '0 2px 10px rgba(0,0,0,0.10)',
      } : undefined}
    >
      {/* v24.19: Tutorial-CTA mittig im Header (nur Landing) — ersetzt die
          frühere Bubble unten auf der Landing Page. Absolut zentriert, damit
          die Position unabhängig von den Breiten links/rechts wirklich mittig
          sitzt. */}
      {isLanding && !tutorialCtaHidden && (
        <div
          style={{
            position: 'absolute', left: '50%', top: '50%', transform: 'translate(-50%, -50%)',
            display: 'inline-flex', alignItems: 'center', gap: 4, whiteSpace: 'nowrap',
            background: 'var(--dex-green)', borderRadius: 999,
            boxShadow: '0 2px 8px rgba(0,0,0,0.10)', maxWidth: 'min(46vw, 460px)',
          }}
        >
          <button
            type="button"
            onClick={openTutorial}
            style={{
              display: 'inline-flex', alignItems: 'center', gap: 8, whiteSpace: 'nowrap',
              background: 'transparent', color: '#fff',
              padding: '9px 6px 9px 18px', borderRadius: 999,
              fontSize: '0.95rem', lineHeight: 1.2, fontFamily: 'inherit',
              border: 'none', cursor: 'pointer', overflow: 'hidden',
            }}
            title={locale === 'de' ? 'Geführtes Tutorial starten' : 'Start the guided tutorial'}
          >
            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>
              {locale === 'de'
                ? <><strong>Neu hier?</strong> Starte das DEX Tutorial</>
                : <><strong>New here?</strong> Start the DEX tutorial</>}
            </span>
          </button>
          <button
            type="button"
            onClick={dismissTutorialCta}
            aria-label={locale === 'de' ? 'Tutorial-Hinweis ausblenden' : 'Hide tutorial hint'}
            title={locale === 'de' ? 'Ausblenden (nicht mehr anzeigen)' : 'Hide (do not show again)'}
            style={{
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              width: 26, height: 26, marginRight: 6, borderRadius: '50%',
              background: 'rgba(255,255,255,0.2)', color: '#fff',
              border: 'none', cursor: 'pointer', fontSize: '1rem', lineHeight: 1, fontFamily: 'inherit',
            }}
          >
            ×
          </button>
        </div>
      )}
      <div className="header-left">
        {isLanding ? (
          <div className="header-logo">
            {/* v24.17: Header-Variante des offiziellen Deloitte-Logos (Repo-Asset)
                statt Text-Wortmarke. */}
            {DELOITTE_LOGO_HEADER
              ? <img src={DELOITTE_LOGO_HEADER} alt="Deloitte" style={{ height: 30, width: 'auto', display: 'block' }} />
              : <>Deloitte<span>.</span></>}
          </div>
        ) : (
          <>
            {/* v22.28: „Zurück"-Beschriftung neben dem Chevron — die runde
                Icon-Box allein war als Zurück-Navigation nicht klar genug. */}
            <button
              className="back-btn"
              onClick={() => navigate(isStart ? 'landing' : 'start')}
              aria-label={locale === 'de' ? 'Zurück' : 'Back'}
              style={{ width: 'auto', borderRadius: 999, padding: '0 16px 0 10px', fontSize: '0.85rem', fontWeight: 600 }}
            >
              <ChevronLeft size={20} /> {locale === 'de' ? 'Zurück' : 'Back'}
            </button>
            <span className="header-title" style={{ border: 'none', paddingLeft: 0, fontWeight: 500 }}>
              {getTitle()}
            </span>
          </>
        )}
      </div>
      {/* v22.50: Globale Such-Leiste — ersetzt das frühere Check-in-Icon im
          Header. Self-gated (nur Admin/Organizer eigener Events). Auf der
          Landing Page ausgeblendet, weil dort der Boot-/Logo-Look gilt. */}
      {!isLanding && <GlobalSearch />}
      <div className="header-right">
        {/* v26: Grüner „Hast du Fragen?"-Button — Ticketsystem für alle User.
            v26.34: jetzt auch auf der Landing Page im Header sichtbar. */}
        <QuestionButton isMobile={isMobile} />
        {/* v9.29: Refresh-Button im Header — ersetzt die alten in-page
            Aktualisieren-Buttons in AdminPage. Nur auf Seiten anzeigen, auf
            denen ein Refresh sinnvoll ist (Admin Center, Event-Liste,
            Meine Events, Teilnehmer-Liste). Triggert ein globales Event,
            das von der jeweiligen Page abgegriffen wird. */}
        {(currentPage === 'admin' || currentPage === 'register' || currentPage === 'my-events' || currentPage === 'participants') && (
          <button
            className="header-icon-btn"
            onClick={() => { window.dispatchEvent(new CustomEvent('dex-refresh-page')); }}
            title={locale === 'de' ? 'Aktualisieren' : 'Refresh'}
            aria-label={locale === 'de' ? 'Aktualisieren' : 'Refresh'}
            style={isMobile ? {} : { width: 'auto', padding: '0 12px', display: 'inline-flex', alignItems: 'center', gap: 8 }}
          >
            <RefreshCw size={18} />
            {!isMobile && (
              <span style={{ fontSize: '0.85rem', fontWeight: 500, lineHeight: 1 }}>
                {locale === 'de' ? 'Aktualisieren' : 'Refresh'}
              </span>
            )}
          </button>
        )}
        {/* v24.69: „Demo: als User testen" aus dem User-Menü in den Header
            verlegt (nur echte Admins). Öffnet den Impersonate-Dialog. */}
        {originalIsAdmin && (
          <button
            className="header-icon-btn"
            onClick={() => setShowImpersonate(true)}
            title={locale === 'de' ? 'Demo: als User testen' : 'Demo: test as a user'}
            style={isMobile ? {} : { width: 'auto', padding: '0 12px', display: 'inline-flex', alignItems: 'center', gap: 8 }}
          >
            <Users size={18} />
            {!isMobile && (
              <span style={{ fontSize: '0.85rem', fontWeight: 500, lineHeight: 1 }}>
                {locale === 'de' ? 'Demo' : 'Demo'}
              </span>
            )}
          </button>
        )}
        {/* v24.22: „Über die App" — von der Landing Page in den Header verlegt,
            links neben „Handbuch". Öffnet dasselbe Info-Modal. */}
        <button
          className="header-icon-btn"
          onClick={() => setShowAbout(true)}
          title={locale === 'de' ? 'Über die App' : 'About the app'}
          style={isMobile ? {} : { width: 'auto', padding: '0 12px', display: 'inline-flex', alignItems: 'center', gap: 8 }}
        >
          <Info size={18} />
          {!isMobile && (
            <span style={{ fontSize: '0.85rem', fontWeight: 500, lineHeight: 1 }}>
              {locale === 'de' ? 'Über die App' : 'About the app'}
            </span>
          )}
        </button>
        <button
          className="header-icon-btn"
          onClick={() => navigate('manual')}
          title={t('header.manual')}
          style={{
            ...(currentPage === 'manual' ? { background: 'var(--dex-gray-200)' } : {}),
            // v6.36: Auf Desktop zusätzlich den Text "Handbuch" / "Handbook"
            // neben dem Icon, weil das Icon allein nicht selbsterklärend ist.
            // v7.2: explizit flex + alignItems:center, damit Icon und Text
            // vertikal auf derselben Baseline sitzen (vorher lag der Text
            // leicht versetzt, weil der Button aus der CSS-Klasse keine
            // Flex-Alignment-Regel bekommt, wenn wir width:auto setzen).
            ...(isMobile ? {} : { width: 'auto', padding: '0 12px', display: 'inline-flex', alignItems: 'center', gap: 8 }),
          }}
        >
          <Book size={20} />
          {!isMobile && (
            <span style={{ fontSize: '0.85rem', fontWeight: 500, lineHeight: 1 }}>
              {t('header.manual')}
            </span>
          )}
        </button>
        {/* v7.26: Sprach-Toggle DE/EN — lässt den User auch im laufenden
            Tool zwischen Deutsch und Englisch wechseln. Visuell ein kleiner
            Pill-Toggle im Header-Style. */}
        <div
          role="group"
          aria-label={locale === 'de' ? 'Sprache wechseln' : 'Switch language'}
          style={{
            display: 'inline-flex', alignItems: 'center',
            background: 'var(--dex-gray-100, #f3f4f6)',
            borderRadius: 999, padding: 2, gap: 2,
            height: 30, alignSelf: 'center',
          }}
        >
          <button
            type="button"
            onClick={() => handleLangClick('de')}
            title="Deutsch"
            style={{
              padding: '3px 10px', borderRadius: 999,
              border: 'none', cursor: 'pointer',
              fontSize: '0.72rem', fontWeight: 700, fontFamily: 'inherit',
              background: pickerLang === 'de' ? '#fff' : 'transparent',
              color: pickerLang === 'de' ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-500)',
              boxShadow: pickerLang === 'de' ? '0 1px 3px rgba(0,0,0,0.08)' : 'none',
              transition: 'all 0.15s ease',
            }}
          >DE</button>
          <button
            type="button"
            onClick={() => handleLangClick('en')}
            title="English"
            style={{
              padding: '3px 10px', borderRadius: 999,
              border: 'none', cursor: 'pointer',
              fontSize: '0.72rem', fontWeight: 700, fontFamily: 'inherit',
              background: pickerLang === 'en' ? '#fff' : 'transparent',
              color: pickerLang === 'en' ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-500)',
              boxShadow: pickerLang === 'en' ? '0 1px 3px rgba(0,0,0,0.08)' : 'none',
              transition: 'all 0.15s ease',
            }}
          >EN</button>
        </div>
        {/* v24.11: Hinweis-Chip RECHTS neben dem Picker, grün — „Organizer set
            the language ... for this registration form". */}
        {showRegLangHint && (
          <span
            title={forcedRegLang === 'de'
              ? 'Dieses Anmeldeformular wird auf Deutsch angezeigt.'
              : 'This registration form is shown in English.'}
            style={{
              display: 'inline-flex', alignItems: 'center', gap: 5, alignSelf: 'center',
              background: 'rgba(134,188,37,0.14)', color: 'var(--dex-green-dark, #4a7c1f)',
              border: '1px solid var(--dex-green, #86bc25)', borderRadius: 999,
              padding: '3px 10px', fontSize: '0.72rem', fontWeight: 600, whiteSpace: 'nowrap',
            }}
          >
            <Icon iconName="Globe" style={{ fontSize: 12 }} />
            {regLangHintText}
          </span>
        )}
        {/* v24.84: Settings-(Zahnrad-)Icon im Header entfernt — die Rollen-
            verwaltung/Settings ist über die Admin-Kachel (Admin-Hub) erreichbar. */}
        {/* User-Avatar mit Initialen + Popup.
            v22.23: data-tour-Anker auf dem Wrapper — die Klasse .header-avatar
            existiert nur im Initialen-Fallback (mit Profilfoto rendert ein
            <img>), der Tutorial-Spotlight braucht aber beide Varianten. */}
        <div style={{ position: 'relative' }} data-tour="header-avatar">
          {photoUrl ? (
            <img
              src={photoUrl}
              alt={`${currentUser.firstName} ${currentUser.surname}`}
              onClick={() => setShowPopup(!showPopup)}
              style={{ width: 36, height: 36, borderRadius: '50%', objectFit: 'cover', cursor: 'pointer' }}
              onError={(e) => { (e.target as HTMLImageElement).style.display = 'none'; }}
            />
          ) : (
            <div
              className="header-avatar"
              title={`${currentUser.firstName} ${currentUser.surname}`}
              onClick={() => setShowPopup(!showPopup)}
              style={{ cursor: 'pointer' }}
            >
              {currentUser.firstName ? currentUser.firstName[0] : ''}{currentUser.surname ? currentUser.surname[0] : ''}
            </div>
          )}
          {showPopup && (
            <div style={{
              position: 'absolute', right: 0, top: '100%', marginTop: 8,
              background: '#fff', borderRadius: 'var(--dex-radius-lg, 12px)',
              boxShadow: '0 8px 32px rgba(0,0,0,0.18)', padding: '20px 24px',
              minWidth: 260, zIndex: 1000,
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 16 }}>
                {photoUrl ? (
                  <img src={photoUrl} alt="Profile" style={{ width: 48, height: 48, borderRadius: '50%', objectFit: 'cover' }} />
                ) : (
                  <div style={{
                    width: 48, height: 48, borderRadius: '50%',
                    background: 'linear-gradient(135deg, #86bc25, #0076a8)',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    color: '#fff', fontWeight: 700, fontSize: '1.1rem',
                  }}>
                    {currentUser.firstName ? currentUser.firstName[0] : ''}{currentUser.surname ? currentUser.surname[0] : ''}
                  </div>
                )}
                <div>
                  <div style={{ fontWeight: 600, fontSize: '1rem' }}>{currentUser.firstName} {currentUser.surname}</div>
                  <div style={{ color: '#666', fontSize: '0.85rem' }}>{currentUser.email}</div>
                </div>
              </div>
              {currentUser.location && (
                <div style={{ fontSize: '0.85rem', color: '#555', marginBottom: 8 }}>
                  Location: {currentUser.location}
                </div>
              )}
              <div style={{ fontSize: '0.85rem', marginBottom: 16 }}>
                <span style={{
                  display: 'inline-block', padding: '2px 10px', borderRadius: 12,
                  background: rc.bg, color: rc.color,
                  fontSize: '0.8rem', fontWeight: 500,
                }}>
                  {currentUserRole}
                </span>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                <button
                  className="btn btn-secondary btn-block"
                  style={{ fontSize: '0.85rem' }}
                  onClick={() => { setShowPopup(false); navigate('profile'); }}
                >
                  {t('profile.viewfull')}
                </button>
                {/* v24.69: „Rollenverwaltung" (jetzt eigene Admin-Hub-Kachel) und
                    „Demo: als User testen" (jetzt eigener Header-Button) sind aus
                    dem User-Menü entfernt. */}
              </div>
              <div
                title="Page-ID — bei UI-Anfragen kannst du diese ID nennen, dann finde ich die Seite sofort."
                style={{
                  marginTop: 14, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)',
                  fontSize: '0.7rem', color: 'var(--dex-gray-500)',
                  display: 'flex', alignItems: 'center', gap: 6, fontFamily: 'monospace',
                }}
              >
                <span style={{ color: 'var(--dex-gray-400)' }}>Page-ID:</span>
                <span style={{ fontWeight: 600, color: 'var(--dex-gray-700)' }}>{pageIdLabel}</span>
              </div>
            </div>
          )}
        </div>
      </div>
      <ImpersonateModal open={showImpersonate} onClose={() => setShowImpersonate(false)} />
      <LandingInfoModal open={showAbout} locale={locale === 'de' ? 'de' : 'en'} onClose={() => setShowAbout(false)} />
    </header>
    </div>
  );
}
