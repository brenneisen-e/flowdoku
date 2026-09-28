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
import { ChevronLeft, Check, Book, RefreshCw, Info, Users, Menu } from './Icons';
import { Icon } from '@fluentui/react/lib/Icon';
import ImpersonateModal from './ImpersonateModal';
import LandingInfoModal from './LandingInfoModal';
import GlobalSearch from './GlobalSearch';
import QuestionButton from './QuestionButton';
import { cx, ensureDexUiStyles } from './dexUi';
import { useTutorial } from './tutorial/TutorialGuide';
import { useIsMobile } from '../utils/useIsMobile';
// v28.98: „Zurück" waehrend eines laufenden Speichervorgangs sperren.
import { isSaveInProgress, subscribeSaveInProgress } from '../utils/saveGuard';

export default function Header(): React.ReactElement {
  // v32.3: Das Burger-Menü nutzt dex-ui-menuitem — das Stylesheet muss auch
  // dann da sein, wenn noch kein Modal es eingefügt hat (idempotent).
  ensureDexUiStyles();
  // v28.98: Laeuft gerade ein Speichervorgang? Dann ist „Zurück" gesperrt —
  // ein Abbruch mittendrin hinterlaesst ein halb angelegtes Event.
  const [saveBusy, setSaveBusy] = React.useState<boolean>(isSaveInProgress());
  React.useEffect(() => subscribeSaveInProgress(setSaveBusy), []);
  const { currentPage, navigate, goBack, selectedEventId } = useNavigation();
  const { currentUser, photoUrl } = useCurrentUser();
  const { currentUserRole, originalIsAdmin, previewAsUser, setPreviewAsUser } = useRoles();
  // v30.43: Hover für den Ansicht-Wechselschalter. Inline-Styles können kein
  // :hover; ohne Reaktion liest sich der Schalter als Beschriftung.
  // v32.2.1: Demo-Menü (Ansicht + Impersonate), schließt bei Klick daneben.
  const [demoOpen, setDemoOpen] = React.useState(false);
  const demoRef = React.useRef<HTMLDivElement | null>(null);
  React.useEffect(() => {
    if (!demoOpen) return undefined;
    const onDown = (e: MouseEvent): void => { if (demoRef.current && !demoRef.current.contains(e.target as Node)) setDemoOpen(false); };
    const onKey = (e: KeyboardEvent): void => { if (e.key === 'Escape') setDemoOpen(false); };
    window.addEventListener('mousedown', onDown);
    window.addEventListener('keydown', onKey);
    return () => { window.removeEventListener('mousedown', onDown); window.removeEventListener('keydown', onKey); };
  }, [demoOpen]);
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
  const isLanding = currentPage === 'landing';

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
    'feedback-overview': 'feedback-overview',
    'intro-onepager': 'intro-onepager',
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
  // v26.36: Zentraler Mobile-Hook (spiegelt das frühere inline-matchMedia inkl.
  // window.__dexForceMobile für die Handbuch-Preview) — appweit dieselbe Grenze.
  const isMobile = useIsMobile();

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
      case 'feedback-overview': return locale === 'de' ? 'Feedback der Organizer' : 'Organizer feedback';
      case 'intro-onepager': return locale === 'de' ? 'Einführungs-Onepager' : 'Introduction one-pager';
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

  // v7.26: Sprach-Toggle DE/EN — lässt den User auch im laufenden Tool
  // zwischen Deutsch und Englisch wechseln. v32.2.1: als Konstante, weil er
  // am Rechner in der Mitte und am Handy rechts steht.
  const langToggle = (
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
  );

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
      {/* v26.36: Auf Mobile ausgeblendet — die absolut zentrierte CTA-Pille
          überlappt sonst das rechte Icon-Cluster und macht Buttons unklickbar. */}
      {/* v30.25: Admins sehen die „Neu hier?"-Pille nicht mehr. Sie hing
          bisher nur an `isLanding` — also an der Seite, NICHT daran, ob
          jemand neu ist; ein Admin bekam sie bei jedem Start angeboten,
          bis er sie einmal per × wegklickte (localStorage, pro Browser).
          Für Admins gibt es das Tutorial stattdessen als Kachel im Admin
          Center; Organizer und Teilnehmer behalten die Pille (für sie ist
          es der einzige Einstieg). */}
      <div className="header-left" style={isMobile ? undefined : { flex: '1 1 0', minWidth: 0 }}>
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
              // v31.59: Zurück heißt zurück — über den Stack, nicht hart auf die
              // Startseite. Ohne Rückweg (Deep-Link, F5) greift `fallbackFor`.
              onClick={() => { if (!saveBusy) goBack(); }}
              disabled={saveBusy}
              aria-label={locale === 'de' ? 'Zurück' : 'Back'}
              title={saveBusy
                ? (locale === 'de'
                  ? 'Es wird gerade gespeichert — bitte warte, bis der Vorgang durch ist. Ein Abbruch mittendrin hinterlässt ein halb angelegtes Event.'
                  : 'Saving is in progress — please wait. Leaving now would leave the event half-created.')
                : undefined}
              style={{
                width: 'auto', borderRadius: 999, padding: '0 16px 0 10px', fontSize: '0.85rem', fontWeight: 600, flexShrink: 0, whiteSpace: 'nowrap',
                ...(saveBusy ? { opacity: 0.5, cursor: 'not-allowed' } : {}),
              }}
            >
              <ChevronLeft size={20} /> {locale === 'de' ? 'Zurück' : 'Back'}
            </button>
            {/* v32.3: eine Zeile, notfalls gekürzt — seit der Header eine Mitte
                hat, ist links weniger Platz, und in der Vorschau (Rahmen) brach
                „Aktuelle Events | Verfügbar an deinem Standort" wortweise um. */}
            <span className="header-title" title={getTitle()} style={{ border: 'none', paddingLeft: 0, fontWeight: 500, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', minWidth: 0, flex: '0 1 auto' }}>
              {getTitle()}
            </span>
          </>
        )}
      </div>
      {/* v22.50: Globale Such-Leiste — ersetzt das frühere Check-in-Icon im
          Header. Self-gated (nur Admin/Organizer eigener Events). Auf der
          Landing Page ausgeblendet, weil dort der Boot-/Logo-Look gilt. */}
      {/* v32.2.1: Mitte des Headers (Nutzer-Ansage 28.09.2026: „die
          Sprachauswahl horizontal zentriert in der Mitte, daneben den
          Fragen-Button"). Eine Gruppe aus Suche (außer Landing), „Neu hier?"
          (nur Landing), Sprachauswahl und „Hast du Fragen?". Links und rechts
          teilen sich den Rest je zur Hälfte (flex 1 1 0) — so steht die
          Gruppe wirklich in der Mitte, solange Platz ist. Die „Neu hier?"-
          Pille lag vorher absolut in der Mitte; sie wandert in die Gruppe,
          sonst lägen beide übereinander. Auf dem Handy bleibt alles rechts. */}
      {isMobile ? (!isLanding && <GlobalSearch />) : (
        <div className="header-center" style={{ flex: isLanding ? '0 1 auto' : '0 1 780px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 12, minWidth: 0 }}>
          {!isLanding && <GlobalSearch />}
          {/* v32.3: Die Pille „Neu hier? Starte das DEX Tutorial" ist weg — der
              Einstieg steckt jetzt im Zwischendialog von „Hast du Fragen?"
              (Nutzer-Ansage 28.09.2026). */}
          {langToggle}
          <span style={{ flexShrink: 0, whiteSpace: 'nowrap', display: 'inline-flex' }}><QuestionButton isMobile={isMobile} onAbout={() => setShowAbout(true)} /></span>
        </div>
      )}
      {/* v32.3: Aktualisieren, Menü und Profil schließen direkt an „Hast du
          Fragen?" an (Nutzer-Ansage 28.09.2026) statt am rechten Rand zu
          kleben — auf breiten Schirmen lag dazwischen eine halbe Seite. Die
          Gruppe bleibt flex 1 1 0, damit die Mitte weiter mittig steht. */}
      <div className="header-right" style={isMobile ? undefined : { flex: '1 1 0', justifyContent: 'flex-start', paddingLeft: 12, minWidth: 0 }}>
        {/* v26: Grüner „Hast du Fragen?"-Button — Ticketsystem für alle User.
            v26.34: jetzt auch auf der Landing Page im Header sichtbar. */}
        {isMobile && <QuestionButton isMobile={isMobile} onAbout={() => setShowAbout(true)} />}
        {/* v9.29: Refresh-Button im Header. v32.2.1: nur das Symbol (Nutzer-
            Ansage 28.09.2026: „Aktualisieren braucht man nicht ausschreiben,
            da reicht das Symbol"); der Name steht im Tooltip. */}
        {(currentPage === 'admin' || currentPage === 'register' || currentPage === 'my-events' || currentPage === 'participants') && (
          <button
            className="header-icon-btn"
            onClick={() => { window.dispatchEvent(new CustomEvent('dex-refresh-page')); }}
            title={locale === 'de' ? 'Aktualisieren' : 'Refresh'}
            aria-label={locale === 'de' ? 'Aktualisieren' : 'Refresh'}
          >
            <RefreshCw size={18} />
          </button>
        )}
        {/* v32.2.1: Steht die User-Ansicht an, sagt der Header es — sonst
            verschwände der Zustand im Menü, und man wundert sich über
            fehlende Kacheln. Ein Klick schaltet zurück. */}
        {previewAsUser && (currentUserRole === 'Organizer' || originalIsAdmin) && (
          <button
            type="button"
            className="dex-ui-chip is-active"
            onClick={() => setPreviewAsUser(false)}
            title={locale === 'de' ? 'Zurück zur Organizer-Ansicht' : 'Back to the organizer view'}
            style={{ whiteSpace: 'nowrap', flexShrink: 0 }}
          >
            {locale === 'de' ? 'User-Ansicht · zurück' : 'User view · back'}
          </button>
        )}
        {isMobile && langToggle}
        {/* v32.2.1: Burger-Menü statt vier Knöpfen (Nutzer-Ansage 28.09.2026:
            „ein Burger-Menü, in dem Handbuch, Über die App und Demo drin
            sind"). Demo umfasst die Ansicht (Organizer/User, v30.43) und für
            echte Admins „Als bestimmte Person testen" (Impersonate, v24.69).
            Sichtbarkeit von Demo an der ECHTEN Rolle — in der User-Ansicht
            sind isAdmin/isOrganizer abgesenkt, der Eintrag darf dann nicht
            verschwinden. */}
        <div ref={demoRef} style={{ position: 'relative', flexShrink: 0 }}>
          <button
            type="button"
            className="header-icon-btn"
            aria-haspopup="menu"
            aria-expanded={demoOpen}
            aria-label={locale === 'de' ? 'Menü' : 'Menu'}
            title={locale === 'de' ? 'Menü: Handbuch, Über die App, Demo' : 'Menu: manual, about the app, demo'}
            onClick={() => setDemoOpen(o => !o)}
            style={demoOpen || currentPage === 'manual' ? { background: 'var(--dex-gray-200)' } : undefined}
          >
            <Menu size={20} />
          </button>
          {demoOpen && (
            <div role="menu" style={{
              position: 'absolute', top: 'calc(100% + 8px)', right: 0, zIndex: 1200, width: 300,
              background: '#fff', border: '1px solid var(--dex-gray-200, #e1e1e1)', borderRadius: 12,
              boxShadow: '0 12px 32px rgba(0,0,0,0.14)', padding: 8,
            }}>
              {([
                { key: 'manual', icon: <Book size={16} />, title: t('header.manual'), sub: locale === 'de' ? 'Anleitungen zu allen Funktionen.' : 'Guides for every feature.', onClick: () => navigate('manual') },
                { key: 'about', icon: <Info size={16} />, title: locale === 'de' ? 'Über die App' : 'About the app', sub: locale === 'de' ? 'Wofür DEX gedacht ist, Ablauf und Tutorial.' : 'What DEX is for, the flow and the tutorial.', onClick: () => setShowAbout(true) },
              ]).map(it => (
                <button key={it.key} type="button" role="menuitem" className="dex-ui-menuitem"
                  onClick={() => { setDemoOpen(false); it.onClick(); }}>
                  <span style={{ width: 16, flexShrink: 0, color: 'var(--dex-gray-600)', paddingTop: 2 }}>{it.icon}</span>
                  <span style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                    <span style={{ fontSize: '0.86rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>{it.title}</span>
                    <span style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)', lineHeight: 1.4 }}>{it.sub}</span>
                  </span>
                </button>
              ))}
              {(currentUserRole === 'Organizer' || originalIsAdmin) && (<>
                <div style={{ height: 1, background: 'var(--dex-gray-200)', margin: '6px 4px' }} />
                <div style={{ padding: '6px 10px 4px', fontSize: '0.68rem', fontWeight: 700, letterSpacing: '0.06em', textTransform: 'uppercase', color: 'var(--dex-gray-500)' }}>
                  {locale === 'de' ? 'Demo · Ansicht' : 'Demo · view'}
                </div>
                {([
                  { key: 'org' as const, active: !previewAsUser, title: locale === 'de' ? 'Organizer-Ansicht' : 'Organizer view', sub: locale === 'de' ? 'Mit allen Hinweisen und Rechten.' : 'With all notices and rights.' },
                  { key: 'user' as const, active: previewAsUser, title: locale === 'de' ? 'User-Ansicht' : 'User view', sub: locale === 'de' ? 'So sehen Teilnehmer die App — reine Ansicht, anmelden geht darin nicht.' : 'How attendees see the app — view only, no registering.' },
                ]).map(it => (
                  <button key={it.key} type="button" role="menuitemradio" aria-checked={it.active} className={cx('dex-ui-menuitem', it.active && 'is-active')}
                    onClick={() => { setPreviewAsUser(it.key === 'user'); setDemoOpen(false); }}>
                    <span style={{ width: 16, flexShrink: 0, color: 'var(--dex-green-dark, #4a7c1f)', paddingTop: 2 }}>{it.active && <Check size={14} />}</span>
                    <span style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                      <span style={{ fontSize: '0.86rem', fontWeight: it.active ? 700 : 600, color: 'var(--dex-gray-800)' }}>{it.title}</span>
                      <span style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)', lineHeight: 1.4 }}>{it.sub}</span>
                    </span>
                  </button>
                ))}
                {originalIsAdmin && (
                  <button type="button" role="menuitem" className="dex-ui-menuitem"
                    onClick={() => { setDemoOpen(false); setShowImpersonate(true); }}>
                    <span style={{ width: 16, flexShrink: 0, color: 'var(--dex-gray-600)', paddingTop: 2 }}><Users size={16} /></span>
                    <span style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                      <span style={{ fontSize: '0.86rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>{locale === 'de' ? 'Als bestimmte Person testen' : 'Test as a specific person'}</span>
                      <span style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)', lineHeight: 1.4 }}>{locale === 'de' ? 'Die App mit Rolle und Standort einer anderen Person ansehen.' : 'See the app with another person’s role and location.'}</span>
                    </span>
                  </button>
                )}
              </>)}
            </div>
          )}
        </div>
        {/* v24.11: Hinweis-Chip RECHTS neben dem Picker, grün — „Organizer set
            the language ... for this registration form". */}
        {showRegLangHint && !isMobile && (
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
      <LandingInfoModal open={showAbout} locale={locale === 'de' ? 'de' : 'en'} onClose={() => setShowAbout(false)} onStartTutorial={openTutorial} />
    </header>
    </div>
  );
}
