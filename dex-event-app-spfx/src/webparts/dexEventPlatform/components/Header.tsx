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
import { useEvents } from '../context/EventContext';
import { useLanguage } from '../context/LanguageContext';
import { ChevronLeft, Settings, Book, QrCode, RefreshCw } from './Icons';
import { Icon } from '@fluentui/react/lib/Icon';
import ImpersonateModal from './ImpersonateModal';

export default function Header(): React.ReactElement {
  const { currentPage, navigate, selectedEventId } = useNavigation();
  const { currentUser, photoUrl } = useCurrentUser();
  const { currentUserRole, isAdmin, isOrganizer, originalIsAdmin } = useRoles();
  const [showImpersonate, setShowImpersonate] = React.useState(false);
  const { events } = useEvents();
  // Check-In-Button (Admin / Organizer / QR-Scanner): schneller Einstieg in den
  // QR-Scanner ohne vorher ein konkretes Event auszuwählen. CheckInPage liest
  // das Event aus dem gescannten QR-Code selbst (`DEX|<eventNumber>|<email>`).
  // v6.26: Zugriff auch für User, die per E-Mail in event.qrScannerEmails
  // mindestens eines Events eingetragen sind — ohne globale Organizer-Rolle.
  // v9.18: Co-Organizer pro Event ebenfalls Check-In-fähig.
  // v11.46: Seit v9.20 hat der Wizard nur einen Organizer-Picker, der alle
  // (Haupt-Organizer wie Co-Organizer) in event.organizerEmails schreibt.
  // Prüfung deshalb analog zu AdminPage.isOrganizerFor: organizerEmails ODER
  // coOrganizerEmails (Backward-Compat für alte Events).
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isQRScannerOfAny = !!currentEmailLc && (events || []).some(
    e => (e.qrScannerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc),
  );
  const isOrganizerOfAnyEvent = !!currentEmailLc && (events || []).some(e => {
    const inOrg = (e.organizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
    if (inOrg) return true;
    return (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
  });
  const canCheckIn = isAdmin || isOrganizer || isQRScannerOfAny || isOrganizerOfAnyEvent;
  const { t, locale, setLocale } = useLanguage();
  const [showPopup, setShowPopup] = React.useState(false);
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
  const regLangHintText = forcedRegLang === 'de' ? 'Anmeldung auf Deutsch' : 'Registration in English';

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
      case 'admin': return t('header.admin');
      case 'role-matrix': return t('header.rolematrix');
      case 'participants': return t('header.participants');
      case 'flowcharts': return t('header.flowcharts');
      case 'check-in': return t('header.checkin');
      case 'manual': return t('header.manual');
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

  return (
    <header className="header">
      <div className="header-left">
        {isLanding ? (
          <div className="header-logo">
            Deloitte<span>.</span>
          </div>
        ) : (
          <>
            <button className="back-btn" onClick={() => navigate(isStart ? 'landing' : 'start')} aria-label="Back">
              <ChevronLeft size={20} />
            </button>
            <span className="header-title" style={{ border: 'none', paddingLeft: 0, fontWeight: 500 }}>
              {getTitle()}
            </span>
          </>
        )}
      </div>
      <div className="header-right">
        {canCheckIn && isLanding && isMobile && (
          <button
            type="button"
            onClick={() => navigate('check-in')}
            aria-label={t('header.checkin')}
            style={{
              background: 'var(--dex-green)', color: '#fff',
              border: 'none', borderRadius: 12,
              padding: '6px 10px',
              fontSize: '0.78rem', fontWeight: 600, lineHeight: 1.2,
              cursor: 'pointer', marginRight: 6,
              boxShadow: '0 2px 8px rgba(134,188,37,0.28)',
              fontFamily: 'inherit', whiteSpace: 'nowrap',
              position: 'relative',
            }}
          >
            {t('header.checkin.bubble') || 'Jetzt einchecken'}
            <span style={{
              position: 'absolute', top: '50%', right: -6, transform: 'translateY(-50%)',
              width: 0, height: 0,
              borderTop: '6px solid transparent',
              borderBottom: '6px solid transparent',
              borderLeft: '6px solid var(--dex-green)',
            }} />
          </button>
        )}
        {/* v13.10: Check-In-Button auch auf Desktop sichtbar, damit reine
            QR-Scanner (ohne Admin-Center-Zugang) den manuellen
            Namens-Such-Check-In auf der Check-In-Seite erreichen. Der
            Kamera-Scanner selbst macht am Desktop wenig Sinn (keine
            Rückkamera), die Teilnehmer-Einchecken-Liste direkt darunter
            funktioniert dort aber genauso wie auf dem Handy. */}
        {canCheckIn && (
          <button
            className="header-icon-btn"
            onClick={() => navigate('check-in')}
            title={t('header.checkin')}
            style={{
              ...(currentPage === 'check-in' ? { background: 'var(--dex-gray-200)' } : {}),
              ...(isMobile ? {} : { width: 'auto', padding: '0 12px', display: 'inline-flex', alignItems: 'center', gap: 8 }),
            }}
            aria-label={t('header.checkin')}
          >
            <QrCode size={20} />
            {!isMobile && (
              <span style={{ fontSize: '0.85rem', fontWeight: 500, lineHeight: 1 }}>
                {t('header.checkin')}
              </span>
            )}
          </button>
        )}
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
        {/* v18.40: Hinweis-Chip „Registration in English/Deutsch" jetzt RECHTS,
            direkt links neben dem Sprach-Picker (statt neben dem Titel). */}
        {showRegLangHint && (
          <span
            title={forcedRegLang === 'de'
              ? 'Dieses Anmeldeformular wird auf Deutsch angezeigt.'
              : 'This registration form is shown in English.'}
            style={{
              display: 'inline-flex', alignItems: 'center', gap: 5, alignSelf: 'center',
              background: 'var(--dex-gray-100, #f0f0ee)', color: 'var(--dex-gray-600, #555)',
              border: '1px solid var(--dex-gray-200, #e0e0e0)', borderRadius: 999,
              padding: '3px 10px', fontSize: '0.72rem', fontWeight: 600, whiteSpace: 'nowrap',
            }}
          >
            <Icon iconName="Globe" style={{ fontSize: 12 }} />
            {regLangHintText}
          </span>
        )}
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
            onClick={() => setLocale('de')}
            title="Deutsch"
            style={{
              padding: '3px 10px', borderRadius: 999,
              border: 'none', cursor: 'pointer',
              fontSize: '0.72rem', fontWeight: 700, fontFamily: 'inherit',
              background: locale === 'de' ? '#fff' : 'transparent',
              color: locale === 'de' ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-500)',
              boxShadow: locale === 'de' ? '0 1px 3px rgba(0,0,0,0.08)' : 'none',
              transition: 'all 0.15s ease',
            }}
          >DE</button>
          <button
            type="button"
            onClick={() => setLocale('en')}
            title="English"
            style={{
              padding: '3px 10px', borderRadius: 999,
              border: 'none', cursor: 'pointer',
              fontSize: '0.72rem', fontWeight: 700, fontFamily: 'inherit',
              background: locale === 'en' ? '#fff' : 'transparent',
              color: locale === 'en' ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-500)',
              boxShadow: locale === 'en' ? '0 1px 3px rgba(0,0,0,0.08)' : 'none',
              transition: 'all 0.15s ease',
            }}
          >EN</button>
        </div>
        {!isLanding && (
          <button
            className="header-icon-btn"
            onClick={() => navigate('settings')}
            title="Settings"
            style={currentPage === 'settings' ? { background: 'var(--dex-gray-200)' } : {}}
          >
            <Settings size={20} />
          </button>
        )}
        {/* User-Avatar mit Initialen + Popup */}
        <div style={{ position: 'relative' }}>
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
                {isAdmin && (
                  <button
                    className="btn btn-primary btn-block"
                    style={{ fontSize: '0.85rem' }}
                    onClick={() => { setShowPopup(false); navigate('settings'); }}
                  >
                    <Settings size={14} /> {t('settings.rolemanagement')}
                  </button>
                )}
                {/* v12.7: Demo-Modus-Toggle nur für echte Admins
                    (originalIsAdmin), damit der Button auch sichtbar bleibt,
                    wenn die App durch laufende Impersonation kurzzeitig
                    'isAdmin=false' meldet. */}
                {originalIsAdmin && (
                  <button
                    className="btn btn-secondary btn-block"
                    style={{ fontSize: '0.85rem' }}
                    onClick={() => { setShowPopup(false); setShowImpersonate(true); }}
                  >
                    {t('header.demoUser') || 'Demo: als User testen'}
                  </button>
                )}
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
    </header>
  );
}
