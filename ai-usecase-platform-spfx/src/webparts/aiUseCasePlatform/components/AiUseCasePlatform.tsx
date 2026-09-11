/**
 * Die Anwendung.
 *
 * Aufbau wie `DexEventPlatform.tsx`: Die Provider schachteln von aussen nach
 * innen (Sprache → Dialoge → Rollen → Navigation → Daten), und AppContent
 * rendert INNERHALB davon — nur so kennt es Sprache und Rolle.
 *
 * Die Seiten werden per `React.lazy` nachgeladen. Der Platzhalter dabei ist
 * bewusst kein „…": In DEX standen drei Punkte fuer den groessten Chunk, und
 * auf langsamer Leitung war das von „haengt" nicht zu unterscheiden
 * (v31.9.5). Hier steht ein Ring, der Name des Bereichs und der Hinweis,
 * dass das einmalig passiert.
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './AiUseCasePlatform.module.scss';
import { ensureDexUiStyles, cx } from './dexUi';
import StartPage from './StartPage';
import LandingPage from './LandingPage';
import { Settings, Users } from './Icons';
import { LanguageProvider, useLanguage } from '../context/LanguageContext';
import { DialogProvider } from '../context/DialogContext';
import { RoleProvider, useRoles } from '../context/RoleContext';
import { NavigationProvider, useNavigation } from '../context/NavigationContext';
import { UseCaseProvider } from '../context/UseCaseContext';
import { APP_NAME } from '../constants';
import { APP_VERSION } from '../version';
import { DELOITTE_LOGO_HEADER } from '../data/brandLogos';

const UseCaseDetailPage = React.lazy(() => import('./UseCaseDetailPage'));
const ManagePage = React.lazy(() => import('./ManagePage'));
const RolePage = React.lazy(() => import('./RolePage'));

export interface IAiUseCasePlatformProps {
  context: WebPartContext;
}

/** Platzhalter, waehrend ein Bereich einmalig nachgeladen wird. */
function LazyFallback(props: { name: string }): React.ReactElement {
  const { t } = useLanguage();
  return (
    <div className="dex-ui-empty" role="status" aria-live="polite">
      <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
      <div className="dex-ui-empty-title" style={{ marginTop: 14 }}>{props.name} {t('wird geladen …', 'is loading …')}</div>
      <div className="dex-ui-empty-desc">{t('Das passiert beim ersten Aufruf einmalig.', 'This happens once, on first use.')}</div>
    </div>
  );
}

/**
 * Die Hoehen-Kette der Huelle — v1.2.
 *
 * Kartiert an DEX (`DexEventPlatform.tsx` Z. 776-819). Sie besteht aus vier
 * Gliedern, und **drei davon stehen nicht im Stylesheet**:
 *
 *  1. `html, body { overflow:hidden; height:100vh }` — hier injiziert, nicht
 *     im SCSS. Ohne das scrollt das SharePoint-Fenster mit und der Header
 *     wandert aus dem Bild.
 *  2. `.app-layout` bekommt eine PIXEL-Hoehe (`innerHeight - rect.top`,
 *     mindestens 400).
 *  3. `.main-content` nimmt den Rest ueber `flex: 1` (steht im SCSS).
 *  4. `.landing { height:100% }` loest sich erst gegen diese dann definite
 *     Hoehe auf.
 *
 * Faellt Glied 2 weg, kollabiert Glied 4: `.app-layout` hat nur
 * `min-height: 400px` — die Landing Page waere 400 px hoch, egal wie gross
 * das Fenster ist. Genau so sah die erste Fassung aus.
 *
 * Die zwei Nachzieher (500/1500 ms) sind kein Aberglaube: Das
 * SharePoint-Chrome baut sich nach dem ersten Render noch um, und `rect.top`
 * stimmt erst danach.
 */
function useShellHeight(): React.RefObject<HTMLDivElement> {
  const layoutRef = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    const STYLE_ID = 'aiuc-no-scroll';
    if (!document.getElementById(STYLE_ID)) {
      const st = document.createElement('style');
      st.id = STYLE_ID;
      st.textContent = 'html, body { overflow: hidden !important; height: 100vh !important; }';
      document.head.appendChild(st);
    }
    const setHeight = (): void => {
      const el = layoutRef.current;
      if (!el) return;
      const top = el.getBoundingClientRect().top;
      el.style.height = `${Math.max(window.innerHeight - top, 400)}px`;
    };
    setHeight();
    window.addEventListener('resize', setHeight);
    const t1 = window.setTimeout(setHeight, 500);
    const t2 = window.setTimeout(setHeight, 1500);
    return () => {
      window.removeEventListener('resize', setHeight);
      window.clearTimeout(t1);
      window.clearTimeout(t2);
      const st = document.getElementById(STYLE_ID);
      if (st && st.parentNode) st.parentNode.removeChild(st);
    };
  }, []);
  return layoutRef;
}

function AppContent(): React.ReactElement {
  ensureDexUiStyles();
  const { t, isDe, setLocale } = useLanguage();
  const { currentPage, currentUseCaseId, navigate } = useNavigation();
  const { isAdmin, isKurator, isRolesLoading, rolesReadStatus, currentUserRole, previewAsUser, setPreviewAsUser } = useRoles();
  const layoutRef = useShellHeight();

  const seitenName =
    currentPage === 'detail' ? t('Use Case', 'Use case')
      : currentPage === 'verwaltung' ? t('Verwaltung', 'Management')
        : currentPage === 'rollen' ? t('Rollen', 'Roles')
          : APP_NAME;
  const istLanding = currentPage === 'landing';

  // v1.2: Der Seitenwechsel muss den Scroller zuruecksetzen, der WIRKLICH
  // scrollt. In DEX setzt der Effekt `scrollTop` auf window, body und
  // `.app-layout` — also genau nicht auf `.main-content` und `.landing`, die
  // als einzige scrollen; deshalb macht `RegistrationPage` es sich dort
  // selbst. Hier steht es von Anfang an an der richtigen Stelle.
  React.useEffect(() => {
    const el = layoutRef.current?.querySelector('.main-content');
    if (el) el.scrollTop = 0;
  }, [currentPage, currentUseCaseId, layoutRef]);

  const banner = (
    <>
      {/* Der Fall, den DEX teuer gelernt hat: Die Person steht in der
          Rollenliste, darf sie aber nicht lesen — dann ist ihre Rolle
          wirkungslos, und das muss dastehen statt still zu wirken. */}
      {!isRolesLoading && rolesReadStatus === 'forbidden' && (
        <div className="dex-ui-callout dex-ui-callout--warn" role="status" style={{ margin: '12px 24px 0' }}>
          <span>
            {t('Deine Rolle konnte nicht geprüft werden — dir fehlt das Leserecht auf der Rollenliste. Falls du eigentlich Kurator oder Admin bist: Ein Admin muss dir das Leserecht nachsetzen.',
              'Your role could not be checked — you lack read access to the roles list. If you are meant to be a curator or admin, an admin has to grant it.')}
          </span>
        </div>
      )}
      {previewAsUser && (
        <div className="dex-ui-callout dex-ui-callout--info" role="status" style={{ margin: '12px 24px 0' }}>
          <span>
            {t('Du siehst die Plattform gerade als normaler Nutzer.', 'You are viewing the platform as a regular user.')}{' '}
            <button type="button" className="dex-ui-textbtn" onClick={() => setPreviewAsUser(false)}>
              {t('Vorschau beenden', 'End preview')}
            </button>
          </span>
        </div>
      )}
    </>
  );

  const seite = (
    <React.Suspense fallback={<LazyFallback name={seitenName} />}>
      {currentPage === 'landing' && <LandingPage />}
      {currentPage === 'start' && <StartPage />}
      {currentPage === 'detail' && <UseCaseDetailPage useCaseId={currentUseCaseId} />}
      {currentPage === 'verwaltung' && <ManagePage editId={currentUseCaseId} />}
      {currentPage === 'rollen' && <RolePage />}
    </React.Suspense>
  );

  return (
    <div className={styles.dexApp} lang={isDe ? 'de-DE' : 'en-GB'}>
      <div className="app-layout" ref={layoutRef}>
        <header className="header">
          {/* Links. Auf der Landing Page das Logo, sonst der Zurueck-Knopf mit
              dem Seitennamen — genau die Verzweigung, die DEX an `isLanding`
              haengt. Zwei Dinge an derselben Stelle, nie beide. */}
          <div className="header-left">
            {istLanding ? (
              <button
                type="button"
                className="header-logo dex-ui-btn-reset"
                onClick={() => navigate('landing')}
                aria-label={t('Zum Startbildschirm', 'Back to the start screen')}
                style={{ cursor: 'pointer', display: 'flex', alignItems: 'center' }}
              >
                <img src={DELOITTE_LOGO_HEADER} alt="Deloitte" style={{ height: 30, width: 'auto', display: 'block' }} />
              </button>
            ) : (
              <>
                <button
                  type="button"
                  className="back-btn"
                  onClick={() => navigate(currentPage === 'start' ? 'landing' : 'start')}
                  aria-label={t('Zurück', 'Back')}
                  // Von rund auf Pille: Ein Pfeil ohne Wort ist auf dem Handy
                  // nicht als „zurueck" lesbar. Der Inline-Stil ueberschreibt
                  // `border-radius:50%` und die feste Breite aus dem SCSS.
                  style={{ width: 'auto', borderRadius: 999, padding: '0 16px 0 10px', gap: 6, fontSize: '0.85rem' }}
                >
                  <span aria-hidden="true">‹</span>
                  {t('Zurück', 'Back')}
                </button>
                <span
                  className="header-title"
                  // Ohne Kuerzung bricht der Titel bei 1280 px auf drei Zeilen
                  // und ragt aus dem 64-px-Kasten — in DEX gemessen
                  // (scrollHeight 68 bei clientHeight 63).
                  style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: 'min(42vw, 420px)' }}
                >{seitenName}</span>
              </>
            )}
          </div>

          <div className="header-right">
            {/* Sprache. Eine Quelle fuer die ganze App — nicht die
                Browsersprache, die in DEX zweimal auseinandergelaufen ist. */}
            <div className="dex-ui-tabs" role="group" aria-label={t('Sprache', 'Language')}>
              <button type="button" className={cx('dex-ui-tab', isDe && 'is-active')} onClick={() => setLocale('de')}>DE</button>
              <button type="button" className={cx('dex-ui-tab', !isDe && 'is-active')} onClick={() => setLocale('en')}>EN</button>
            </div>

            {isKurator && (
              <button type="button" className="header-icon-btn" onClick={() => navigate('verwaltung')} title={t('Use Cases pflegen', 'Manage use cases')} aria-label={t('Use Cases pflegen', 'Manage use cases')}>
                <Settings size={18} />
              </button>
            )}
            {isAdmin && (
              <button type="button" className="header-icon-btn" onClick={() => navigate('rollen')} title={t('Rollen verwalten', 'Manage roles')} aria-label={t('Rollen verwalten', 'Manage roles')}>
                <Users size={18} />
              </button>
            )}
          </div>
        </header>

        {banner}

        {/* `display:flex` und `flex-direction:column` stehen INLINE, nicht im
            SCSS — und sie sind nicht kosmetisch: `.page-container` hat
            `max-width:100%; margin:0 auto`. In einem Block-Container ist das
            volle Breite, als Flex-Item mit auto-Cross-Margins dagegen
            shrink-to-fit UND zentriert. Wer das hier wegnimmt, verliert die
            Zentrierung jeder Seite, ohne eine Regel geaendert zu haben.
            (In DEX genauso: `DexEventPlatform.tsx` Z. 1176.) */}
        <main className="main-content" style={{ position: 'relative', display: 'flex', flexDirection: 'column' }}>
          {/* Die Landing Page rendert OHNE `.page-container`. Sie bringt ihre
              Abstaende selbst mit (`.landing__hero { padding }`) und braucht
              `height:100%`, um den grauen Grund bis zum unteren Rand zu
              ziehen — ein gepolsterter Wrapper darum macht daraus eine Karte
              mit weissem Rand, und genau so sah es vorher aus. */}
          {istLanding ? seite : <div className="page-container">{seite}</div>}
        </main>

        {/* Dritter Flex-Sohn, ausserhalb des Scrollers: `flex-shrink: 0`,
            sonst quetscht ihn `.main-content` bei langen Seiten weg. */}
        <footer style={{ padding: '8px 16px', textAlign: 'center', flexShrink: 0, borderTop: '1px solid var(--dex-gray-200)', background: 'var(--dex-white)' }}>
          <span className="dex-ui-muted" style={{ fontSize: '0.72rem' }}>
            {APP_NAME} v{APP_VERSION}
            {!isRolesLoading && ` · ${t('Deine Rolle', 'Your role')}: ${currentUserRole}`}
            {isKurator && !previewAsUser && (
              <>
                {' · '}
                <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={() => setPreviewAsUser(true)}>
                  {t('Als Nutzer ansehen', 'View as user')}
                </button>
              </>
            )}
          </span>
        </footer>
      </div>
    </div>
  );
}

export default function AiUseCasePlatform(props: IAiUseCasePlatformProps): React.ReactElement {
  // v1.1: Der SPFx-Context als Fenster-Merker — dasselbe Muster wie
  // `__dexSpfxContext` in DEX. Komponenten, die tief im Baum sitzen und den
  // Context nur einmal brauchen (die Begruessung der Landing Page), holen ihn
  // sich hier, statt ihn durch fuenf Ebenen durchzureichen.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).__aiucSpfxContext = props.context;
  return (
    <LanguageProvider>
      <DialogProvider>
        <RoleProvider context={props.context}>
          <NavigationProvider>
            <UseCaseProvider context={props.context}>
              <AppContent />
            </UseCaseProvider>
          </NavigationProvider>
        </RoleProvider>
      </DialogProvider>
    </LanguageProvider>
  );
}
