/**
 * Die Anwendung.
 *
 * Aufbau wie `DexEventPlatform.tsx`: Die Provider schachteln von aussen nach
 * innen (Sprache → Dialoge → Nutzer → Rollen → Navigation → Daten → Suche →
 * Hilfe), und AppContent rendert INNERHALB davon — nur so kennt es Sprache,
 * Person und Rolle.
 *
 * Die Seiten werden per `React.lazy` nachgeladen. Der Platzhalter dabei ist
 * bewusst kein „…": In DEX standen drei Punkte fuer den groessten Chunk, und
 * auf langsamer Leitung war das von „haengt" nicht zu unterscheiden
 * (v31.9.5). Hier steht ein Ring, der Name des Bereichs und der Hinweis,
 * dass das einmalig passiert.
 *
 * v1.3: Die Kopfzeile ist eine eigene Komponente (`Header`, Vorbild DEX), und
 * zwischen Landing Page und Kachelwand liegt die Start-Übersicht
 * (`StartPage`). Seiten-IDs: landing → start → usecases → detail; dazu
 * studio, rollen, protokoll.
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './AiUseCasePlatform.module.scss';
import { ensureDexUiStyles } from './dexUi';
import Header from './Header';
import StartPage from './StartPage';
import UseCasesPage from './UseCasesPage';
import LandingPage from './LandingPage';
import { LanguageProvider, useLanguage } from '../context/LanguageContext';
import { DialogProvider } from '../context/DialogContext';
import { UserProvider } from '../context/UserContext';
import { RoleProvider, useRoles } from '../context/RoleContext';
import { NavigationProvider, useNavigation } from '../context/NavigationContext';
import { UseCaseProvider } from '../context/UseCaseContext';
import { SucheProvider } from '../context/SucheContext';
import { HilfeProvider } from '../context/HilfeContext';
import { APP_NAME } from '../constants';
import { APP_VERSION } from '../version';
import { rolleAnzeige } from '../utils/rollen';
import ErrorBoundary from './ErrorBoundary';

const UseCaseDetailPage = React.lazy(() => import('./UseCaseDetailPage'));
const ManagePage = React.lazy(() => import('./ManagePage'));
const RolePage = React.lazy(() => import('./RolePage'));
const LogPage = React.lazy(() => import('./LogPage'));

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
  const { t, isDe } = useLanguage();
  const { currentPage, currentUseCaseId } = useNavigation();
  const { isRolesLoading, currentUserRole, rolesReadStatus } = useRoles();
  const layoutRef = useShellHeight();

  const seitenName =
    currentPage === 'detail' ? t('Use Case', 'Use case')
      : currentPage === 'usecases' ? 'Use Cases'
        : currentPage === 'studio' ? 'Use Case Studio'
          : currentPage === 'rollen' ? t('Rollenverwaltung', 'Role management')
            : currentPage === 'protokoll' ? t('Protokoll', 'Log')
              : APP_NAME;
  const istLanding = currentPage === 'landing';
  // Feste Höchstbreite für die schmalen Seiten. `.page-container` ist ein Flex-Kind mit
  // `margin: 0 auto` und damit shrink-to-fit: Ohne Breite war die Pflegeseite 528 px schmal, das
  // Protokoll eine andere Breite und die Kachelwand voll breit (Sichtprüfung, 29.09.2026).
  // Die Kachelwand und die Start-Übersicht bringen ihre Breite selbst mit.
  const seitenBreite = currentPage === 'rollen' ? 1100
    : (currentPage === 'detail' || currentPage === 'studio' || currentPage === 'protokoll') ? 900 : 0;

  // v1.2: Der Seitenwechsel muss den Scroller zuruecksetzen, der WIRKLICH
  // scrollt. In DEX setzt der Effekt `scrollTop` auf window, body und
  // `.app-layout` — also genau nicht auf `.main-content` und `.landing`, die
  // als einzige scrollen; deshalb macht `RegistrationPage` es sich dort
  // selbst. Hier steht es von Anfang an an der richtigen Stelle.
  React.useEffect(() => {
    const el = layoutRef.current?.querySelector('.main-content');
    if (el) el.scrollTop = 0;
  }, [currentPage, currentUseCaseId, layoutRef]);

  const seite = (
    <ErrorBoundary isDe={isDe} resetKey={`${currentPage}:${currentUseCaseId || 0}`}>
      <React.Suspense fallback={<LazyFallback name={seitenName} />}>
        {currentPage === 'landing' && <LandingPage />}
        {currentPage === 'start' && <StartPage />}
        {currentPage === 'usecases' && <UseCasesPage />}
        {/* `key`: Beim Wechsel von einem Use Case zum nächsten (Suche in der Kopfzeile) bleibt die
            Seite sonst dieselbe Komponente — ein offenes iframe und „Link kopiert" liefen mit
            hinüber (Review 29.09.2026). */}
        {currentPage === 'detail' && <UseCaseDetailPage key={currentUseCaseId} useCaseId={currentUseCaseId} />}
        {currentPage === 'studio' && <ManagePage editId={currentUseCaseId} />}
        {currentPage === 'rollen' && <RolePage />}
        {currentPage === 'protokoll' && <LogPage useCaseId={currentUseCaseId} />}
      </React.Suspense>
    </ErrorBoundary>
  );

  return (
    <div className={styles.dexApp} lang={isDe ? 'de-DE' : 'en-GB'}>
      <div className="app-layout" ref={layoutRef}>
        <Header />

        {/* Hier stand bis v1.2 eine Warnleiste „Deine Rolle konnte nicht geprüft
            werden", sobald die Rollenliste mit 403 antwortete. Das ist aber
            der NORMALFALL jedes gewöhnlichen Nutzers: Die Liste hat eigene
            Rechte (Owners und vergebene Personen), wer nicht darin steht,
            darf sie nicht lesen. Die Leiste stand damit auf jeder Seite bei
            jedem Nutzer und versprach „ein Admin muss dir das Leserecht
            nachsetzen" — und ein Admin, der das ernst nimmt, öffnet die
            Adressliste für alle (Review-Fund 7, 29.09.2026).
            Wie in DEX steht der Hinweis jetzt NUR an der Kachel „Use Case
            Studio" (StartPage), wo die Frage „bin ich Organizer?" gestellt
            wird. */}

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
          {istLanding ? seite : (
            <div className="page-container" style={seitenBreite ? { width: '100%', maxWidth: seitenBreite } : undefined}>{seite}</div>
          )}
        </main>

        {/* Dritter Flex-Sohn, ausserhalb des Scrollers: `flex-shrink: 0`,
            sonst quetscht ihn `.main-content` bei langen Seiten weg.
            Die Vorschau „als Nutzer ansehen" liegt seit v1.3 im Menue der
            Kopfzeile (wie in DEX), nicht mehr hier — ein Weg, nicht zwei. */}
        <footer style={{ padding: '8px 16px', textAlign: 'center', flexShrink: 0, borderTop: '1px solid var(--dex-gray-200)', background: 'var(--dex-white)' }}>
          <span className="dex-ui-muted" style={{ fontSize: '0.72rem' }}>
            {APP_NAME} v{APP_VERSION}
            {!isRolesLoading && ` · ${t('Deine Rolle', 'Your role')}: ${rolleAnzeige(currentUserRole, rolesReadStatus, t)}`}
          </span>
        </footer>
      </div>
    </div>
  );
}

export default function AiUseCasePlatform(props: IAiUseCasePlatformProps): React.ReactElement {
  // Bis v1.3 stand hier der SPFx-Context als Fenster-Merker (`window.__aiucSpfxContext`,
  // wie `__dexSpfxContext` in DEX). Niemand las ihn — aber jedes Skript auf der Seite
  // hätte damit `spHttpClient` und den Token-Anbieter in der Hand gehabt (Sicherheits-Review
  // 29.09.2026). Der UserContext liefert den Namen, alles andere geht über die Provider.
  return (
    <LanguageProvider>
      <DialogProvider>
        <UserProvider context={props.context}>
          <RoleProvider context={props.context}>
            <NavigationProvider>
              <UseCaseProvider context={props.context}>
                <SucheProvider>
                  <HilfeProvider>
                    <AppContent />
                  </HilfeProvider>
                </SucheProvider>
              </UseCaseProvider>
            </NavigationProvider>
          </RoleProvider>
        </UserProvider>
      </DialogProvider>
    </LanguageProvider>
  );
}
