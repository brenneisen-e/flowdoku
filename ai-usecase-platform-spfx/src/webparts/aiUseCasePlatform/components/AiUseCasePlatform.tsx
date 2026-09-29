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
import { rolleLabel } from '../utils/rollen';

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
  const { isRolesLoading, rolesReadStatus, currentUserRole } = useRoles();
  const layoutRef = useShellHeight();

  const seitenName =
    currentPage === 'detail' ? t('Use Case', 'Use case')
      : currentPage === 'usecases' ? 'Use Cases'
        : currentPage === 'studio' ? 'Use Case Studio'
          : currentPage === 'rollen' ? t('Rollenverwaltung', 'Role management')
            : currentPage === 'protokoll' ? t('Protokoll', 'Log')
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

  const seite = (
    <React.Suspense fallback={<LazyFallback name={seitenName} />}>
      {currentPage === 'landing' && <LandingPage />}
      {currentPage === 'start' && <StartPage />}
      {currentPage === 'usecases' && <UseCasesPage />}
      {currentPage === 'detail' && <UseCaseDetailPage useCaseId={currentUseCaseId} />}
      {currentPage === 'studio' && <ManagePage editId={currentUseCaseId} />}
      {currentPage === 'rollen' && <RolePage />}
      {currentPage === 'protokoll' && <LogPage useCaseId={currentUseCaseId} />}
    </React.Suspense>
  );

  return (
    <div className={styles.dexApp} lang={isDe ? 'de-DE' : 'en-GB'}>
      <div className="app-layout" ref={layoutRef}>
        <Header />

        {/* Der Fall, den DEX teuer gelernt hat: Die Person steht in der
            Rollenliste, darf sie aber nicht lesen — dann ist ihre Rolle
            wirkungslos, und das muss dastehen statt still zu wirken. */}
        {!isRolesLoading && rolesReadStatus === 'forbidden' && (
          <div className="dex-ui-callout dex-ui-callout--warn" role="status" style={{ margin: '12px 24px 0' }}>
            <span>
              {t('Deine Rolle konnte nicht geprüft werden — dir fehlt das Leserecht auf der Rollenliste. Falls du eigentlich Use Case Organizer oder Admin bist: Ein Admin muss dir das Leserecht nachsetzen.',
                'Your role could not be checked — you lack read access to the roles list. If you are meant to be a Use Case Organizer or admin, an admin has to grant it.')}
            </span>
          </div>
        )}

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
            sonst quetscht ihn `.main-content` bei langen Seiten weg.
            Die Vorschau „als Nutzer ansehen" liegt seit v1.3 im Menue der
            Kopfzeile (wie in DEX), nicht mehr hier — ein Weg, nicht zwei. */}
        <footer style={{ padding: '8px 16px', textAlign: 'center', flexShrink: 0, borderTop: '1px solid var(--dex-gray-200)', background: 'var(--dex-white)' }}>
          <span className="dex-ui-muted" style={{ fontSize: '0.72rem' }}>
            {APP_NAME} v{APP_VERSION}
            {!isRolesLoading && ` · ${t('Deine Rolle', 'Your role')}: ${rolleLabel(currentUserRole)}`}
          </span>
        </footer>
      </div>
    </div>
  );
}

export default function AiUseCasePlatform(props: IAiUseCasePlatformProps): React.ReactElement {
  // v1.1: Der SPFx-Context als Fenster-Merker — dasselbe Muster wie
  // `__dexSpfxContext` in DEX. Komponenten, die tief im Baum sitzen und den
  // Context nur einmal brauchen, holen ihn sich hier, statt ihn durch fuenf
  // Ebenen durchzureichen. (Seit v1.3 liest die Begruessung den Namen ueber
  // `UserContext`; der Merker bleibt fuer kuenftige Stellen.)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).__aiucSpfxContext = props.context;
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
