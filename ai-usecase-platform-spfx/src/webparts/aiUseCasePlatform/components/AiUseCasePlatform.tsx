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
import { Settings, Users } from './Icons';
import { LanguageProvider, useLanguage } from '../context/LanguageContext';
import { DialogProvider } from '../context/DialogContext';
import { RoleProvider, useRoles } from '../context/RoleContext';
import { NavigationProvider, useNavigation } from '../context/NavigationContext';
import { UseCaseProvider } from '../context/UseCaseContext';
import { APP_NAME } from '../constants';
import { APP_VERSION } from '../version';

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

function AppContent(): React.ReactElement {
  ensureDexUiStyles();
  const { t, isDe, setLocale } = useLanguage();
  const { currentPage, currentUseCaseId, navigate } = useNavigation();
  const { isAdmin, isKurator, isRolesLoading, rolesReadStatus, currentUserRole, previewAsUser, setPreviewAsUser } = useRoles();

  const seitenName =
    currentPage === 'detail' ? t('Use Case', 'Use case')
      : currentPage === 'verwaltung' ? t('Verwaltung', 'Management')
        : currentPage === 'rollen' ? t('Rollen', 'Roles')
          : APP_NAME;

  return (
    <div className={styles.dexApp}>
      <div className="app-container">
        <header className="app-header">
          <button
            type="button"
            className="header-logo dex-ui-btn-reset"
            onClick={() => navigate('start')}
            aria-label={t('Zur Übersicht', 'Back to overview')}
            style={{ cursor: 'pointer' }}
          >
            <strong>AI Use Case</strong> <span>Platform</span>
          </button>

          <div className="dex-ui-toolbar-spacer" />

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
        </header>

        <main className="main-content">
          <div className="content-wrapper" style={{ maxWidth: 1200, margin: '0 auto', padding: '20px 16px 48px' }}>

            {/* Der Fall, den DEX teuer gelernt hat: Die Person steht in der
                Rollenliste, darf sie aber nicht lesen — dann ist ihre Rolle
                wirkungslos, und das muss dastehen statt still zu wirken. */}
            {!isRolesLoading && rolesReadStatus === 'forbidden' && (
              <div className="dex-ui-callout dex-ui-callout--warn" role="status" style={{ marginBottom: 16 }}>
                <span>
                  {t('Deine Rolle konnte nicht geprüft werden — dir fehlt das Leserecht auf der Rollenliste. Falls du eigentlich Kurator oder Admin bist: Ein Admin muss dir das Leserecht nachsetzen.',
                    'Your role could not be checked — you lack read access to the roles list. If you are meant to be a curator or admin, an admin has to grant it.')}
                </span>
              </div>
            )}

            {previewAsUser && (
              <div className="dex-ui-callout dex-ui-callout--info" role="status" style={{ marginBottom: 16 }}>
                <span>
                  {t('Du siehst die Plattform gerade als normaler Nutzer.', 'You are viewing the platform as a regular user.')}{' '}
                  <button type="button" className="dex-ui-textbtn" onClick={() => setPreviewAsUser(false)}>
                    {t('Vorschau beenden', 'End preview')}
                  </button>
                </span>
              </div>
            )}

            <React.Suspense fallback={<LazyFallback name={seitenName} />}>
              {currentPage === 'start' && <StartPage />}
              {currentPage === 'detail' && <UseCaseDetailPage useCaseId={currentUseCaseId} />}
              {currentPage === 'verwaltung' && <ManagePage editId={currentUseCaseId} />}
              {currentPage === 'rollen' && <RolePage />}
            </React.Suspense>
          </div>
        </main>

        <footer style={{ padding: '10px 16px', textAlign: 'center' }}>
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
