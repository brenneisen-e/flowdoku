/**
 * Die Kopfzeile — Aufbau wie `Header.tsx` in DEX.
 *
 * Auf der Landing Page das Deloitte-Logo, auf jeder anderen Seite links der
 * Seitenname. In der Mitte eine Gruppe: „Zurück", die Suche, die Sprache und
 * „Hast du Fragen?". Rechts das Burger-Menü und das Profilbild mit Popup.
 *
 * Bewusst NICHT aus DEX übernommen:
 *  - das JS-Anheften der Kopfzeile (`position: fixed` beim Scrollen): Dort
 *    scrollt die SharePoint-Seite; hier hat die Hülle eine feste Höhe und
 *    scrollt nur `.main-content` — die Kopfzeile bleibt von selbst stehen
 *    (`useShellHeight` in `AiUseCasePlatform`).
 *  - Aktualisieren-Knopf, Handbuch, Impersonate, Tutorial: gibt es hier nicht.
 *
 * Die Vorschau „als Nutzer ansehen" steht im Menü (Admins und Organizer), und
 * solange sie läuft, sagt die Kopfzeile es mit einem Chip — sonst verschwände
 * der Zustand im Menü, und man wunderte sich über fehlende Kacheln.
 */

import * as React from 'react';
import { ChevronLeft, Check, Info, Menu, MessageSquare } from './Icons';
import { cx } from './dexUi';
import UseCaseSearch from './UseCaseSearch';
import { useNavigation, Page } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { useCurrentUser } from '../context/UserContext';
import { useHilfe } from '../context/HilfeContext';
import { useIsMobile } from '../utils/useIsMobile';
import { rolleLabel } from '../utils/rollen';
import { DELOITTE_LOGO_HEADER } from '../data/brandLogos';

/** Klick außerhalb schließt — für Menü und Profil-Popup. */
function useKlickDaneben(ref: React.RefObject<HTMLElement>, offen: boolean, zu: () => void): void {
  React.useEffect(() => {
    if (!offen) return undefined;
    const onDown = (e: MouseEvent): void => { if (ref.current && !ref.current.contains(e.target as Node)) zu(); };
    const onKey = (e: KeyboardEvent): void => { if (e.key === 'Escape') zu(); };
    window.addEventListener('mousedown', onDown);
    window.addEventListener('keydown', onKey);
    return () => { window.removeEventListener('mousedown', onDown); window.removeEventListener('keydown', onKey); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [offen]);
}

const ROLLEN_FARBEN: Record<string, { bg: string; color: string }> = {
  Admin: { bg: '#e8f5e9', color: '#2e7d32' },
  Organizer: { bg: '#e3f2fd', color: '#1565c0' },
  User: { bg: '#f5f5f5', color: '#666' },
};

export default function Header(): React.ReactElement {
  const { currentPage, navigate, goBack, canGoBack } = useNavigation();
  const { t, isDe, setLocale } = useLanguage();
  const { currentUser, photoUrl } = useCurrentUser();
  const { currentUserRole, originalIsOrganizer, previewAsUser, setPreviewAsUser } = useRoles();
  const { openKontakt, openAbout } = useHilfe();
  const isMobile = useIsMobile();

  const [menueOffen, setMenueOffen] = React.useState(false);
  const [profilOffen, setProfilOffen] = React.useState(false);
  const menueRef = React.useRef<HTMLDivElement | null>(null);
  const profilRef = React.useRef<HTMLDivElement | null>(null);
  useKlickDaneben(menueRef, menueOffen, () => setMenueOffen(false));
  useKlickDaneben(profilRef, profilOffen, () => setProfilOffen(false));
  // Das Foto kann fehlen (404) — dann die Initialen. Zustand statt
  // `style.display = 'none'`: Ein versteckter <img> lässt einen leeren Kreis stehen.
  const [fotoKaputt, setFotoKaputt] = React.useState(false);

  const istLanding = currentPage === 'landing';

  const titel = ((): string => {
    switch (currentPage) {
      case 'usecases': return 'Use Cases';
      case 'detail': return t('Use Case', 'Use case');
      case 'studio': return 'Use Case Studio';
      case 'rollen': return t('Rollenverwaltung', 'Role management');
      case 'protokoll': return t('Protokoll', 'Log');
      // Auf der Start-Übersicht steht links nichts: Der Name steht schon im
      // Deloitte-Logo der Landing Page, und die Kacheln tragen ihre eigenen Titel.
      default: return '';
    }
  })();

  // Zurück heißt zurück — über den Verlauf. Ohne Verlauf (Deep-Link auf einen
  // Use Case, Neuladen) gäbe `goBack` nichts her und der Knopf wäre tot; dann
  // geht es eine Ebene höher.
  const zurueck = (): void => {
    if (canGoBack) { goBack(); return; }
    const hoeher: Page = currentPage === 'detail' ? 'usecases' : currentPage === 'start' ? 'landing' : 'start';
    navigate(hoeher);
  };

  const initialen = `${currentUser.firstName ? currentUser.firstName[0] : ''}${currentUser.surname ? currentUser.surname[0] : ''}` || (currentUser.email ? currentUser.email[0].toUpperCase() : '');
  const rc = ROLLEN_FARBEN[currentUserRole] || ROLLEN_FARBEN.User;

  const sprachwahl = (
    <div
      role="group"
      aria-label={t('Sprache wechseln', 'Switch language')}
      style={{ display: 'inline-flex', alignItems: 'center', background: 'var(--dex-gray-100, #f3f4f6)', borderRadius: 999, padding: 2, gap: 2, height: 30, alignSelf: 'center' }}
    >
      {(['de', 'en'] as const).map(l => {
        const aktiv = (l === 'de') === isDe;
        return (
          <button
            key={l}
            type="button"
            onClick={() => setLocale(l)}
            title={l === 'de' ? 'Deutsch' : 'English'}
            aria-pressed={aktiv}
            style={{
              padding: '3px 10px', borderRadius: 999, border: 'none', cursor: 'pointer',
              fontSize: '0.72rem', fontWeight: 700, fontFamily: 'inherit',
              background: aktiv ? '#fff' : 'transparent',
              color: aktiv ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-500)',
              boxShadow: aktiv ? '0 1px 3px rgba(0,0,0,0.08)' : 'none',
              transition: 'all 0.15s ease',
            }}
          >{l.toUpperCase()}</button>
        );
      })}
    </div>
  );

  const zurueckKnopf = (
    <button
      type="button"
      className="back-btn"
      onClick={zurueck}
      aria-label={t('Zurück', 'Back')}
      style={{ width: 'auto', borderRadius: 999, padding: '0 16px 0 10px', fontSize: '0.85rem', fontWeight: 600, flexShrink: 0, whiteSpace: 'nowrap' }}
    >
      <ChevronLeft size={20} /> {t('Zurück', 'Back')}
    </button>
  );

  // Der grüne Knopf. Auf dem Handy nur das Symbol — dort ist kein Platz für den Text.
  const fragenKnopf = (
    <button
      type="button"
      onClick={() => openKontakt('frage')}
      title={t('Hast du Fragen?', 'Any questions?')}
      aria-label={t('Hast du Fragen?', 'Any questions?')}
      style={{
        display: 'inline-flex', alignItems: 'center', gap: 8, flexShrink: 0, whiteSpace: 'nowrap',
        background: 'var(--dex-green, #86bc25)', color: '#fff', border: 'none', borderRadius: 10,
        padding: isMobile ? '7px 9px' : '7px 14px', fontWeight: 700, fontSize: '0.85rem', fontFamily: 'inherit', cursor: 'pointer',
      }}
    >
      <MessageSquare size={18} strokeWidth={2} />
      {!isMobile && t('Hast du Fragen?', 'Any questions?')}
    </button>
  );

  const menuePunkte: Array<{ key: string; icon: React.ReactNode; titel: string; sub: string; aktiv?: boolean; radio?: boolean; onClick: () => void }> = [
    { key: 'about', icon: <Info size={16} />, titel: t('Über die App', 'About the app'), sub: t('Wofür die Plattform da ist, wie sie läuft, wer was darf.', 'What the platform is for, how it works, who may do what.'), onClick: openAbout },
  ];
  const ansichtPunkte: Array<{ key: string; titel: string; sub: string; aktiv: boolean; onClick: () => void }> = [
    { key: 'org', titel: t('Organizer-Ansicht', 'Organizer view'), sub: t('Mit allen Rechten und Hinweisen.', 'With all rights and notices.'), aktiv: !previewAsUser, onClick: () => setPreviewAsUser(false) },
    { key: 'user', titel: t('User-Ansicht', 'User view'), sub: t('So sehen normale Nutzer die Plattform — reine Ansicht.', 'How regular users see the platform — view only.'), aktiv: previewAsUser, onClick: () => setPreviewAsUser(true) },
  ];

  return (
    <header className="header">
      <div className="header-left" style={isMobile ? undefined : { flex: '1 1 0', minWidth: 0 }}>
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
            {isMobile && zurueckKnopf}
            {titel && (
              <span className="header-title" title={titel} style={{ border: 'none', paddingLeft: 0, fontWeight: 500, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', minWidth: 0, flex: '0 1 auto' }}>
                {titel}
              </span>
            )}
          </>
        )}
      </div>

      {/* Mitte: Auf dem Rechner eine Gruppe; links und rechts teilen sich den
          Rest je zur Hälfte (`flex: 1 1 0`), damit die Gruppe wirklich mittig
          steht. Auf dem Handy bleibt alles rechts. */}
      {!isMobile && (
        <div className="header-center" style={{ flex: istLanding ? '0 1 auto' : '0 1 780px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 12, minWidth: 0 }}>
          {!istLanding && zurueckKnopf}
          {!istLanding && <UseCaseSearch />}
          {sprachwahl}
          {fragenKnopf}
        </div>
      )}

      <div className="header-right" style={isMobile ? undefined : { flex: '1 1 0', justifyContent: 'flex-start', paddingLeft: 12, minWidth: 0 }}>
        {isMobile && fragenKnopf}

        {previewAsUser && originalIsOrganizer && (
          <button
            type="button"
            className="dex-ui-chip is-active"
            onClick={() => setPreviewAsUser(false)}
            title={t('Zurück zur Organizer-Ansicht', 'Back to the organizer view')}
            style={{ whiteSpace: 'nowrap', flexShrink: 0 }}
          >
            {t('User-Ansicht · zurück', 'User view · back')}
          </button>
        )}

        {isMobile && sprachwahl}

        <div ref={menueRef} style={{ position: 'relative', flexShrink: 0 }}>
          <button
            type="button"
            className="header-icon-btn"
            aria-haspopup="menu"
            aria-expanded={menueOffen}
            aria-label={t('Menü', 'Menu')}
            title={t('Menü: Über die App, Ansicht', 'Menu: about the app, view')}
            onClick={() => setMenueOffen(o => !o)}
            style={menueOffen ? { background: 'var(--dex-gray-200)' } : undefined}
          >
            <Menu size={20} />
          </button>
          {menueOffen && (
            <div role="menu" style={{
              position: 'absolute', top: 'calc(100% + 8px)', right: 0, zIndex: 1200, width: 300, maxWidth: 'calc(100vw - 24px)',
              background: '#fff', border: '1px solid var(--dex-gray-200, #e1e1e1)', borderRadius: 12,
              boxShadow: '0 12px 32px rgba(0,0,0,0.14)', padding: 8,
            }}>
              {menuePunkte.map(it => (
                <button key={it.key} type="button" role="menuitem" className="dex-ui-menuitem" onClick={() => { setMenueOffen(false); it.onClick(); }}>
                  <span style={{ width: 16, flexShrink: 0, color: 'var(--dex-gray-600)', paddingTop: 2 }}>{it.icon}</span>
                  <span style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                    <span style={{ fontSize: '0.86rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>{it.titel}</span>
                    <span style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)', lineHeight: 1.4 }}>{it.sub}</span>
                  </span>
                </button>
              ))}
              {/* Nur wer wirklich Organizer oder Admin ist. In der User-Ansicht
                  sind `isAdmin`/`isOrganizer` abgesenkt — der Eintrag darf dann
                  nicht verschwinden, sonst kommt man nicht zurück. */}
              {originalIsOrganizer && (
                <>
                  <div style={{ height: 1, background: 'var(--dex-gray-200)', margin: '6px 4px' }} />
                  <div style={{ padding: '6px 10px 4px', fontSize: '0.68rem', fontWeight: 700, letterSpacing: '0.06em', textTransform: 'uppercase', color: 'var(--dex-gray-500)' }}>
                    {t('Ansicht', 'View')}
                  </div>
                  {ansichtPunkte.map(it => (
                    <button
                      key={it.key}
                      type="button"
                      role="menuitemradio"
                      aria-checked={it.aktiv}
                      className={cx('dex-ui-menuitem', it.aktiv && 'is-active')}
                      onClick={() => { setMenueOffen(false); it.onClick(); }}
                    >
                      <span style={{ width: 16, flexShrink: 0, color: 'var(--dex-green-dark, #4a7c1f)', paddingTop: 2 }}>{it.aktiv && <Check size={14} />}</span>
                      <span style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                        <span style={{ fontSize: '0.86rem', fontWeight: it.aktiv ? 700 : 600, color: 'var(--dex-gray-800)' }}>{it.titel}</span>
                        <span style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)', lineHeight: 1.4 }}>{it.sub}</span>
                      </span>
                    </button>
                  ))}
                </>
              )}
            </div>
          )}
        </div>

        <div ref={profilRef} style={{ position: 'relative' }}>
          {photoUrl && !fotoKaputt ? (
            <img
              src={photoUrl}
              alt={currentUser.displayName}
              onClick={() => setProfilOffen(o => !o)}
              onError={() => setFotoKaputt(true)}
              style={{ width: 36, height: 36, borderRadius: '50%', objectFit: 'cover', cursor: 'pointer', display: 'block' }}
            />
          ) : (
            <div
              className="header-avatar"
              title={currentUser.displayName}
              onClick={() => setProfilOffen(o => !o)}
              style={{ cursor: 'pointer' }}
            >{initialen}</div>
          )}
          {profilOffen && (
            <div style={{
              position: 'absolute', right: 0, top: '100%', marginTop: 8,
              background: '#fff', borderRadius: 'var(--dex-radius-lg, 12px)',
              boxShadow: '0 8px 32px rgba(0,0,0,0.18)', padding: '20px 24px',
              minWidth: 260, maxWidth: 'calc(100vw - 24px)', zIndex: 1200,
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 16 }}>
                {photoUrl && !fotoKaputt ? (
                  <img src={photoUrl} alt="" style={{ width: 48, height: 48, borderRadius: '50%', objectFit: 'cover' }} />
                ) : (
                  <div style={{
                    width: 48, height: 48, borderRadius: '50%', background: 'linear-gradient(135deg, #86bc25, #0076a8)',
                    display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontWeight: 700, fontSize: '1.1rem',
                  }}>{initialen}</div>
                )}
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 600, fontSize: '1rem' }}>{currentUser.displayName}</div>
                  <div style={{ color: '#666', fontSize: '0.85rem', overflowWrap: 'anywhere' }}>{currentUser.email}</div>
                </div>
              </div>
              <div style={{ fontSize: '0.85rem', marginBottom: 14 }}>
                <span style={{ display: 'inline-block', padding: '2px 10px', borderRadius: 12, background: rc.bg, color: rc.color, fontSize: '0.8rem', fontWeight: 500 }}>
                  {rolleLabel(currentUserRole)}
                </span>
              </div>
              {/* Für Rückfragen: die Seiten-Kennung, wie in DEX. Wer „auf der
                  Seite X geht etwas nicht" schreibt, ist damit eindeutig. */}
              <div
                title={t('Bei Rückfragen kannst du diese Kennung nennen.', 'You can quote this ID when you ask about a page.')}
                style={{ paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)', fontSize: '0.7rem', color: 'var(--dex-gray-500)', display: 'flex', alignItems: 'center', gap: 6, fontFamily: 'monospace' }}
              >
                <span style={{ color: 'var(--dex-gray-400)' }}>Page-ID:</span>
                <span style={{ fontWeight: 600, color: 'var(--dex-gray-700)' }}>{currentPage}</span>
              </div>
            </div>
          )}
        </div>
      </div>
    </header>
  );
}
