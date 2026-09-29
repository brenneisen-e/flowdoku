/**
 * Die Suche in der Kopfzeile — Vorbild `GlobalSearch` in DEX.
 *
 * Sie ist die EINZIGE Suche der Plattform. Auf der Kachelwand filtert sie die
 * Kacheln direkt (dort ist das Ergebnis die Seite selbst, ein Ergebnisfenster
 * darüber wäre doppelt); auf jeder anderen Seite zeigt sie unter dem Feld die
 * besten Treffer und führt mit Enter oder „Alle Treffer" auf die Kachelwand.
 *
 * Zwei Zustände, die nicht zusammenfallen dürfen: „keine Treffer" und „die
 * Use Cases konnten nicht geladen werden". Bei einem Lesefehler ist die Liste
 * leer, und „nichts gefunden" wäre eine Aussage über Daten, die niemand
 * gelesen hat — dieselbe Regel wie überall in dieser App.
 */

import * as React from 'react';
import { Search, X } from './Icons';
import { cx } from './dexUi';
import { useSuche } from '../context/SucheContext';
import { useUseCases } from '../context/UseCaseContext';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';
import { useRoles } from '../context/RoleContext';
import { passtZurSuche, suchTokens } from '../utils/suche';

/** Wie viele Treffer das Fenster zeigt — mehr ist die Kachelwand. */
const MAX_TREFFER = 6;

export default function UseCaseSearch(): React.ReactElement {
  const { suche, setSuche } = useSuche();
  const { useCases, ladeStatus } = useUseCases();
  const { navigate, currentPage } = useNavigation();
  const { t } = useLanguage();
  const { isOrganizer } = useRoles();

  const [offen, setOffen] = React.useState(false);
  const wrapRef = React.useRef<HTMLDivElement | null>(null);

  // Auf der Kachelwand ist die Seite das Ergebnis.
  const aufWand = currentPage === 'usecases';

  const treffer = React.useMemo(() => {
    const tokens = suchTokens(suche);
    if (tokens.length === 0) return [];
    return useCases
      .filter(u => u.status !== 'Archiviert' || isOrganizer)
      .filter(u => passtZurSuche(u, tokens));
  }, [suche, useCases, isOrganizer]);

  // Klick daneben und Escape schließen das Fenster.
  React.useEffect(() => {
    if (!offen) return undefined;
    const onDown = (e: MouseEvent): void => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) setOffen(false);
    };
    window.addEventListener('mousedown', onDown);
    return () => window.removeEventListener('mousedown', onDown);
  }, [offen]);

  const zurWand = (): void => {
    setOffen(false);
    if (!aufWand) navigate('usecases');
  };

  const fensterSichtbar = offen && !aufWand && suche.trim().length > 0;

  return (
    <div ref={wrapRef} className="dex-ui-searchbar" style={{ position: 'relative', flex: '1 1 220px', minWidth: 0, maxWidth: 420 }}>
      <span className="dex-ui-searchbar-icon"><Search size={15} /></span>
      <input
        type="text"
        className="dex-ui-input dex-ui-input--sm"
        style={{ paddingRight: 34, borderRadius: 999 }}
        value={suche}
        placeholder={t('Suche', 'Search')}
        aria-label={t('Use Cases durchsuchen', 'Search use cases')}
        onChange={e => { setSuche(e.target.value); setOffen(true); }}
        onFocus={() => setOffen(true)}
        onKeyDown={e => {
          if (e.key === 'Enter' && suche.trim()) { e.preventDefault(); zurWand(); }
          else if (e.key === 'Escape') { setOffen(false); }
        }}
      />
      {suche && (
        <button
          type="button"
          className="dex-ui-iconbtn"
          aria-label={t('Suche leeren', 'Clear search')}
          onClick={() => { setSuche(''); setOffen(false); }}
          style={{ position: 'absolute', right: 3, top: '50%', transform: 'translateY(-50%)', width: 28, height: 28 }}
        ><X size={14} /></button>
      )}

      {fensterSichtbar && (
        <div
          role="listbox"
          style={{
            position: 'absolute', top: 'calc(100% + 8px)', left: 0, right: 0, zIndex: 1200,
            background: '#fff', border: '1px solid var(--dex-gray-200, #e1e1e1)', borderRadius: 12,
            boxShadow: '0 12px 32px rgba(0,0,0,0.14)', padding: 6, minWidth: 280,
          }}
        >
          {ladeStatus === 'fehler' && (
            <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" role="status">
              {t('Die Use Cases konnten nicht geladen werden — die Suche kann gerade nichts finden.', 'The use cases could not be loaded — search cannot find anything right now.')}
            </div>
          )}
          {ladeStatus === 'laedt' && (
            <div className="dex-ui-help" style={{ padding: '8px 10px', margin: 0 }}>{t('Use Cases werden geladen …', 'Loading use cases …')}</div>
          )}
          {ladeStatus === 'ok' && treffer.length === 0 && (
            <div className="dex-ui-help" style={{ padding: '8px 10px', margin: 0 }}>{t('Kein Treffer.', 'No match.')}</div>
          )}
          {treffer.slice(0, MAX_TREFFER).map(u => (
            <button
              key={u.id}
              type="button"
              role="option"
              aria-selected={false}
              className="dex-ui-menuitem"
              onClick={() => { setOffen(false); navigate('detail', u.id); }}
            >
              <span style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                <span style={{ fontSize: '0.86rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>{u.titel}</span>
                <span style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)' }}>
                  {u.bereich || t('kein Bereich', 'no area')}
                  {' · '}
                  <span className={cx('dex-ui-pill', 'dex-ui-pill--sm', u.status === 'Live' ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')}>
                    {u.status === 'InArbeit' ? t('In Arbeit', 'In progress') : u.status === 'Geplant' ? t('Geplant', 'Planned') : u.status === 'Live' ? 'Live' : t('Archiviert', 'Archived')}
                  </span>
                </span>
              </span>
            </button>
          ))}
          {ladeStatus === 'ok' && treffer.length > MAX_TREFFER && (
            <button type="button" className="dex-ui-menuitem" onClick={zurWand}>
              <span style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                {t(`Alle ${treffer.length} Treffer anzeigen`, `Show all ${treffer.length} matches`)}
              </span>
            </button>
          )}
        </div>
      )}
    </div>
  );
}
