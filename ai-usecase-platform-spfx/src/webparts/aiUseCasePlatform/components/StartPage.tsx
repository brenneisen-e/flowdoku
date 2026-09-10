/**
 * Die Kachelwand — die Startseite der Plattform.
 *
 * Reihenfolge nach dem UI-Leitfaden: erst der Kopf mit der Aussage, dann
 * das Werkzeug zum Finden (Suche + Bereichsfilter), dann die Kacheln. Kein
 * Werbekasten davor, kein Karussell — wer hier landet, sucht eine Demo.
 */

import * as React from 'react';
import { cx, ensureDexUiStyles } from './dexUi';
import UseCaseCard from './UseCaseCard';
import { Search, X } from './Icons';
import { useUseCases } from '../context/UseCaseContext';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { APP_NAME, APP_SUBTITLE_DE, APP_SUBTITLE_EN } from '../constants';

export default function StartPage(): React.ReactElement {
  // Das Stylesheet stellt die SEITE sicher, nicht die Kachel — sonst
  // injiziert jede der zwanzig Kacheln dieselbe Pruefung (DEX-Konvention).
  ensureDexUiStyles();

  const { useCases, ladeStatus, letzterStatus, letzterFehler, fehlendeSpalten, reload, bereiche } = useUseCases();
  const { t, isDe } = useLanguage();
  const { navigate } = useNavigation();
  const { isKurator } = useRoles();

  const [suche, setSuche] = React.useState('');
  const [bereich, setBereich] = React.useState<string>('');
  const [nurLive, setNurLive] = React.useState(false);

  /**
   * Suchvergleich ohne Bindestriche und Umlaute.
   *
   * Genau die Falle aus DEX v31.10: Dort fand die Aktionssuche „email" nicht,
   * weil die Titel „E-Mail" heissen und roh verglichen wurde. Hier heissen
   * Use Cases „Credit-Memo-Agent" und „Anlage-Advisory" — wer „credit memo"
   * tippt, muss sie finden.
   */
  const norm = (s: string): string => (s || '')
    .toLowerCase()
    .replace(/ä/g, 'a').replace(/ö/g, 'o').replace(/ü/g, 'u').replace(/ß/g, 'ss')
    .replace(/[^a-z0-9]/g, '');

  const gefiltert = React.useMemo(() => {
    const tokens = suche.trim().split(/\s+/).map(norm).filter(Boolean);
    return useCases
      .filter(u => u.status !== 'Archiviert' || isKurator)
      .filter(u => !nurLive || u.status === 'Live')
      .filter(u => !bereich || u.bereich === bereich)
      .filter(u => {
        if (tokens.length === 0) return true;
        const heu = norm(u.titel) + ' ' + norm(u.kurzbeschreibung) + ' ' + norm(u.bereich) + ' ' + norm(u.schlagworte.join(' '));
        return tokens.every(tok => heu.indexOf(tok) >= 0);
      });
  }, [useCases, suche, bereich, nurLive, isKurator]);

  const liveAnzahl = useCases.filter(u => u.status === 'Live').length;

  return (
    <div>
      <div className="dex-ui-page-head">
        <div>
          <h1 className="dex-ui-page-head-title">{APP_NAME}</h1>
          <p className="dex-ui-page-head-meta">
            {isDe ? APP_SUBTITLE_DE : APP_SUBTITLE_EN}
            {ladeStatus === 'ok' && ` · ${liveAnzahl} ${t('Demos verfügbar', 'demos available')}`}
          </p>
        </div>
      </div>

      {/* Finden. Steht vor den Kacheln, weil man bei zwanzig Demos sucht
          statt scrollt — und nach den Kacheln waere es unauffindbar. */}
      <div className="dex-ui-toolbar" style={{ marginBottom: 16 }}>
        <div className="dex-ui-searchbar">
          <span className="dex-ui-searchbar-icon"><Search size={15} /></span>
          <input
            type="text"
            id="uc-suche"
            className="dex-ui-input dex-ui-input--sm"
            value={suche}
            onChange={e => setSuche(e.target.value)}
            placeholder={t('Use Case suchen …', 'Search use case …')}
            aria-label={t('Use Case suchen', 'Search use case')}
            style={{ paddingRight: 34 }}
          />
          {suche && (
            <button
              type="button"
              className="dex-ui-iconbtn"
              onClick={() => setSuche('')}
              aria-label={t('Suche leeren', 'Clear search')}
              style={{ position: 'absolute', right: 3, top: '50%', transform: 'translateY(-50%)', width: 32, height: 32 }}
            ><X size={14} /></button>
          )}
        </div>

        {bereiche.length > 0 && (
          <select
            id="uc-bereich"
            className="dex-ui-select dex-ui-select--sm"
            value={bereich}
            onChange={e => setBereich(e.target.value)}
            aria-label={t('Bereich', 'Area')}
          >
            <option value="">{t('Alle Bereiche', 'All areas')}</option>
            {bereiche.map(b => <option key={b} value={b}>{b}</option>)}
          </select>
        )}

        <button
          type="button"
          className={cx('dex-ui-chip', nurLive && 'is-active')}
          aria-pressed={nurLive}
          onClick={() => setNurLive(v => !v)}
        >{t('Nur aufrufbare', 'Only callable')}</button>

        <span className="dex-ui-toolbar-spacer" />

        {isKurator && (
          <button type="button" className="btn btn-secondary" onClick={() => navigate('verwaltung')}>
            {t('Use Cases pflegen', 'Manage use cases')}
          </button>
        )}
      </div>

      {/* Drei Zustaende, drei Meldungen. „Keine Use Cases" bei einem
          Lesefehler waere eine Aussage ueber die Daten, die niemand geprueft
          hat — die Lehre aus DEX v30.37. */}
      {ladeStatus === 'laedt' && (
        <div className="dex-ui-empty">
          <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
          <div className="dex-ui-empty-desc" style={{ marginTop: 12 }}>{t('Use Cases werden geladen …', 'Loading use cases …')}</div>
        </div>
      )}

      {ladeStatus === 'fehler' && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
          <span>
            <strong>{t('Die Use Cases konnten nicht geladen werden.', 'The use cases could not be loaded.')}</strong>
            <br />
            {letzterStatus === 403
              ? t('Dir fehlt das Leserecht auf der Liste — bitte melde dich bei einem Admin der Plattform.',
                'You do not have read access to the list — please contact a platform admin.')
              : letzterStatus === 404
                ? t('Die Liste gibt es (noch) nicht. Lade die Seite neu — die App legt sie beim Start selbst an.',
                  'The list does not exist (yet). Reload the page — the app creates it on start.')
                : letzterStatus === 400
                  ? t('SharePoint hat die Abfrage abgelehnt. Meist fehlt eine Spalte, die beim Anlegen der Liste nicht entstanden ist.',
                    'SharePoint rejected the query. Usually a column is missing that was not created with the list.')
                  : t('Das ist ein Lade-Fehler, keine leere Plattform. Versuch es gleich noch einmal.',
                    'This is a loading error, not an empty platform. Please try again shortly.')}
            {letzterStatus > 0 && <span className="dex-ui-muted"> (HTTP {letzterStatus})</span>}
            {/* Der Klartext von SharePoint. Ohne ihn raet beim naechsten Mal
                wieder jemand — mit ihm steht die Ursache im Bild. */}
            {letzterFehler && (
              <>
                <br />
                <span className="dex-ui-muted" style={{ fontSize: '0.78rem' }}>{letzterFehler}</span>
              </>
            )}
            {fehlendeSpalten.length > 0 && (
              <>
                <br />
                <span className="dex-ui-muted" style={{ fontSize: '0.78rem' }}>
                  {t('Nicht angelegte Spalten: ', 'Columns not created: ')}{fehlendeSpalten.join(', ')}
                </span>
              </>
            )}
            <br />
            <button type="button" className="dex-ui-textbtn" style={{ marginTop: 8 }} onClick={() => { void reload(); }}>
              {t('Erneut versuchen', 'Try again')}
            </button>
          </span>
        </div>
      )}

      {ladeStatus === 'ok' && gefiltert.length === 0 && (
        <div className="dex-ui-empty">
          <div className="dex-ui-empty-title">
            {useCases.length === 0
              ? t('Noch keine Use Cases', 'No use cases yet')
              : t('Kein Treffer', 'No match')}
          </div>
          <div className="dex-ui-empty-desc">
            {useCases.length === 0
              ? t('Sobald die ersten Agent-Demos stehen, erscheinen sie hier als Kacheln.',
                'As soon as the first agent demos exist, they appear here as tiles.')
              : t('Für diese Suche gibt es nichts. Lösch den Filter oder such nach etwas anderem.',
                'Nothing matches this search. Clear the filter or search for something else.')}
          </div>
          {useCases.length > 0 && (
            <button type="button" className="dex-ui-empty-action" onClick={() => { setSuche(''); setBereich(''); setNurLive(false); }}>
              {t('Filter zurücksetzen', 'Reset filters')}
            </button>
          )}
          {useCases.length === 0 && isKurator && (
            <button type="button" className="dex-ui-empty-action" onClick={() => navigate('verwaltung')}>
              {t('Ersten Use Case anlegen', 'Create the first use case')}
            </button>
          )}
        </div>
      )}

      {ladeStatus === 'ok' && gefiltert.length > 0 && (
        <div
          style={{
            display: 'grid',
            // `auto-fill` mit Mindestbreite: Auf dem Handy wird daraus von
            // selbst eine Spalte, ohne eigene Media-Query.
            gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))',
            gap: 16,
          }}
        >
          {gefiltert.map(uc => (
            <UseCaseCard key={uc.id} useCase={uc} onOpen={id => navigate('detail', id)} />
          ))}
        </div>
      )}
    </div>
  );
}
