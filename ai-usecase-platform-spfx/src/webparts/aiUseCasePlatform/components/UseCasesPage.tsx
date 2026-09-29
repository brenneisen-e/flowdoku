/**
 * Die Kachelwand — alle Use Cases als Kacheln (Kachel „Use Cases" auf der
 * Start-Übersicht).
 *
 * Reihenfolge nach dem UI-Leitfaden: erst der Kopf mit der Aussage, dann
 * das Werkzeug zum Finden (Bereichsfilter), dann die Kacheln. Kein
 * Werbekasten davor, kein Karussell — wer hier landet, sucht eine Demo.
 *
 * v1.3: Die Suche sitzt in der Kopfzeile (wie bei DEX) und ist die einzige.
 * Die Wand liest denselben Suchtext und filtert danach; ein zweites Suchfeld
 * auf der Seite wäre ein zweiter Bedienweg für dieselbe Auswahl. Nur auf dem
 * Handy steht das Feld hier, weil die Kopfzeile dort keinen Platz dafür hat —
 * es ist dann das einzige sichtbare, und es bedient denselben Text.
 */

import * as React from 'react';
import { cx, ensureDexUiStyles } from './dexUi';
import UseCaseCard from './UseCaseCard';
import { Search, X } from './Icons';
import { useUseCases } from '../context/UseCaseContext';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useSuche } from '../context/SucheContext';
import { useIsMobile } from '../utils/useIsMobile';
import { passtZurSuche, suchTokens } from '../utils/suche';
import { APP_SUBTITLE_DE, APP_SUBTITLE_EN } from '../constants';
import { START_USE_CASES } from '../data/startUseCases';
import { startbestandMeldung } from '../utils/startbestandText';

export default function UseCasesPage(): React.ReactElement {
  // Das Stylesheet stellt die SEITE sicher, nicht die Kachel — sonst
  // injiziert jede der zwanzig Kacheln dieselbe Prüfung (DEX-Konvention).
  ensureDexUiStyles();

  const { useCases, ladeStatus, letzterStatus, letzterFehler, fehlendeSpalten, aktualisierungFehler, startbestandTeilweise, reload, bereiche, seedStartUseCases } = useUseCases();
  const { t, isDe } = useLanguage();
  const { navigate } = useNavigation();
  const { isOrganizer } = useRoles();

  const { suche, setSuche } = useSuche();
  const isMobile = useIsMobile();
  const [bereich, setBereich] = React.useState<string>('');
  const [nurLive, setNurLive] = React.useState(false);
  const [seedBusy, setSeedBusy] = React.useState(false);
  const [seedMeldung, setSeedMeldung] = React.useState('');

  // v1.1: Die Start-Use-Cases von Hand anlegen. Meldet den GRUND, wenn nichts
  // entsteht — „hat nicht geklappt" lässt den Nutzer raten, ob es an den
  // Rechten oder an der App lag. v1.3: Text und Grund kommen aus dem Ergebnis des
  // Laufs (`startbestandMeldung`), nicht aus einem State von vor dem Lauf.
  const seedJetzt = async (): Promise<void> => {
    setSeedBusy(true);
    setSeedMeldung('');
    try {
      const erg = await seedStartUseCases();
      const m = startbestandMeldung(erg, t);
      if (m.art !== 'ok') setSeedMeldung(m.text);
    } finally {
      setSeedBusy(false);
    }
  };

  const gefiltert = React.useMemo(() => {
    // Der Vergleich steht in `utils/suche` — dieselbe Regel wie das
    // Ergebnisfenster der Kopfzeile, sonst fände das eine, was das andere nicht findet.
    const tokens = suchTokens(suche);
    return useCases
      .filter(u => u.status !== 'Archiviert' || isOrganizer)
      .filter(u => !nurLive || u.status === 'Live')
      .filter(u => !bereich || u.bereich === bereich)
      .filter(u => passtZurSuche(u, tokens));
  }, [useCases, suche, bereich, nurLive, isOrganizer]);

  const liveAnzahl = useCases.filter(u => u.status === 'Live').length;

  return (
    <div>
      <div className="dex-ui-page-head">
        <div>
          <h1 className="dex-ui-page-head-title">Use Cases</h1>
          <p className="dex-ui-page-head-meta">
            {isDe ? APP_SUBTITLE_DE : APP_SUBTITLE_EN}
            {ladeStatus === 'ok' && ` · ${liveAnzahl} ${t('Demos verfügbar', 'demos available')}`}
          </p>
        </div>
      </div>

      {/* Finden. Steht vor den Kacheln, weil man bei zwanzig Demos sucht
          statt scrollt — und nach den Kacheln wäre es unauffindbar. */}
      <div className="dex-ui-toolbar" style={{ marginBottom: 16 }}>
        {/* Nur auf dem Handy: Die Kopfzeile hat dort keinen Platz für ein Feld.
            Es bedient denselben Suchtext wie das Feld der Kopfzeile. */}
        {isMobile && (
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
        )}

        {/* Auf dem Rechner sagt ein Chip, WONACH gefiltert ist — sonst wundert
            man sich über eine kurze Liste, ohne dass in der Seite ein
            Suchfeld zu sehen wäre, das den Grund trägt. */}
        {!isMobile && suche.trim() && (
          <button
            type="button"
            className="dex-ui-chip is-active"
            onClick={() => setSuche('')}
            title={t('Suche aufheben', 'Clear search')}
          >
            {t('Suche', 'Search')}: „{suche.trim()}“ <X size={12} />
          </button>
        )}

        {bereiche.length > 0 && (
          <select
            id="uc-bereich"
            className="dex-ui-select dex-ui-select--sm"
            // Ohne Suchfeld daneben (seit v1.3 sitzt sie in der Kopfzeile) dehnte sich die
            // Auswahl über die ganze Zeile.
            style={{ width: 'auto', minWidth: 200, maxWidth: 280 }}
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

        {isOrganizer && (
          <button type="button" className="btn btn-secondary" onClick={() => navigate('studio')}>
            Use Case Studio
          </button>
        )}
      </div>

      {/* Drei Zustände, drei Meldungen. „Keine Use Cases" bei einem
          Lesefehler wäre eine Aussage über die Daten, die niemand geprüft
          hat — die Lehre aus DEX v30.37. */}
      {ladeStatus === 'ok' && aktualisierungFehler && (
        <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" role="status" style={{ marginBottom: 12 }}>
          <span className="dex-ui-callout-body">
            {t('Der aktuelle Stand konnte nicht nachgeladen werden — du siehst den zuletzt geladenen. ',
              'The current state could not be reloaded — you see the last loaded one. ')}
            <button type="button" className="dex-ui-textbtn" onClick={() => { void reload(); }}>{t('Erneut versuchen', 'Try again')}</button>
          </span>
        </div>
      )}

      {isOrganizer && startbestandTeilweise && (
        <div className="dex-ui-callout dex-ui-callout--warn" role="status" style={{ marginBottom: 12 }}>
          <span className="dex-ui-callout-body">
            {startbestandMeldung(startbestandTeilweise, t).text}{' '}
            {/* Der Knopf nur, wo etwas fehlt: Bei „vollständig, aber Merker nicht gespeichert" gäbe es
                nichts nachzulegen. */}
            {startbestandTeilweise.angelegt < startbestandTeilweise.fehlend && (
              <button type="button" className="dex-ui-textbtn" disabled={seedBusy} onClick={() => { void seedJetzt(); }}>
                {seedBusy ? t('Wird angelegt …', 'Creating …') : t('Fehlende anlegen', 'Add missing ones')}
              </button>
            )}
          </span>
        </div>
      )}

      {/* Die Antwort auf einen Klick auf „Fehlende anlegen" bzw. den Knopf im Leerzustand. Sie stand
          nur im Leerzustand — beim Klick im Banner (Kacheln vorhanden) blieb der Knopf ohne Rückmeldung. */}
      {isOrganizer && seedMeldung && (
        <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" role="status" style={{ marginBottom: 12 }}>
          <span className="dex-ui-callout-body">{seedMeldung}</span>
        </div>
      )}

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
            {/* Der Klartext von SharePoint nur für Organizer und Admins: Er trägt Interna
                (Korrelations-Ids, Spaltennamen), die einem normalen Nutzer nichts helfen. */}
            {isOrganizer && letzterFehler && (
              <>
                <br />
                <span className="dex-ui-muted" style={{ fontSize: '0.78rem' }}>{letzterFehler}</span>
              </>
            )}
            {isOrganizer && fehlendeSpalten.length > 0 && (
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
          {useCases.length === 0 && isOrganizer && (
            <div className="dex-ui-inline dex-ui-empty-action" style={{ justifyContent: 'center', flexWrap: 'wrap', gap: 10 }}>
              {/* v1.1: Das Netz unter der automatischen Erstbefüllung. Die
                  läuft nur, wenn auch das Protokoll lesbar ist; genau dann
                  stand die Plattform vorher leer da, ohne Weg sie zu füllen. */}
              <button
                type="button"
                className="btn btn-primary"
                disabled={seedBusy}
                onClick={() => { void seedJetzt(); }}
              >
                {seedBusy
                  ? t('Wird angelegt …', 'Creating …')
                  : t(`Die ${START_USE_CASES.length} Start-Use-Cases anlegen`, `Add the ${START_USE_CASES.length} starter use cases`)}
              </button>
              <button type="button" className="btn btn-secondary" onClick={() => navigate('studio')}>
                {t('Eigenen Use Case anlegen', 'Create my own use case')}
              </button>
            </div>
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
