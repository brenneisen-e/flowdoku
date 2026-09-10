/**
 * Die Detailseite eines Use Cases.
 *
 * Reihenfolge nach dem UI-Leitfaden: erst wer/was (Kopf), dann die HANDLUNG
 * (Demo starten), dann die Erklaerung, dann die Ressourcen. Der Start-Knopf
 * steht oben, weil er der Grund ist, warum jemand hier ist.
 */

import * as React from 'react';
import { cx, ensureDexUiStyles } from './dexUi';
import { bewertungPunkte } from './UseCaseCard';
import { ChevronLeft, ExternalLink, Link2 as LinkIcon, FileText, Book, Video, AlertCircle, Pencil } from './Icons';
import { useUseCases } from '../context/UseCaseContext';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { UseCase } from '../types';

interface RessourceZeile {
  key: string;
  url: string;
  titel: string;
  beschreibung: string;
  icon: React.ReactElement;
}

export default function UseCaseDetailPage(props: { useCaseId?: number }): React.ReactElement {
  ensureDexUiStyles();

  const { useCases, ladeStatus } = useUseCases();
  const { t, isDe } = useLanguage();
  const { navigate, goBack, canGoBack } = useNavigation();
  const { isKurator } = useRoles();

  // ALLE Hooks stehen vor dem ersten frühen Return. In DEX hat ein Hook
  // hinter einem Return die Hook-Reihenfolge zerrissen und nach jeder
  // Anmeldung einen weissen Bildschirm erzeugt (React #300, v30.3).
  const [eingebettetOffen, setEingebettetOffen] = React.useState(false);

  const uc: UseCase | undefined = useCases.filter(u => u.id === props.useCaseId)[0];

  if (ladeStatus === 'laedt') {
    return (
      <div className="dex-ui-empty">
        <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
      </div>
    );
  }

  if (!uc) {
    return (
      <div className="dex-ui-empty">
        <div className="dex-ui-empty-title">{t('Dieser Use Case ist nicht (mehr) da', 'This use case is not (or no longer) here')}</div>
        <div className="dex-ui-empty-desc">
          {t('Er wurde gelöscht, oder der Link zeigt auf eine Nummer, die es nicht gibt.',
            'It was deleted, or the link points to an id that does not exist.')}
        </div>
        <button type="button" className="dex-ui-empty-action" onClick={() => navigate('start')}>
          {t('Zur Übersicht', 'Back to overview')}
        </button>
      </div>
    );
  }

  const start = uc.ressourcen.deployment;
  const kannStarten = uc.status === 'Live' && !!start;

  const ressourcen: RessourceZeile[] = [
    { key: 'code', url: uc.ressourcen.sourceCode, titel: t('Source Code', 'Source code'), beschreibung: t('Das Repository zur Demo.', 'The demo repository.'), icon: <LinkIcon size={16} /> },
    { key: 'guide', url: uc.ressourcen.deploymentGuide, titel: t('Deployment-Guide', 'Deployment guide'), beschreibung: t('Wie du die Demo selbst aufsetzt.', 'How to set the demo up yourself.'), icon: <FileText size={16} /> },
    { key: 'wiki', url: uc.ressourcen.wiki, titel: t('Wiki / Doku', 'Wiki / docs'), beschreibung: t('Fachliche Beschreibung des Use Cases.', 'The functional description of the use case.'), icon: <Book size={16} /> },
    { key: 'video', url: uc.ressourcen.video, titel: t('Use-Case-Video', 'Use case video'), beschreibung: t('Aufzeichnung — falls die Live-Demo mal klemmt.', 'Recording — in case the live demo fails.'), icon: <Video size={16} /> },
  ];
  const vorhandene = ressourcen.filter(r => !!r.url);

  const starte = (): void => {
    if (!kannStarten) return;
    if (uc.aufrufArt === 'eingebettet') { setEingebettetOffen(true); return; }
    // `noopener` ist Pflicht: Ohne das kann die geoeffnete Seite ueber
    // `window.opener` auf diese hier zugreifen.
    window.open(start, '_blank', 'noopener,noreferrer');
  };

  return (
    <div>
      <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" style={{ marginBottom: 12 }} onClick={() => (canGoBack ? goBack() : navigate('start'))}>
        <ChevronLeft size={14} /> {t('Zurück', 'Back')}
      </button>

      <div className="dex-ui-page-head">
        <div>
          {uc.bereich && <p className="dex-ui-page-head-meta" style={{ marginBottom: 4 }}>{uc.bereich}</p>}
          <h1 className="dex-ui-page-head-title">{uc.titel}</h1>
        </div>
        <div className="dex-ui-page-head-actions">
          {isKurator && (
            <button type="button" className="dex-ui-textbtn" onClick={() => navigate('verwaltung', uc.id)}>
              <Pencil size={14} /> {t('Bearbeiten', 'Edit')}
            </button>
          )}
        </div>
      </div>

      {/* Die Handlung, derentwegen jemand hier ist. */}
      <div className="dex-ui-section">
        {kannStarten ? (
          <>
            <button type="button" className="btn btn-primary" onClick={starte} style={{ minHeight: 44 }}>
              {uc.aufrufArt === 'eingebettet'
                ? t('Demo hier öffnen', 'Open demo here')
                : <>{t('Demo starten', 'Start demo')} <ExternalLink size={15} /></>}
            </button>
            <p className="dex-ui-help" style={{ marginTop: 8 }}>
              {uc.aufrufArt === 'eingebettet'
                ? t('Die Demo läuft eingebettet auf dieser Seite.', 'The demo runs embedded on this page.')
                : t('Öffnet in einem neuen Tab.', 'Opens in a new tab.')}
            </p>
          </>
        ) : (
          <div className="dex-ui-callout dex-ui-callout--neutral">
            <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
            <span>
              {uc.status === 'Geplant' && t('Dieser Use Case ist geplant — es gibt noch keine Demo zum Starten.', 'This use case is planned — there is no demo to start yet.')}
              {uc.status === 'InArbeit' && t('Die Demo wird gerade gebaut. Sobald sie steht, erscheint hier der Start-Knopf.', 'The demo is being built. The start button appears here as soon as it is ready.')}
              {uc.status === 'Archiviert' && t('Dieser Use Case ist archiviert.', 'This use case is archived.')}
              {uc.status === 'Live' && !start && t('Der Use Case steht auf „Live", aber es ist kein Deployment-Link hinterlegt. Das ist ein Pflegefehler — ein Kurator kann ihn nachtragen.', 'The use case is marked "Live", but no deployment link is stored. A curator can add it.')}
            </span>
          </div>
        )}
      </div>

      {/* Eingebettet: erst nach dem Klick. Ein iframe, das beim Seitenaufbau
          mitlaedt, kostet jeden Besucher Ladezeit fuer etwas, das die
          meisten nicht oeffnen. */}
      {eingebettetOffen && kannStarten && (
        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{t('Demo', 'Demo')}</div>
          <div className="dex-ui-inline" style={{ marginBottom: 8 }}>
            <button type="button" className="dex-ui-textbtn" onClick={() => window.open(start, '_blank', 'noopener,noreferrer')}>
              <ExternalLink size={14} /> {t('In neuem Tab öffnen', 'Open in a new tab')}
            </button>
            <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={() => setEingebettetOffen(false)}>
              {t('Schließen', 'Close')}
            </button>
          </div>
          <iframe
            src={start}
            title={uc.titel}
            style={{ width: '100%', height: '70vh', minHeight: 420, border: '1px solid var(--dex-gray-200, #e8e8e8)', borderRadius: 'var(--dex-radius, 12px)', background: '#fff' }}
          />
          {/* Der Satz steht hier, weil ein leeres iframe sonst wie ein Fehler
              der Plattform aussieht. Permissions Policy reicht Rechte nur
              abwaerts durch: Was das SharePoint-WebView nicht hat, kann keine
              eingebettete Seite gewinnen — daran sind in DEX fuenf Anlaeufe
              fuer den Kamera-Scan gescheitert. Dazu verbieten viele Ziele das
              Einbetten selbst per X-Frame-Options. */}
          <p className="dex-ui-help" style={{ marginTop: 8 }}>
            {t('Bleibt der Bereich leer, erlaubt die Demo das Einbetten nicht — dann hilft „In neuem Tab öffnen". Demos mit Kamera oder Mikrofon funktionieren nur im eigenen Tab.',
              'If this area stays empty, the demo does not allow embedding — use "Open in a new tab". Demos using camera or microphone only work in their own tab.')}
          </p>
        </div>
      )}

      {uc.kurzbeschreibung && (
        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{t('Worum es geht', 'What this is about')}</div>
          <p style={{ margin: 0 }}>{uc.kurzbeschreibung}</p>
          {uc.beschreibung && (
            <div
              style={{ marginTop: 12 }}
              // Der Text kommt aus der eigenen SharePoint-Liste und wird nur
              // von Kuratoren gepflegt — dieselbe Vertrauensstellung wie die
              // Event-Beschreibung in DEX.
              dangerouslySetInnerHTML={{ __html: uc.beschreibung }}
            />
          )}
        </div>
      )}

      <div className="dex-ui-section">
        <div className="dex-ui-section-title">{t('Bewertung', 'Assessment')}</div>
        <div className="dex-ui-kpi-row">
          {[
            { label: t('Sales-Relevanz', 'Sales relevance'), wert: uc.salesRelevanz },
            { label: t('Machbarkeit', 'Feasibility'), wert: uc.machbarkeit },
            { label: t('Demo-Tauglichkeit', 'Demo readiness'), wert: uc.demoTauglichkeit },
          ].map(k => (
            <div key={k.label} className={cx('dex-ui-kpi', k.wert === 'Hoch' ? 'dex-ui-kpi--green' : k.wert === 'unbewertet' ? 'dex-ui-kpi--gray' : 'dex-ui-kpi--blue')}>
              <div className="dex-ui-kpi-value" style={{ letterSpacing: '0.1em' }}>{bewertungPunkte(k.wert)}</div>
              <div className="dex-ui-kpi-label">{k.label}</div>
              <div className="dex-ui-kpi-sub">
                {k.wert === 'unbewertet' ? t('noch nicht bewertet', 'not assessed yet') : k.wert}
              </div>
            </div>
          ))}
        </div>
      </div>

      <div className="dex-ui-section">
        <div className="dex-ui-section-title">{t('Ressourcen', 'Resources')}</div>
        {vorhandene.length === 0 ? (
          <p className="dex-ui-help" style={{ margin: 0 }}>
            {isDe
              ? 'Für diesen Use Case ist noch nichts hinterlegt — Source Code, Guide, Wiki und Video kommen, sobald das Team sie liefert.'
              : 'Nothing is stored for this use case yet — source code, guide, wiki and video follow once the team delivers them.'}
          </p>
        ) : (
          <div className="dex-ui-stack">
            {vorhandene.map(r => (
              <a
                key={r.key}
                className="dex-ui-row dex-ui-row--bordered dex-ui-rowbtn"
                href={r.url}
                target="_blank"
                rel="noopener noreferrer"
                style={{ textDecoration: 'none', color: 'inherit' }}
              >
                <span className="dex-ui-action-icon" aria-hidden="true">{r.icon}</span>
                <span className="dex-ui-row-main">
                  <span className="dex-ui-row-title">{r.titel}</span>
                  <span className="dex-ui-row-sub">{r.beschreibung}</span>
                </span>
                <span className="dex-ui-row-actions"><ExternalLink size={15} /></span>
              </a>
            ))}
          </div>
        )}
      </div>

      {(uc.betreuerNamen.length > 0 || uc.geaendertVon) && (
        <p className="dex-ui-help">
          {uc.betreuerNamen.length > 0 && <>{t('Betreut von', 'Maintained by')} {uc.betreuerNamen.join(', ')}. </>}
          {uc.geaendertVon && <>{t('Zuletzt geändert von', 'Last changed by')} {uc.geaendertVon}.</>}
        </p>
      )}
    </div>
  );
}
