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
import { ChevronLeft, ExternalLink, Link2 as LinkIcon, FileText, Book, Video, AlertCircle, Pencil, Copy, Check } from './Icons';
import { useUseCases } from '../context/UseCaseContext';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { UseCase } from '../types';
import { linkZumUseCase } from '../constants';
import { sichereUrl, sichereMail, bereinigeHtml } from '../utils/sicher';
import { bewertungText } from '../utils/anzeige';

interface RessourceZeile {
  key: string;
  url: string;
  titel: string;
  beschreibung: string;
  icon: React.ReactElement;
}

export default function UseCaseDetailPage(props: { useCaseId?: number }): React.ReactElement {
  ensureDexUiStyles();

  const { useCases, ladeStatus, letzterStatus, reload } = useUseCases();
  const { t, isDe } = useLanguage();
  const { navigate, goBack, canGoBack } = useNavigation();
  const { isOrganizer } = useRoles();

  // ALLE Hooks stehen vor dem ersten frühen Return. In DEX hat ein Hook
  // hinter einem Return die Hook-Reihenfolge zerrissen und nach jeder
  // Anmeldung einen weissen Bildschirm erzeugt (React #300, v30.3).
  const [eingebettetOffen, setEingebettetOffen] = React.useState(false);
  // Link kopieren: '' = noch nichts, 'kopiert' = in der Zwischenablage,
  // 'manuell' = die Zwischenablage ging nicht, der Link steht zum Markieren da.
  const [linkStatus, setLinkStatus] = React.useState<'' | 'kopiert' | 'manuell'>('');

  const gefunden: UseCase | undefined = useCases.filter(u => u.id === props.useCaseId)[0];
  // Archiviertes sehen nur Organizer — wie auf der Kachelwand. Die Kachel ist
  // dort ausgeblendet, aber ein Deep-Link (`?uc=`) oder ein Merker im Browser
  // führt trotzdem hierher, und „Nur für Organizer sichtbar" (Pflegeseite)
  // gilt sonst nur für die Wand.
  const uc: UseCase | undefined = gefunden && (gefunden.status !== 'Archiviert' || isOrganizer) ? gefunden : undefined;

  if (ladeStatus === 'laedt') {
    return (
      <div className="dex-ui-empty" role="status" aria-live="polite">
        <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
        <div className="dex-ui-empty-desc" style={{ marginTop: 12 }}>{t('Use Case wird geladen …', 'Loading use case …')}</div>
      </div>
    );
  }

  // Ein Lesefehler ist keine Aussage über den Use Case. Bei einem Deep-Link
  // ist diese Seite oft die ERSTE, die lädt — schlägt das Lesen fehl, ist die
  // Liste leer, und „nicht (mehr) da" wäre eine Behauptung über Daten, die
  // niemand gelesen hat.
  if (!uc && ladeStatus === 'fehler') {
    return (
      <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
        <span>
          <strong>{t('Der Use Case konnte nicht geladen werden.', 'The use case could not be loaded.')}</strong>
          <br />
          {letzterStatus === 403
            ? t('Dir fehlt das Leserecht auf der Liste — bitte melde dich bei einem Admin der Plattform.', 'You do not have read access to the list — please contact a platform admin.')
            : t('Das ist ein Lesefehler, kein gelöschter Use Case. Versuch es gleich noch einmal.', 'This is a read error, not a deleted use case. Please try again shortly.')}
          {letzterStatus > 0 && <span className="dex-ui-muted"> (HTTP {letzterStatus})</span>}
          <br />
          <button type="button" className="dex-ui-textbtn" style={{ marginTop: 8 }} onClick={() => { void reload(); }}>{t('Erneut versuchen', 'Try again')}</button>
          {' '}
          <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" style={{ marginTop: 8 }} onClick={() => navigate('usecases')}>{t('Zur Übersicht', 'Back to overview')}</button>
        </span>
      </div>
    );
  }

  if (!uc) {
    return (
      <div className="dex-ui-empty">
        <div className="dex-ui-empty-title">{t('Dieser Use Case ist nicht (mehr) da', 'This use case is not (or no longer) here')}</div>
        <div className="dex-ui-empty-desc">
          {t('Er wurde gelöscht oder archiviert, oder der Link zeigt auf eine Nummer, die es nicht gibt.',
            'It was deleted or archived, or the link points to an id that does not exist.')}
        </div>
        <button type="button" className="dex-ui-empty-action" onClick={() => navigate('usecases')}>
          {t('Zur Übersicht', 'Back to overview')}
        </button>
      </div>
    );
  }

  // Alle Links kommen aus der Liste — und die ist in SharePoint auch direkt beschreibbar.
  // Nur absolute http(s)-Adressen gehen weiter; `javascript:` in einem iframe liefe im Origin
  // dieser SharePoint-Seite, mit der Sitzung der Person (Sicherheits-Review 29.09.2026).
  const start = sichereUrl(uc.ressourcen.deployment);
  const startUngueltig = !!uc.ressourcen.deployment && !start;
  const kannStarten = uc.status === 'Live' && !!start;
  // Liegt das Ziel auf DERSELBEN Herkunft wie diese Seite (ein Link auf die SharePoint-Site
  // selbst), darf es nicht eingebettet werden: `allow-scripts` + `allow-same-origin` zusammen
  // höben die Sandbox dort auf. Es öffnet dann im neuen Tab.
  let gleicheHerkunft = false;
  try {
    const a = document.createElement('a');
    a.href = start;
    gleicheHerkunft = !!start && a.protocol + '//' + a.host === window.location.protocol + '//' + window.location.host;
  } catch { gleicheHerkunft = true; }
  const einbetten = uc.aufrufArt === 'eingebettet' && !gleicheHerkunft;

  const ressourcen: RessourceZeile[] = [
    { key: 'code', url: sichereUrl(uc.ressourcen.sourceCode), titel: t('Source Code', 'Source code'), beschreibung: t('Das Repository zur Demo.', 'The demo repository.'), icon: <LinkIcon size={16} /> },
    { key: 'guide', url: sichereUrl(uc.ressourcen.deploymentGuide), titel: t('Deployment-Guide', 'Deployment guide'), beschreibung: t('Wie du die Demo selbst aufsetzt.', 'How to set the demo up yourself.'), icon: <FileText size={16} /> },
    { key: 'wiki', url: sichereUrl(uc.ressourcen.wiki), titel: t('Wiki / Doku', 'Wiki / docs'), beschreibung: t('Fachliche Beschreibung des Use Cases.', 'The functional description of the use case.'), icon: <Book size={16} /> },
    { key: 'video', url: sichereUrl(uc.ressourcen.video), titel: t('Use-Case-Video', 'Use case video'), beschreibung: t('Aufzeichnung — falls die Live-Demo mal klemmt.', 'Recording — in case the live demo fails.'), icon: <Video size={16} /> },
  ];
  const vorhandene = ressourcen.filter(r => !!r.url);
  // Werte, die in der Liste stehen, aber keine gültige Adresse sind (früher nahm die Pflegeseite alles an):
  // Sie fallen aus der Anzeige — Organizer erfahren wenigstens hier, welche und warum.
  const rohLinks: Array<{ name: string; wert: string }> = [
    { name: 'Source Code', wert: uc.ressourcen.sourceCode }, { name: 'Deployment-Guide', wert: uc.ressourcen.deploymentGuide },
    { name: 'Wiki', wert: uc.ressourcen.wiki }, { name: 'Video', wert: uc.ressourcen.video },
  ];
  const ungueltigeLinks = rohLinks.filter(r => !!r.wert && !sichereUrl(r.wert)).map(r => r.name);

  const starte = (): void => {
    if (!kannStarten) return;
    if (einbetten) { setEingebettetOffen(true); return; }
    // `noopener` ist Pflicht: Ohne das kann die geoeffnete Seite ueber
    // `window.opener` auf diese hier zugreifen.
    window.open(start, '_blank', 'noopener,noreferrer');
  };

  const link = linkZumUseCase(uc.id);
  const kopiereLink = (): void => {
    try {
      // Fehlt in der SharePoint-Handy-App (Permissions Policy reicht Rechte
      // nur abwärts durch) — dann bleibt der Link zum Markieren stehen, statt
      // dass der Knopf stumm nichts tut.
      void navigator.clipboard.writeText(link).then(() => setLinkStatus('kopiert'), () => setLinkStatus('manuell'));
    } catch { setLinkStatus('manuell'); }
  };

  return (
    <div>
      <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" style={{ marginBottom: 12 }} onClick={() => (canGoBack ? goBack() : navigate('usecases'))}>
        <ChevronLeft size={14} /> {t('Zurück', 'Back')}
      </button>

      <div className="dex-ui-page-head">
        <div>
          {uc.bereich && <p className="dex-ui-page-head-meta" style={{ marginBottom: 4 }}>{uc.bereich}</p>}
          <h1 className="dex-ui-page-head-title">{uc.titel}</h1>
        </div>
        <div className="dex-ui-page-head-actions">
          <button type="button" className="dex-ui-textbtn" onClick={kopiereLink} title={t('Link zu diesem Use Case kopieren', 'Copy the link to this use case')}>
            {linkStatus === 'kopiert' ? <><Check size={14} /> {t('Link kopiert', 'Link copied')}</> : <><Copy size={14} /> {t('Link kopieren', 'Copy link')}</>}
          </button>
          {isOrganizer && (
            <button type="button" className="dex-ui-textbtn" onClick={() => navigate('protokoll', uc.id)}>
              {t('Verlauf', 'History')}
            </button>
          )}
          {isOrganizer && (
            <button type="button" className="dex-ui-textbtn" onClick={() => navigate('studio', uc.id)}>
              <Pencil size={14} /> {t('Bearbeiten', 'Edit')}
            </button>
          )}
        </div>
      </div>

      {linkStatus === 'manuell' && (
        <div className="dex-ui-callout dex-ui-callout--neutral dex-ui-callout--sm" style={{ marginBottom: 12 }}>
          <span style={{ flex: 1, minWidth: 0 }}>
            {t('Die Zwischenablage ist hier nicht erreichbar — markier den Link und kopier ihn von Hand:', 'The clipboard is not available here — select the link and copy it by hand:')}
            <input
              type="text"
              readOnly
              className="dex-ui-input dex-ui-input--sm"
              style={{ marginTop: 6 }}
              value={link}
              aria-label={t('Link zu diesem Use Case', 'Link to this use case')}
              onFocus={e => e.target.select()}
            />
          </span>
        </div>
      )}

      {/* Das Kachelbild, groß. Nur wenn eines hochgeladen ist — ohne Bild
          würde ein leerer Kasten nach Ladefehler aussehen. */}
      {sichereUrl(uc.bildUrl) && (
        <div
          role="img"
          aria-label={uc.titel}
          style={{
            width: 96, height: 96, borderRadius: 20, marginBottom: 16,
            border: '1px solid var(--dex-gray-200, #e8e8e8)',
            background: `#fff center/cover no-repeat url("${sichereUrl(uc.bildUrl).replace(/"/g, '%22')}")`,
          }}
        />
      )}

      {/* Die Handlung, derentwegen jemand hier ist. */}
      <div className="dex-ui-section">
        {kannStarten ? (
          <>
            <button type="button" className="btn btn-primary" onClick={starte} style={{ minHeight: 44 }}>
              {einbetten
                ? t('Demo hier öffnen', 'Open demo here')
                : <>{t('Demo starten', 'Start demo')} <ExternalLink size={15} /></>}
            </button>
            <p className="dex-ui-help" style={{ marginTop: 8 }}>
              {einbetten
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
              {uc.status === 'Live' && startUngueltig && t('Der Use Case steht auf „Live", aber der hinterlegte Deployment-Link ist keine gültige Adresse (sie muss mit https:// beginnen). Ein Organizer kann ihn korrigieren.', 'The use case is marked "Live", but the stored deployment link is not a valid address (it must start with https://). An organizer can correct it.')}
              {uc.status === 'Live' && !start && !startUngueltig && t('Der Use Case steht auf „Live", aber es ist kein Deployment-Link hinterlegt. Das ist ein Pflegefehler — ein Organizer kann ihn nachtragen.', 'The use case is marked "Live", but no deployment link is stored. An organizer can add it.')}
            </span>
          </div>
        )}
      </div>

      {/* Eingebettet: erst nach dem Klick. Ein iframe, das beim Seitenaufbau
          mitlaedt, kostet jeden Besucher Ladezeit fuer etwas, das die
          meisten nicht oeffnen. */}
      {eingebettetOffen && kannStarten && einbetten && (
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
          {/* sandbox: Die Demo darf Skripte, Formulare, Pop-ups und Downloads — aber die Seite
              darüber NICHT umleiten (kein `allow-top-navigation`) und nicht an deren Origin
              (kein Zugriff auf SharePoint). `allow-same-origin` heißt hier nur: Die Demo behält
              ihren EIGENEN Origin (Cookies, localStorage — das Passwort-Gate des KI-Arbeitsplatzes
              braucht das). Für Ziele auf dieser Herkunft ist Einbetten oben ausgeschlossen. */}
          <iframe
            src={start}
            title={uc.titel}
            sandbox="allow-scripts allow-same-origin allow-forms allow-popups allow-popups-to-escape-sandbox allow-modals allow-downloads"
            referrerPolicy="no-referrer"
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
              // Die Beschreibung hat in der Pflegeseite gar kein Feld — sie steht nur direkt in
              // SharePoint, wo jede Person mit Bearbeiten-Recht auf der Liste schreiben darf. Sie
              // ist deshalb NICHT vertrauenswürdig: `bereinigeHtml` lässt nur einfache
              // Auszeichnungen (Absatz, Fett, Liste, Link) durch — kein Skript, kein Ereignis-
              // Attribut, kein Bild (gespeichertes XSS, Sicherheits-Review 29.09.2026).
              dangerouslySetInnerHTML={{ __html: bereinigeHtml(uc.beschreibung) }}
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
                {k.wert === 'unbewertet' ? t('noch nicht bewertet', 'not assessed yet') : bewertungText(k.wert, t)}
              </div>
            </div>
          ))}
        </div>
      </div>

      <div className="dex-ui-section">
        <div className="dex-ui-section-title">{t('Ressourcen', 'Resources')}</div>
        {isOrganizer && ungueltigeLinks.length > 0 && (
          <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" style={{ marginBottom: 10 }}>
            <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
            <span className="dex-ui-callout-body">
              {t(`Nicht angezeigt, weil keine gültige Adresse (sie muss mit https:// beginnen): ${ungueltigeLinks.join(', ')}. Im Use Case Studio korrigieren.`,
                `Not shown because it is not a valid address (it must start with https://): ${ungueltigeLinks.join(', ')}. Correct it in the Use Case Studio.`)}
            </span>
          </div>
        )}
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

      {(uc.betreuerEmails.length > 0 || uc.geaendertVon) && (
        <p className="dex-ui-help">
          {uc.betreuerEmails.length > 0 && (
            <>
              {t('Betreut von', 'Maintained by')}{' '}
              {uc.betreuerEmails.map((mail, i) => (
                <React.Fragment key={mail}>
                  {i > 0 && ', '}
                  {/* Nur eine schlichte Adresse wird zum Link: `a@b.de?bcc=…` hängt in `mailto:`
                      Empfänger und Text an. Sonst steht der Name ohne Link da. */}
                  {sichereMail(mail)
                    ? <a href={`mailto:${sichereMail(mail)}`} style={{ color: 'inherit' }}>{uc.betreuerNamen[i] || mail}</a>
                    : <span>{uc.betreuerNamen[i] || mail}</span>}
                </React.Fragment>
              ))}
              {'. '}
            </>
          )}
          {uc.geaendertVon && <>{t('Zuletzt geändert von', 'Last changed by')} {uc.geaendertVon}.</>}
        </p>
      )}
    </div>
  );
}
