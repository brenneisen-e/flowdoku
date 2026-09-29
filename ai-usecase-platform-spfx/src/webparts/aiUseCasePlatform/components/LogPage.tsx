/**
 * Das Änderungsprotokoll — wer hat wann was an den Use Cases getan.
 *
 * Die Liste `AIUC_Log` wurde bis v1.2 geschrieben und nirgends angezeigt: Wer
 * wissen wollte, warum eine Kachel plötzlich auf „Archiviert" stand, musste in
 * SharePoint die Listenansicht öffnen. Ein Protokoll, das nur der Admin mit
 * Site-Zugriff lesen kann, ist ein Beleg, aber keine Hilfe im Alltag.
 *
 * Sichtbar für Use Case Organizer und Admins. Die Einträge tragen E-Mail-Adressen von
 * Personen; für normale Nutzer gibt es keinen Grund, sie zu lesen.
 *
 * Wie überall in dieser App drei Zustände: `laedt`, `ok`, `fehler`. Ein
 * Lesefehler ist kein leeres Protokoll — „noch nichts passiert" und „darf ich
 * nicht sehen" dürfen nicht gleich aussehen.
 */

import * as React from 'react';
import { cx, ensureDexUiStyles } from './dexUi';
import { ChevronLeft, RefreshCw } from './Icons';
import { useRoles } from '../context/RoleContext';
import { useUseCases } from '../context/UseCaseContext';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { LogEintrag } from '../types';

type LogStatus = 'laedt' | 'ok' | 'fehler';

/** Die Aktionen, die die App schreibt — mit lesbarer Bezeichnung. */
function aktionLabel(aktion: string, isDe: boolean): string {
  switch (aktion) {
    case 'angelegt': return isDe ? 'Angelegt' : 'Created';
    case 'geaendert': return isDe ? 'Geändert' : 'Changed';
    case 'geloescht': return isDe ? 'Gelöscht' : 'Deleted';
    case 'loeschen-fehlgeschlagen': return isDe ? 'Löschen fehlgeschlagen' : 'Delete failed';
    case 'erstbefuellung': return isDe ? 'Erstbefüllung' : 'Initial fill';
    // Rollenänderungen (v1.3): Sie gehören keinem Use Case, stehen aber im selben
    // Protokoll — wer wann jemanden zum Admin gemacht hat, muss auffindbar sein.
    case 'rolle-vergeben': return isDe ? 'Rolle vergeben' : 'Role assigned';
    case 'rolle-geaendert': return isDe ? 'Rolle geändert' : 'Role changed';
    case 'rolle-entfernt': return isDe ? 'Rolle entfernt' : 'Role removed';
    // Eine Aktion, die diese Version nicht kennt (aus einer späteren oder von
    // Hand geschrieben): roh anzeigen, nie verschlucken.
    default: return aktion || '—';
  }
}

function aktionFarbe(aktion: string): string {
  if (aktion === 'angelegt' || aktion === 'erstbefuellung' || aktion === 'rolle-vergeben') return 'dex-ui-pill--green';
  if (aktion === 'geloescht' || aktion === 'loeschen-fehlgeschlagen' || aktion === 'rolle-entfernt') return 'dex-ui-pill--red';
  if (aktion === 'geaendert' || aktion === 'rolle-geaendert') return 'dex-ui-pill--blue';
  return 'dex-ui-pill--gray';
}

/** Datum und Uhrzeit in Berliner Zeit — die Zeit, in der die Leute arbeiten. */
function zeit(iso: string, isDe: boolean): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  try {
    return d.toLocaleString(isDe ? 'de-DE' : 'en-GB', {
      day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit',
      timeZone: 'Europe/Berlin',
    });
  } catch {
    return iso;
  }
}

export default function LogPage(props: { useCaseId?: number }): React.ReactElement {
  ensureDexUiStyles();

  const { isOrganizer, service } = useRoles();
  const { useCases } = useUseCases();
  const { t, isDe } = useLanguage();
  const { navigate, goBack, canGoBack } = useNavigation();

  // ALLE Hooks stehen vor dem ersten frühen Return (React #300, DEX v30.3).
  const [eintraege, setEintraege] = React.useState<LogEintrag[]>([]);
  const [status, setStatus] = React.useState<LogStatus>('laedt');
  const [httpStatus, setHttpStatus] = React.useState(0);
  const [fehlertext, setFehlertext] = React.useState('');
  const [nurDieser, setNurDieser] = React.useState(!!props.useCaseId);

  const laden = React.useCallback(async (): Promise<void> => {
    setStatus('laedt');
    const rows = await service.getLog();
    setHttpStatus(service.lastLogReadStatus);
    setFehlertext(service.lastLogReadError);
    if (rows === null) {
      // Bestehenden Stand STEHEN LASSEN: ein Lesefehler ist keine Aussage
      // über die Daten. Die Liste bleibt, was sie war; das Banner sagt warum.
      setStatus('fehler');
      return;
    }
    setEintraege(rows);
    setStatus('ok');
  }, [service]);

  React.useEffect(() => {
    if (isOrganizer) void laden();
  }, [isOrganizer, laden]);

  if (!isOrganizer) {
    return (
      <div className="dex-ui-empty">
        <div className="dex-ui-empty-title">{t('Nur für Organizer', 'Organizers only')}</div>
        <div className="dex-ui-empty-desc">
          {t('Das Protokoll enthält die Adressen der Personen, die Use Cases pflegen. Es ist für Use Case Organizer und Admins sichtbar.',
            'The log contains the addresses of the people who maintain use cases. It is visible to Use Case Organizers and admins.')}
        </div>
        <button type="button" className="dex-ui-empty-action" onClick={() => navigate('start')}>
          {t('Zur Übersicht', 'Back to overview')}
        </button>
      </div>
    );
  }

  const titelVon = (id: number): string => {
    if (!id) return t('Plattform', 'Platform');
    const uc = useCases.filter(u => u.id === id)[0];
    // Gelöscht oder nie gesehen: die Nummer statt einer erfundenen Bezeichnung.
    return uc ? uc.titel : `#${id}`;
  };

  /**
   * Die Aktionen, bei denen das Detailfeld den TITEL trägt (`angelegt`,
   * `geloescht`, `loeschen-fehlgeschlagen`). Bei einem gelöschten Use Case ist
   * das der einzige Ort, an dem der Name noch steht — die Zeile zeigt ihn dann
   * als Titel, statt nur „#12".
   */
  const detailIstTitel = (aktion: string): boolean =>
    aktion === 'angelegt' || aktion === 'geloescht' || aktion === 'loeschen-fehlgeschlagen';

  const zeilenTitel = (e: LogEintrag): string => {
    const bekannt = e.useCaseId ? useCases.filter(u => u.id === e.useCaseId)[0] : undefined;
    if (bekannt) return bekannt.titel;
    if (e.useCaseId && detailIstTitel(e.aktion) && e.detail) return e.detail;
    return titelVon(e.useCaseId);
  };

  const gefiltert = nurDieser && props.useCaseId
    ? eintraege.filter(e => e.useCaseId === props.useCaseId)
    : eintraege;

  return (
    <div>
      <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" style={{ marginBottom: 12 }} onClick={() => (canGoBack ? goBack() : navigate('studio'))}>
        <ChevronLeft size={14} /> {t('Zurück', 'Back')}
      </button>

      <div className="dex-ui-page-head">
        <div>
          <h1 className="dex-ui-page-head-title">{t('Protokoll', 'Log')}</h1>
          <p className="dex-ui-page-head-meta">
            {status === 'ok'
              ? `${gefiltert.length} ${t('Einträge', 'entries')}${eintraege.length >= 500 ? ` · ${t('die letzten 500', 'the latest 500')}` : ''}`
              : t('Wer hat wann was geändert', 'Who changed what and when')}
          </p>
        </div>
        <div className="dex-ui-page-head-actions">
          <button type="button" className="dex-ui-textbtn" disabled={status === 'laedt'} onClick={() => { void laden(); }}>
            <RefreshCw size={14} /> {t('Neu laden', 'Reload')}
          </button>
        </div>
      </div>

      {props.useCaseId ? (
        <div className="dex-ui-toolbar" style={{ marginBottom: 16 }}>
          <button
            type="button"
            className={cx('dex-ui-chip', nurDieser && 'is-active')}
            aria-pressed={nurDieser}
            onClick={() => setNurDieser(v => !v)}
          >
            {t('Nur', 'Only')} „{titelVon(props.useCaseId)}“
          </button>
        </div>
      ) : null}

      {status === 'fehler' && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert" style={{ marginBottom: 16 }}>
          <span>
            <strong>{t('Das Protokoll konnte nicht gelesen werden.', 'The log could not be read.')}</strong>
            <br />
            {httpStatus === 403
              ? t('Dir fehlt das Leserecht auf der Protokoll-Liste. Ein Admin kann es nachsetzen (Rolle neu vergeben).',
                'You lack read access to the log list. An admin can add it (assign the role again).')
              : httpStatus === 404
                ? t('Die Protokoll-Liste gibt es (noch) nicht. Lade die Seite neu — die App legt sie beim Start selbst an.',
                  'The log list does not exist (yet). Reload the page — the app creates it on start.')
                : t('Das ist ein Lesefehler, kein leeres Protokoll. Versuch es gleich noch einmal.',
                  'This is a read error, not an empty log. Please try again shortly.')}
            {httpStatus > 0 && <span className="dex-ui-muted"> (HTTP {httpStatus})</span>}
            {fehlertext && (
              <>
                <br />
                <span className="dex-ui-muted" style={{ fontSize: '0.78rem' }}>{fehlertext}</span>
              </>
            )}
          </span>
        </div>
      )}

      {status === 'laedt' && eintraege.length === 0 && (
        <div className="dex-ui-empty">
          <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
          <div className="dex-ui-empty-desc" style={{ marginTop: 12 }}>{t('Protokoll wird geladen …', 'Loading log …')}</div>
        </div>
      )}

      {status === 'ok' && gefiltert.length === 0 && (
        <div className="dex-ui-empty">
          <div className="dex-ui-empty-title">{t('Noch nichts protokolliert', 'Nothing logged yet')}</div>
          <div className="dex-ui-empty-desc">
            {t('Sobald jemand einen Use Case anlegt, ändert oder löscht, steht es hier.',
              'As soon as someone creates, changes or deletes a use case, it shows up here.')}
          </div>
        </div>
      )}

      {gefiltert.length > 0 && (
        <div className="dex-ui-stack">
          {gefiltert.map(e => (
            <div key={e.id} className="dex-ui-row dex-ui-row--bordered dex-ui-row--static">
              <span className="dex-ui-row-main">
                <span className="dex-ui-row-title dex-ui-row-title--wrap">
                  <span className={cx('dex-ui-pill', 'dex-ui-pill--sm', aktionFarbe(e.aktion))}>{aktionLabel(e.aktion, isDe)}</span>
                  {zeilenTitel(e)}
                </span>
                {e.detail && !detailIstTitel(e.aktion) && (
                  <span className="dex-ui-row-sub">{e.detail}</span>
                )}
                <span className="dex-ui-row-sub">
                  {e.wer || t('unbekannt', 'unknown')}{e.wann ? ` · ${zeit(e.wann, isDe)}` : ''}
                </span>
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
