/**
 * v1.1 — Der Startbildschirm.
 *
 * Nutzer-Ansage 11.09.2026: „die Landing Page die auch das DEX Orb Logo
 * nimmt, also mit Startscreen, dann eine Kachelübersicht und dann kann ich
 * meine bestehenden Use Cases öffnen." Bis v1.0.1 fiel die App sofort auf die
 * Kachelwand — ohne Begrüßung, ohne Gesicht, und bei leerer Liste ohne jede
 * Aussage darüber, was sie eigentlich ist.
 *
 * Aufbau 1:1 wie `LandingPage` in DEX: Orb, Begrüßung, ein großer Knopf,
 * darunter die Einstiegskarten. Der Orb ist dieselbe Komponente
 * (`DexLogo`, Canvas, keine Abhängigkeiten) — das ist der Wiedererkennungs-
 * wert, den die Nutzer-Ansage meint.
 *
 * Der Knopf ist NICHT gesperrt, solange geladen wird. Wer auf einen toten
 * Knopf klickt, hält die App für kaputt; die Kachelwand dahinter zeigt ihren
 * eigenen Ladezustand und kann das besser.
 */

import * as React from 'react';
import DexLogo from './DexLogo';
import { ChevronRight, Settings, Users } from './Icons';
import { ensureDexUiStyles, cx } from './dexUi';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { resolveMyDisplayName } from '../utils/displayName';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useUseCases } from '../context/UseCaseContext';
import { APP_NAME, APP_SUBTITLE_DE, APP_SUBTITLE_EN } from '../constants';

export default function LandingPage(): React.ReactElement {
  ensureDexUiStyles();
  const { t, isDe } = useLanguage();
  const { navigate } = useNavigation();
  const { isAdmin, isKurator } = useRoles();
  const { useCases, ladeStatus } = useUseCases();

  // Der Vorname fuer die Begruessung. `resolveMyDisplayName` raeumt dabei die
  // Claim-Schreibweise weg („i:0#.f|membership|…"), die SharePoint in manchen
  // Mandanten als displayName liefert — in DEX war genau das der Grund, warum
  // die Begruessung einmal eine Anmelde-Kennung anredete.
  const [vorname, setVorname] = React.useState('');
  React.useEffect(() => {
    let weg = false;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__aiucSpfxContext as WebPartContext | undefined;
    if (!ctx) return undefined;
    resolveMyDisplayName(ctx)
      .then(n => { if (!weg) setVorname((n || '').trim().split(' ')[0] || ''); })
      .catch(() => { /* ohne Namen gruesst die Seite eben ohne Namen */ });
    return () => { weg = true; };
  }, []);

  const stunde = new Date().getHours();
  const gruss = stunde < 11
    ? t('Guten Morgen', 'Good morning')
    : stunde < 18 ? t('Hallo', 'Hello') : t('Guten Abend', 'Good evening');
  const live = useCases.filter(u => u.status === 'Live').length;
  const gesamt = useCases.length;

  return (
    <div className="landing">
      <div className="landing__hero">
        <div className="landing__card">
          <div className="landing__orb">
            <DexLogo title={APP_NAME} motion="oscillate" style={{ width: '100%' }} />
          </div>

          <div className="landing__text">
            <h1>{gruss}{vorname ? <>, <strong>{vorname}</strong></> : ''}.</h1>
            <p>
              {isDe
                ? <>Willkommen bei der <strong>AI Use Case Platform</strong> — unserem {APP_SUBTITLE_DE}. Alle Agent-Demos an einer Stelle: Kachel anklicken, Kurzbeschreibung lesen, Demo starten.</>
                : <>Welcome to the <strong>AI Use Case Platform</strong> — our {APP_SUBTITLE_EN}. All agent demos in one place: click a tile, read the summary, start the demo.</>}
            </p>
          </div>

          <button type="button" className="btn btn-primary" onClick={() => navigate('start')}>
            {t('Use Cases ansehen', 'Browse use cases')} <ChevronRight size={16} />
          </button>

          {/* Was dahinter liegt, in einem Satz — und ehrlich, wenn es gerade
              nicht lesbar war. „0 Demos" bei einem Lesefehler wäre eine
              Aussage über die Plattform, die niemand geprüft hat. */}
          <div className="dex-ui-muted" style={{ fontSize: '0.82rem', textAlign: 'center' }}>
            {ladeStatus === 'laedt' && t('Die Übersicht wird geladen …', 'Loading the overview …')}
            {ladeStatus === 'fehler' && t('Wie viele Demos es gerade gibt, konnte nicht gelesen werden — die Übersicht sagt dir den Grund.',
              'How many demos there are right now could not be read — the overview tells you why.')}
            {ladeStatus === 'ok' && (gesamt === 0
              ? t('Noch keine Demos hinterlegt.', 'No demos on file yet.')
              : t(`${gesamt} Use Cases · davon ${live} sofort aufrufbar`, `${gesamt} use cases · ${live} callable right away`))}
          </div>
        </div>
      </div>

      {/* Einstiegskarten wie auf der DEX-Startseite: der eine Weg für alle,
          die Pflege-Wege nur für die, die sie benutzen dürfen. */}
      <div className={cx('start-grid', (isKurator || isAdmin) && 'start-grid--with-admin')} style={{ paddingBottom: 32 }}>
        <button type="button" className="card start-card" onClick={() => navigate('start')}>
          <span className="start-card__icon"><ChevronRight size={40} /></span>
          <h2>{t('Use Cases', 'Use cases')}</h2>
          <p>{t('Die Kachelübersicht aller Demos — suchen, filtern, öffnen.',
            'The tile overview of all demos — search, filter, open.')}</p>
        </button>

        {isKurator && (
          <button type="button" className="card start-card" onClick={() => navigate('verwaltung')}>
            <span className="start-card__icon"><Settings size={40} /></span>
            <h2>{t('Use Cases pflegen', 'Manage use cases')}</h2>
            <p>{t('Neue Demos anlegen, Beschreibungen und Links pflegen, Reihenfolge ändern.',
              'Add demos, maintain descriptions and links, change the order.')}</p>
          </button>
        )}

        {isAdmin && (
          <button type="button" className="card start-card" onClick={() => navigate('rollen')}>
            <span className="start-card__icon"><Users size={40} /></span>
            <h2>{t('Rollen', 'Roles')}</h2>
            <p>{t('Wer darf pflegen, wer darf verwalten — und die Rechte dazu prüfen.',
              'Who may curate, who may administrate — and check the permissions.')}</p>
          </button>
        )}
      </div>
    </div>
  );
}
