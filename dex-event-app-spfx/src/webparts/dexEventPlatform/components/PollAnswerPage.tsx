/**
 * v31.12 — Die Antwortseite der Event-Umfrage.
 *
 * Erreicht über den Deep-Link `#action=umfrage&e=<Event-Nr>[&a=<Index>]` aus
 * der Umfrage-Mail. Steht `a` in der URL, ist diese Antwort schon gewählt,
 * BEVOR die Seite etwas fragt — das ist der ganze Sinn der Kästchen in der
 * Mail: Ein Klick, fertig. Alles Weitere (zweiter Grund, Freitext) ist ein
 * Angebot, keine Bedingung.
 *
 * Deshalb speichert die Seite die Vorauswahl SOFORT und sagt das auch
 * („Deine Antwort ist gespeichert."). Wer danach den Tab schließt, hat
 * geantwortet. Ein „Bitte jetzt noch auf Absenden klicken" wäre genau der
 * zusätzliche Schritt, an dem Rückläufe sterben.
 *
 * Hook-Reihenfolge: Diese Seite hat frühe Returns (Laden, Fehler, keine
 * Umfrage). ALLE Hooks stehen davor — dieselbe Regel wie in
 * `RegistrationPage` (v30.3/v30.4), und seit v30.41 erzwingt ESLint sie.
 */

import * as React from 'react';
import { Check, AlertCircle } from './Icons';
import { cx, ensureDexUiStyles } from './dexUi';
import { deepLinkParams } from '../utils/deepLink';
import { useLanguage } from '../context/LanguageContext';
import { useCurrentUser } from '../context/UserContext';
import { EventService } from '../services/EventService';
import type { PollConfig } from '../services/events/poll';

type Status = 'laedt' | 'bereit' | 'keine' | 'fehler' | 'zu';

const PollAnswerPage: React.FC<{ onLeave?: (_page: 'start' | 'my-events') => void }> = ({ onLeave }) => {
  ensureDexUiStyles();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const { currentUser } = useCurrentUser();
  const t = (de: string, en: string): string => (isDe ? de : en);

  const [status, setStatus] = React.useState<Status>('laedt');
  const [poll, setPoll] = React.useState<PollConfig | null>(null);
  const [gewaehlt, setGewaehlt] = React.useState<string[]>([]);
  const [freitext, setFreitext] = React.useState('');
  const [answerId, setAnswerId] = React.useState(0);
  const [gespeichert, setGespeichert] = React.useState(false);
  const [speichert, setSpeichert] = React.useState(false);
  const [speicherFehler, setSpeicherFehler] = React.useState(false);

  const didRun = React.useRef(false);
  const email = (currentUser.email || '').trim().toLowerCase();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const svc = React.useMemo<EventService | null>(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    return ctx ? new EventService(ctx) : null;
  }, []);

  /**
   * Speichern — wird sowohl vom Sofort-Klick aus der Mail als auch vom
   * „Speichern"-Knopf gerufen. Der Rückgabewert ist die neue Item-Id, damit
   * der zweite Aufruf die Zeile ergänzt statt eine zweite anzulegen.
   */
  const speichern = React.useCallback(async (
    optionen: string[],
    text: string,
    vorhandeneId: number,
    eventNumber: number,
  ): Promise<boolean> => {
    if (!svc || !email) return false;
    const ok = await svc.savePollAnswer(eventNumber, email, vorhandeneId, optionen, text);
    if (!ok) return false;
    if (vorhandeneId === 0) {
      // Die Id der frisch angelegten Zeile nachlesen — ohne sie legt der
      // nächste Klick eine ZWEITE Antwort derselben Person an.
      const nach = await svc.getMyPollAnswer(eventNumber, email);
      if (nach.answer) setAnswerId(nach.answer.id);
    }
    return true;
  }, [svc, email]);

  React.useEffect(() => {
    if (didRun.current) return;
    didRun.current = true;
    (async () => {
      if (!svc) { setStatus('fehler'); return; }
      const p = deepLinkParams();
      const nr = parseInt(p.get('e') || '', 10);
      if (isNaN(nr)) { setStatus('keine'); return; }

      const cfg = await svc.getPoll(nr);
      if (!cfg.ok) { setStatus('fehler'); return; }
      if (!cfg.poll) { setStatus('keine'); return; }
      setPoll(cfg.poll);
      if (!cfg.poll.aktiv) { setStatus('zu'); return; }

      // Die eigene Antwort zuerst lesen. Ein Lesefehler ist hier KEINE leere
      // Antwort: Würden wir mit 0 weitermachen, legt der Klick eine zweite
      // Zeile an und die Person hätte zwei Stimmen.
      const mein = await svc.getMyPollAnswer(nr, email);
      if (!mein.ok) { setStatus('fehler'); return; }

      let start = mein.answer ? mein.answer.optionen : [];
      const id = mein.answer ? mein.answer.id : 0;
      setAnswerId(id);
      setFreitext(mein.answer ? mein.answer.freitext : '');

      // Die Vorauswahl aus der Mail — sofort speichern.
      const aRaw = p.get('a');
      const aIdx = aRaw === null ? -1 : parseInt(aRaw, 10);
      const gewaehlteOption = (!isNaN(aIdx) && aIdx >= 0 && aIdx < cfg.poll.optionen.length)
        ? cfg.poll.optionen[aIdx]
        : '';
      if (gewaehlteOption) {
        start = cfg.poll.mehrfach
          ? (start.indexOf(gewaehlteOption) >= 0 ? start : start.concat([gewaehlteOption]))
          : [gewaehlteOption];
        setGewaehlt(start);
        setStatus('bereit');
        const ok = await speichern(start, mein.answer ? mein.answer.freitext : '', id, nr);
        if (ok) setGespeichert(true); else setSpeicherFehler(true);
        return;
      }
      setGewaehlt(start);
      setGespeichert(!!mein.answer);
      setStatus('bereit');
    })().catch(() => setStatus('fehler'));
  }, [svc, email, speichern]);

  const umschalten = React.useCallback((opt: string): void => {
    setGespeichert(false);
    setSpeicherFehler(false);
    setGewaehlt(prev => {
      if (!poll) return prev;
      if (!poll.mehrfach) return prev.indexOf(opt) >= 0 ? [] : [opt];
      return prev.indexOf(opt) >= 0 ? prev.filter(o => o !== opt) : prev.concat([opt]);
    });
  }, [poll]);

  const jetztSpeichern = React.useCallback(async (): Promise<void> => {
    if (!poll) return;
    setSpeichert(true);
    setSpeicherFehler(false);
    const ok = await speichern(gewaehlt, freitext, answerId, poll.eventNumber);
    setSpeichert(false);
    if (ok) setGespeichert(true); else setSpeicherFehler(true);
  }, [poll, gewaehlt, freitext, answerId, speichern]);

  // ---- ab hier erst die Anzeige; alle Hooks stehen oben ----------------

  const verlassen = (seite: 'start' | 'my-events'): void => { if (onLeave) onLeave(seite); };

  const rahmen = (inhalt: React.ReactNode): React.ReactElement => (
    <div style={{ maxWidth: 620, margin: '0 auto', padding: '24px 16px' }}>{inhalt}</div>
  );

  if (status === 'laedt') {
    return rahmen(
      <div className="dex-ui-empty">
        <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
        <div className="dex-ui-empty-desc" style={{ marginTop: 12 }}>
          {t('Umfrage wird geladen …', 'Loading poll …')}
        </div>
      </div>,
    );
  }

  if (status === 'fehler') {
    return rahmen(
      <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
        <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
        <span className="dex-ui-callout-body">
          <strong>{t('Die Umfrage konnte nicht geladen werden.', 'The poll could not be loaded.')}</strong>
          <br />
          {t('Das ist ein Ladefehler, keine geschlossene Umfrage — deine Antwort ist nicht verloren. Lade die Seite bitte neu.',
            'This is a loading error, not a closed poll — your answer is not lost. Please reload the page.')}
        </span>
      </div>,
    );
  }

  if (status === 'keine') {
    return rahmen(
      <div className="dex-ui-empty">
        <div className="dex-ui-empty-title">{t('Keine Umfrage zu diesem Event', 'No poll for this event')}</div>
        <div className="dex-ui-empty-desc">
          {t('Der Link gehört zu einem Event, zu dem es keine Umfrage (mehr) gibt.',
            'This link belongs to an event that has no poll (any more).')}
        </div>
        <button type="button" className="btn btn-secondary dex-ui-empty-action" onClick={() => verlassen('start')}>
          {t('Zur Startseite', 'Go to start page')}
        </button>
      </div>,
    );
  }

  if (status === 'zu' || !poll) {
    return rahmen(
      <div className="dex-ui-empty">
        <div className="dex-ui-empty-title">{t('Die Umfrage ist beendet', 'The poll is closed')}</div>
        <div className="dex-ui-empty-desc">
          {t('Danke für dein Interesse — hier wird nichts mehr erfasst.',
            'Thanks for your interest — nothing is recorded here any more.')}
        </div>
        <button type="button" className="btn btn-secondary dex-ui-empty-action" onClick={() => verlassen('start')}>
          {t('Zur Startseite', 'Go to start page')}
        </button>
      </div>,
    );
  }

  return rahmen(
    <>
      <div className="dex-ui-page-head">
        <div>
          <h1 className="dex-ui-page-head-title">{t('Kurze Rückmeldung', 'A quick word')}</h1>
          <p className="dex-ui-page-head-meta">{poll.frage}</p>
        </div>
      </div>

      {gespeichert && !speicherFehler && (
        <div className="dex-ui-callout dex-ui-callout--success" style={{ marginBottom: 14 }}>
          <span className="dex-ui-callout-icon"><Check size={16} /></span>
          <span className="dex-ui-callout-body">
            <strong>{t('Deine Antwort ist gespeichert.', 'Your answer is saved.')}</strong>
            <br />
            {poll.mehrfach
              ? t('Du kannst weitere Gründe anhaken oder etwas dazuschreiben — musst du aber nicht.',
                'You can tick further reasons or add a comment — but you do not have to.')
              : t('Du kannst deine Auswahl noch ändern oder etwas dazuschreiben — musst du aber nicht.',
                'You can still change your choice or add a comment — but you do not have to.')}
          </span>
        </div>
      )}

      {speicherFehler && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert" style={{ marginBottom: 14 }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span className="dex-ui-callout-body">
            {t('Deine Antwort konnte nicht gespeichert werden. Bitte versuch es gleich noch einmal.',
              'Your answer could not be saved. Please try again shortly.')}
          </span>
        </div>
      )}

      <div className="dex-ui-section">
        <div className="dex-ui-section-title">
          {poll.mehrfach
            ? t('Was trifft zu? (Mehrfachnennung möglich)', 'What applies? (more than one possible)')
            : t('Was trifft zu?', 'What applies?')}
        </div>
        <div className="dex-ui-stack" style={{ gap: 8 }}>
          {poll.optionen.map((opt, i) => {
            const aktiv = gewaehlt.indexOf(opt) >= 0;
            return (
              <button
                key={i}
                type="button"
                className={cx('dex-ui-choice', poll.mehrfach && 'dex-ui-choice--multi', aktiv && 'is-active')}
                aria-pressed={aktiv}
                onClick={() => umschalten(opt)}
              >
                <span className="dex-ui-choice-body">
                  <span className="dex-ui-choice-label">{opt}</span>
                </span>
                <span className="dex-ui-choice-check">{aktiv && <Check size={12} />}</span>
              </button>
            );
          })}
        </div>
      </div>

      <div className="dex-ui-section">
        <label className="dex-ui-label" htmlFor="poll-freitext">
          {t('Möchtest du etwas dazuschreiben?', 'Would you like to add anything?')}
          <span className="dex-ui-label-optional">{t(' (optional)', ' (optional)')}</span>
        </label>
        <textarea
          id="poll-freitext"
          className="dex-ui-textarea"
          rows={4}
          value={freitext}
          onChange={e => { setFreitext(e.target.value); setGespeichert(false); }}
          placeholder={t('Zum Beispiel: Was hätte dir geholfen, dabei zu sein?',
            'For example: what would have helped you to be there?')}
        />
      </div>

      <div className="dex-ui-inline" style={{ gap: 10, marginTop: 4 }}>
        <button
          type="button"
          className="btn btn-primary"
          disabled={speichert || (gewaehlt.length === 0 && !freitext.trim())}
          onClick={() => { void jetztSpeichern(); }}
        >
          {speichert ? t('Wird gespeichert …', 'Saving …') : t('Antwort speichern', 'Save answer')}
        </button>
        <button type="button" className="btn btn-secondary" onClick={() => verlassen('start')}>
          {t('Fertig', 'Done')}
        </button>
      </div>

      {/* Was DEX zeigt und was SharePoint trotzdem weiß — das gehört hierhin,
          nicht in eine Datenschutzseite, die niemand öffnet. */}
      <p className="dex-ui-muted" style={{ fontSize: '0.78rem', marginTop: 18, lineHeight: 1.5 }}>
        {poll.anonym
          ? t('Die Auswertung ist anonym: Den Organisierenden zeigt DEX nur, wie oft etwas gewählt wurde, und die Freitexte ohne Absender. Gespeichert wird zusätzlich, DASS du geantwortet hast — sonst würdest du weiter erinnert.',
            'The evaluation is anonymous: DEX only shows the organisers how often each answer was picked, plus the free-text comments without a sender. It does record THAT you answered — otherwise you would keep getting reminders.')
          : t('Deine Antwort ist für die Organisierenden dieses Events mit deinem Namen sichtbar.',
            'Your answer is visible to the organisers of this event together with your name.')}
      </p>
    </>,
  );
};

export default PollAnswerPage;
