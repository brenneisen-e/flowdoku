/**
 * v31.23 — Die Feedback-Seite für Organizer.
 *
 * Erreicht über den Deep-Link `#action=feedback&e=<Event-Id>` aus der
 * Nachbereitungs-Mail („Danke für euer Event!"). Beim Absenden geht eine
 * Mail an das DEX-Team (`utils/dexFeedback`).
 *
 * ## Warum hier NICHT sofort gespeichert wird
 *
 * Die Umfrage-Seite (`PollAnswerPage`, v31.12) speichert die Vorauswahl aus
 * der Mail sofort, noch bevor sie etwas fragt — dort ist jeder zusätzliche
 * Schritt ein verlorener Rücklauf. Hier ist es umgekehrt: Der Knopf in der
 * Mail trägt keine Antwort, es gibt nichts vorzumerken, und eine halb
 * ausgefüllte Rückmeldung an ein Team-Postfach ist schlechter als keine.
 * Deshalb ein klares Absenden — und danach ist es raus.
 *
 * ## Hook-Reihenfolge
 *
 * Diese Seite hat frühe Returns (lädt, Event unbekannt, abgeschickt). ALLE
 * Hooks stehen davor — dieselbe Regel wie in `RegistrationPage` (v30.3/v30.4
 * White Screen) und `PollAnswerPage`; seit v30.41 erzwingt ESLint sie.
 */

import * as React from 'react';
import { AlertCircle, Check, Send } from './Icons';
import { cx, ensureDexUiStyles } from './dexUi';
import { deepLinkParams } from '../utils/deepLink';
import { useLanguage } from '../context/LanguageContext';
import { useCurrentUser } from '../context/UserContext';
import { useEvents } from '../context/EventContext';
import { EventService } from '../services/EventService';
import { wrapTemplate } from '../services/EmailTemplates';
import {
  DEX_FEEDBACK_TO, FEEDBACK_BESSER, FEEDBACK_GUT, FeedbackOption,
  buildFeedbackMailInner, feedbackHatInhalt, feedbackMailSubject,
} from '../utils/dexFeedback';

const FeedbackPage: React.FC<{ onLeave?: (_page: 'start' | 'my-events') => void }> = ({ onLeave }) => {
  ensureDexUiStyles();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const t = (de: string, en: string): string => (isDe ? de : en);
  const { currentUser } = useCurrentUser();
  const { events, isEventsLoading } = useEvents();

  const [gut, setGut] = React.useState<string[]>([]);
  const [besser, setBesser] = React.useState<string[]>([]);
  const [freitextGut, setFreitextGut] = React.useState('');
  const [freitextBesser, setFreitextBesser] = React.useState('');
  const [sendet, setSendet] = React.useState(false);
  const [fertig, setFertig] = React.useState(false);
  const [fehler, setFehler] = React.useState('');

  const eventId = React.useMemo(() => {
    try { return (deepLinkParams().get('e') || '').trim(); } catch { return ''; }
  }, []);
  const ev = React.useMemo(
    () => events.filter(e => String(e.id) === eventId)[0] || null,
    [events, eventId],
  );

  const svc = React.useMemo<EventService | null>(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    return ctx ? new EventService(ctx) : null;
  }, []);

  const toggle = (setter: React.Dispatch<React.SetStateAction<string[]>>, id: string): void => {
    setter(prev => (prev.indexOf(id) >= 0 ? prev.filter(x => x !== id) : prev.concat(id)));
  };

  const eingabe = { gut, besser, freitextGut, freitextBesser };
  const hatInhalt = feedbackHatInhalt(eingabe);

  const absenden = async (): Promise<void> => {
    if (!ev || !svc || sendet || !hatInhalt) return;
    setSendet(true);
    setFehler('');
    try {
      const inner = buildFeedbackMailInner(eingabe, { title: ev.title, id: String(ev.id), startDate: ev.startDate }, {
        name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim(), email: currentUser.email || '',
      });
      const body = wrapTemplate('#86bc25', 'DEX-Feedback', ev.title, inner);
      // v31.23: Der Rückgabewert ENTSCHEIDET. `queueEmail` wirft nicht — wer
      // nur den catch prüft, meldet „Danke!", während die Zeile nie
      // geschrieben wurde (die Lehre aus v30.67).
      const ok = await svc.queueEmail(
        feedbackMailSubject(ev.title),
        DEX_FEEDBACK_TO, 'DEX Team', body,
        'DexFeedback', ev.title, String(ev.id),
        // Der Organizer bekommt sich selbst in Kopie — sonst hat er keinen
        // Beleg dafür, was er gesagt hat, und das Team kann per „Antworten
        // an alle" direkt zurückfragen.
        (currentUser.email || '') || undefined,
      );
      if (!ok) {
        setFehler(t(
          'Das Feedback konnte gerade nicht abgeschickt werden. Bitte versuch es in einem Moment noch einmal — deine Eingaben bleiben stehen.',
          'The feedback could not be sent right now. Please try again in a moment — your input is preserved.',
        ));
        setSendet(false);
        return;
      }
      setFertig(true);
    } catch {
      setFehler(t(
        'Das Feedback konnte gerade nicht abgeschickt werden. Bitte versuch es in einem Moment noch einmal — deine Eingaben bleiben stehen.',
        'The feedback could not be sent right now. Please try again in a moment — your input is preserved.',
      ));
      setSendet(false);
    }
  };

  /** Eine Antwortmöglichkeit als anklickbare Zeile (Leitfaden 4). */
  const zeile = (o: FeedbackOption, aktiv: boolean, onClick: () => void): React.ReactElement => (
    <label
      key={o.id}
      className={cx('dex-ui-row', 'dex-ui-rowbtn', aktiv && 'is-active')}
      style={{ cursor: 'pointer', alignItems: 'center', gap: 10 }}
    >
      <input
        type="checkbox"
        className="dex-ui-checkbox"
        checked={aktiv}
        onChange={onClick}
        style={{ flexShrink: 0 }}
      />
      <span className="dex-ui-row-main">
        <span className="dex-ui-row-title" style={{ fontWeight: aktiv ? 600 : 400 }}>
          {isDe ? o.de : o.en}
        </span>
      </span>
    </label>
  );

  // ---- ab hier frühe Returns; oberhalb steht KEIN Hook mehr ----

  if (!eventId) {
    return (
      <div className="page-container" style={{ maxWidth: 720 }}>
        <div className="dex-ui-callout dex-ui-callout--warn" role="status">
          <AlertCircle size={16} />
          <span>{t('Dieser Link ist unvollständig — es fehlt das Event. Bitte den Knopf in der Mail noch einmal anklicken.',
            'This link is incomplete — the event is missing. Please click the button in the email again.')}</span>
        </div>
      </div>
    );
  }

  if (isEventsLoading && !ev) {
    return (
      <div className="page-container" style={{ maxWidth: 720 }}>
        <div className="dex-ui-empty" role="status" aria-live="polite">
          <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
          <div className="dex-ui-empty-title" style={{ marginTop: 14 }}>{t('Einen Moment …', 'One moment …')}</div>
        </div>
      </div>
    );
  }

  if (!ev) {
    return (
      <div className="page-container" style={{ maxWidth: 720 }}>
        <div className="dex-ui-callout dex-ui-callout--warn" role="status">
          <AlertCircle size={16} />
          {/* Nicht „Event gibt es nicht": Am wahrscheinlichsten ist, dass die
              Person das Event nicht (mehr) sehen darf oder es gelöscht wurde.
              Eine Fehlermeldung, die den falschen Grund nennt, schickt den
              Leser auf die falsche Suche (CLAUDE.md: Guard-Meldungen). */}
          <span>{t('Dieses Event ist für dich gerade nicht sichtbar — womöglich wurde es gelöscht, oder du bist nicht mehr als Organizer eingetragen. Dein Feedback nehmen wir trotzdem gern: schreib uns einfach an ',
            'This event is not visible to you right now — it may have been deleted, or you are no longer listed as an organizer. We would still like your feedback: just write to us at ')}
            <a href={`mailto:${DEX_FEEDBACK_TO}`}>{DEX_FEEDBACK_TO}</a>.</span>
        </div>
      </div>
    );
  }

  if (fertig) {
    return (
      <div className="page-container" style={{ maxWidth: 720 }}>
        <div className="dex-ui-card" style={{ padding: 28, textAlign: 'center' }}>
          <div style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><Check size={40} /></div>
          <h2 style={{ margin: '10px 0 6px' }}>{t('Danke!', 'Thank you!')}</h2>
          <p className="dex-ui-muted" style={{ margin: '0 0 18px' }}>
            {t('Dein Feedback ist beim DEX-Team. Eine Kopie liegt in deinem Postfach — falls du etwas ergänzen möchtest, antworte einfach darauf.',
              'Your feedback is with the DEX team. A copy is in your mailbox — if you want to add anything, just reply to it.')}
          </p>
          <div className="dex-ui-inline" style={{ justifyContent: 'center', gap: 8 }}>
            <button type="button" className="btn btn-secondary" onClick={() => onLeave && onLeave('my-events')}>
              {t('Zu meinen Events', 'To my events')}
            </button>
            <button type="button" className="btn btn-primary" onClick={() => onLeave && onLeave('start')}>
              {t('Zur Startseite', 'To the start page')}
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="page-container" style={{ maxWidth: 720 }}>
      <div className="dex-ui-card" style={{ padding: 28 }}>
        <h2 className="dex-step-head-title" style={{ margin: '0 0 4px' }}>
          {t('Wie lief es mit DEX?', 'How did it go with DEX?')}
        </h2>
        <p className="dex-ui-muted" style={{ margin: '0 0 4px' }}>
          {t('Dein Event ', 'Your event ')}<strong>{ev.title}</strong>{t(' ist vorbei.', ' is over.')}{' '}
          {t('Zwei Minuten Rückmeldung helfen uns, die App besser zu machen. Hak an, was zutrifft — schreiben musst du nichts.',
            'Two minutes of feedback help us make the app better. Tick what applies — you do not have to write anything.')}
        </p>
        <p className="dex-ui-muted" style={{ margin: '0 0 8px', fontSize: '0.78rem' }}>
          {t('Geht an das DEX-Team (', 'Goes to the DEX team (')}{DEX_FEEDBACK_TO}
          {t('), mit dir in Kopie.', '), with a copy to you.')}
        </p>

        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{t('Was lief gut?', 'What went well?')}</div>
          <div className="dex-ui-section-desc">{t('Mehrfachauswahl.', 'Select all that apply.')}</div>
          <div className="dex-ui-card dex-ui-card--list">
            {FEEDBACK_GUT.map(o => zeile(o, gut.indexOf(o.id) >= 0, () => toggle(setGut, o.id)))}
          </div>
          <label className="dex-ui-label" htmlFor="fb-gut" style={{ marginTop: 10, display: 'block' }}>
            {t('Möchtest du etwas ergänzen?', 'Anything to add?')}{' '}
            <span className="dex-ui-label-optional">{t('optional', 'optional')}</span>
          </label>
          <textarea
            id="fb-gut"
            className="dex-ui-input"
            rows={3}
            value={freitextGut}
            onChange={e => setFreitextGut(e.target.value)}
            placeholder={t('z.B. was dir besonders geholfen hat', 'e.g. what helped you most')}
            style={{ resize: 'vertical' }}
          />
        </div>

        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{t('Was können wir verbessern?', 'What can we improve?')}</div>
          <div className="dex-ui-section-desc">{t('Mehrfachauswahl.', 'Select all that apply.')}</div>
          <div className="dex-ui-card dex-ui-card--list">
            {FEEDBACK_BESSER.map(o => zeile(o, besser.indexOf(o.id) >= 0, () => toggle(setBesser, o.id)))}
          </div>
          <label className="dex-ui-label" htmlFor="fb-besser" style={{ marginTop: 10, display: 'block' }}>
            {t('Was genau war es?', 'What exactly was it?')}{' '}
            <span className="dex-ui-label-optional">{t('optional', 'optional')}</span>
          </label>
          <textarea
            id="fb-besser"
            className="dex-ui-input"
            rows={3}
            value={freitextBesser}
            onChange={e => setFreitextBesser(e.target.value)}
            placeholder={t('Je konkreter, desto eher können wir es ändern.', 'The more specific, the more likely we can change it.')}
            style={{ resize: 'vertical' }}
          />
        </div>

        {fehler && (
          <div className="dex-ui-callout dex-ui-callout--warn" role="status" style={{ marginTop: 16 }}>
            <AlertCircle size={16} /><span>{fehler}</span>
          </div>
        )}

        <div className="dex-ui-inline" style={{ marginTop: 22, gap: 8, alignItems: 'center' }}>
          <button
            type="button"
            className="btn btn-primary"
            disabled={sendet || !hatInhalt}
            style={{ opacity: (sendet || !hatInhalt) ? 0.55 : 1 }}
            onClick={() => { void absenden(); }}
          >
            <Send size={16} /> {sendet ? t('Wird gesendet …', 'Sending …') : t('Feedback absenden', 'Send feedback')}
          </button>
          <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={() => onLeave && onLeave('start')}>
            {t('Später', 'Later')}
          </button>
          {/* Sagt, WARUM der Knopf noch nicht geht — ein grauer Knopf ohne
              Grund ist die Meldung, die niemand lesen kann. */}
          {!hatInhalt && (
            <span className="dex-ui-muted" style={{ fontSize: '0.78rem' }}>
              {t('Hak mindestens eine Antwort an oder schreib etwas.', 'Tick at least one answer or write something.')}
            </span>
          )}
        </div>
      </div>
    </div>
  );
};

export default FeedbackPage;
