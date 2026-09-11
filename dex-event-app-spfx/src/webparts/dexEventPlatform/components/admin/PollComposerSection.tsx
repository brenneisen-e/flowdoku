/**
 * v31.12 — Umfrage in der Massenmail.
 *
 * Sitzt im Kopfbereich von „E-Mail versenden", zwischen „Wie sieht der Kopf
 * aus?" und „Bevor du sendest" — Nutzer-Ansage 11.09.2026: Die Umfrage soll
 * über die Aktion „E-Mail versenden" abrufbar sein, nicht als eigener Weg
 * daneben. Ein zweiter Weg zur selben Mail wäre die Sucherei, die DEX
 * an anderer Stelle schon einmal gekostet hat (zwei Reiter-Reihen, v28.88).
 *
 * Warum der Block IN DEN TEXT eingefügt wird statt beim Senden angehängt:
 * Vorschau, Testmail und Versand lesen alle denselben `emailBody`. Ein
 * Anhängen erst beim Senden hiesse, dass die Vorschau etwas anderes zeigt
 * als die Mail — genau der Fehler, den CLAUDE.md beim Outlook-Standardtext
 * beschreibt („nie mit eigener Logik, sondern mit derselben Quelle"). Der
 * Preis: Der Organizer kann den Block im Editor verändern. Das ist gewollt —
 * er sieht, was rausgeht, und kann Text dazwischenschreiben.
 *
 * Der „Einfügen"-Knopf ist auch der Speicher-Punkt: Ohne gespeicherte Frage
 * zeigen die Links in der Mail eine leere Umfrage. Erst speichern, dann
 * einfügen — und wenn das Speichern scheitert, wird NICHTS eingefügt.
 */

import * as React from 'react';
import { Check, Plus, Trash2, AlertCircle } from '../Icons';
import { cx } from '../dexUi';
import { DeloitteEvent } from '../../types';
import { EventService } from '../../services/EventService';
import { POLL_VORLAGEN_DE, PollConfig } from '../../services/events/poll';
import { buildPollMailBody } from '../../utils/pollEmail';
import { APP_URL } from '../../services/EmailTemplates';
import { useCurrentUser } from '../../context/UserContext';

export interface PollComposerSectionProps {
  selectedEvent: DeloitteEvent;
  eventServiceRef: EventService;
  isDe: boolean;
  disabled: boolean;
  /** Der aktuelle Mailtext — der Block wird hinten angehängt. */
  emailBody: string;
  setEmailBody: React.Dispatch<React.SetStateAction<string>>;
}

const VORLAGEN_EN: Array<{ titel: string; frage: string; optionen: string[] }> = [
  {
    titel: 'Understand cancellations',
    frage: 'Sorry you could not make it — what was the reason?',
    optionen: ['Start block too early', 'Fell ill', 'Important client meeting', 'Travel too complicated', 'No longer interested', 'Other reason'],
  },
  {
    titel: 'Feedback on the event',
    frage: 'How did you like the event?',
    optionen: ['Very good', 'Good', 'So-so', 'Not good'],
  },
  {
    titel: 'Wish for next time',
    frage: 'What should we do differently next time?',
    optionen: ['A different date', 'A different location', 'More time to network', 'A shorter programme', 'Everything is fine as it is'],
  },
];

export const PollComposerSection: React.FC<PollComposerSectionProps> = (p) => {
  const { selectedEvent, eventServiceRef, isDe, disabled, emailBody, setEmailBody } = p;
  // Die eigene Adresse für „Meine eigene Antwort löschen" — über den Context
  // statt als Prop, damit der Composer-Aufrufer nichts durchreichen muss.
  const { currentUser } = useCurrentUser();
  const myEmail = (currentUser.email || '').trim().toLowerCase();
  const t = (de: string, en: string): string => (isDe ? de : en);
  const vorlagen = isDe ? POLL_VORLAGEN_DE : VORLAGEN_EN;

  const [an, setAn] = React.useState(false);
  const [frage, setFrage] = React.useState('');
  const [optionen, setOptionen] = React.useState<string[]>([]);
  const [mehrfach, setMehrfach] = React.useState(true);
  const [anonym, setAnonym] = React.useState(true);
  const [pollId, setPollId] = React.useState(0);
  const [busy, setBusy] = React.useState(false);
  const [meldung, setMeldung] = React.useState('');
  const [meldungOk, setMeldungOk] = React.useState(true);
  const [geladen, setGeladen] = React.useState(false);

  const eventNumber = selectedEvent.eventNumber || 0;

  // Eine bereits angelegte Umfrage übernehmen, statt eine zweite danebenzubauen.
  React.useEffect(() => {
    let weg = false;
    (async () => {
      if (!eventServiceRef || !eventNumber) { setGeladen(true); return; }
      const r = await eventServiceRef.getPoll(eventNumber);
      if (weg) return;
      if (r.ok && r.poll) {
        setPollId(r.poll.id);
        setFrage(r.poll.frage);
        setOptionen(r.poll.optionen);
        setMehrfach(r.poll.mehrfach);
        setAnonym(r.poll.anonym);
        setAn(true);
      }
      setGeladen(true);
    })().catch(() => setGeladen(true));
    return () => { weg = true; };
  }, [eventServiceRef, eventNumber]);

  const vorlageNehmen = (idx: number): void => {
    const v = vorlagen[idx];
    if (!v) return;
    setFrage(v.frage);
    setOptionen(v.optionen.slice());
    setMeldung('');
  };

  const optionSetzen = (i: number, wert: string): void => {
    setOptionen(prev => prev.map((o, j) => (j === i ? wert : o)));
  };

  const einfuegen = async (): Promise<void> => {
    setBusy(true);
    setMeldung('');
    try {
      const sauber = optionen.map(o => o.trim()).filter(Boolean);
      if (!frage.trim() || sauber.length < 2) {
        setMeldungOk(false);
        setMeldung(t('Es braucht eine Frage und mindestens zwei Antwortmöglichkeiten.',
          'A question and at least two answer options are needed.'));
        return;
      }
      await eventServiceRef.ensurePollLists();
      const cfg: PollConfig = {
        id: pollId,
        eventNumber,
        eventId: selectedEvent.id,
        frage: frage.trim(),
        optionen: sauber,
        mehrfach,
        anonym,
        aktiv: true,
      };
      const ok = await eventServiceRef.savePoll(cfg);
      if (!ok) {
        // Ohne gespeicherte Frage zeigen die Links in der Mail eine leere
        // Umfrage — deshalb wird hier NICHTS eingefügt.
        setMeldungOk(false);
        setMeldung(t('Die Umfrage konnte nicht gespeichert werden — es wurde nichts in die Mail eingefügt. Bitte versuch es noch einmal.',
          'The poll could not be saved — nothing was inserted into the email. Please try again.'));
        return;
      }
      if (pollId === 0) {
        const nach = await eventServiceRef.getPoll(eventNumber);
        if (nach.ok && nach.poll) setPollId(nach.poll.id);
      }
      const block = buildPollMailBody({
        appUrl: APP_URL,
        eventNumber,
        frage: frage.trim(),
        optionen: sauber,
        mehrfach,
        anonym,
        isDe,
      });
      setEmailBody(prev => `${prev || ''}\n${block}`);
      setMeldungOk(true);
      setMeldung(t('Die Umfrage steht jetzt unten im Text. Sie ist auch in der Testmail — klick darin ruhig ein Kästchen an.',
        'The poll is now at the end of the text. It is in the test email too — feel free to tick a box there.'));
    } finally {
      setBusy(false);
    }
  };

  const testantwortLoeschen = async (): Promise<void> => {
    setBusy(true);
    setMeldung('');
    try {
      const ok = await eventServiceRef.deletePollAnswer(eventNumber, myEmail);
      setMeldungOk(ok);
      setMeldung(ok
        ? t('Deine eigene Antwort ist gelöscht — die Auswertung zählt sie nicht mehr mit.',
          'Your own answer is deleted — the evaluation no longer counts it.')
        : t('Deine Antwort konnte nicht gelöscht werden. Bitte versuch es noch einmal.',
          'Your answer could not be deleted. Please try again.'));
    } finally {
      setBusy(false);
    }
  };

  const bereit = !!frage.trim() && optionen.filter(o => o.trim()).length >= 2;

  return (
    <div>
      <div className="dex-ui-section-title">{t('Soll die Mail eine Umfrage enthalten?', 'Should the email contain a poll?')}</div>

      <label className={cx('dex-ui-toggle-row', an && 'is-active')} htmlFor="poll-an">
        <input
          id="poll-an"
          type="checkbox"
          className="dex-ui-checkbox"
          checked={an}
          disabled={disabled || !geladen}
          onChange={e => { setAn(e.target.checked); setMeldung(''); }}
        />
        <span className="dex-ui-toggle-row-body">
          <span className="dex-ui-toggle-row-title">{t('Ja, mit Antwort-Kästchen zum Anklicken', 'Yes, with clickable answer boxes')}</span>
          <span className="dex-ui-toggle-row-desc">
            {t('Jede Antwortmöglichkeit wird ein Kästchen in der Mail. Ein Klick genügt — die Antwort ist sofort gespeichert, ohne dass jemand etwas absenden muss.',
              'Each answer option becomes a box in the email. One click is enough — the answer is saved right away, nobody has to submit anything.')}
          </span>
        </span>
      </label>

      {an && (
        <div className="dex-ui-stack" style={{ gap: 14, marginTop: 12 }}>
          <div>
            <div className="dex-ui-label">{t('Vorlage übernehmen', 'Use a template')}</div>
            <div className="dex-ui-inline" style={{ flexWrap: 'wrap' }}>
              {vorlagen.map((v, i) => (
                <button key={i} type="button" className="dex-ui-chip" disabled={disabled || busy} onClick={() => vorlageNehmen(i)}>
                  {v.titel}
                </button>
              ))}
            </div>
            <div className="dex-ui-help">
              {t('Die Vorlage füllt Frage und Antworten aus — beides kannst du danach frei ändern.',
                'The template fills in the question and the answers — you can change both afterwards.')}
            </div>
          </div>

          <div>
            <label className="dex-ui-label" htmlFor="poll-frage">
              {t('Was möchtest du wissen?', 'What do you want to know?')}
              <span className="dex-ui-label-required">*</span>
            </label>
            <input
              id="poll-frage"
              className="dex-ui-input"
              value={frage}
              disabled={disabled || busy}
              onChange={e => { setFrage(e.target.value); setMeldung(''); }}
              placeholder={t('Zum Beispiel: Schade, dass du nicht dabei warst — woran lag es?',
                'For example: Sorry you could not make it — what was the reason?')}
            />
          </div>

          <div>
            <div className="dex-ui-label">
              {t('Antwortmöglichkeiten', 'Answer options')}
              <span className="dex-ui-label-required">*</span>
            </div>
            <div className="dex-ui-stack" style={{ gap: 6 }}>
              {optionen.map((o, i) => (
                <div key={i} className="dex-ui-inline" style={{ gap: 6 }}>
                  <input
                    className="dex-ui-input dex-ui-input--sm"
                    value={o}
                    disabled={disabled || busy}
                    onChange={e => optionSetzen(i, e.target.value)}
                    aria-label={`${t('Antwort', 'Answer')} ${i + 1}`}
                    style={{ flex: 1 }}
                  />
                  <button
                    type="button"
                    className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                    disabled={disabled || busy}
                    onClick={() => setOptionen(prev => prev.filter((_, j) => j !== i))}
                    aria-label={t('Antwort entfernen', 'Remove answer')}
                    title={t('Antwort entfernen', 'Remove answer')}
                  ><Trash2 size={14} /></button>
                </div>
              ))}
            </div>
            <button
              type="button"
              className="dex-ui-textbtn"
              style={{ marginLeft: -8, marginTop: 4 }}
              disabled={disabled || busy}
              onClick={() => setOptionen(prev => prev.concat(['']))}
            ><Plus size={13} /> {t('Antwortmöglichkeit hinzufügen', 'Add an answer option')}</button>
          </div>

          <label className={cx('dex-ui-toggle-row', mehrfach && 'is-active')} htmlFor="poll-mehrfach">
            <input
              id="poll-mehrfach" type="checkbox" className="dex-ui-checkbox"
              checked={mehrfach} disabled={disabled || busy}
              onChange={e => setMehrfach(e.target.checked)}
            />
            <span className="dex-ui-toggle-row-body">
              <span className="dex-ui-toggle-row-title">{t('Mehrere Antworten erlauben', 'Allow more than one answer')}</span>
              <span className="dex-ui-toggle-row-desc">
                {t('„Krank geworden" UND „wichtiger Kundentermin" ist ein echter Fall — ohne Mehrfachnennung muss sich die Person für eine Hälfte entscheiden.',
                  '“Fell ill” AND “important client meeting” is a real case — without multiple answers people have to pick one half.')}
              </span>
            </span>
          </label>

          <label className={cx('dex-ui-toggle-row', anonym && 'is-active')} htmlFor="poll-anonym">
            <input
              id="poll-anonym" type="checkbox" className="dex-ui-checkbox"
              checked={anonym} disabled={disabled || busy}
              onChange={e => setAnonym(e.target.checked)}
            />
            <span className="dex-ui-toggle-row-body">
              <span className="dex-ui-toggle-row-title">{t('Anonym auswerten', 'Evaluate anonymously')}</span>
              <span className="dex-ui-toggle-row-desc">
                {t('DEX zeigt dir dann nur Summen und die Freitexte ohne Absender. Gespeichert bleibt, WER geantwortet hat — sonst könntest du niemanden erinnern. Wer Vollzugriff auf die SharePoint-Website hat, kann dort trotzdem sehen, wer eine Zeile geschrieben hat; echte Anonymität bräuchte einen Flow, der die Antwort unter einem Dienstkonto schreibt.',
                  'DEX then shows you only totals and the free-text comments without a sender. It still records WHO answered — otherwise you could not remind anyone. Anyone with full control of the SharePoint site can still see who wrote a row; true anonymity would need a flow writing the answer under a service account.')}
              </span>
            </span>
          </label>

          {meldung && (
            <div className={cx('dex-ui-callout', meldungOk ? 'dex-ui-callout--success' : 'dex-ui-callout--warn')} role="status">
              <span className="dex-ui-callout-icon">{meldungOk ? <Check size={15} /> : <AlertCircle size={15} />}</span>
              <span className="dex-ui-callout-body">{meldung}</span>
            </div>
          )}

          <div className="dex-ui-inline" style={{ gap: 10, flexWrap: 'wrap' }}>
            <button
              type="button"
              className="btn btn-secondary dex-ui-btn-sm"
              disabled={disabled || busy || !bereit}
              onClick={() => { void einfuegen(); }}
            >
              <Plus size={14} /> {busy ? t('Einen Moment …', 'One moment …') : t('Umfrage in die Mail einfügen', 'Insert the poll into the email')}
            </button>
            <button
              type="button"
              className="dex-ui-textbtn dex-ui-textbtn--muted"
              disabled={disabled || busy || !myEmail}
              onClick={() => { void testantwortLoeschen(); }}
            >
              {t('Meine eigene Antwort löschen', 'Delete my own answer')}
            </button>
          </div>
          <div className="dex-ui-help">
            {t('Der Knopf speichert die Umfrage und hängt die Kästchen unten an den Mailtext — dort siehst du sie in der Vorschau und kannst sie verschieben. Die Testmail enthält dieselben Links: Ein Klick darin zählt als echte Antwort, deshalb der zweite Knopf.',
              'The button saves the poll and appends the boxes to the end of the email text — you see them in the preview and can move them. The test email carries the same links: a click there counts as a real answer, hence the second button.')}
          </div>
          {emailBody.indexOf('action=umfrage') >= 0 && (
            <div className="dex-ui-pill dex-ui-pill--green" style={{ alignSelf: 'flex-start' }}>
              <Check size={12} /> {t('Umfrage steht im Mailtext', 'Poll is in the email text')}
            </div>
          )}
        </div>
      )}
    </div>
  );
};
