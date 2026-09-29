/**
 * „Hast du Fragen?" und „Organizer werden?" — ein Dialog, zwei Anlässe.
 *
 * In DEX ist „Hast du Fragen?" ein Ticketsystem: Die Frage landet in einer
 * SharePoint-Liste, Power-User beantworten sie in der App, Power-Automate-
 * Flows schicken die Mails. Das ist hier bewusst NICHT nachgebaut. Eine
 * Demo-Plattform mit einer Handvoll Organizern braucht keine Warteschlange;
 * sie braucht, dass eine Frage bei einem Menschen ankommt. Deshalb öffnet der
 * Dialog die Mail an die Ansprechperson (`KONTAKT_EMAIL`), vorausgefüllt mit
 * Anlass und Absender.
 *
 * Zwei Absicherungen, weil `mailto:` nicht überall etwas öffnet: Die Adresse
 * steht sichtbar im Dialog und lässt sich kopieren — im SharePoint-Handy-
 * Browser gibt es oft kein Mailprogramm, das den Link annimmt, und dann steht
 * die Person vor einem Knopf, der nichts tut.
 *
 * Wächst der Bedarf, ist der nächste Schritt eine Liste `AIUC_Fragen` mit
 * derselben Oberfläche; der Dialog behält dann Anlass und Text und ersetzt nur
 * `senden`.
 */

import * as React from 'react';
import Modal from './Modal';
import { cx } from './dexUi';
import { Copy, Check, Mail } from './Icons';
import { useLanguage } from '../context/LanguageContext';
import { useCurrentUser } from '../context/UserContext';
import { KONTAKT_EMAIL, APP_NAME } from '../constants';
import { APP_VERSION } from '../version';

export type KontaktArt = 'frage' | 'organizer';

/** Die Anlässe der Frage. „Organizer werden" hat keine Wahl — der Anlass steht fest. */
type Anlass = 'frage' | 'idee' | 'fehler' | 'vorschlag';

/** `mailto:` ist bei den meisten Mailprogrammen auf rund 2 000 Zeichen begrenzt. */
const MAX_TEXT = 1200;

export default function ContactModal(props: { open: boolean; art: KontaktArt; onClose: () => void }): React.ReactElement | null {
  const { open, art, onClose } = props;
  const { t, isDe } = useLanguage();
  const { currentUser } = useCurrentUser();

  const [anlass, setAnlass] = React.useState<Anlass>('frage');
  const [text, setText] = React.useState('');
  const [kopiert, setKopiert] = React.useState(false);

  // Beim Öffnen leer anfangen — sonst steht beim nächsten Mal der Text der
  // vorigen Frage da.
  React.useEffect(() => {
    if (open) { setAnlass('frage'); setText(''); setKopiert(false); }
  }, [open, art]);

  const ANLAESSE: Array<{ w: Anlass; titel: string; hilfe: string }> = [
    { w: 'frage', titel: t('Ich habe eine Frage', 'I have a question'), hilfe: t('Zur Bedienung, zu einem Use Case oder zur Plattform.', 'About using the platform or a use case.') },
    { w: 'vorschlag', titel: t('Ich möchte einen Use Case vorschlagen', 'I want to suggest a use case'), hilfe: t('Eine Demo, die auf die Plattform gehört.', 'A demo that belongs on the platform.') },
    { w: 'idee', titel: t('Ich habe eine Idee', 'I have an idea'), hilfe: t('Für die Plattform selbst.', 'For the platform itself.') },
    { w: 'fehler', titel: t('Mir ist ein Fehler aufgefallen', 'I noticed a problem'), hilfe: t('Etwas geht nicht oder sieht falsch aus.', 'Something does not work or looks wrong.') },
  ];

  const betreff = art === 'organizer'
    ? `${APP_NAME}: ${isDe ? 'Anfrage Use Case Organizer' : 'Request to become a Use Case Organizer'}`
    : `${APP_NAME}: ${ANLAESSE.filter(a => a.w === anlass)[0].titel}`;

  const vorgabe = art === 'organizer'
    ? t('Ich möchte Use Cases auf der Plattform einstellen und pflegen.\n\nMein Vorhaben:\n', 'I would like to add and maintain use cases on the platform.\n\nWhat I have in mind:\n')
    : '';

  const absender = [currentUser.displayName, currentUser.email].filter(Boolean).join(' · ');
  const koerper = `${(text || vorgabe).slice(0, MAX_TEXT)}\n\n—\n${isDe ? 'Von' : 'From'}: ${absender || '—'}\n${APP_NAME} v${APP_VERSION}`;
  const link = `mailto:${KONTAKT_EMAIL}?subject=${encodeURIComponent(betreff)}&body=${encodeURIComponent(koerper)}`;

  const kopieren = (): void => {
    try {
      // Kann in der SharePoint-Handy-App fehlen (Permissions Policy) — dann
      // bleibt die sichtbare Adresse der Weg, und der Knopf tut nichts Falsches.
      void navigator.clipboard.writeText(KONTAKT_EMAIL).then(
        () => setKopiert(true),
        () => setKopiert(false),
      );
    } catch { setKopiert(false); }
  };

  const titel = art === 'organizer' ? t('Use Case Organizer werden', 'Become a Use Case Organizer') : t('Hast du Fragen?', 'Any questions?');
  const untertitel = art === 'organizer'
    ? t('Organizer legen Use Cases an und pflegen sie. Schreib kurz, was du einstellen möchtest.', 'Organizers create and maintain use cases. Tell us briefly what you would like to add.')
    : t('Schreib uns — die Mail öffnet sich vorausgefüllt in deinem Mailprogramm.', 'Write to us — the email opens prefilled in your mail client.');

  return (
    <Modal
      open={open}
      onClose={onClose}
      maxWidth={560}
      ariaLabel={titel}
      title={titel}
      subtitle={untertitel}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={onClose}>{t('Schließen', 'Close')}</button>
        {/* Kein `onClose` am Link: Öffnet sich kein Mailprogramm (Handy-Browser, kein Handler,
            abgebrochene Rückfrage), waren Text und Notausgang (sichtbare Adresse, Kopieren) genau
            dann weg, wenn man sie brauchte. Der Dialog bleibt stehen, bis man ihn selbst schließt. */}
        <a className="btn btn-primary" href={link} style={{ textDecoration: 'none' }}>
          <Mail size={16} /> {t('E-Mail schreiben', 'Write email')}
        </a>
      </>}
    >
      {art === 'frage' && (
        <div className="dex-ui-field">
          <span className="dex-ui-label">{t('Worum geht es?', 'What is it about?')}</span>
          <div style={{ display: 'grid', gap: 8 }}>
            {ANLAESSE.map(a => (
              <button key={a.w} type="button" className={cx('dex-ui-choice', anlass === a.w && 'is-active')} style={{ padding: '10px 12px', alignItems: 'center' }} aria-pressed={anlass === a.w} onClick={() => setAnlass(a.w)}>
                <span className="dex-ui-choice-body" style={{ display: 'block' }}>
                  <span className="dex-ui-choice-title" style={{ display: 'block' }}>{a.titel}</span>
                  <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{a.hilfe}</span>
                </span>
                <span className="dex-ui-choice-check">{anlass === a.w && <Check size={12} />}</span>
              </button>
            ))}
          </div>
        </div>
      )}

      <label className="dex-ui-field">
        <span className="dex-ui-label">{art === 'organizer' ? t('Was möchtest du einstellen?', 'What would you like to add?') : t('Was möchtest du uns sagen?', 'What would you like to tell us?')}</span>
        <textarea
          className="dex-ui-textarea"
          rows={5}
          maxLength={MAX_TEXT}
          value={text}
          placeholder={vorgabe.replace(/\n+$/, '')}
          onChange={e => setText(e.target.value)}
        />
        <span className="dex-ui-help">{t('Dein Name und deine Adresse werden mitgeschickt.', 'Your name and address are sent along.')}</span>
      </label>

      {/* Für den Fall, dass sich kein Mailprogramm öffnet. */}
      <div className="dex-ui-callout dex-ui-callout--neutral dex-ui-callout--sm">
        <span style={{ flex: 1, minWidth: 0 }}>
          {t('Öffnet sich kein Mailprogramm, schreib direkt an ', 'If no mail client opens, write directly to ')}
          <strong style={{ overflowWrap: 'anywhere' }}>{KONTAKT_EMAIL}</strong>
        </span>
        <button type="button" className="dex-ui-textbtn" onClick={kopieren}>
          {kopiert ? <><Check size={14} /> {t('Kopiert', 'Copied')}</> : <><Copy size={14} /> {t('Adresse kopieren', 'Copy address')}</>}
        </button>
      </div>
    </Modal>
  );
}
