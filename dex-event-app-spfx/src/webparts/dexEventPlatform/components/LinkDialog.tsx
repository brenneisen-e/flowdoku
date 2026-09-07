/**
 * v30.51: Link-Dialog des HTML-Editors — Ziel UND Anzeige-Text in EINEM
 * Fenster, dazu die Wahl zwischen Web-Adresse und E-Mail.
 *
 * Vorher waren es zwei nacheinander aufpoppende Prompts: erst die URL, dann
 * — und nur bei leerer Auswahl — der Anzeige-Text. Wer Text markiert hatte,
 * bekam die Frage nach dem Anzeige-Text nie zu sehen und konnte ihn also
 * auch nicht ändern. Zwei Dialoge hintereinander sind außerdem genau die
 * Bedienung, bei der man nach dem ersten „Übernehmen" glaubt, fertig zu sein.
 *
 * **E-Mail ist bewusst ein eigener Modus und kein Ratespiel.** Aus einer
 * Eingabe zu erkennen, ob `nils@deloitte.de` als `mailto:` gemeint ist,
 * geht meistens gut und dann einmal daneben — und ein Link, der im
 * Browser statt im Mailprogramm landet, fällt erst beim Empfänger auf.
 * Der Modus setzt `mailto:` selbst; die Adresse bleibt im Feld genau so
 * stehen, wie sie eingegeben wurde.
 *
 * Der Dialog schließt NICHT über den Hintergrund (`backdropClose={false}`) —
 * nur über Abbrechen, Übernehmen oder Escape. Sonst wirft ein Klick daneben
 * die Eingabe weg, was hier besonders leicht passiert: Wer den vorbelegten
 * Text markiert, zieht dabei fast zwangsläufig über den Rand der Karte
 * hinaus (s. Modal, v30.51).
 */
import * as React from 'react';
import Modal from './Modal';
// v31.2: Der Editor drumherum spricht längst beide Sprachen (HtmlEditorModal,
// isDe); der Link-Dialog war bis hier deutsch fest verdrahtet. `useLocaleSafe`
// braucht keinen Provider-Zwang und fällt auf Deutsch zurück.
import { useLocaleSafe } from '../context/LanguageContext';
import { cx } from './dexUi';
import { AlertCircle, Link2, Mail, Trash2 } from './Icons';

export interface LinkDialogResult {
  /** Fertige href — bei E-Mail bereits mit `mailto:`. */
  href: string;
  /** Anzeige-Text. Leer = Aufrufer setzt die href als Text ein. */
  text: string;
}

export interface LinkDialogProps {
  open: boolean;
  /** Vorbelegung: bestehende href (beim Bearbeiten eines Links). */
  initialHref?: string;
  /** Vorbelegung: markierter Text bzw. Text des bestehenden Links. */
  initialText?: string;
  /** Es wird ein BESTEHENDER Link bearbeitet — dann gibt es „Link entfernen". */
  editing?: boolean;
  /** Der Anzeige-Text lässt sich nicht ändern (Auswahl über mehrere Absätze). */
  textLocked?: boolean;
  onCancel: () => void;
  onApply: (r: LinkDialogResult) => void;
  /** Nur bei `editing` — Link entfernen, Text behalten. */
  onRemove?: () => void;
}

const MAILTO = 'mailto:';

export default function LinkDialog(props: LinkDialogProps): React.ReactElement | null {
  const { open, initialHref = '', initialText = '', editing = false, textLocked = false } = props;

  const [mode, setMode] = React.useState<'web' | 'mail'>('web');
  const [target, setTarget] = React.useState('');
  const [text, setText] = React.useState('');
  const [error, setError] = React.useState('');

  // Beim Öffnen aus der Vorbelegung befüllen. Ein bestehender `mailto:`-Link
  // öffnet den Dialog im E-Mail-Modus — sonst stünde das Präfix im Feld und
  // wäre beim Speichern doppelt.
  React.useEffect(() => {
    if (!open) return;
    const href = (initialHref || '').trim();
    if (href.toLowerCase().indexOf(MAILTO) === 0) {
      setMode('mail');
      setTarget(href.slice(MAILTO.length));
    } else {
      setMode('web');
      setTarget(href || 'https://');
    }
    setText(initialText || '');
    setError('');
  }, [open, initialHref, initialText]);

  // v31.2: Hinter den bestehenden Hooks, vor dem frühen Return — die
  // Reihenfolge der alten Hooks bleibt damit unangetastet.
  const isDe = useLocaleSafe() === 'de';
  const t = (de: string, en: string): string => (isDe ? de : en);

  if (!open) return null;

  const apply = (): void => {
    const raw = target.trim();
    if (mode === 'mail') {
      // Absichtlich milde Prüfung: Es geht darum, den Vertipper „nils@" oder
      // eine versehentlich eingefügte URL zu erwischen, nicht darum, die
      // Adress-Syntax nachzubauen.
      if (raw.indexOf('@') < 1 || raw.indexOf('.', raw.indexOf('@')) < 0 || /\s/.test(raw)) {
        setError(t(
          'Das sieht nicht nach einer E-Mail-Adresse aus — erwartet wird z.B. b2runkoeln@deloitte.de.',
          'That does not look like an email address — expected something like b2runkoeln@deloitte.de.',
        ));
        return;
      }
      props.onApply({ href: MAILTO + raw, text: text.trim() });
      return;
    }
    if (!raw || raw === 'https://' || raw === 'http://') {
      setError(t(
        'Trag die Adresse der Seite ein, auf die der Link führen soll.',
        'Enter the address of the page the link should open.',
      ));
      return;
    }
    // Ohne Schema landet der Link relativ zur SharePoint-Seite — der
    // häufigste stille Fehler bei kopierten Adressen wie „www.b2run.de".
    const href = /^(https?:|mailto:|tel:|#)/i.test(raw) ? raw : `https://${raw}`;
    props.onApply({ href, text: text.trim() });
  };

  // v31.2: Die Wahl Web/E-Mail ist eine Entscheidung mit Folge (Browser oder
  // Mailprogramm) — deshalb zwei Kacheln mit je einer Zeile „was das heißt"
  // statt zweier Wörter in einer Pill (Leitfaden 2b).
  const choice = (m: 'web' | 'mail', icon: React.ReactNode, title: string, desc: string): React.ReactElement => (
    <button
      type="button"
      className={cx('dex-ui-choice', mode === m && 'is-active')}
      onClick={() => {
        setError('');
        setMode(prev => {
          if (prev === m) return prev;
          // Beim Wechsel das Feld leeren statt eine halbe Adresse stehen zu
          // lassen — `https://` als E-Mail-Adresse ist nur Arbeit für den
          // Nutzer.
          setTarget(m === 'mail' ? '' : 'https://');
          return m;
        });
      }}
      aria-pressed={mode === m}
    >
      <span className="dex-ui-choice-icon">{icon}</span>
      <span className="dex-ui-choice-body">
        <span className="dex-ui-choice-title">{title}</span>
        <span className="dex-ui-choice-desc">{desc}</span>
      </span>
    </button>
  );

  const isMail = mode === 'mail';

  return (
    <Modal
      open
      onClose={props.onCancel}
      backdropClose={false}
      maxWidth={520}
      ariaLabel="Link"
      title={editing ? t('Link bearbeiten', 'Edit link') : t('Link einfügen', 'Insert link')}
      subtitle={t('Wohin soll der Link führen — und was soll dort stehen?', 'Where should the link go — and what should it say?')}
      icon={<Link2 size={20} />}
      footer={<>
        {editing && props.onRemove && (
          // v31.2: Links außen und mit Abstand zu „Abbrechen" — der Text bleibt
          // stehen, nur die Verlinkung geht weg, deshalb ohne Rückfrage.
          <span className="dex-ui-modal-foot-left">
            <button type="button" className="dex-ui-textbtn dex-ui-textbtn--danger" onClick={props.onRemove}>
              <Trash2 size={15} />
              {t('Link entfernen', 'Remove link')}
            </button>
          </span>
        )}
        <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={props.onCancel}>
          {t('Abbrechen', 'Cancel')}
        </button>
        <button type="button" className="btn btn-primary dex-ui-btn-sm" onClick={apply}>
          {t('Übernehmen', 'Apply')}
        </button>
      </>}
    >
      <div className="dex-ui-grid-2" role="group" aria-label={t('Art des Links', 'Link type')}>
        {choice('web', <Link2 size={18} />, t('Web-Adresse', 'Web address'),
          t('Öffnet eine Seite im Browser.', 'Opens a page in the browser.'))}
        {choice('mail', <Mail size={18} />, t('E-Mail', 'Email'),
          t('Öffnet das Mailprogramm mit dieser Adresse.', 'Opens the mail app with this address.'))}
      </div>

      <div className="dex-ui-field">
        <label className="dex-ui-label" htmlFor="dex-link-target">
          {isMail
            ? t('An welche E-Mail-Adresse soll geschrieben werden?', 'Which email address should be written to?')
            : t('Auf welche Seite soll der Link führen?', 'Which page should the link open?')}
        </label>
        <input
          id="dex-link-target"
          className="dex-ui-input"
          value={target}
          onChange={e => { setTarget(e.target.value); setError(''); }}
          onKeyDown={e => { if (e.key === 'Enter') apply(); }}
          placeholder={isMail ? 'b2runkoeln@deloitte.de' : 'https://www.b2run.de/koeln'}
          autoFocus
        />
        <div className="dex-ui-help">
          {isMail
            ? t('Das mailto: setzen wir selbst — trag nur die Adresse ein.', 'We add mailto: for you — just enter the address.')
            : t('Fehlt https:// am Anfang, ergänzen wir es beim Übernehmen.', 'If https:// is missing, we add it when you apply.')}
        </div>
      </div>

      <div className="dex-ui-field">
        <label className="dex-ui-label" htmlFor="dex-link-text">
          {t('Was soll als Link-Text stehen?', 'What should the link text say?')}
          {!textLocked && <span className="dex-ui-label-optional">{t('(optional)', '(optional)')}</span>}
        </label>
        <input
          id="dex-link-text"
          className="dex-ui-input"
          value={text}
          disabled={textLocked}
          onChange={e => setText(e.target.value)}
          onKeyDown={e => { if (e.key === 'Enter') apply(); }}
          placeholder={isMail ? t('Schreib uns', 'Write to us') : t('Zur Anmeldung', 'Go to registration')}
        />
        <div className="dex-ui-help">
          {textLocked
            ? t(
              'Deine Auswahl geht über mehrere Absätze — der markierte Text bleibt unverändert und wird nur verlinkt.',
              'Your selection spans several paragraphs — the marked text stays as it is and only gets linked.',
            )
            : t('Leer gelassen erscheint die Adresse selbst als Text.', 'Leave it empty and the address itself is shown.')}
        </div>
      </div>

      {error && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>{error}</span>
        </div>
      )}
    </Modal>
  );
}
