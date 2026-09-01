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

const labelStyle: React.CSSProperties = {
  display: 'block', fontSize: '0.78rem', fontWeight: 600,
  color: 'var(--dex-gray-700, #444)', marginBottom: 4,
};

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

  if (!open) return null;

  const apply = (): void => {
    const raw = target.trim();
    if (mode === 'mail') {
      // Absichtlich milde Prüfung: Es geht darum, den Vertipper „nils@" oder
      // eine versehentlich eingefügte URL zu erwischen, nicht darum, die
      // Adress-Syntax nachzubauen.
      if (raw.indexOf('@') < 1 || raw.indexOf('.', raw.indexOf('@')) < 0 || /\s/.test(raw)) {
        setError('Das sieht nicht nach einer E-Mail-Adresse aus — erwartet wird z.B. b2runkoeln@deloitte.de.');
        return;
      }
      props.onApply({ href: MAILTO + raw, text: text.trim() });
      return;
    }
    if (!raw || raw === 'https://' || raw === 'http://') {
      setError('Bitte trage die Adresse der Seite ein, auf die der Link führen soll.');
      return;
    }
    // Ohne Schema landet der Link relativ zur SharePoint-Seite — der
    // häufigste stille Fehler bei kopierten Adressen wie „www.b2run.de".
    const href = /^(https?:|mailto:|tel:|#)/i.test(raw) ? raw : `https://${raw}`;
    props.onApply({ href, text: text.trim() });
  };

  const seg = (m: 'web' | 'mail', label: string): React.ReactElement => (
    <button
      type="button"
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
      style={{
        border: 'none', cursor: 'pointer', padding: '7px 16px',
        borderRadius: 999, fontSize: '0.82rem', fontWeight: 700,
        background: mode === m ? 'var(--dex-green, #86bc25)' : 'transparent',
        color: mode === m ? '#fff' : 'var(--dex-gray-700, #444)',
        transition: 'background 140ms ease, color 140ms ease',
      }}
    >
      {label}
    </button>
  );

  return (
    <Modal open onClose={props.onCancel} backdropClose={false} maxWidth={520} ariaLabel="Link">
      <h3 style={{ margin: 0, fontSize: '1.02rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
        {editing ? 'Link bearbeiten' : 'Link einfügen'}
      </h3>

      <div style={{
        display: 'inline-flex', gap: 4, padding: 3, alignSelf: 'flex-start',
        background: 'var(--dex-gray-100, #f5f5f5)', borderRadius: 999,
      }}>
        {seg('web', 'Web-Adresse')}
        {seg('mail', 'E-Mail')}
      </div>

      <div>
        <label style={labelStyle} htmlFor="dex-link-target">
          {mode === 'mail' ? 'E-Mail-Adresse' : 'Link-Adresse (URL)'}
        </label>
        <input
          id="dex-link-target"
          className="form-input"
          value={target}
          onChange={e => { setTarget(e.target.value); setError(''); }}
          onKeyDown={e => { if (e.key === 'Enter') apply(); }}
          placeholder={mode === 'mail' ? 'b2runkoeln@deloitte.de' : 'https://www.b2run.de/koeln'}
          autoFocus
          style={{ fontSize: '0.9rem', padding: '9px 12px', width: '100%' }}
        />
      </div>

      <div>
        <label style={labelStyle} htmlFor="dex-link-text">
          Angezeigter Text {textLocked ? '' : <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}>(leer = die Adresse selbst)</span>}
        </label>
        <input
          id="dex-link-text"
          className="form-input"
          value={text}
          disabled={textLocked}
          onChange={e => setText(e.target.value)}
          onKeyDown={e => { if (e.key === 'Enter') apply(); }}
          placeholder={mode === 'mail' ? 'Schreib uns' : 'Zur Anmeldung'}
          style={{ fontSize: '0.9rem', padding: '9px 12px', width: '100%' }}
        />
        {textLocked && (
          <p style={{ margin: '5px 0 0', fontSize: '0.74rem', color: 'var(--dex-gray-500)', lineHeight: 1.45 }}>
            Deine Auswahl geht über mehrere Absätze — der markierte Text bleibt dann unverändert und wird
            nur verlinkt.
          </p>
        )}
      </div>

      {error && (
        <p style={{
          margin: 0, padding: '8px 12px', borderRadius: 8,
          background: 'rgba(218,41,28,0.08)', color: 'var(--dex-red, #da291c)',
          fontSize: '0.8rem', lineHeight: 1.45,
        }}>{error}</p>
      )}

      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 4, flexWrap: 'wrap' }}>
        {editing && props.onRemove && (
          <button
            type="button"
            className="btn btn-secondary"
            onClick={props.onRemove}
            style={{ fontSize: '0.88rem', padding: '9px 18px', marginRight: 'auto' }}
          >
            Link entfernen
          </button>
        )}
        <button
          type="button"
          className="btn btn-secondary"
          onClick={props.onCancel}
          style={{ fontSize: '0.88rem', padding: '9px 18px' }}
        >
          Abbrechen
        </button>
        <button
          type="button"
          className="btn btn-primary"
          onClick={apply}
          style={{ fontSize: '0.88rem', padding: '9px 18px' }}
        >
          Übernehmen
        </button>
      </div>
    </Modal>
  );
}
