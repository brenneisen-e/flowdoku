/**
 * v30.52: Die Auswahl „DEX-Logo | Event-Foto" für den Mail-Kopf — dieselbe
 * Bedienung in Massenmail, Einladungsmail und QR-Mail.
 *
 * Vorher stand diese Reiter-Reihe zweimal wortgleich in `AdminPage` (einmal je
 * Mail-Dialog) und in der QR-Mail gar nicht. Die reine Logik (Zustand, Maße,
 * Einsetzen des Bildes) liegt in `utils/mailHeaderImage` — sie wird auch beim
 * automatischen Versand gebraucht, wo es keine Oberfläche gibt.
 *
 * Die Felder für Breite und Innenabstand stehen bewusst weiter im
 * „HEADER-BILD"-Block des `HtmlEditorModal`: direkt neben der Live-Vorschau,
 * die sie sofort zeigt. Beide Stellen arbeiten auf demselben Objekt.
 */
import * as React from 'react';
import { MailHeaderImage } from '../../utils/mailHeaderImage';
import { Calendar, Check, Mail, Pencil } from '../Icons';
import { cx } from '../dexUi';

export interface MailHeaderImageChooserProps {
  value: MailHeaderImage;
  onChange: (next: MailHeaderImage) => void;
  /** Event-Foto als Base64. Leer = „Event-Foto" ist nicht wählbar. */
  eventPhotoB64: string;
  disabled?: boolean;
  /** Öffnet den Zuschneiden-Dialog. Fehlt er, entfällt der Knopf. */
  onCrop?: () => void;
  isDe: boolean;
}

/**
 * Die Auswahl „Standard-Logo | Event-Foto" plus Zuschneiden — identisch in
 * jedem Mail-Dialog. Die Maße stehen im „HEADER-BILD"-Block des Editors (s. oben).
 */
export default function MailHeaderImageChooser(props: MailHeaderImageChooserProps): React.ReactElement {
  const { value, onChange, eventPhotoB64, disabled, isDe } = props;
  const noPhoto = isDe ? 'Dieses Event hat kein Bild hinterlegt.' : 'This event has no image set.';
  // v31.2: Aus der schmalen Reiter-Reihe werden zwei Kacheln mit Vorschau und
  // einer Zeile Folge — der frühere Statussatz unter der Reihe („Das Event-Foto
  // erscheint im Mail-Kopf") steht jetzt in der Kachel selbst, wo er beim
  // Entscheiden gelesen wird, nicht erst danach.
  const opts: Array<{ key: 'logo' | 'event'; label: string; desc: string; enabled: boolean; icon: React.ReactNode }> = [
    {
      key: 'logo', enabled: true, icon: <Mail size={18} />,
      label: isDe ? 'Standard-Logo' : 'Default logo',
      desc: isDe ? 'Das DEX-Logo — oder dein Mail-Logo, wenn das Event eins hat.' : 'The DEX logo — or your mail logo if the event has one.',
    },
    {
      key: 'event', enabled: !!eventPhotoB64,
      icon: eventPhotoB64
        ? <img src={eventPhotoB64} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
        : <Calendar size={18} strokeWidth={2} />,
      label: isDe ? 'Event-Foto' : 'Event photo',
      desc: eventPhotoB64
        ? (isDe ? 'Das Foto des Events erscheint oben in der Mail.' : 'The event photo appears at the top of the email.')
        : noPhoto,
    },
  ];
  return (
    <div className="dex-ui-stack" style={{ gap: 8, marginBottom: 10 }}>
      <div className="dex-ui-label" style={{ marginBottom: 0 }}>
        {isDe ? 'Welches Bild steht oben in der Mail?' : 'Which image sits at the top of the email?'}
      </div>
      <div className="dex-ui-grid-2" style={{ gap: 10 }}>
        {opts.map(opt => {
          const active = value.hero === opt.key;
          return (
            <button
              key={opt.key}
              type="button"
              className={cx('dex-ui-choice', active && 'is-active', !opt.enabled && 'is-disabled')}
              disabled={disabled || !opt.enabled}
              aria-pressed={active}
              onClick={() => { if (opt.enabled) onChange({ ...value, hero: opt.key }); }}
              title={!opt.enabled ? noPhoto : undefined}
              style={{ padding: '10px 12px' }}
            >
              <span className="dex-ui-choice-icon" style={{ overflow: 'hidden' }} aria-hidden="true">{opt.icon}</span>
              <span className="dex-ui-choice-body">
                <span className="dex-ui-choice-title" style={{ display: 'block' }}>{opt.label}</span>
                <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{opt.desc}</span>
              </span>
              <span className="dex-ui-choice-check">{active && <Check size={12} />}</span>
            </button>
          );
        })}
      </div>
      {/* v31.2: Zuschneiden gehört zum Foto — deshalb direkt unter der Wahl und
          nur, wenn das Foto gewählt ist (ein Knopf im Knopf wäre kein gültiges HTML). */}
      {value.hero === 'event' && !!eventPhotoB64 && props.onCrop && (
        <div className="dex-ui-inline">
          <button type="button" className="dex-ui-textbtn" disabled={disabled} onClick={props.onCrop}>
            <Pencil size={14} />{isDe ? 'Foto zuschneiden' : 'Crop photo'}
          </button>
          <span className="dex-ui-muted">{isDe ? 'Ausschnitt wählen, bevor die Mail rausgeht.' : 'Pick the crop before the email goes out.'}</span>
        </div>
      )}
    </div>
  );
}
