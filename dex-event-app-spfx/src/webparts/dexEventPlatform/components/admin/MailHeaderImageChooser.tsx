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
 * Die Auswahl „DEX-Logo | Event-Foto" plus Zuschneiden — identisch in jedem
 * Mail-Dialog. Die Maße stehen im „HEADER-BILD"-Block des Editors (s. oben).
 */
export default function MailHeaderImageChooser(props: MailHeaderImageChooserProps): React.ReactElement {
  const { value, onChange, eventPhotoB64, disabled, isDe } = props;
  const opts: Array<{ key: 'logo' | 'event'; label: string; enabled: boolean }> = [
    { key: 'logo', label: isDe ? 'DEX-Logo' : 'DEX logo', enabled: true },
    { key: 'event', label: isDe ? 'Event-Foto' : 'Event photo', enabled: !!eventPhotoB64 },
  ];
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', marginBottom: 10 }}>
      <span style={{ fontSize: '0.78rem', fontWeight: 700, color: 'var(--dex-gray-600)' }}>
        {isDe ? 'Bild im Mail-Kopf:' : 'Header image:'}
      </span>
      <div style={{ display: 'inline-flex', borderRadius: 8, overflow: 'hidden', border: '1px solid var(--dex-gray-300)' }}>
        {opts.map(opt => {
          const active = value.hero === opt.key;
          return (
            <button
              key={opt.key}
              type="button"
              disabled={disabled || !opt.enabled}
              onClick={() => { if (opt.enabled) onChange({ ...value, hero: opt.key }); }}
              title={!opt.enabled ? (isDe ? 'Dieses Event hat kein Bild hinterlegt.' : 'This event has no image set.') : undefined}
              style={{
                padding: '6px 14px', fontSize: '0.78rem', border: 'none',
                cursor: (opt.enabled && !disabled) ? 'pointer' : 'not-allowed',
                background: active ? 'var(--dex-green)' : 'transparent',
                color: active ? '#fff' : (opt.enabled ? 'var(--dex-gray-600)' : 'var(--dex-gray-400)'),
                fontWeight: active ? 700 : 500, opacity: opt.enabled ? 1 : 0.6,
              }}
            >
              {opt.label}
            </button>
          );
        })}
      </div>
      {value.hero === 'event' && !!eventPhotoB64 && props.onCrop && (
        <button
          type="button"
          disabled={disabled}
          onClick={props.onCrop}
          style={{
            padding: '5px 12px', fontSize: '0.75rem', fontWeight: 600,
            cursor: disabled ? 'not-allowed' : 'pointer',
            border: '1px solid var(--dex-green, #86bc25)', borderRadius: 6,
            background: 'rgba(134,188,37,0.10)', color: 'var(--dex-green-dark, #4a7c1f)',
          }}
        >
          {isDe ? 'Foto zuschneiden' : 'Crop photo'}
        </button>
      )}
      <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
        {value.hero === 'event'
          ? (isDe ? 'Das Event-Foto erscheint im Mail-Kopf.' : 'The event photo is shown in the header.')
          : (isDe ? 'Standard-Bild (DEX-Logo bzw. dein Mail-Logo).' : 'Default image (DEX logo or your mail logo).')}
      </span>
    </div>
  );
}
