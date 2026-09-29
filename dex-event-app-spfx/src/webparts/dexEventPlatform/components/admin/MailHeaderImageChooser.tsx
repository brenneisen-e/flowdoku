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
import { KOPF_RUND, KOPF_VOLLE_BREITE, MailHeaderImage, bildMasse, istVolleBreite, kopfMasseFuerBild } from '../../utils/mailHeaderImage';
import { Calendar, Check, ImageIcon, Mail, Pencil, Trash2 } from '../Icons';
import { cx } from '../dexUi';

export interface MailHeaderImageChooserProps {
  value: MailHeaderImage;
  onChange: (next: MailHeaderImage) => void;
  /** Event-Foto als Base64. Leer = „Event-Foto" ist nicht wählbar. */
  eventPhotoB64: string;
  /** v32.33: Mail-Logo des Events (EmailImageBase64/_eventLogo). Hat das
   *  Event eins, IST das sein Bild für Mails — der Flow setzt es für
   *  {{ORB_URL}} ein (hero 'logo'). */
  mailLogoB64?: string;
  disabled?: boolean;
  /** Öffnet den Zuschneiden-Dialog. Fehlt er, entfällt der Knopf. */
  onCrop?: () => void;
  /** v31.9.7: Für DIESE Mail hochgeladenes Bild (Base64). Leer = keins. */
  customB64?: string;
  /** Fehlt der Rückruf, entfällt die Kachel „Eigenes Bild" ganz — die QR-Mail
   *  speichert ihren Kopf dauerhaft und hat dafür das Mail-Logo des Events. */
  onPickCustom?: (file: File) => void;
  onRemoveCustom?: () => void;
  /** Läuft gerade die Kompression? */
  customBusy?: boolean;
  /** Ergebnis der letzten Auswahl (Größe bzw. Ablehnungsgrund). */
  customNote?: string;
  isDe: boolean;
}

/**
 * Die Auswahl „Standard-Logo | Event-Foto" plus Zuschneiden — identisch in
 * jedem Mail-Dialog. Die Maße stehen im „HEADER-BILD"-Block des Editors (s. oben).
 */
export default function MailHeaderImageChooser(props: MailHeaderImageChooserProps): React.ReactElement {
  const { value, onChange, eventPhotoB64, disabled, isDe, customB64, onPickCustom, onRemoveCustom, customBusy, customNote, mailLogoB64 } = props;
  const fileRef = React.useRef<HTMLInputElement | null>(null);
  const noPhoto = isDe ? 'Dieses Event hat kein Bild hinterlegt.' : 'This event has no image set.';
  // v31.2: Aus der schmalen Reiter-Reihe werden zwei Kacheln mit Vorschau und
  // einer Zeile Folge — der frühere Statussatz unter der Reihe („Das Event-Foto
  // erscheint im Mail-Kopf") steht jetzt in der Kachel selbst, wo er beim
  // Entscheiden gelesen wird, nicht erst danach.
  // v32.33: Drei Fragen statt Technik-Begriffen (Nutzer-Ansage 29.09.2026:
  // „wenn ein Event ein Event-Foto hat, ist Event-Foto Standard — sonst DEX-
  // Logo oder die Möglichkeit, ein eigenes Foto hochzuladen"). „Standard-Logo"
  // hieß bis v32.32 je nach Event das Mail-Logo ODER der Orb — man sah der
  // Kachel nicht an, was kommt. Jetzt:
  //  - „Event-Bild" = das Mail-Logo des Events (hero 'logo', vom Flow
  //    eingesetzt), sonst das Event-Foto (hero 'event', eingebacken);
  //  - „DEX-Logo" = immer der Orb (hero 'orb', eingebacken; ohne Mail-Logo
  //    setzt ihn auch der Flow, dann zählt hero 'logo' ebenfalls als Orb);
  //  - „Eigenes Bild" wie bisher.
  const eventBild = mailLogoB64 || eventPhotoB64;
  type Kachel = 'eventbild' | 'orb' | 'custom';
  const aktiv: Kachel = value.hero === 'custom' ? 'custom'
    : value.hero === 'orb' ? 'orb'
    : value.hero === 'event' ? 'eventbild'
    : (mailLogoB64 ? 'eventbild' : 'orb');
  const opts: Array<{ key: Kachel; label: string; desc: string; enabled: boolean; icon: React.ReactNode }> = [
    {
      key: 'eventbild', enabled: !!eventBild,
      icon: eventBild
        ? <img src={eventBild} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
        : <Calendar size={18} strokeWidth={2} />,
      label: isDe ? 'Event-Bild' : 'Event image',
      desc: eventBild
        ? (isDe ? 'Das Bild dieses Events steht oben in der Mail.' : 'This event’s image sits at the top of the email.')
        : noPhoto,
    },
    {
      key: 'orb', enabled: true, icon: <Mail size={18} />,
      label: isDe ? 'DEX-Logo' : 'DEX logo',
      desc: isDe ? 'Das runde DEX-Logo, 300 px breit.' : 'The round DEX logo, 300 px wide.',
    },
  ];
  // v31.9.7: Ein Bild nur für DIESE Mail. Nutzer-Frage 10.09.2026: „warum kann
  // ich kein eigenes Foto auswählen bzw. hochladen für den Header?" — es gab
  // keinen Grund, die Kachel fehlte einfach. Sie erscheint nur dort, wo der
  // Aufrufer den Rückruf mitgibt: Die QR-Mail speichert ihren Kopf dauerhaft,
  // ein Bild ohne Speicherort wäre dort eine Auswahl, die beim nächsten Öffnen
  // weg ist.
  if (onPickCustom) {
    opts.push({
      key: 'custom', enabled: true,
      icon: customB64
        ? <img src={customB64} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
        : <ImageIcon size={18} />,
      label: isDe ? 'Eigenes Bild' : 'Own image',
      desc: customB64
        ? (isDe ? 'Dein hochgeladenes Bild steht oben in der Mail.' : 'Your uploaded image sits at the top of the email.')
        : (isDe ? 'Ein Bild nur für diese Mail hochladen.' : 'Upload an image just for this email.'),
    });
  }
  return (
    <div className="dex-ui-stack" style={{ gap: 8, marginBottom: 10 }}>
      <div className="dex-ui-label" style={{ marginBottom: 0 }}>
        {isDe ? 'Welches Bild steht oben in der Mail?' : 'Which image sits at the top of the email?'}
      </div>
      <div className="dex-ui-grid-2" style={{ gap: 10 }}>
        {/* v31.79: Die Größe steht als VIERTE Kachel direkt bei der Bildwahl
            (Nutzer-Ansage 22.09.2026: „kann man die Schnellauswahl der Breite
            nicht einfach hier als 4. Kachel machen … das hier hochziehen …
            Volle Breite und Standard (300 px)"). Vorher lag sie als Aufklapper
            „Kopfbild" weiter unten im Editor — Bild wählen oben, Größe unten,
            dazwischen der ganze Mailkopf. Die Rechnung dahinter ist dieselbe
            wie in der Formregel: Banner → 600/0/0, sonst 300/24/24. */}
        {opts.map(opt => {
          const active = aktiv === opt.key;
          return (
            <button
              key={opt.key}
              type="button"
              className={cx('dex-ui-choice', active && 'is-active', !opt.enabled && 'is-disabled')}
              disabled={disabled || !opt.enabled}
              aria-pressed={active}
              onClick={() => {
                if (!opt.enabled) return;
                // Ohne Bild führt der Klick direkt zur Dateiauswahl —
                // eine Kachel auszuwählen, die nichts zeigt, wäre ein
                // Zwischenschritt ohne Zweck.
                if (opt.key === 'custom' && !customB64) { fileRef.current?.click(); return; }
                // v31.76: Beim Wechsel auf ein FOTO richten sich die Maße nach
                // seiner Form: rund/quadratisch → 300 px, sonst volle Breite
                // (kopfMasseFuerBild). Vorher blieben die Maße des Mail-Logos
                // (600/0/0) stehen, und der runde Event-Kreis füllte die ganze
                // Mail. Das Standard-Logo behält seine Maße.
                if (opt.key === 'orb') { onChange({ ...value, hero: 'orb', ...KOPF_RUND }); return; }
                if (opt.key === 'eventbild') {
                  const hero = mailLogoB64 ? 'logo' : 'event';
                  void bildMasse(eventBild).then(m => onChange({ ...value, hero, ...kopfMasseFuerBild(m.width, m.height) }));
                  return;
                }
                void bildMasse(customB64 || '').then(m => onChange({ ...value, hero: 'custom', ...kopfMasseFuerBild(m.width, m.height) }));
              }}
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
        {(() => {
          const vollOn = istVolleBreite(value);
          const standardOn = value.width === KOPF_RUND.width && value.paddingV === KOPF_RUND.paddingV && value.paddingH === KOPF_RUND.paddingH;
          const setze = (w: number, pv: number, ph: number): void => onChange({ ...value, width: w, paddingV: pv, paddingH: ph });
          const clamp = (raw: string, max: number, fallback: number): number => {
            const n = parseInt(raw, 10);
            return isNaN(n) ? fallback : Math.max(0, Math.min(max, n));
          };
          const lbl: React.CSSProperties = { display: 'flex', flexDirection: 'column', gap: 3, fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-gray-600)' };
          const inp: React.CSSProperties = { width: 72, fontSize: '0.82rem' };
          return (
            <div className="dex-ui-card dex-ui-card--soft" style={{ padding: '10px 12px', display: 'flex', flexDirection: 'column', gap: 8 }}>
              <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', gap: 8 }}>
                <span className="dex-ui-choice-title">{isDe ? 'Wie groß im Kopf?' : 'How large in the header?'}</span>
                <span className="dex-ui-muted" style={{ whiteSpace: 'nowrap' }}>{vollOn ? (isDe ? 'Volle Breite' : 'Full width') : `${value.width} px`}</span>
              </div>
              <div className="dex-ui-inline">
                <button
                  type="button"
                  className={cx('dex-ui-chip', vollOn && 'is-active')}
                  disabled={disabled}
                  onClick={() => setze(KOPF_VOLLE_BREITE.width, KOPF_VOLLE_BREITE.paddingV, KOPF_VOLLE_BREITE.paddingH)}
                  title={isDe ? 'Bild füllt den Kopf über die volle Breite (Höhe passt sich an) — für breite Bilder' : 'Image fills the header edge to edge (height adjusts) — for wide images'}
                >{vollOn && <Check size={12} />}{isDe ? 'Volle Breite' : 'Full width'}</button>
                <button
                  type="button"
                  className={cx('dex-ui-chip', standardOn && 'is-active')}
                  disabled={disabled}
                  onClick={() => setze(KOPF_RUND.width, KOPF_RUND.paddingV, KOPF_RUND.paddingH)}
                  title={isDe ? '300 px breit, mittig, mit Rand — für runde und quadratische Bilder' : '300 px wide, centered, with spacing — for round and square images'}
                >{standardOn && <Check size={12} />}{isDe ? 'Standard (300 px)' : 'Default (300 px)'}</button>
              </div>
              <div style={{ display: 'flex', alignItems: 'flex-end', gap: 10, flexWrap: 'wrap' }}>
                <label style={lbl}>{isDe ? 'Breite' : 'Width'}
                  <input className="dex-ui-input dex-ui-input--sm" type="number" min={80} max={600} step={10} value={value.width} disabled={disabled} onChange={e => onChange({ ...value, width: Math.max(80, clamp(e.target.value, 600, value.width)) })} style={inp} />
                </label>
                <label style={lbl}>{isDe ? 'Seitlich' : 'Sides'}
                  <input className="dex-ui-input dex-ui-input--sm" type="number" min={0} max={80} step={2} value={value.paddingH} disabled={disabled} onChange={e => onChange({ ...value, paddingH: clamp(e.target.value, 80, value.paddingH) })} style={inp} />
                </label>
                <label style={lbl}>{isDe ? 'Oben/unten' : 'Top/bottom'}
                  <input className="dex-ui-input dex-ui-input--sm" type="number" min={0} max={80} step={2} value={value.paddingV} disabled={disabled} onChange={e => onChange({ ...value, paddingV: clamp(e.target.value, 80, value.paddingV) })} style={inp} />
                </label>
              </div>
            </div>
          );
        })()}
      </div>
      {/* v31.9.7: Das Dateifeld liegt außerhalb der Kacheln — ein `<input>` in
          einem `<button>` wäre kein gültiges HTML. Die Kachel löst es aus. */}
      {onPickCustom && (
        <input
          ref={fileRef}
          type="file"
          accept="image/*"
          style={{ display: 'none' }}
          onChange={(e) => {
            const f = e.target.files && e.target.files[0];
            // Wert leeren, sonst löst dieselbe Datei beim zweiten Mal kein
            // `change` aus und der Knopf wirkt kaputt.
            e.target.value = '';
            if (f) onPickCustom(f);
          }}
        />
      )}
      {value.hero === 'custom' && onPickCustom && (
        <div className="dex-ui-inline">
          <button type="button" className="dex-ui-textbtn" disabled={disabled || customBusy} onClick={() => fileRef.current?.click()}>
            <ImageIcon size={14} />{customB64 ? (isDe ? 'Anderes Bild wählen' : 'Choose another image') : (isDe ? 'Bild auswählen' : 'Choose image')}
          </button>
          {!!customB64 && onRemoveCustom && (
            <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" disabled={disabled || customBusy} onClick={onRemoveCustom}>
              <Trash2 size={14} />{isDe ? 'Entfernen' : 'Remove'}
            </button>
          )}
          {/* v31.10: `flex: 1 1 200px` — der Satz stand ohne Basis neben den
              Knöpfen und wurde auf dem Handy zu einer 130 px breiten Textsäule.
              Mit der Basis rutscht er als Ganzes in die nächste Zeile, sobald es
              eng wird, und bleibt auf dem Desktop daneben stehen. */}
          <span className="dex-ui-muted" style={{ flex: '1 1 200px' }}>
            {customBusy
              ? (isDe ? 'Bild wird verkleinert …' : 'Compressing image …')
              : (customNote || (isDe ? 'Wird fest in die Mail eingebacken — der Empfänger muss nichts nachladen.' : 'Baked into the email — the recipient does not have to load anything.'))}
          </span>
        </div>
      )}
      {/* v31.2: Zuschneiden gehört zum Foto — deshalb direkt unter der Wahl und
          nur, wenn das Foto gewählt ist (ein Knopf im Knopf wäre kein gültiges HTML). */}
      {value.hero === 'event' && !!eventPhotoB64 && props.onCrop && (
        <div className="dex-ui-inline">
          <button type="button" className="dex-ui-textbtn" disabled={disabled} onClick={props.onCrop}>
            <Pencil size={14} />{isDe ? 'Foto zuschneiden' : 'Crop photo'}
          </button>
          <span className="dex-ui-muted" style={{ flex: '1 1 200px' }}>{isDe ? 'Ausschnitt wählen, bevor die Mail rausgeht.' : 'Pick the crop before the email goes out.'}</span>
        </div>
      )}
    </div>
  );
}
