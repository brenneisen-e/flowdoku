/**
 * Zuschnitt für das Kachelbild.
 *
 * Bewusst KEIN Port des Zuschnitt-Dialogs aus DEX: Der kann Kreis, Rechteck und
 * frei wählbare Seitenverhältnisse, weil ein Event-Bild in Mail, Anmeldeseite
 * und Outlook in verschiedenen Formen steht. Eine Use-Case-Kachel zeigt ihr
 * Bild in genau EINER Form, 16:9 (`aspectRatio` in `UseCaseCard`). Also gibt es
 * hier genau eine Form — und damit nichts zu wählen, was man falsch wählen
 * könnte.
 *
 * Was der Dialog trotzdem können muss:
 *  - Zoom nach unten bis „ganzes Bild sichtbar". Ein quadratisches Logo würde
 *    bei „füllt den Rahmen" oben und unten abgeschnitten; mit dem Regler auf
 *    Anschlag steht es ganz da, auf weißem Grund.
 *  - Verschieben per Maus/Finger UND per Pfeiltasten. Ein Zuschnitt, der nur
 *    mit der Maus geht, ist für Tastaturnutzer keiner.
 *  - Die Vorschau IST das Ergebnis: Sie zeichnet mit derselben Funktion wie der
 *    Export, nur kleiner. Was man sieht, wird gespeichert.
 *
 * Ergebnis ist ein JPEG (1280 × 720, weißer Grund) — ein Foto als PNG wöge
 * mehrere hundert KB, und die Kachelwand lädt zwanzig davon.
 */

import * as React from 'react';
import Modal from './Modal';
import { useLanguage } from '../context/LanguageContext';

interface Props {
  open: boolean;
  /** Das gewählte Bild als Data-URL (so ist die Zeichenfläche nie „tainted"). */
  src: string;
  onClose: () => void;
  onApply: (datei: File, vorschau: string) => void;
}

const VORSCHAU_W = 640;
const AUSGABE_W = 1280;
const ASPEKT = 16 / 9;
const MAX_ZOOM = 4;
const SCHRITT = 0.08;

interface Nat { w: number; h: number }

/**
 * Wo das Bild in einer Fläche `w × h` liegt.
 *
 * `zoom` ist relativ zu „füllt den Rahmen" (1 = die kürzere Seite passt
 * genau). `ox`/`oy` sind auf −1…1 normiert: 0 = mittig, ±1 = bis zum Rand des
 * Überstands. Weil sie normiert sind, gelten sie in der Vorschau UND im Export
 * gleich — sonst wäre der Ausschnitt in der Ausgabe ein anderer als der, den
 * man eingestellt hat.
 */
function lage(nat: Nat, zoom: number, ox: number, oy: number, w: number, h: number): { x: number; y: number; dw: number; dh: number; uebX: number; uebY: number } {
  const fuellt = Math.max(w / nat.w, h / nat.h);
  const s = fuellt * zoom;
  const dw = nat.w * s;
  const dh = nat.h * s;
  const uebX = Math.max(0, dw - w) / 2;
  const uebY = Math.max(0, dh - h) / 2;
  return { x: (w - dw) / 2 + ox * uebX, y: (h - dh) / 2 + oy * uebY, dw, dh, uebX, uebY };
}

function zeichne(ctx: CanvasRenderingContext2D, img: HTMLImageElement, nat: Nat, zoom: number, ox: number, oy: number, w: number, h: number): void {
  // Weißer Grund: Bei „ganzes Bild" bleibt Rand, und transparente PNGs würden
  // beim JPEG-Export sonst schwarz.
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, w, h);
  const g = lage(nat, zoom, ox, oy, w, h);
  ctx.drawImage(img, g.x, g.y, g.dw, g.dh);
}

const klemme = (v: number, lo: number, hi: number): number => Math.min(hi, Math.max(lo, v));

export default function ImageCropModal(props: Props): React.ReactElement | null {
  const { open, src, onClose, onApply } = props;
  const { t } = useLanguage();

  const [nat, setNat] = React.useState<Nat | null>(null);
  const [fehler, setFehler] = React.useState('');
  const [zoom, setZoom] = React.useState(1);
  const [ox, setOx] = React.useState(0);
  const [oy, setOy] = React.useState(0);
  const imgRef = React.useRef<HTMLImageElement | null>(null);
  const canvasRef = React.useRef<HTMLCanvasElement | null>(null);
  const dragRef = React.useRef<{ x: number; y: number; ox: number; oy: number } | null>(null);

  // Bild laden, Stand zurücksetzen. Der Dialog ist beim Schließen abgehängt
  // (`Modal` rendert bei `open=false` nichts) — ohne Zurücksetzen stünde beim
  // nächsten Bild noch der Zoom des vorigen da.
  React.useEffect(() => {
    if (!open || !src) return undefined;
    let abgebrochen = false;
    setNat(null);
    setFehler('');
    setZoom(1);
    setOx(0);
    setOy(0);
    const img = new Image();
    img.onload = () => {
      if (abgebrochen) return;
      imgRef.current = img;
      setNat({ w: img.naturalWidth, h: img.naturalHeight });
    };
    img.onerror = () => {
      if (abgebrochen) return;
      setFehler(t('Dieses Bild lässt sich nicht lesen. Nimm ein JPEG, PNG oder WebP.', 'This image cannot be read. Use a JPEG, PNG or WebP.'));
    };
    img.src = src;
    return () => { abgebrochen = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, src]);

  // „Ganzes Bild sichtbar" — der kleinste Zoom, bei dem nichts abgeschnitten wird.
  const minZoom = nat
    ? Math.min(VORSCHAU_W / nat.w, VORSCHAU_W / ASPEKT / nat.h) / Math.max(VORSCHAU_W / nat.w, VORSCHAU_W / ASPEKT / nat.h)
    : 1;

  // Vorschau zeichnen — bei jeder Änderung.
  React.useEffect(() => {
    const c = canvasRef.current;
    const img = imgRef.current;
    if (!open || !c || !img || !nat) return;
    const ctx = c.getContext('2d');
    if (!ctx) return;
    zeichne(ctx, img, nat, zoom, ox, oy, c.width, c.height);
  }, [open, nat, zoom, ox, oy]);

  const bewege = (dx: number, dy: number): void => {
    // In Einheiten des Überstands: ±1 ist der Rand. Ohne Überstand in einer
    // Richtung (Bild schmaler als der Rahmen) gibt es dort nichts zu verschieben.
    if (!nat) return;
    const g = lage(nat, zoom, ox, oy, VORSCHAU_W, VORSCHAU_W / ASPEKT);
    if (g.uebX > 0) setOx(v => klemme(v + dx / g.uebX, -1, 1));
    if (g.uebY > 0) setOy(v => klemme(v + dy / g.uebY, -1, 1));
  };

  const onPointerDown = (e: React.PointerEvent<HTMLCanvasElement>): void => {
    dragRef.current = { x: e.clientX, y: e.clientY, ox, oy };
    try { e.currentTarget.setPointerCapture(e.pointerId); } catch { /* ältere Browser */ }
  };
  const onPointerMove = (e: React.PointerEvent<HTMLCanvasElement>): void => {
    const d = dragRef.current;
    const c = canvasRef.current;
    if (!d || !c || !nat) return;
    // CSS-Pixel → Zeichenflächen-Einheiten: Auf dem Handy wird die Fläche
    // kleiner dargestellt als sie intern ist.
    const f = VORSCHAU_W / c.getBoundingClientRect().width;
    const g = lage(nat, zoom, 0, 0, VORSCHAU_W, VORSCHAU_W / ASPEKT);
    if (g.uebX > 0) setOx(klemme(d.ox + ((e.clientX - d.x) * f) / g.uebX, -1, 1));
    if (g.uebY > 0) setOy(klemme(d.oy + ((e.clientY - d.y) * f) / g.uebY, -1, 1));
  };
  const onPointerUp = (): void => { dragRef.current = null; };

  const onKeyDown = (e: React.KeyboardEvent<HTMLCanvasElement>): void => {
    const schritt = 24;
    if (e.key === 'ArrowLeft') { bewege(-schritt, 0); e.preventDefault(); }
    else if (e.key === 'ArrowRight') { bewege(schritt, 0); e.preventDefault(); }
    else if (e.key === 'ArrowUp') { bewege(0, -schritt); e.preventDefault(); }
    else if (e.key === 'ArrowDown') { bewege(0, schritt); e.preventDefault(); }
    else if (e.key === '+' || e.key === '=') { setZoom(z => klemme(z + SCHRITT, minZoom, MAX_ZOOM)); e.preventDefault(); }
    else if (e.key === '-') { setZoom(z => klemme(z - SCHRITT, minZoom, MAX_ZOOM)); e.preventDefault(); }
  };

  const uebernehmen = (): void => {
    const img = imgRef.current;
    if (!img || !nat) return;
    const c = document.createElement('canvas');
    c.width = AUSGABE_W;
    c.height = Math.round(AUSGABE_W / ASPEKT);
    const ctx = c.getContext('2d');
    if (!ctx) {
      setFehler(t('Der Browser kann das Bild nicht verarbeiten.', 'The browser cannot process the image.'));
      return;
    }
    zeichne(ctx, img, nat, zoom, ox, oy, c.width, c.height);
    try {
      c.toBlob(blob => {
        if (!blob) {
          setFehler(t('Das Bild ließ sich nicht umwandeln.', 'The image could not be converted.'));
          return;
        }
        onApply(new File([blob], 'kachel.jpg', { type: 'image/jpeg' }), c.toDataURL('image/jpeg', 0.85));
      }, 'image/jpeg', 0.85);
    } catch {
      setFehler(t('Das Bild ließ sich nicht umwandeln.', 'The image could not be converted.'));
    }
  };

  return (
    <Modal
      open={open}
      onClose={onClose}
      maxWidth={560}
      // Über dem Pflegedialog, aus dem er geöffnet wird — die Ebene ist hier
      // Absicht und kein Zufall der Einhäng-Reihenfolge (siehe `Modal.zIndex`).
      zIndex={10000}
      ariaLabel={t('Bild zuschneiden', 'Crop image')}
      title={t('Bild zuschneiden', 'Crop image')}
      subtitle={t('So sieht die Kachel aus. Ziehen verschiebt das Bild, der Regler zoomt.', 'This is how the tile looks. Drag to move the image, use the slider to zoom.')}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={onClose}>{t('Abbrechen', 'Cancel')}</button>
        <button type="button" className="btn btn-primary" disabled={!nat} onClick={uebernehmen}>{t('Übernehmen', 'Apply')}</button>
      </>}
    >
      {fehler && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert"><span>{fehler}</span></div>
      )}

      {!fehler && (
        <>
          <canvas
            ref={canvasRef}
            width={VORSCHAU_W}
            height={Math.round(VORSCHAU_W / ASPEKT)}
            tabIndex={0}
            role="img"
            aria-label={t('Vorschau der Kachel. Mit den Pfeiltasten verschieben, mit Plus und Minus zoomen.', 'Tile preview. Move with the arrow keys, zoom with plus and minus.')}
            onPointerDown={onPointerDown}
            onPointerMove={onPointerMove}
            onPointerUp={onPointerUp}
            onPointerCancel={onPointerUp}
            onKeyDown={onKeyDown}
            style={{
              width: '100%', height: 'auto', display: 'block',
              borderRadius: 12, border: '1px solid var(--dex-gray-200, #e8e8e8)',
              // Ohne das scrollt auf dem Handy die Seite mit, statt das Bild zu bewegen.
              touchAction: 'none', cursor: 'grab', background: '#fff',
            }}
          />
          <label className="dex-ui-field" style={{ marginTop: 4 }}>
            <span className="dex-ui-label">{t('Zoom', 'Zoom')}</span>
            <input
              type="range"
              className="dex-ui-range"
              min={minZoom}
              max={MAX_ZOOM}
              step={0.01}
              value={zoom}
              disabled={!nat}
              onChange={e => setZoom(klemme(parseFloat(e.target.value) || 1, minZoom, MAX_ZOOM))}
            />
            <span className="dex-ui-help">
              {t('Ganz links ist das ganze Bild sichtbar, mit weißem Rand. Weiter rechts füllt es die Kachel.', 'Fully left shows the whole image with a white border. Further right fills the tile.')}
            </span>
          </label>
        </>
      )}
    </Modal>
  );
}
