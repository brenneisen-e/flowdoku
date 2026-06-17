import * as React from 'react';
import Modal from './Modal';

/**
 * v23.15: Bild-Editor zum Zuschneiden des Event-Bildes.
 *
 * - Zoom (Slider) + Verschieben (Ziehen mit der Maus) positionieren das Bild
 *   im quadratischen Zuschnitt-Rahmen.
 * - Form: Rechteck/Quadrat ODER Kreis (Kreis schneidet die Ecken transparent
 *   weg — praktisch, um z.B. einen schwarzen Rand loszuwerden).
 * - „Übernehmen" rendert den sichtbaren Ausschnitt auf ein Canvas und liefert
 *   das Ergebnis als Data-URL + File zurück (PNG, damit der transparente
 *   Kreis-Hintergrund erhalten bleibt).
 *
 * Hinweis: Bei bereits hochgeladenen Bildern (SharePoint-URL) kann das Canvas
 * aus CORS-Gründen „tainted" sein — dann scheitert das Exportieren und wir
 * zeigen einen Hinweis (Bild bitte neu auswählen). Bei frisch ausgewählten
 * Bildern (Data-URL) tritt das nicht auf.
 */
interface Props {
  open: boolean;
  src: string;
  isDe: boolean;
  onClose: () => void;
  onApply: (dataUrl: string, file: File) => void;
}

const FRAME = 320; // Anzeige-Kantenlänge des Zuschnitt-Rahmens (px)
const OUT = 700;   // Ausgabe-Kantenlänge (px)

export default function ImageCropModal({ open, src, isDe, onClose, onApply }: Props): React.ReactElement | null {
  const [shape, setShape] = React.useState<'rect' | 'circle'>('circle');
  const [zoom, setZoom] = React.useState(1);
  const [offset, setOffset] = React.useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const [nat, setNat] = React.useState<{ w: number; h: number } | null>(null);
  const [error, setError] = React.useState('');
  const dragRef = React.useRef<{ startX: number; startY: number; baseX: number; baseY: number } | null>(null);

  // Bild-Naturgröße laden, State zurücksetzen beim Öffnen.
  React.useEffect(() => {
    if (!open || !src) return;
    setError('');
    setZoom(1);
    setOffset({ x: 0, y: 0 });
    setNat(null);
    const img = new Image();
    img.onload = () => setNat({ w: img.naturalWidth || 1, h: img.naturalHeight || 1 });
    img.onerror = () => setError(isDe ? 'Bild konnte nicht geladen werden.' : 'Could not load image.');
    img.src = src;
  }, [open, src, isDe]);

  if (!open) return null;

  // baseScale = „cover" bei Zoom 1 (Bild füllt den Rahmen vollständig).
  const baseScale = nat ? Math.max(FRAME / nat.w, FRAME / nat.h) : 1;
  const eff = baseScale * zoom;
  const dw = nat ? nat.w * eff : FRAME;
  const dh = nat ? nat.h * eff : FRAME;
  const imgLeft = (FRAME - dw) / 2 + offset.x;
  const imgTop = (FRAME - dh) / 2 + offset.y;

  const onPointerDown = (e: React.MouseEvent): void => {
    dragRef.current = { startX: e.clientX, startY: e.clientY, baseX: offset.x, baseY: offset.y };
  };
  const onPointerMove = (e: React.MouseEvent): void => {
    if (!dragRef.current) return;
    setOffset({
      x: dragRef.current.baseX + (e.clientX - dragRef.current.startX),
      y: dragRef.current.baseY + (e.clientY - dragRef.current.startY),
    });
  };
  const endDrag = (): void => { dragRef.current = null; };

  const apply = (): void => {
    if (!nat) return;
    try {
      const canvas = document.createElement('canvas');
      canvas.width = OUT;
      canvas.height = OUT;
      const ctx = canvas.getContext('2d');
      if (!ctx) { setError(isDe ? 'Zuschneiden nicht möglich.' : 'Cropping not possible.'); return; }
      const k = OUT / FRAME;
      if (shape === 'circle') {
        ctx.save();
        ctx.beginPath();
        ctx.arc(OUT / 2, OUT / 2, OUT / 2, 0, Math.PI * 2);
        ctx.closePath();
        ctx.clip();
      }
      const img = new Image();
      img.onload = () => {
        try {
          ctx.drawImage(img, imgLeft * k, imgTop * k, dw * k, dh * k);
          if (shape === 'circle') ctx.restore();
          const dataUrl = canvas.toDataURL('image/png');
          // Data-URL → File (für den Upload-Pfad).
          const byteString = atob(dataUrl.split(',')[1]);
          const ab = new ArrayBuffer(byteString.length);
          const ia = new Uint8Array(ab);
          for (let i = 0; i < byteString.length; i++) ia[i] = byteString.charCodeAt(i);
          const blob = new Blob([ab], { type: 'image/png' });
          const file = new File([blob], 'event-image.png', { type: 'image/png' });
          onApply(dataUrl, file);
        } catch {
          setError(isDe
            ? 'Export fehlgeschlagen (vermutlich CORS bei einem bereits hochgeladenen Bild). Bitte das Bild neu auswählen und dann zuschneiden.'
            : 'Export failed (likely CORS on an already uploaded image). Please re-select the image and then crop.');
        }
      };
      img.onerror = () => setError(isDe ? 'Bild konnte nicht geladen werden.' : 'Could not load image.');
      img.crossOrigin = 'anonymous';
      img.src = src;
    } catch {
      setError(isDe ? 'Zuschneiden fehlgeschlagen.' : 'Cropping failed.');
    }
  };

  return (
    <Modal open={open} onClose={onClose} maxWidth={420} padding={24} ariaLabel={isDe ? 'Bild zuschneiden' : 'Crop image'}>
      <h3 style={{ marginTop: 0, marginBottom: 4, color: 'var(--dex-green-dark, #4a7c1f)' }}>
        {isDe ? 'Bild zuschneiden' : 'Crop image'}
      </h3>
      <p style={{ marginTop: 0, fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
        {isDe
          ? 'Ziehen zum Verschieben, Slider zum Zoomen. Mit „Kreis" werden die Ecken transparent zugeschnitten (z.B. um einen schwarzen Rand zu entfernen).'
          : 'Drag to move, slider to zoom. „Circle" crops the corners transparently (e.g. to remove a black border).'}
      </p>

      {/* Zuschnitt-Vorschau */}
      <div
        onMouseDown={onPointerDown}
        onMouseMove={onPointerMove}
        onMouseUp={endDrag}
        onMouseLeave={endDrag}
        style={{
          width: FRAME, height: FRAME, margin: '0 auto', position: 'relative', overflow: 'hidden',
          borderRadius: shape === 'circle' ? '50%' : 12,
          background: '#0000000d', cursor: 'grab', userSelect: 'none',
          boxShadow: 'inset 0 0 0 2px var(--dex-green, #86bc25)',
          touchAction: 'none',
        }}
      >
        {nat && (
          <img
            src={src}
            alt={isDe ? 'Zuschnitt-Vorschau' : 'Crop preview'}
            draggable={false}
            style={{ position: 'absolute', left: imgLeft, top: imgTop, width: dw, height: dh, maxWidth: 'none', pointerEvents: 'none' }}
          />
        )}
      </div>

      {/* Form-Wahl */}
      <div style={{ display: 'flex', gap: 8, justifyContent: 'center', marginTop: 14 }}>
        <button type="button" className={`btn ${shape === 'circle' ? 'btn-primary' : 'btn-secondary'}`} style={{ fontSize: '0.82rem' }} onClick={() => setShape('circle')}>
          {isDe ? 'Kreis' : 'Circle'}
        </button>
        <button type="button" className={`btn ${shape === 'rect' ? 'btn-primary' : 'btn-secondary'}`} style={{ fontSize: '0.82rem' }} onClick={() => setShape('rect')}>
          {isDe ? 'Quadrat' : 'Square'}
        </button>
      </div>

      {/* Zoom */}
      <div style={{ marginTop: 14 }}>
        <label style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>{isDe ? 'Zoom' : 'Zoom'}</label>
        <input
          type="range" min={1} max={4} step={0.01} value={zoom}
          onChange={e => setZoom(parseFloat(e.target.value))}
          style={{ width: '100%' }}
        />
      </div>

      {error && <p style={{ color: 'var(--dex-red, #c00)', fontSize: '0.8rem', marginTop: 8 }}>{error}</p>}

      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 18 }}>
        <button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
        <button type="button" className="btn btn-primary" onClick={apply} disabled={!nat}>{isDe ? 'Übernehmen' : 'Apply'}</button>
      </div>
    </Modal>
  );
}
