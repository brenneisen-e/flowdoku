import * as React from 'react';
import Modal from './Modal';

/**
 * v23.15/v23.17: Bild-Editor zum Zuschneiden des Event-Bildes.
 *
 * - Live-Canvas-Vorschau: zeigt GENAU das Ergebnis, das gespeichert und überall
 *   (Anmeldeseite, Karte, Mail) verwendet wird — so kann der Organizer Zoom,
 *   Position und Form direkt beurteilen.
 * - Zoom (Slider) + Verschieben (Maus-Drag).
 * - Form: Rechteck/Quadrat ODER Kreis. Beim Kreis zusätzlich ein „Rand"-Regler:
 *   0 = Kreis füllt den Rahmen (Ecken transparent), höher = Kreis kleiner und
 *   zentriert mit WEISSEM Rand außen herum (sieht „von weiter weg" sauber aus).
 * - „Übernehmen" liefert das Ergebnis als PNG-Data-URL + File zurück (PNG, damit
 *   transparente Kreis-Ecken erhalten bleiben).
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

const FRAME = 320; // Anzeige-Kantenlänge der Vorschau (px)
const OUT = 700;   // Ausgabe-Kantenlänge (px)

export default function ImageCropModal({ open, src, isDe, onClose, onApply }: Props): React.ReactElement | null {
  const [shape, setShape] = React.useState<'rect' | 'circle'>('circle');
  const [zoom, setZoom] = React.useState(1);
  const [padding, setPadding] = React.useState(0); // 0..0.35 — weißer Rand um den Kreis
  const [offset, setOffset] = React.useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const [nat, setNat] = React.useState<{ w: number; h: number } | null>(null);
  const [error, setError] = React.useState('');
  const imgRef = React.useRef<HTMLImageElement | null>(null);
  const canvasRef = React.useRef<HTMLCanvasElement | null>(null);
  const dragRef = React.useRef<{ startX: number; startY: number; baseX: number; baseY: number } | null>(null);

  // Bild laden, State zurücksetzen beim Öffnen.
  React.useEffect(() => {
    if (!open || !src) return;
    setError('');
    setZoom(1);
    setPadding(0);
    setOffset({ x: 0, y: 0 });
    setNat(null);
    imgRef.current = null;
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => { imgRef.current = img; setNat({ w: img.naturalWidth || 1, h: img.naturalHeight || 1 }); };
    img.onerror = () => setError(isDe ? 'Bild konnte nicht geladen werden.' : 'Could not load image.');
    img.src = src;
  }, [open, src, isDe]);

  // Zeichen-Routine — identisch für Live-Vorschau (size=FRAME) und Export (size=OUT).
  const drawTo = React.useCallback((canvas: HTMLCanvasElement | null, size: number): void => {
    if (!canvas || !nat || !imgRef.current) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    canvas.width = size;
    canvas.height = size;
    ctx.clearRect(0, 0, size, size);
    const isCircle = shape === 'circle';
    const pad = isCircle ? padding : 0;
    // Weißer Hintergrund nur, wenn ein Kreis-Rand gewünscht ist (sonst
    // transparente Ecken bzw. randloses Rechteck).
    if (isCircle && pad > 0) {
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, size, size);
    }
    // Bild-Geometrie in FRAME-Einheiten berechnen, dann auf `size` skalieren.
    const baseScaleF = Math.max(FRAME / nat.w, FRAME / nat.h);
    const effF = baseScaleF * zoom;
    const dwF = nat.w * effF;
    const dhF = nat.h * effF;
    const imgLeftF = (FRAME - dwF) / 2 + offset.x;
    const imgTopF = (FRAME - dhF) / 2 + offset.y;
    const k = size / FRAME;
    ctx.save();
    if (isCircle) {
      const r = (size / 2) * (1 - pad);
      ctx.beginPath();
      ctx.arc(size / 2, size / 2, r, 0, Math.PI * 2);
      ctx.closePath();
      ctx.clip();
    }
    ctx.drawImage(imgRef.current, imgLeftF * k, imgTopF * k, dwF * k, dhF * k);
    ctx.restore();
  }, [nat, shape, padding, zoom, offset]);

  // Live-Vorschau neu zeichnen, wenn sich etwas ändert.
  React.useEffect(() => {
    drawTo(canvasRef.current, FRAME);
  }, [drawTo]);

  if (!open) return null;

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
    if (!nat || !imgRef.current) return;
    try {
      const canvas = document.createElement('canvas');
      drawTo(canvas, OUT);
      const dataUrl = canvas.toDataURL('image/png');
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

  return (
    <Modal open={open} onClose={onClose} maxWidth={440} padding={24} ariaLabel={isDe ? 'Bild zuschneiden' : 'Crop image'}>
      <h3 style={{ marginTop: 0, marginBottom: 4, color: 'var(--dex-green-dark, #4a7c1f)' }}>
        {isDe ? 'Bild zuschneiden' : 'Crop image'}
      </h3>
      <p style={{ marginTop: 0, fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
        {isDe
          ? 'So wird das Bild gespeichert und überall angezeigt (Anmeldeseite, Karte, Mail). Ziehen zum Verschieben, Slider zum Zoomen.'
          : 'This is exactly how the image will be saved and shown everywhere (registration page, card, mail). Drag to move, slider to zoom.'}
      </p>

      {/* Live-Canvas-Vorschau = exaktes Ergebnis */}
      <div style={{ display: 'flex', justifyContent: 'center' }}>
        <canvas
          ref={canvasRef}
          width={FRAME}
          height={FRAME}
          onMouseDown={onPointerDown}
          onMouseMove={onPointerMove}
          onMouseUp={endDrag}
          onMouseLeave={endDrag}
          style={{
            width: FRAME, height: FRAME, cursor: 'grab', userSelect: 'none',
            borderRadius: 12, boxShadow: 'inset 0 0 0 1px var(--dex-gray-200)',
            background: '#f3f3f1', touchAction: 'none',
          }}
        />
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
        <input type="range" min={1} max={4} step={0.01} value={zoom} onChange={e => setZoom(parseFloat(e.target.value))} style={{ width: '100%' }} />
      </div>

      {/* Kreis-Rand (nur im Kreis-Modus) */}
      {shape === 'circle' && (
        <div style={{ marginTop: 10 }}>
          <label style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
            {isDe ? 'Weißer Rand um den Kreis' : 'White margin around the circle'}
          </label>
          <input type="range" min={0} max={0.35} step={0.01} value={padding} onChange={e => setPadding(parseFloat(e.target.value))} style={{ width: '100%' }} />
        </div>
      )}

      {error && <p style={{ color: 'var(--dex-red, #c00)', fontSize: '0.8rem', marginTop: 8 }}>{error}</p>}

      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 18 }}>
        <button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
        <button type="button" className="btn btn-primary" onClick={apply} disabled={!nat}>{isDe ? 'Übernehmen' : 'Apply'}</button>
      </div>
    </Modal>
  );
}
