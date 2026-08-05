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
 * v27.5: Optionaler `allowAspect`-Modus für Kopfbilder (Mail-/Outlook-Header):
 *   frei wählbares Seitenverhältnis (Quadrat / 3:2 / 16:9 / Banner). Dadurch
 *   kann ein Querformat-Foto zu einem breiten Banner zugeschnitten werden und
 *   der Organizer kann oben/unten wegschneiden (Bild vertikal verschieben).
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
  /** v23.25: optionaler Zusatzblock (z.B. „Darstellung pro Ansicht") —
      wird unter den Zuschnitt-Reglern, über den Aktions-Buttons gerendert. */
  children?: React.ReactNode;
  /** v27.5: Seitenverhältnis frei wählbar (für Kopfbilder). Blendet die
      Kreis/Quadrat-Wahl aus und zeigt stattdessen Aspekt-Presets. */
  allowAspect?: boolean;
  /** v27.5: Start-Seitenverhältnis (Breite/Höhe) im allowAspect-Modus.
      Default 16/9 (breites Banner). Ignoriert, wenn allowAspect nicht gesetzt. */
  defaultAspect?: number;
  /** v28.10: „Empfohlen"-Kennzeichnung am Kreis-Zuschnitt + Hinweiszeile —
      fürs Event-Bild, das auf der Anmeldeseite als Kreis oben mittig in die
      Karte eingebaut wird. */
  recommendCircle?: boolean;
}

const FRAME = 320; // Anzeige-Breite der Vorschau (px)
const OUT = 700;   // Ausgabe-Breite (px)

// v27.5: Aspekt-Presets für Kopfbilder (Breite : Höhe).
const ASPECT_PRESETS: Array<{ a: number; de: string; en: string }> = [
  { a: 1, de: 'Quadrat', en: 'Square' },
  { a: 3 / 2, de: '3:2', en: '3:2' },
  { a: 16 / 9, de: '16:9', en: '16:9' },
  { a: 5 / 2, de: 'Banner', en: 'Banner' },
];

export default function ImageCropModal({ open, src, isDe, onClose, onApply, children, allowAspect, defaultAspect, recommendCircle }: Props): React.ReactElement | null {
  const [shape, setShape] = React.useState<'rect' | 'circle'>('circle');
  const [aspect, setAspect] = React.useState<number>(allowAspect ? (defaultAspect || 16 / 9) : 1);
  const [zoom, setZoom] = React.useState(1);
  const [padding, setPadding] = React.useState(0); // 0..0.35 — weißer Rand um den Kreis
  const [offset, setOffset] = React.useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const [nat, setNat] = React.useState<{ w: number; h: number } | null>(null);
  const [error, setError] = React.useState('');
  const imgRef = React.useRef<HTMLImageElement | null>(null);
  const canvasRef = React.useRef<HTMLCanvasElement | null>(null);
  // v26.x: `scale` = FRAME-Einheiten pro gerendertem CSS-Pixel. Auf dem Handy
  // wird die Canvas responsiv kleiner als FRAME (320) gerendert; die Drag-Deltas
  // (CSS-Pixel) müssen daher in FRAME-Einheiten umgerechnet werden, damit der
  // Offset (der in FRAME-Einheiten in drawTo einfließt) weiterhin stimmt und der
  // Zuschnitt korrekt auf das Ausgabebild (OUT) abgebildet wird.
  const dragRef = React.useRef<{ startX: number; startY: number; baseX: number; baseY: number; scale: number } | null>(null);

  // v27.5: im allowAspect-Modus ist die Form immer ein Rechteck (kein Kreis).
  const isCircle = !allowAspect && shape === 'circle' && aspect === 1;

  // Bild laden, State zurücksetzen beim Öffnen.
  React.useEffect(() => {
    if (!open || !src) return;
    setError('');
    setZoom(1);
    setPadding(0);
    setOffset({ x: 0, y: 0 });
    setAspect(allowAspect ? (defaultAspect || 16 / 9) : 1);
    setNat(null);
    imgRef.current = null;
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => { imgRef.current = img; setNat({ w: img.naturalWidth || 1, h: img.naturalHeight || 1 }); };
    img.onerror = () => setError(isDe ? 'Bild konnte nicht geladen werden.' : 'Could not load image.');
    img.src = src;
  }, [open, src, isDe, allowAspect, defaultAspect]);

  // Zeichen-Routine — identisch für Live-Vorschau (sizeW=FRAME) und Export (sizeW=OUT).
  // v27.5: nicht mehr quadratisch — Höhe ergibt sich aus dem Seitenverhältnis.
  const drawTo = React.useCallback((canvas: HTMLCanvasElement | null, sizeW: number): void => {
    if (!canvas || !nat || !imgRef.current) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const a = aspect || 1;
    const sizeH = Math.round(sizeW / a);
    const frameW = FRAME;
    const frameH = FRAME / a;
    canvas.width = sizeW;
    canvas.height = sizeH;
    ctx.clearRect(0, 0, sizeW, sizeH);
    const pad = isCircle ? padding : 0;
    // Weißer Hintergrund nur, wenn ein Kreis-Rand gewünscht ist (sonst
    // transparente Ecken bzw. randloses Rechteck).
    if (isCircle && pad > 0) {
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, sizeW, sizeH);
    }
    // Bild-Geometrie in FRAME-Einheiten berechnen, dann auf `sizeW` skalieren.
    const baseScaleF = Math.max(frameW / nat.w, frameH / nat.h);
    const effF = baseScaleF * zoom;
    const dwF = nat.w * effF;
    const dhF = nat.h * effF;
    const imgLeftF = (frameW - dwF) / 2 + offset.x;
    const imgTopF = (frameH - dhF) / 2 + offset.y;
    const k = sizeW / frameW; // gleicher Faktor für x und y (frameH/sizeH == frameW/sizeW)
    ctx.save();
    if (isCircle) {
      const r = (sizeW / 2) * (1 - pad);
      ctx.beginPath();
      ctx.arc(sizeW / 2, sizeH / 2, r, 0, Math.PI * 2);
      ctx.closePath();
      ctx.clip();
    }
    ctx.drawImage(imgRef.current, imgLeftF * k, imgTopF * k, dwF * k, dhF * k);
    ctx.restore();
  }, [nat, isCircle, padding, zoom, offset, aspect]);

  // Live-Vorschau neu zeichnen, wenn sich etwas ändert.
  React.useEffect(() => {
    drawTo(canvasRef.current, FRAME);
  }, [drawTo]);

  if (!open) return null;

  const onPointerDown = (e: React.MouseEvent): void => {
    // Skalierungsfaktor aus der tatsächlich gerenderten Canvas-Breite ableiten
    // (FRAME-Einheiten pro CSS-Pixel). Auf dem Desktop ist die Breite == FRAME,
    // also scale == 1 (unverändertes Verhalten); auf dem Handy < FRAME → scale > 1.
    const rect = e.currentTarget.getBoundingClientRect();
    const scale = rect.width > 0 ? FRAME / rect.width : 1;
    dragRef.current = { startX: e.clientX, startY: e.clientY, baseX: offset.x, baseY: offset.y, scale };
  };
  const onPointerMove = (e: React.MouseEvent): void => {
    if (!dragRef.current) return;
    const { scale } = dragRef.current;
    setOffset({
      x: dragRef.current.baseX + (e.clientX - dragRef.current.startX) * scale,
      y: dragRef.current.baseY + (e.clientY - dragRef.current.startY) * scale,
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
    <Modal open={open} onClose={onClose} maxWidth={560} padding={24} ariaLabel={isDe ? 'Bild zuschneiden' : 'Crop image'}>
      <h3 style={{ marginTop: 0, marginBottom: 4, color: 'var(--dex-green-dark, #4a7c1f)' }}>
        {isDe ? 'Bild zuschneiden' : 'Crop image'}
      </h3>
      <p style={{ marginTop: 0, fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
        {allowAspect
          ? (isDe
            ? 'Wähle ein Seitenverhältnis und ziehe das Bild in Position (z.B. oben/unten wegschneiden). So erscheint es im Mail-/Outlook-Kopf.'
            : 'Pick an aspect ratio and drag the image into place (e.g. crop off top/bottom). This is how it appears in the mail/Outlook header.')
          : (isDe
            ? 'So erscheint das Bild auf der Anmeldeseite und der Event-Karte. In den E-Mails und im Outlook-Termin wird automatisch das unbeschnittene Originalfoto verwendet — dort ist der Kopf rechteckig. Ziehen zum Verschieben, Slider zum Zoomen.'
            : 'This is how the image appears on the registration page and the event card. Emails and the Outlook invite automatically use the uncropped original photo — their header is rectangular. Drag to move, slider to zoom.')}
      </p>

      {/* Live-Canvas-Vorschau = exaktes Ergebnis.
          Das Backing-Store bleibt FRAME-basiert (Zeichen-/Export-Mathematik
          unverändert); per CSS wird die Canvas nur responsiv verkleinert, damit
          sie auf schmalen Karten (Handy ~295px) nicht überläuft. Die Drag-Math
          rechnet die CSS-Pixel über den Skalierungsfaktor in FRAME-Einheiten um. */}
      <div style={{ display: 'flex', justifyContent: 'center' }}>
        <canvas
          ref={canvasRef}
          width={FRAME}
          height={Math.round(FRAME / (aspect || 1))}
          onMouseDown={onPointerDown}
          onMouseMove={onPointerMove}
          onMouseUp={endDrag}
          onMouseLeave={endDrag}
          style={{
            width: '100%', maxWidth: FRAME, aspectRatio: String(aspect || 1), height: 'auto',
            cursor: 'grab', userSelect: 'none',
            borderRadius: 12, boxShadow: 'inset 0 0 0 1px var(--dex-gray-200)',
            background: '#f3f3f1', touchAction: 'none',
          }}
        />
      </div>

      {/* v27.5: Seitenverhältnis-Wahl (nur Kopfbild-Modus) ODER Form-Wahl. */}
      {allowAspect ? (
        <div style={{ display: 'flex', gap: 8, justifyContent: 'center', marginTop: 14, flexWrap: 'wrap' }}>
          {ASPECT_PRESETS.map(p => (
            <button
              key={p.a}
              type="button"
              className={`btn ${Math.abs(aspect - p.a) < 0.001 ? 'btn-primary' : 'btn-secondary'}`}
              style={{ fontSize: '0.82rem' }}
              onClick={() => setAspect(p.a)}
            >
              {isDe ? p.de : p.en}
            </button>
          ))}
        </div>
      ) : (
        <>
        <div style={{ display: 'flex', gap: 8, justifyContent: 'center', marginTop: 14 }}>
          <button type="button" className={`btn ${shape === 'circle' ? 'btn-primary' : 'btn-secondary'}`} style={{ fontSize: '0.82rem', position: 'relative' }} onClick={() => setShape('circle')}>
            {isDe ? 'Kreis' : 'Circle'}
            {recommendCircle && (
              <span style={{ marginLeft: 6, padding: '1px 7px', borderRadius: 999, background: 'var(--dex-orange, #ed8b00)', color: '#fff', fontSize: '0.62rem', fontWeight: 800, letterSpacing: 0.4, textTransform: 'uppercase', verticalAlign: 'middle' }}>
                {isDe ? 'Empfohlen' : 'Recommended'}
              </span>
            )}
          </button>
          <button type="button" className={`btn ${shape === 'rect' ? 'btn-primary' : 'btn-secondary'}`} style={{ fontSize: '0.82rem' }} onClick={() => setShape('rect')}>
            {isDe ? 'Quadrat' : 'Square'}
          </button>
        </div>
        {recommendCircle && (
          <p style={{ margin: '8px 0 0', fontSize: '0.75rem', color: 'var(--dex-gray-500)', textAlign: 'center', lineHeight: 1.45 }}>
            {isDe
              ? 'Tipp: Der kreisförmige Zuschnitt wird empfohlen — das Bild wird auf der Anmeldeseite als Kreis oben mittig in die Event-Karte eingebaut.'
              : 'Tip: the circular crop is recommended — the image is embedded as a circle at the top center of the event card on the registration page.'}
          </p>
        )}
        </>
      )}

      {/* Zoom */}
      <div style={{ marginTop: 14 }}>
        <label style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>{isDe ? 'Zoom' : 'Zoom'}</label>
        <input type="range" min={1} max={4} step={0.01} value={zoom} onChange={e => setZoom(parseFloat(e.target.value))} style={{ width: '100%' }} />
      </div>

      {/* Kreis-Rand (nur im Kreis-Modus) */}
      {isCircle && (
        <div style={{ marginTop: 10 }}>
          <label style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
            {isDe ? 'Weißer Rand um den Kreis' : 'White margin around the circle'}
          </label>
          <input type="range" min={0} max={0.35} step={0.01} value={padding} onChange={e => setPadding(parseFloat(e.target.value))} style={{ width: '100%' }} />
        </div>
      )}

      {error && <p style={{ color: 'var(--dex-red, #c00)', fontSize: '0.8rem', marginTop: 8 }}>{error}</p>}

      {/* v23.25: Zusatzblock (Darstellung pro Ansicht), abgesetzt mit Trennlinie. */}
      {children && (
        <div style={{ marginTop: 18, paddingTop: 16, borderTop: '1px solid var(--dex-gray-200)' }}>
          {children}
        </div>
      )}

      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 18 }}>
        <button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
        <button type="button" className="btn btn-primary" onClick={apply} disabled={!nat}>{isDe ? 'Übernehmen' : 'Apply'}</button>
      </div>
    </Modal>
  );
}
