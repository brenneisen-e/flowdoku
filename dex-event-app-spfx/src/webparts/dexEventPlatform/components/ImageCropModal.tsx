import * as React from 'react';
import Modal from './Modal';
import { cx } from './dexUi';
import { AlertCircle, Check, Info } from './Icons';

/**
 * v23.15/v23.17: Bild-Editor zum Zuschneiden des Event-Bildes.
 *
 * - Live-Canvas-Vorschau: zeigt GENAU das Ergebnis, das gespeichert und überall
 *   (Anmeldeseite, Karte, Mail) verwendet wird — so kann der Organizer Zoom,
 *   Position und Form direkt beurteilen.
 * - Zoom (Slider) + Verschieben (Maus-Drag).
 * - Form: Nicht beschneiden, Rechteck/Quadrat ODER Kreis. Beim Kreis zusätzlich
 *   ein „Rand"-Regler: 0 = Kreis füllt den Rahmen (Ecken transparent), höher =
 *   Kreis kleiner und zentriert mit WEISSEM Rand außen herum.
 *
 * v31.9.7: „Nicht beschneiden" ist die Vorauswahl (Nutzer-Ansage 10.09.2026:
 *   „eventfoto braucht auch die auswahl nicht beschneiden.. und das ist
 *   default.. bei der anmeldeseite ist es ja eh automatisch ein kreis").
 *   Der Grund: Die Anmeldeseite und die Event-Karte zeigen das Foto ohnehin in
 *   ihrer eigenen Form; ein zusätzlicher Zuschnitt im Dialog schnitt nur ein
 *   zweites Mal ab. In diesem Modus behält das Bild sein eigenes
 *   Seitenverhältnis, Zoom und Verschieben entfallen — es gibt nichts zu
 *   wählen, wenn nichts weggeschnitten wird.
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

/** v31.9.7: `none` = das Bild bleibt, wie es ist (Vorauswahl). */
type Shape = 'none' | 'circle' | 'rect';

// v31.2: Symbole nur für diesen Dialog — Icons.tsx wird zentral gepflegt und
// hat (noch) kein Zuschneide-Symbol; ein lokales SVG hält die Datei
// unabhängig, statt im Portal auf die Fluent-Schrift zu warten.
const CropIcon = ({ size = 20 }: { size?: number }): React.ReactElement => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M6 2v14a2 2 0 0 0 2 2h14" /><path d="M18 22V8a2 2 0 0 0-2-2H2" />
  </svg>
);
const ShapeIcon = ({ kind }: { kind: Shape }): React.ReactElement => (
  <svg width={18} height={18} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} aria-hidden="true">
    {kind === 'circle' && <circle cx="12" cy="12" r="9" />}
    {kind === 'rect' && <rect x="3" y="3" width="18" height="18" rx="3" />}
    {/* „Nicht beschneiden": ein Foto-Rahmen mit Inhalt — das ganze Bild bleibt. */}
    {kind === 'none' && <><rect x="2" y="5" width="20" height="14" rx="2" /><path d="M2 16l5-5 4 4 3-3 8 8" /><circle cx="8" cy="10" r="1.4" /></>}
  </svg>
);

// v27.5: Aspekt-Presets für Kopfbilder (Breite : Höhe).
const ASPECT_PRESETS: Array<{ a: number; de: string; en: string }> = [
  { a: 1, de: 'Quadrat', en: 'Square' },
  { a: 3 / 2, de: '3:2', en: '3:2' },
  { a: 16 / 9, de: '16:9', en: '16:9' },
  { a: 5 / 2, de: 'Banner', en: 'Banner' },
];

export default function ImageCropModal({ open, src, isDe, onClose, onApply, children, allowAspect, defaultAspect, recommendCircle }: Props): React.ReactElement | null {
  const [shape, setShape] = React.useState<Shape>('none');
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
  // v31.9.7: „Nicht beschneiden" — das Bild behält sein eigenes Seitenverhältnis
  // und wird vollständig übernommen. Solange die Maße noch nicht bekannt sind
  // (`nat === null`), bleibt es beim quadratischen Rahmen; sobald das Bild
  // geladen ist, richtet sich der Rahmen nach ihm.
  const noCrop = !allowAspect && shape === 'none';
  const effAspect = (noCrop && nat ? nat.w / nat.h : aspect) || 1;

  // Bild laden, State zurücksetzen beim Öffnen.
  React.useEffect(() => {
    if (!open || !src) return;
    setError('');
    setZoom(1);
    setPadding(0);
    setOffset({ x: 0, y: 0 });
    // v31.9.7: Auch die Form zurück auf die Vorauswahl. Ohne das trüge das
    // nächste Bild den Zuschnitt des vorigen — der Dialog gehört zum Bild,
    // nicht zur Sitzung.
    setShape('none');
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
    const a = effAspect;
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
    // v31.9.7: Ohne Zuschnitt sind Zoom und Versatz neutral — sonst könnte man
    // über den Regler doch wieder etwas wegschneiden, obwohl die Kachel das
    // Gegenteil verspricht.
    const z = noCrop ? 1 : zoom;
    const offX = noCrop ? 0 : offset.x;
    const offY = noCrop ? 0 : offset.y;
    const baseScaleF = Math.max(frameW / nat.w, frameH / nat.h);
    const effF = baseScaleF * z;
    const dwF = nat.w * effF;
    const dhF = nat.h * effF;
    const imgLeftF = (frameW - dwF) / 2 + offX;
    const imgTopF = (frameH - dhF) / 2 + offY;
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
  }, [nat, isCircle, noCrop, padding, zoom, offset, effAspect]);

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

  // v31.2: Regler mit Wert daneben — ein Slider ohne Zahl sagt nicht, wie weit
  // man schon ist, und „4,0×" erklärt die Grenze von selbst. Ein <label> um
  // Titel und Eingabe, damit der Klick auf den Titel den Regler fokussiert.
  const slider = (title: string, value: string, input: React.ReactElement, help?: string): React.ReactElement => (
    <label className="dex-ui-field dex-ui-fade-in" style={{ display: 'block', marginTop: 12 }}>
      <span className="dex-ui-label" style={{ justifyContent: 'space-between' }}>
        <span>{title}</span><span className="dex-ui-pill dex-ui-pill--gray">{value}</span>
      </span>
      {React.cloneElement(input, { style: { width: '100%', accentColor: '#86bc25', cursor: 'pointer', margin: 0 } })}
      {help && <span className="dex-ui-help" style={{ display: 'block' }}>{help}</span>}
    </label>
  );
  const shapeChoice = (kind: Shape): React.ReactElement => {
    const on = shape === kind;
    const title = kind === 'none'
      ? (isDe ? 'Nicht beschneiden' : 'Do not crop')
      : kind === 'circle' ? (isDe ? 'Kreis' : 'Circle') : (isDe ? 'Quadrat' : 'Square');
    const desc = kind === 'none'
      ? (isDe ? 'Das Foto bleibt, wie es ist — die Anmeldeseite rundet es ohnehin selbst.' : 'The photo stays as it is — the registration page rounds it off by itself.')
      : kind === 'circle'
        ? (recommendCircle
          ? (isDe ? 'Sitzt rund oben mittig in der Event-Karte.' : 'Sits round at the top centre of the event card.')
          : (isDe ? 'Runder Ausschnitt, Ecken bleiben transparent.' : 'Round crop, corners stay transparent.'))
        : (isDe ? 'Rechteckiger Ausschnitt.' : 'Rectangular crop.');
    return (
      <button type="button" className={cx('dex-ui-choice', on && 'is-active')} style={{ padding: '10px 12px', alignItems: 'center' }} aria-pressed={on} onClick={() => setShape(kind)}>
        <span className="dex-ui-choice-icon" style={{ width: 30, height: 30 }}><ShapeIcon kind={kind} /></span>
        <span className="dex-ui-choice-body" style={{ display: 'block' }}>
          <span className="dex-ui-choice-title" style={{ display: 'block' }}>{title}</span>
          <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{desc}</span>
        </span>
        <span className="dex-ui-choice-check">{on && <Check size={12} />}</span>
      </button>
    );
  };

  return (
    <Modal
      open={open} onClose={onClose} maxWidth={560} ariaLabel={isDe ? 'Bild zuschneiden' : 'Crop image'}
      title={isDe ? 'Bild zuschneiden' : 'Crop image'}
      icon={<CropIcon size={20} />}
      subtitle={allowAspect
        ? (isDe
          ? 'Wähle das Seitenverhältnis und ziehe das Bild in Position — so erscheint es im Mail- und Outlook-Kopf.'
          : 'Pick an aspect ratio and drag the image into place — this is how it appears in the mail and Outlook header.')
        : (isDe
          ? 'So erscheint das Bild auf der Anmeldeseite und der Event-Karte.'
          : 'This is how the image appears on the registration page and the event card.')}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
        <button type="button" className="btn btn-primary" onClick={apply} disabled={!nat}>{isDe ? 'Übernehmen' : 'Apply'}</button>
      </>}
    >
      {/* Live-Canvas-Vorschau = exaktes Ergebnis. Das Backing-Store bleibt
          FRAME-basiert (Zeichen-/Export-Mathematik unverändert); per CSS wird die
          Canvas nur responsiv verkleinert (Handy ~295px). Die Drag-Math rechnet
          CSS-Pixel über den Skalierungsfaktor in FRAME-Einheiten um.
          v31.2: Zoom steht direkt unter der Vorschau — Verschieben und Zoomen
          sind EINE Frage („welcher Ausschnitt?"), die Form eine andere. */}
      <div className="dex-ui-section" style={{ margin: 0 }}>
        <div className="dex-ui-section-title">{noCrop ? (isDe ? 'Vorschau' : 'Preview') : (isDe ? 'Ausschnitt' : 'Crop area')}</div>
        <div style={{ display: 'flex', justifyContent: 'center' }}>
          <canvas
            ref={canvasRef} width={FRAME} height={Math.round(FRAME / effAspect)}
            onMouseDown={noCrop ? undefined : onPointerDown} onMouseMove={noCrop ? undefined : onPointerMove}
            onMouseUp={noCrop ? undefined : endDrag} onMouseLeave={noCrop ? undefined : endDrag}
            title={noCrop ? undefined : (isDe ? 'Ziehen, um das Bild zu verschieben' : 'Drag to move the image')}
            style={{
              width: '100%', maxWidth: FRAME, aspectRatio: String(effAspect), height: 'auto',
              cursor: noCrop ? 'default' : 'grab', userSelect: 'none',
              borderRadius: 12, boxShadow: 'inset 0 0 0 1px var(--dex-gray-200)', background: '#f3f3f1', touchAction: 'none',
            }}
          />
        </div>
        <p className="dex-ui-help" style={{ textAlign: 'center', margin: '8px 0 0' }}>
          {noCrop
            ? (isDe ? 'Das ganze Foto wird übernommen — nichts wird weggeschnitten.' : 'The whole photo is used — nothing is cropped off.')
            : allowAspect
              ? (isDe ? 'Ziehe das Bild in Position, z.B. um oben oder unten etwas wegzuschneiden.' : 'Drag the image into place, e.g. to crop off the top or bottom.')
              : (isDe ? 'Ziehe das Bild mit der Maus in Position — die Vorschau zeigt genau das Ergebnis.' : 'Drag the image into place — the preview shows exactly what you get.')}
        </p>
        {/* v31.9.7: Ohne Zuschnitt gibt es keinen Zoom — ein Regler, der nur
            wieder etwas abschneiden kann, widerspricht der Kachel darüber. */}
        {!noCrop && slider('Zoom', `${zoom.toFixed(1).replace('.', isDe ? ',' : '.')}×`,
          <input type="range" min={1} max={4} step={0.01} value={zoom} onChange={e => setZoom(parseFloat(e.target.value))} />)}
      </div>

      {/* v27.5: Seitenverhältnis-Wahl (nur Kopfbild-Modus) ODER Form-Wahl.
          v30.98: „Empfohlen"-Badge und Tipp entfallen (Nutzer-Ansage 07.09.2026:
          „nimm das Empfohlen hier raus") — Kreis bleibt die Vorauswahl, ohne
          Wertung. `recommendCircle` bleibt als Prop und steuert nur noch die
          Erklärzeile der Kreis-Kachel (Event-Karte vs. allgemeiner Kreis).
          v31.2: Der Kreis-Rand steht direkt unter der Form-Wahl, weil er nur
          zum Kreis gehört — vorher stand der Zoom dazwischen. */}
      <div className="dex-ui-section" style={{ margin: 0 }}>
        <div className="dex-ui-section-title">{allowAspect ? (isDe ? 'Seitenverhältnis' : 'Aspect ratio') : (isDe ? 'Zuschnitt' : 'Crop')}</div>
        {allowAspect ? (
          <div className="dex-ui-tabs" role="group" aria-label={isDe ? 'Seitenverhältnis' : 'Aspect ratio'}>
            {ASPECT_PRESETS.map(p => (
              <button key={p.a} type="button" className={cx('dex-ui-tab', Math.abs(aspect - p.a) < 0.001 && 'is-active')} onClick={() => setAspect(p.a)}>
                {isDe ? p.de : p.en}
              </button>
            ))}
          </div>
        ) : (
          // v31.9.7: Drei Kacheln untereinander statt nebeneinander — bei 560 px
          // Dialogbreite bliebe je Kachel keine Zeile für die Erklärung übrig.
          <div style={{ display: 'grid', gap: 10 }}>{shapeChoice('none')}{shapeChoice('circle')}{shapeChoice('rect')}</div>
        )}
        {isCircle && slider(
          isDe ? 'Weißer Rand um den Kreis' : 'White margin around the circle', `${Math.round(padding * 100)} %`,
          <input type="range" min={0} max={0.35} step={0.01} value={padding} onChange={e => setPadding(parseFloat(e.target.value))} />,
          isDe ? '0 % = der Kreis füllt den Rahmen; mehr = kleinerer Kreis mit weißem Rand außen.' : '0 % = the circle fills the frame; more = smaller circle with a white margin outside.')}
      </div>

      {/* v31.2: Der Hinweis auf das Original in Mail und Termin stand vorher im
          Einleitungssatz — als eigener Kasten liest man ihn erst, wenn man ihn
          braucht, und die Einleitung bleibt eine Zeile. */}
      {/* v31.9.7: Nur noch bei einem echten Zuschnitt. Wer „Nicht beschneiden"
          gewählt hat, bekäme sonst den Hinweis, dass anderswo das Original
          benutzt wird — das ist dann überall dasselbe Bild. */}
      {!allowAspect && !noCrop && (
        <div className="dex-ui-callout dex-ui-callout--neutral">
          <span className="dex-ui-callout-icon"><Info size={16} /></span>
          <span>
            {isDe
              ? 'E-Mails und der Outlook-Termin nutzen automatisch das unbeschnittene Originalfoto — dort ist der Kopf rechteckig.'
              : 'Emails and the Outlook invite automatically use the uncropped original photo — their header is rectangular.'}
          </span>
        </div>
      )}

      {error && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>{error}</span>
        </div>
      )}

      {/* v23.25: Zusatzblock (Darstellung pro Ansicht), abgesetzt mit Trennlinie. */}
      {children && (
        <div style={{ paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
          {children}
        </div>
      )}
    </Modal>
  );
}
