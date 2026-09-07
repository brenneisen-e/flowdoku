/**
 * ImageAnnotateModal (v26.10.0)
 *
 * Zeigt ein Bild (z.B. den angehängten Screenshot) GROSS und lässt den Nutzer
 * mit der Maus transparente Markierungs-Rechtecke darüberziehen — damit z.B.
 * der/die Antwortende sofort sieht, welche Stelle gemeint ist. Beim Übernehmen
 * werden die Markierungen fest in ein neues PNG gerendert (Canvas), das das
 * Original-Bild ersetzt — kein zusätzliches Datenmodell nötig, die Markierung
 * ist Teil des Bildes.
 *
 * Die Screenshots stammen aus dem eigenen Tab/Blob (html2canvas bzw. lokale
 * Datei) → same-origin, das Canvas ist nicht „tainted", der Export funktioniert.
 */
import * as React from 'react';
import Modal from './Modal';
import { cx } from './dexUi';
import { Icon } from '@fluentui/react/lib/Icon';

interface Box { x: number; y: number; w: number; h: number; } // Anteile 0..1

export default function ImageAnnotateModal(props: {
  open: boolean;
  src: string;
  fileName?: string;
  isDe?: boolean;
  onClose: () => void;
  onSave: (file: File) => void;
}): React.ReactElement | null {
  const { open, src, fileName = 'screenshot.png', isDe = true } = props;
  const [boxes, setBoxes] = React.useState<Box[]>([]);
  const [draft, setDraft] = React.useState<Box | null>(null);
  const [busy, setBusy] = React.useState(false);
  const imgRef = React.useRef<HTMLImageElement>(null);
  const dragStart = React.useRef<{ x: number; y: number } | null>(null);

  // Bei jedem Öffnen / Bildwechsel zurücksetzen.
  React.useEffect(() => { if (open) { setBoxes([]); setDraft(null); dragStart.current = null; } }, [open, src]);

  if (!open) return null;

  const relPoint = (e: React.MouseEvent): { x: number; y: number } | null => {
    const r = imgRef.current?.getBoundingClientRect();
    if (!r || r.width === 0 || r.height === 0) return null;
    return {
      x: Math.min(1, Math.max(0, (e.clientX - r.left) / r.width)),
      y: Math.min(1, Math.max(0, (e.clientY - r.top) / r.height)),
    };
  };

  const onDown = (e: React.MouseEvent): void => {
    const p = relPoint(e); if (!p) return;
    dragStart.current = p; setDraft({ x: p.x, y: p.y, w: 0, h: 0 });
  };
  const onMove = (e: React.MouseEvent): void => {
    if (!dragStart.current) return;
    const p = relPoint(e); if (!p) return;
    const s = dragStart.current;
    setDraft({ x: Math.min(s.x, p.x), y: Math.min(s.y, p.y), w: Math.abs(p.x - s.x), h: Math.abs(p.y - s.y) });
  };
  const onUp = (): void => {
    if (draft && draft.w > 0.01 && draft.h > 0.01) setBoxes((b) => [...b, draft]);
    setDraft(null); dragStart.current = null;
  };

  const save = async (): Promise<void> => {
    setBusy(true);
    try {
      const img = imgRef.current;
      const natW = img?.naturalWidth || 0;
      const natH = img?.naturalHeight || 0;
      if (!img || !natW || !natH) { setBusy(false); props.onClose(); return; }
      const canvas = document.createElement('canvas');
      canvas.width = natW; canvas.height = natH;
      const ctx = canvas.getContext('2d');
      if (!ctx) { setBusy(false); return; }
      ctx.drawImage(img as CanvasImageSource, 0, 0, natW, natH);
      ctx.lineWidth = Math.max(2, Math.round(natW / 320));
      ctx.strokeStyle = '#ed8b00';
      ctx.fillStyle = 'rgba(237,139,0,0.22)';
      boxes.forEach((b) => {
        const x = b.x * natW, y = b.y * natH, w = b.w * natW, h = b.h * natH;
        ctx.fillRect(x, y, w, h);
        ctx.strokeRect(x, y, w, h);
      });
      const blob: Blob | null = await new Promise((res) => canvas.toBlob((bl) => res(bl), 'image/png'));
      if (blob) {
        const base = fileName.replace(/\.[^.]+$/, '') || 'screenshot';
        props.onSave(new File([blob], `${base}.png`, { type: 'image/png' }));
      }
    } catch { /* best-effort */ }
    setBusy(false);
  };

  const boxStyle = (b: Box): React.CSSProperties => ({
    position: 'absolute', left: `${b.x * 100}%`, top: `${b.y * 100}%`, width: `${b.w * 100}%`, height: `${b.h * 100}%`,
    background: 'rgba(237,139,0,0.22)', border: '2px solid #ed8b00', pointerEvents: 'none', boxSizing: 'border-box',
  });

  // v31.2: Der Status-Text sagt, was der Organizer bisher getan hat und was
  // als Nächstes geht — vorher gab es nur zwei ausgegraute Knöpfe, aus denen
  // man den Zustand erraten musste.
  const n = boxes.length;
  const statusText = n === 0
    ? (isDe ? 'Noch keine Markierung' : 'No marks yet')
    : (isDe ? `${n} Markierung${n === 1 ? '' : 'en'}` : `${n} mark${n === 1 ? '' : 's'}`);

  return (
    <Modal open={open} onClose={props.onClose} maxWidth={920} dismissable={!busy} ariaLabel={isDe ? 'Screenshot markieren' : 'Annotate screenshot'}
      // v31.2: Kopf und Fuß über die Modal-Props — ein Look für alle Dialoge
      // statt eines eigenen <h2> mit eigenen Abständen.
      title={isDe ? 'Stelle im Screenshot markieren' : 'Mark the spot in the screenshot'}
      subtitle={isDe
        ? 'Ziehe mit gedrückter Maustaste ein Rechteck über die Stelle, die du meinst. Mehrere Markierungen sind möglich — beim Übernehmen werden sie fest ins Bild gebacken.'
        : 'Drag a rectangle over the area you mean. You can add several marks — when you apply them, they become part of the image.'}
      icon={<Icon iconName="InsertTextBox" style={{ fontSize: 18 }} />}
      footer={<>
        <button type="button" className="btn btn-secondary" disabled={busy} onClick={props.onClose}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
        <button type="button" className="btn btn-primary" disabled={busy} onClick={save}>
          {busy ? (isDe ? 'Speichert …' : 'Saving …') : (isDe ? 'Ins Bild übernehmen' : 'Apply to image')}
        </button>
      </>}>
      {/* v31.2: Werkzeugleiste direkt über dem Bild — Symbol-Knöpfe mit Hover
          und Tooltip statt zwei Textknöpfen im Fuß, damit der Fuß nur noch
          die Entscheidung (Abbrechen / Übernehmen) trägt. */}
      <div className="dex-ui-inline" style={{ justifyContent: 'space-between' }}>
        <span className={cx('dex-ui-pill', n > 0 ? 'dex-ui-pill--orange' : 'dex-ui-pill--gray')} aria-live="polite">{statusText}</span>
        <span className="dex-ui-inline" style={{ gap: 2 }} role="toolbar" aria-label={isDe ? 'Markierungen bearbeiten' : 'Edit marks'}>
          <button type="button" className="dex-ui-iconbtn" disabled={busy || n === 0} onClick={() => setBoxes((b) => b.slice(0, -1))}
            title={isDe ? 'Letzte Markierung zurücknehmen' : 'Undo last mark'} aria-label={isDe ? 'Letzte Markierung zurücknehmen' : 'Undo last mark'}>
            <Icon iconName="Undo" style={{ fontSize: 16 }} />
          </button>
          <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" disabled={busy || n === 0} onClick={() => setBoxes([])}
            title={isDe ? 'Alle Markierungen entfernen' : 'Remove all marks'} aria-label={isDe ? 'Alle Markierungen entfernen' : 'Remove all marks'}>
            <Icon iconName="Delete" style={{ fontSize: 16 }} />
          </button>
        </span>
      </div>
      <div style={{ maxHeight: '62vh', overflow: 'auto', border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 12, background: 'var(--dex-gray-50,#fafafa)', textAlign: 'center' }}>
        <div style={{ position: 'relative', display: 'inline-block', userSelect: 'none' }}
          onMouseDown={onDown} onMouseMove={onMove} onMouseUp={onUp} onMouseLeave={onUp}>
          <img ref={imgRef} src={src} alt="Screenshot" draggable={false}
            style={{ display: 'block', maxWidth: '100%', cursor: 'crosshair' }} />
          {boxes.map((b, i) => <div key={i} style={boxStyle(b)} />)}
          {/* v31.2: Der noch gezogene Rahmen ist gestrichelt — so sieht man, was
              schon steht und was gerade erst entsteht. */}
          {draft && <div style={{ ...boxStyle(draft), borderStyle: 'dashed' }} />}
        </div>
      </div>
    </Modal>
  );
}
