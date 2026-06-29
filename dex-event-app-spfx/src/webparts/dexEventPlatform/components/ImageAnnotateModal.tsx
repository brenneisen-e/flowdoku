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

  return (
    <Modal open={open} onClose={props.onClose} maxWidth={920} dismissable={!busy} ariaLabel={isDe ? 'Screenshot markieren' : 'Annotate screenshot'}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
        <Icon iconName="InsertTextBox" style={{ fontSize: 18, color: 'var(--dex-green,#86bc25)' }} />
        <h2 style={{ margin: 0, fontSize: '1.1rem' }}>{isDe ? 'Screenshot vergrößern & markieren' : 'Enlarge & mark up screenshot'}</h2>
      </div>
      <p style={{ margin: '0 0 10px', fontSize: '0.84rem', color: 'var(--dex-gray-600,#666)' }}>
        {isDe
          ? 'Ziehe mit gedrückter Maustaste ein Rechteck über die Stelle, die du meinst. Du kannst mehrere Markierungen setzen — sie werden fest ins Bild übernommen.'
          : 'Drag a rectangle over the area you mean. You can add several marks — they are baked into the image.'}
      </p>
      <div style={{ maxHeight: '62vh', overflow: 'auto', border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 8, background: 'var(--dex-gray-50,#fafafa)', textAlign: 'center' }}>
        <div style={{ position: 'relative', display: 'inline-block', userSelect: 'none' }}
          onMouseDown={onDown} onMouseMove={onMove} onMouseUp={onUp} onMouseLeave={onUp}>
          <img ref={imgRef} src={src} alt="Screenshot" draggable={false}
            style={{ display: 'block', maxWidth: '100%', cursor: 'crosshair' }} />
          {boxes.map((b, i) => <div key={i} style={boxStyle(b)} />)}
          {draft && <div style={boxStyle(draft)} />}
        </div>
      </div>
      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8, marginTop: 12, flexWrap: 'wrap' }}>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button className="btn btn-secondary" style={{ padding: '7px 12px' }} disabled={busy || boxes.length === 0} onClick={() => setBoxes((b) => b.slice(0, -1))}>
            {isDe ? 'Letzte Markierung entfernen' : 'Undo last'}
          </button>
          <button className="btn btn-secondary" style={{ padding: '7px 12px' }} disabled={busy || boxes.length === 0} onClick={() => setBoxes([])}>
            {isDe ? 'Alle entfernen' : 'Clear all'}
          </button>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button className="btn btn-secondary" style={{ padding: '7px 16px' }} disabled={busy} onClick={props.onClose}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
          <button className="btn btn-primary" style={{ padding: '7px 16px' }} disabled={busy} onClick={save}>
            {busy ? (isDe ? 'Speichert …' : 'Saving …') : (isDe ? 'Übernehmen' : 'Apply')}
          </button>
        </div>
      </div>
    </Modal>
  );
}
