/**
 * WizardStepPreviewModal (v26.52)
 *
 * Zeigt den ECHTEN Event-Wizard (EventCreationPage mit Demo-Daten, wie die
 * Handbuch-Previews) auf dem gewählten Schritt in einem Modal — statt nur
 * eines Text-Verweises „Schritt 5". Im Editier-Modus (Antwort-Composer der
 * Power-User) kann per Maus-Drag EINE orangene Markierungsbox über die
 * Vorschau gezogen werden („hier klicken"), die als Prozent-Koordinaten
 * relativ zum Vorschau-Inhalt gespeichert wird. Im Viewer-Modus (Fragesteller
 * bzw. beantwortetes Ticket) wird dieselbe Vorschau read-only mit der
 * gespeicherten Markierung angezeigt.
 *
 * WICHTIG: Diese Datei importiert EventCreationPage statisch — Konsumenten
 * MÜSSEN sie per React.lazy laden, sonst wandert der komplette Wizard in den
 * Main-Bundle (EventCreationPage ist in DexEventPlatform.tsx lazy).
 *
 * v31.3: Rahmen, Kopf und Fuß kommen jetzt vom gemeinsamen `Modal`
 * (title/subtitle/icon/footer, siehe docs/ui-leitfaden.md) statt aus einem
 * eigenen Overlay — damit sieht der Dialog aus wie jeder andere, schließt mit
 * Escape und verliert beim Ziehen der Markierung nicht mehr den Fokus (der
 * Backdrop von `Modal` schließt nur, wenn Maus-Runter UND Maus-Hoch auf ihm
 * landen). Die Drag-Mathematik und die Prozent-Koordinaten sind unverändert.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { PreviewContextStack } from '../manual/previews/PreviewProviders';
import Header from '../Header';
import EventCreationPage from '../EventCreationPage';
import Modal from '../Modal';

export interface WizardMarker { x: number; y: number; w: number; h: number; }

interface WizardStepPreviewModalProps {
  /** 1-basierter Wizard-Schritt (identisch zur DexTicket-Konvention). */
  step: number;
  stepLabel?: string;
  isDe: boolean;
  /** true = Composer (Markierung zeichnen), false = reine Ansicht. */
  editable: boolean;
  initialMarker?: WizardMarker | null;
  onClose: () => void;
  /** Nur im Editier-Modus: „Übernehmen" liefert die (ggf. entfernte) Markierung. */
  onSave?: (marker: WizardMarker | null) => void;
}

const PREVIEW_WIDTH = 1024;
const MARKER_COLOR = '#ed8b00';

const clampPct = (v: number): number => Math.max(0, Math.min(100, v));

export default function WizardStepPreviewModal(props: WizardStepPreviewModalProps): React.ReactElement {
  const { step, isDe, editable } = props;
  const [marker, setMarker] = React.useState<WizardMarker | null>(props.initialMarker || null);
  // Aktive Drag-Geste: Startpunkt in Prozent; null = kein Drag.
  const dragStartRef = React.useRef<{ x: number; y: number } | null>(null);
  const [draft, setDraft] = React.useState<WizardMarker | null>(null);
  const contentRef = React.useRef<HTMLDivElement | null>(null);
  const scrollRef = React.useRef<HTMLDivElement | null>(null);

  // Wizard-Schritt SYNCHRON vor dem ersten Render von EventCreationPage
  // setzen (gleicher Mechanismus wie AppPreview im Handbuch): der useState-
  // Initializer in EventCreationPage liest das Flag beim Mount.
  if (typeof window !== 'undefined') {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__dexPreviewInitialStep = Math.max(0, step - 1);
    // Desktop-Vorschau — ein evtl. hängengebliebenes Mobile-Flag stören lassen.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    try { delete (window as any).__dexForceMobile; } catch { /* */ }
  }
  React.useEffect(() => {
    return () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      try { delete (window as any).__dexPreviewInitialStep; } catch { /* */ }
    };
  }, []);

  // Viewer-Modus: nach dem Mount zur Markierung scrollen, damit der
  // Fragesteller die Box ohne Suchen sieht (Timeout: Wizard erst rendern lassen).
  React.useEffect(() => {
    if (editable || !props.initialMarker) return;
    const t = window.setTimeout(() => {
      const sc = scrollRef.current; const ct = contentRef.current;
      if (!sc || !ct || !props.initialMarker) return;
      const targetPx = (props.initialMarker.y / 100) * ct.offsetHeight;
      sc.scrollTop = Math.max(0, targetPx - sc.clientHeight * 0.3);
    }, 350);
    return () => window.clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const pctFromPointer = (e: React.PointerEvent): { x: number; y: number } | null => {
    const ct = contentRef.current;
    if (!ct) return null;
    const r = ct.getBoundingClientRect();
    if (r.width <= 0 || r.height <= 0) return null;
    return { x: clampPct(((e.clientX - r.left) / r.width) * 100), y: clampPct(((e.clientY - r.top) / r.height) * 100) };
  };

  const onPointerDown = (e: React.PointerEvent): void => {
    if (!editable) return;
    const p = pctFromPointer(e);
    if (!p) return;
    dragStartRef.current = p;
    setDraft({ x: p.x, y: p.y, w: 0, h: 0 });
    try { (e.target as Element).setPointerCapture(e.pointerId); } catch { /* */ }
    e.preventDefault();
  };
  const onPointerMove = (e: React.PointerEvent): void => {
    if (!editable || !dragStartRef.current) return;
    const p = pctFromPointer(e);
    if (!p) return;
    const s = dragStartRef.current;
    setDraft({ x: Math.min(s.x, p.x), y: Math.min(s.y, p.y), w: Math.abs(p.x - s.x), h: Math.abs(p.y - s.y) });
  };
  const onPointerUp = (): void => {
    if (!editable || !dragStartRef.current) return;
    dragStartRef.current = null;
    setDraft((d) => {
      // Mini-Drags (< ~0.8% in beiden Richtungen) = versehentlicher Klick →
      // bestehende Markierung NICHT überschreiben.
      if (d && (d.w >= 0.8 || d.h >= 0.8)) {
        setMarker({ x: Math.round(d.x * 100) / 100, y: Math.round(d.y * 100) / 100, w: Math.round(d.w * 100) / 100, h: Math.round(d.h * 100) / 100 });
      }
      return null;
    });
  };

  const shown = draft || marker;
  const title = `${isDe ? 'Event-Wizard · Schritt' : 'Event wizard · step'} ${step}${props.stepLabel ? `: ${props.stepLabel}` : ''}`;

  const renderMarker = (m: WizardMarker): React.ReactElement => (
    <div style={{
      position: 'absolute', left: `${m.x}%`, top: `${m.y}%`, width: `${m.w}%`, height: `${m.h}%`,
      border: `3px solid ${MARKER_COLOR}`, borderRadius: 6, background: 'rgba(237,139,0,0.10)',
      boxShadow: '0 0 0 3px rgba(237,139,0,0.25), 0 2px 10px rgba(0,0,0,0.18)',
      pointerEvents: 'none', boxSizing: 'border-box',
    }}>
      <span style={{
        position: 'absolute', left: -3,
        // Label über der Box; ganz oben am Rand stattdessen darunter, damit
        // es nicht aus dem Vorschau-Container rausläuft.
        ...(m.y < 4 ? { top: '100%', marginTop: 4 } : { bottom: '100%', marginBottom: 4 }),
        background: MARKER_COLOR, color: '#fff', fontSize: 11, fontWeight: 700,
        padding: '2px 8px', borderRadius: 4, whiteSpace: 'nowrap', lineHeight: 1.5,
      }}>{isDe ? 'Hier klicken' : 'Click here'}</span>
    </div>
  );

  return (
    // v31.3: Der Dialog wird AUS einem offenen Ticket-Dialog heraus geöffnet
    // (QuestionButton) und muss darüber liegen. Bei gleicher Ebene entschiede
    // die Einhäng-Reihenfolge der Portale — das ginge heute gut, wäre aber
    // Zufall; die 10500 stand vor v31.3 fest im eigenen Overlay.
    <Modal
      open
      zIndex={10500}
      onClose={props.onClose}
      maxWidth={PREVIEW_WIDTH + 96}
      padding="18px 20px"
      ariaLabel={title}
      title={title}
      icon={<Icon iconName="DocumentManagement" style={{ fontSize: 18 }} />}
      subtitle={
        <span className="dex-ui-inline">
          <Icon iconName={editable ? 'SingleColumnEdit' : 'View'} style={{ fontSize: 13 }} />
          {editable
            ? (isDe ? 'Ziehe mit der Maus einen Rahmen um die Stelle, auf die der Fragesteller klicken soll.' : 'Drag a box around the spot the asker should click.')
            : (isDe ? 'Read-only-Vorschau des echten Wizards mit Demo-Daten — die orange Box zeigt, wo du klicken musst.' : 'Read-only preview of the real wizard with demo data — the orange box shows where to click.')}
        </span>
      }
      footer={editable ? (
        <>
          {/* Links, weil „Markierung entfernen" zum Inhalt gehört; Abbrechen und
              Übernehmen bleiben rechts (Leitfaden: Modal-Fuß). */}
          {marker && (
            <button type="button" className="btn btn-secondary dex-ui-btn-sm dex-ui-modal-foot-left" onClick={() => setMarker(null)}>
              <Icon iconName="EraseTool" style={{ fontSize: 12 }} />{isDe ? 'Markierung entfernen' : 'Remove marker'}
            </button>
          )}
          <button type="button" className="btn btn-secondary" onClick={props.onClose}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
          <button type="button" className="btn btn-primary"
            onClick={() => { if (props.onSave) props.onSave(marker); props.onClose(); }}>
            {isDe ? 'Übernehmen' : 'Apply'}
          </button>
        </>
      ) : (
        <button type="button" className="btn btn-primary" onClick={props.onClose}>{isDe ? 'Schließen' : 'Close'}</button>
      )}
    >
      {/* Vorschau (scrollbar) + Markierungs-Overlay */}
      <div ref={scrollRef}
        style={{ overflow: 'auto', background: 'var(--dex-gray-100,#f5f5f5)', padding: 16, borderRadius: 12, border: '1px solid var(--dex-gray-200,#e8e8e8)', flex: '1 1 auto', minHeight: 0, maxHeight: '70vh' }}>
        <div style={{ width: PREVIEW_WIDTH, margin: '0 auto', background: '#fff', borderRadius: 8, boxShadow: '0 2px 12px rgba(0,0,0,0.10)' }}>
          {/* contentRef = Bezugsrahmen der Prozent-Koordinaten: der komplette
              gerenderte Wizard-Inhalt (scrollt mit, Marker bleibt „kleben"). */}
          <div ref={contentRef} style={{ position: 'relative' }}>
            <div className="dex-preview-scope" style={{ pointerEvents: 'none', userSelect: 'none' }}>
              <PreviewContextStack role="Organizer" page="create-event">
                <Header />
                <EventCreationPage />
              </PreviewContextStack>
            </div>
            <div
              onPointerDown={onPointerDown}
              onPointerMove={onPointerMove}
              onPointerUp={onPointerUp}
              onPointerCancel={onPointerUp}
              style={{ position: 'absolute', inset: 0, cursor: editable ? 'crosshair' : 'default', touchAction: editable ? 'none' : undefined }}
            >
              {shown && (shown.w >= 0.8 || shown.h >= 0.8) && renderMarker(shown)}
            </div>
          </div>
        </div>
      </div>
    </Modal>
  );
}
