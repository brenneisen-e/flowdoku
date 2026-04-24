/**
 * AppPreview — zeigt im Handbuch einen kompakten Vorschau-Button, der per
 * Klick ein Modal mit der echten App-Komponente öffnet. Die Komponente läuft
 * dabei in einem isolierten Context-Stack mit Demo-Daten (siehe
 * PreviewProviders), und Klicks werden durch ein pointer-events:none-Wrap
 * neutralisiert — der User sieht genau das echte Layout der App, kann es
 * aber nicht bedienen (reine Read-only-Vorschau).
 */
import * as React from 'react';
import { PreviewContextStack, PreviewRole } from './PreviewProviders';
import { DeloitteEvent } from '../../../types';

interface AppPreviewProps {
  label: string;
  /** Welche Rolle der Demo-User in diesem Preview hat. Steuert canCheckIn,
   *  isAdmin-Flags etc., damit z.B. die Check-In-Bubble im Header auftaucht. */
  role?: PreviewRole;
  /** Viewport-Breite, mit der die Komponente simuliert wird. Default 430 =
   *  iPhone 15 Pro Max (breit genug, damit die Landing-Card nicht rechts
   *  abgeschnitten wird). Für knappere Previews 412 (Galaxy S23) oder 390
   *  (iPhone Mini). 1024+ rendert das Desktop-Layout (ohne Phone-Frame). */
  width?: number;
  /** Navigation: aktuelle Seite (default 'landing'). */
  page?: string;
  /** Navigation: vorausgewähltes Event (wichtig für Detail-Seiten). */
  selectedEventId?: string;
  /** Zusätzliche Demo-Events über das Default-Office-Event hinaus. */
  extraEvents?: DeloitteEvent[];
  /** Device-Hülle: phone (schwarzer Bezel mit Notch) vs. laptop (Monitor +
   *  schmaler Stand/Tastatur-Base). Default: phone bei width <= 768, sonst laptop. */
  device?: 'phone' | 'laptop' | 'plain';
  /** Nur für EventCreationPage-Previews: initialen Wizard-Schritt (0..6)
   *  gezielt setzen, damit das Handbuch pro Step den passenden Inhalt zeigt. */
  initialStep?: number;
  children: React.ReactNode;
}

export function AppPreview(props: AppPreviewProps): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  const width = props.width || 430;
  // v6.30: Mobile-Flag SYNCHRON setzen (nicht erst im useEffect), damit der
  // Header im Preview schon beim ersten Render die Mobile-Variante rendert.
  // useState-Init in <Header> liest das Flag sonst BEVOR unser useEffect
  // läuft, und die Check-In-Bubble fehlt.
  if (open && width <= 768 && typeof window !== 'undefined') {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__dexForceMobile = true;
  }
  // initialStep für EventCreationPage — genauso synchron wie das Mobile-Flag,
  // damit useState(() => ...) den Wert beim ersten Render lesen kann.
  if (open && typeof props.initialStep === 'number' && typeof window !== 'undefined') {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__dexPreviewInitialStep = props.initialStep;
  }
  // v6.28: Wenn das Modal auf < 768px "Mobile" getrimmt ist, setzen wir
  // das window-Flag `__dexForceMobile`, damit Komponenten wie <Header>
  // im Preview-Content tatsächlich ihre Mobile-Variante rendern — das echte
  // window.matchMedia liest sonst die Desktop-Breite des Handbuchs.
  React.useEffect(() => {
    if (!open) return;
    if (width <= 768) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (window as any).__dexForceMobile = true;
    }
    // v6.29: Globaler Style-Override, solange das Preview-Modal offen ist.
    // Grund: die echte App nutzt clamp(..vw, ..vh..) an vielen Stellen, das
    // rechnet gegen den Browser-Viewport (Desktop = 1920+) und nicht gegen
    // den Handy-Frame-Container im Modal. Deshalb werden Fonts/Orb-Größe/
    // Paddings riesig. Wir zwingen hier im Scope '.dex-preview-scope' die
    // mobile-freundlichen Werte.
    const styleEl = document.createElement('style');
    styleEl.id = 'dex-preview-style-override';
    styleEl.textContent = `
      .dex-preview-scope .landing { overflow: auto !important; height: auto !important; }
      .dex-preview-scope .landing__hero { padding: 16px 12px !important; gap: 16px !important; }
      .dex-preview-scope .landing__card { padding: 20px 16px !important; gap: 16px !important; border-radius: 16px !important; max-width: 100% !important; }
      .dex-preview-scope .landing__text h1 { font-size: 1.35rem !important; line-height: 1.25 !important; }
      .dex-preview-scope .landing__text p { font-size: 0.88rem !important; }
      .dex-preview-scope .landing__orb { width: 180px !important; height: 180px !important; }
    `;
    document.head.appendChild(styleEl);
    return () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      try { delete (window as any).__dexForceMobile; } catch { /* */ }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      try { delete (window as any).__dexPreviewInitialStep; } catch { /* */ }
      const el = document.getElementById('dex-preview-style-override');
      if (el) el.remove();
    };
  }, [open, width]);
  return (
    <>
      <button
        type="button"
        onClick={() => setOpen(true)}
        style={{
          display: 'flex', alignItems: 'center', gap: 10,
          padding: '10px 14px', width: '100%', textAlign: 'left',
          background: 'var(--dex-gray-50)', border: '1px solid var(--dex-gray-200)',
          borderRadius: 10, cursor: 'pointer', fontSize: '0.85rem',
          fontFamily: 'inherit', color: 'var(--dex-gray-700)',
        }}
      >
        <span style={{
          display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
          width: 28, height: 28, borderRadius: 6,
          background: 'var(--dex-green)', color: '#fff', fontSize: '0.75rem', fontWeight: 700,
        }}>👁</span>
        <span style={{ flex: 1 }}>
          <span style={{ fontWeight: 600, display: 'block' }}>{props.label}</span>
          <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
            Vorschau der echten App — klicken zum Öffnen
          </span>
        </span>
      </button>
      {open && (
        <div
          role="dialog" aria-modal="true"
          onClick={() => setOpen(false)}
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            zIndex: 10000, padding: 24,
          }}
        >
          <div
            onClick={e => e.stopPropagation()}
            style={{
              background: 'var(--dex-gray-50)', borderRadius: 16,
              // Bei Mobile-Previews muss Platz für den Phone-Frame-Bezel (24px + padding 24px)
              // sein, sonst wird der Frame rechts abgeschnitten. Desktop-Previews brauchen
              // nur 32px Zusatz-Padding.
              maxWidth: width <= 768 ? Math.min(width + 100, 900) : Math.min(width + 32, 1100),
              width: '100%',
              maxHeight: '90vh', overflow: 'auto',
              boxShadow: '0 16px 48px rgba(0,0,0,0.25)',
              display: 'flex', flexDirection: 'column',
            }}
          >
            <div style={{
              display: 'flex', alignItems: 'center', justifyContent: 'space-between',
              padding: '12px 16px', borderBottom: '1px solid var(--dex-gray-200)',
              background: '#fff',
            }}>
              <div style={{ fontWeight: 600, fontSize: '0.9rem' }}>{props.label}</div>
              <button
                type="button" onClick={() => setOpen(false)}
                style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.25rem', color: 'var(--dex-gray-500)' }}
                aria-label="Schließen"
              >×</button>
            </div>
            {/* Scrollable Inner-Frame simuliert Mobile-/Desktop-Viewport.
                Unter 768px zeichnen wir einen stilisierten Handy-Rahmen mit
                Notch + schwarzem Bezel, damit der Leser sofort erkennt,
                dass es sich um die Mobile-Ansicht handelt. */}
            <div style={{
              padding: '24px 16px', display: 'flex', justifyContent: 'center',
              overflowX: 'auto', background: 'var(--dex-gray-100, #f5f5f5)',
            }}>
              {(() => {
                const device: 'phone' | 'laptop' | 'plain' = props.device
                  || (width <= 768 ? 'phone' : 'laptop');
                const innerFrame = (
                  <div className="dex-preview-scope" style={{ pointerEvents: 'none', userSelect: 'text', width: '100%' }}>
                    <PreviewContextStack role={props.role} page={props.page} selectedEventId={props.selectedEventId} extraEvents={props.extraEvents}>
                      {props.children}
                    </PreviewContextStack>
                  </div>
                );
                if (device === 'phone') {
                  return (
                    <div style={{
                      // Phone-Frame: schwarzer Bezel mit Notch
                      width: width + 24, padding: '38px 12px 18px', borderRadius: 44,
                      background: '#1a1a1a',
                      boxShadow: '0 8px 32px rgba(0,0,0,0.28), inset 0 0 0 2px #333',
                      position: 'relative',
                    }}>
                      <div style={{
                        position: 'absolute', top: 12, left: '50%', transform: 'translateX(-50%)',
                        width: 100, height: 18, background: '#000', borderRadius: 10,
                        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                      }}>
                        <span style={{ width: 6, height: 6, borderRadius: '50%', background: '#222', border: '1px solid #333' }} />
                        <span style={{ width: 38, height: 4, borderRadius: 2, background: '#222' }} />
                      </div>
                      <div style={{
                        width, background: '#fff', borderRadius: 28,
                        minHeight: 620, maxHeight: '72vh',
                        overflowY: 'auto', overflowX: 'hidden',
                      }}>
                        {innerFrame}
                      </div>
                    </div>
                  );
                }
                if (device === 'laptop') {
                  // Laptop-Frame: schmaler Monitor-Bezel oben + seitlich, dicker unten
                  // mit Kamera-Dot am oberen Rand; Stand/Tastatur-Base als breites
                  // abgeflachtes Trapez darunter.
                  return (
                    <div style={{ position: 'relative', paddingBottom: 18 }}>
                      <div style={{
                        width: width + 24, padding: '20px 12px 12px',
                        borderRadius: '14px 14px 6px 6px',
                        background: '#1a1a1a',
                        boxShadow: '0 8px 28px rgba(0,0,0,0.22)',
                        position: 'relative',
                      }}>
                        <div style={{
                          position: 'absolute', top: 7, left: '50%', transform: 'translateX(-50%)',
                          width: 6, height: 6, borderRadius: '50%', background: '#333',
                        }} />
                        <div style={{
                          width, background: '#fff', borderRadius: 4,
                          minHeight: 560, maxHeight: '72vh',
                          overflowY: 'auto', overflowX: 'hidden',
                        }}>
                          {innerFrame}
                        </div>
                      </div>
                      <div style={{
                        position: 'absolute', left: -24, right: -24, bottom: 0, height: 14,
                        background: 'linear-gradient(180deg, #2a2a2a, #111)',
                        borderRadius: '0 0 14px 14px',
                      }} />
                      <div style={{
                        position: 'absolute', left: '40%', right: '40%', bottom: 10, height: 4,
                        background: '#444', borderRadius: '0 0 6px 6px',
                      }} />
                    </div>
                  );
                }
                return (
                  <div style={{
                    width, maxWidth: '100%',
                    background: '#fff', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 12, overflow: 'hidden',
                    boxShadow: '0 2px 12px rgba(0,0,0,0.06)',
                  }}>
                    <div style={{ pointerEvents: 'none', userSelect: 'text' }}>
                      <PreviewContextStack role={props.role} page={props.page} selectedEventId={props.selectedEventId} extraEvents={props.extraEvents}>
                        {props.children}
                      </PreviewContextStack>
                    </div>
                  </div>
                );
              })()}
            </div>
            <div style={{
              padding: '8px 16px', borderTop: '1px solid var(--dex-gray-200)',
              background: '#fff', fontSize: '0.72rem', color: 'var(--dex-gray-500)',
              textAlign: 'center',
            }}>
              Read-only-Vorschau · Demo-Daten · Klicks sind deaktiviert
            </div>
          </div>
        </div>
      )}
    </>
  );
}
