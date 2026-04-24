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
  /** Viewport-Breite, mit der die Komponente simuliert wird. 390 = iPhone-14,
   *  1024 = iPad/Desktop-nah. Erlaubt pro Preview gezielt das Mobile- oder
   *  Desktop-Layout zu erzwingen. */
  width?: number;
  /** Navigation: aktuelle Seite (default 'landing'). */
  page?: string;
  /** Navigation: vorausgewähltes Event (wichtig für Detail-Seiten). */
  selectedEventId?: string;
  /** Zusätzliche Demo-Events über das Default-Office-Event hinaus. */
  extraEvents?: DeloitteEvent[];
  children: React.ReactNode;
}

export function AppPreview(props: AppPreviewProps): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  const width = props.width || 390;
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
    return () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      try { delete (window as any).__dexForceMobile; } catch { /* */ }
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
              maxWidth: Math.min(width + 32, 900), width: '100%',
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
              {width <= 768 ? (
                <div style={{
                  // Phone-Frame: schwarzer Bezel mit Notch
                  width: width + 24, padding: '38px 12px 18px', borderRadius: 44,
                  background: '#1a1a1a',
                  boxShadow: '0 8px 32px rgba(0,0,0,0.28), inset 0 0 0 2px #333',
                  position: 'relative',
                }}>
                  {/* Notch */}
                  <div style={{
                    position: 'absolute', top: 12, left: '50%', transform: 'translateX(-50%)',
                    width: 100, height: 18, background: '#000', borderRadius: 10,
                    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                  }}>
                    <span style={{ width: 6, height: 6, borderRadius: '50%', background: '#222', border: '1px solid #333' }} />
                    <span style={{ width: 38, height: 4, borderRadius: 2, background: '#222' }} />
                  </div>
                  <div style={{
                    width, background: '#fff', borderRadius: 28, overflow: 'hidden',
                    minHeight: 620, maxHeight: '72vh', overflowY: 'auto',
                  }}>
                    <div style={{ pointerEvents: 'none', userSelect: 'text' }}>
                      <PreviewContextStack role={props.role} page={props.page} selectedEventId={props.selectedEventId} extraEvents={props.extraEvents}>
                        {props.children}
                      </PreviewContextStack>
                    </div>
                  </div>
                </div>
              ) : (
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
              )}
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
