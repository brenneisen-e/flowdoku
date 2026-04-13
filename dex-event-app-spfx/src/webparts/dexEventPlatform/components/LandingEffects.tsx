/**
 * Landing Page Hintergrund-Effekte (9 Varianten).
 *
 * Wird nur gerendert wenn ein Admin auf der Landing Page einen Effekt aktiviert.
 * Alle Effekte werden absolut hinter dem eigentlichen Content gelegt (z-index 0)
 * und lassen sich ueber ein kleines Dropdown/Cycle-Button durchschalten.
 *
 * Die Keyframes fuer Blobs/Aurora/Orb liegen zentral im Style-Tag von LandingPage.tsx.
 */

import * as React from 'react';
import { useEvents } from '../context/EventContext';

export const EFFECT_LABELS = [
  'Off',                    // 0
  'Mesh Gradient',          // 1
  'Dot Pattern',            // 2
  'Diagonal Stripes',       // 3
  'Animated Blobs',         // 4
  'Aurora',                 // 5
  'Glass Panels',           // 6
  'Canvas Particles',       // 7
  'Giant Orb',              // 8
  'Event Carousel',         // 9
];

export const EFFECT_COUNT = EFFECT_LABELS.length; // 10 (inkl. Off)

interface Props { effectId: number; }

export default function LandingEffects({ effectId }: Props): React.ReactElement | null {
  if (effectId <= 0) return null;
  switch (effectId) {
    case 1: return <MeshGradient />;
    case 2: return <DotPattern />;
    case 3: return <DiagonalStripes />;
    case 4: return <AnimatedBlobs />;
    case 5: return <Aurora />;
    case 6: return <GlassPanels />;
    case 7: return <CanvasParticles />;
    case 8: return <GiantOrb />;
    case 9: return <EventCarousel />;
    default: return null;
  }
}

const layerStyle: React.CSSProperties = {
  position: 'absolute', inset: 0, zIndex: 0, pointerEvents: 'none', overflow: 'hidden',
};

// 1. Mesh Gradient --------------------------------------------------
function MeshGradient(): React.ReactElement {
  return <div aria-hidden="true" style={{
    ...layerStyle,
    background:
      'radial-gradient(at 20% 20%, rgba(134,188,37,0.35), transparent 50%),' +
      'radial-gradient(at 80% 20%, rgba(0,118,168,0.30), transparent 50%),' +
      'radial-gradient(at 20% 80%, rgba(0,188,212,0.25), transparent 50%),' +
      'radial-gradient(at 80% 80%, rgba(255,235,59,0.20), transparent 50%)',
  }} />;
}

// 2. Dot Pattern ----------------------------------------------------
function DotPattern(): React.ReactElement {
  return <div aria-hidden="true" style={{
    ...layerStyle,
    backgroundImage: 'radial-gradient(circle, rgba(134,188,37,0.25) 1.5px, transparent 1.5px)',
    backgroundSize: '22px 22px',
  }} />;
}

// 3. Diagonal Stripes -----------------------------------------------
function DiagonalStripes(): React.ReactElement {
  return <div aria-hidden="true" style={{
    ...layerStyle,
    background: 'linear-gradient(135deg, #f0f8e8 0%, #e8f4ef 35%, #e8f0f8 70%, #f7f0e8 100%)',
  }} />;
}

// 4. Animated Blobs -------------------------------------------------
function AnimatedBlobs(): React.ReactElement {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const blob: React.CSSProperties = { position: 'absolute', borderRadius: '50%', filter: 'blur(80px)' as any, willChange: 'transform' };
  return (
    <div aria-hidden="true" style={layerStyle}>
      <div style={{ ...blob, top: '-8%', left: '-6%', width: 420, height: 420,
        background: 'radial-gradient(circle, #86bc25 0%, rgba(134,188,37,0) 70%)',
        opacity: 0.55, animation: 'dexBlobA 18s ease-in-out infinite' }} />
      <div style={{ ...blob, top: '10%', right: '-10%', width: 480, height: 480,
        background: 'radial-gradient(circle, #0076a8 0%, rgba(0,118,168,0) 70%)',
        opacity: 0.55, animation: 'dexBlobB 22s ease-in-out infinite' }} />
      <div style={{ ...blob, bottom: '-12%', left: '15%', width: 500, height: 500,
        background: 'radial-gradient(circle, #00bcd4 0%, rgba(0,188,212,0) 70%)',
        opacity: 0.45, animation: 'dexBlobC 20s ease-in-out infinite' }} />
      <div style={{ ...blob, bottom: '-6%', right: '5%', width: 380, height: 380,
        background: 'radial-gradient(circle, #ffeb3b 0%, rgba(255,235,59,0) 70%)',
        opacity: 0.35, animation: 'dexBlobD 26s ease-in-out infinite' }} />
    </div>
  );
}

// 5. Aurora ---------------------------------------------------------
function Aurora(): React.ReactElement {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const band: React.CSSProperties = { position: 'absolute', left: '-10%', right: '-10%', height: 260, filter: 'blur(60px)' as any, opacity: 0.55 };
  return (
    <div aria-hidden="true" style={layerStyle}>
      <div style={{ ...band, top: '5%',
        background: 'linear-gradient(90deg, transparent 0%, #86bc25 25%, #00bcd4 55%, #0076a8 80%, transparent 100%)',
        animation: 'dexAuroraSlow 24s ease-in-out infinite' }} />
      <div style={{ ...band, bottom: '5%', opacity: 0.45,
        background: 'linear-gradient(90deg, transparent 0%, #ffeb3b 20%, #86bc25 50%, #00bcd4 80%, transparent 100%)',
        animation: 'dexAuroraSlow 30s ease-in-out infinite reverse' }} />
    </div>
  );
}

// 6. Glass Panels ---------------------------------------------------
function GlassPanels(): React.ReactElement {
  /* eslint-disable @typescript-eslint/no-explicit-any */
  const panel: React.CSSProperties = {
    position: 'absolute', borderRadius: 24,
    backdropFilter: 'blur(12px)' as any,
    WebkitBackdropFilter: 'blur(12px)' as any,
    border: '1px solid rgba(255,255,255,0.35)',
  };
  /* eslint-enable @typescript-eslint/no-explicit-any */
  return (
    <div aria-hidden="true" style={{
      ...layerStyle,
      // Hintergrund, damit der Glass-Effekt sichtbar wird
      background: 'radial-gradient(at 30% 30%, rgba(134,188,37,0.35), transparent 55%),' +
                  'radial-gradient(at 70% 70%, rgba(0,118,168,0.35), transparent 55%)',
    }}>
      <div style={{ ...panel, top: '12%', left: '8%', width: 280, height: 180, transform: 'rotate(-8deg)',
        background: 'rgba(255,255,255,0.25)' }} />
      <div style={{ ...panel, top: '35%', right: '6%', width: 320, height: 210, transform: 'rotate(6deg)',
        background: 'rgba(255,255,255,0.20)' }} />
      <div style={{ ...panel, bottom: '10%', left: '28%', width: 360, height: 200, transform: 'rotate(-3deg)',
        background: 'rgba(255,255,255,0.28)' }} />
    </div>
  );
}

// 7. Canvas Particles -----------------------------------------------
function CanvasParticles(): React.ReactElement {
  const canvasRef = React.useRef<HTMLCanvasElement>(null);
  React.useEffect(() => {
    const canvas = canvasRef.current; if (!canvas) return;
    const ctx = canvas.getContext('2d'); if (!ctx) return;
    const parent = canvas.parentElement; if (!parent) return;

    let w = 0, h = 0;
    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    const resize = (): void => {
      const rect = parent.getBoundingClientRect();
      w = rect.width; h = rect.height;
      canvas.width = w * dpr; canvas.height = h * dpr;
      canvas.style.width = `${w}px`; canvas.style.height = `${h}px`;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    };
    resize();
    window.addEventListener('resize', resize);

    const N = 55;
    const parts = Array.from({ length: N }, () => ({
      x: Math.random() * w, y: Math.random() * h,
      vx: (Math.random() - 0.5) * 0.35, vy: (Math.random() - 0.5) * 0.35,
      r: 1.2 + Math.random() * 1.6,
    }));

    let raf = 0;
    const tick = (): void => {
      ctx.clearRect(0, 0, w, h);
      for (const p of parts) {
        p.x += p.vx; p.y += p.vy;
        if (p.x < 0 || p.x > w) p.vx *= -1;
        if (p.y < 0 || p.y > h) p.vy *= -1;
      }
      // Linien zwischen nahen Punkten
      ctx.lineWidth = 1;
      for (let i = 0; i < N; i++) {
        for (let j = i + 1; j < N; j++) {
          const dx = parts[i].x - parts[j].x;
          const dy = parts[i].y - parts[j].y;
          const d = Math.sqrt(dx * dx + dy * dy);
          if (d < 140) {
            ctx.strokeStyle = `rgba(134,188,37,${(1 - d / 140) * 0.35})`;
            ctx.beginPath();
            ctx.moveTo(parts[i].x, parts[i].y);
            ctx.lineTo(parts[j].x, parts[j].y);
            ctx.stroke();
          }
        }
      }
      // Punkte
      ctx.fillStyle = 'rgba(134,188,37,0.8)';
      for (const p of parts) {
        ctx.beginPath(); ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2); ctx.fill();
      }
      raf = requestAnimationFrame(tick);
    };
    tick();
    return () => { window.removeEventListener('resize', resize); cancelAnimationFrame(raf); };
  }, []);
  return <div aria-hidden="true" style={layerStyle}><canvas ref={canvasRef} style={{ display: 'block' }} /></div>;
}

// 8. Giant Orb ------------------------------------------------------
function GiantOrb(): React.ReactElement {
  return (
    <div aria-hidden="true" style={layerStyle}>
      <div style={{
        position: 'absolute', top: '50%', left: '50%', width: '90vmin', height: '90vmin',
        transform: 'translate(-50%, -50%)',
        borderRadius: '50%',
        background: 'conic-gradient(from 0deg, #86bc25, #0076a8 90deg, #00bcd4 150deg, #4caf50 210deg, #8bc34a 270deg, #ffeb3b 320deg, #86bc25 360deg)',
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        filter: 'blur(60px)' as any,
        opacity: 0.30,
        animation: 'dexOrbSpin 40s linear infinite',
      }} />
    </div>
  );
}

// 9. Event Carousel -------------------------------------------------
function EventCarousel(): React.ReactElement {
  const { events } = useEvents();
  const now = Date.now();
  const upcoming = events
    .filter(e => e.status === 'Active' && e.startDate && new Date(e.startDate).getTime() >= now)
    .sort((a, b) => new Date(a.startDate).getTime() - new Date(b.startDate).getTime())
    .slice(0, 5);
  const [idx, setIdx] = React.useState(0);
  React.useEffect(() => {
    if (upcoming.length === 0) return;
    const h = setInterval(() => setIdx(i => (i + 1) % upcoming.length), 4000);
    return () => clearInterval(h);
  }, [upcoming.length]);

  if (upcoming.length === 0) {
    return <div aria-hidden="true" style={layerStyle}>
      {/* Fallback: dezenter Mesh-Gradient falls keine Events */}
      <MeshGradient />
    </div>;
  }
  const fmt = (iso: string): string => new Date(iso).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  return (
    <div aria-hidden="true" style={{ ...layerStyle, background: 'linear-gradient(135deg, #f0f8e8 0%, #e8f4ef 100%)' }}>
      <div style={{
        position: 'absolute', left: 0, right: 0, bottom: 24,
        display: 'flex', gap: 14, justifyContent: 'center', padding: '0 16px',
        pointerEvents: 'none',
      }}>
        {upcoming.map((e, i) => (
          <div key={e.id} style={{
            width: i === idx ? 280 : 180, height: i === idx ? 110 : 90,
            borderRadius: 14, overflow: 'hidden', position: 'relative',
            background: e.imageUrl ? `url(${e.imageUrl}) center/cover no-repeat`
              : 'linear-gradient(135deg, #86bc25 0%, #0076a8 100%)',
            boxShadow: i === idx ? '0 8px 24px rgba(0,0,0,0.18)' : '0 2px 8px rgba(0,0,0,0.08)',
            opacity: i === idx ? 1 : 0.7,
            transition: 'all 0.6s ease',
          }}>
            <div style={{
              position: 'absolute', inset: 0,
              background: 'linear-gradient(transparent 50%, rgba(0,0,0,0.7))',
              color: '#fff', padding: 10, display: 'flex', flexDirection: 'column', justifyContent: 'flex-end',
              fontSize: i === idx ? '0.85rem' : '0.72rem', fontWeight: 600,
            }}>
              <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{e.title}</div>
              <div style={{ fontSize: '0.7rem', opacity: 0.85, fontWeight: 400 }}>{fmt(e.startDate)}</div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
