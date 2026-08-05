/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';

/**
 * DexLogo
 * -------
 * Animiertes DEX-Logo als Canvas-Grafik. Keine externen Dependencies,
 * kompatibel mit SPFx (React 17 / TypeScript 4.x, target es5).
 *
 * Rendering: Höhenlinien werden als geschlossene Schleifen auf einer
 * Kugeloberfläche erzeugt, pro Frame um die Y-Achse rotiert und
 * orthographisch projiziert. Die Rückseite wird weggeschnitten.
 */

export interface IDexLogoPalette {
  /** Flächenfarbe der Kugel, im Original durchgehend gleich */
  core: string;
  /** Linienfarbe im Zentrum, gedämpft */
  mid: string;
  /** Linienfarbe im mittleren Bereich */
  line: string;
  /** Linienfarbe am Rand, am hellsten */
  edge: string;
  /** Farbe der Partikel */
  particle: string;
}

/**
 * oscillate – sanftes Pendeln, das Linienzentrum bleibt durchgehend sichtbar
 * rotate    – volle Umdrehung. Das Linienzentrum wandert dabei auf die
 *             Rückseite, das Logo ist zeitweise nicht wiedererkennbar.
 */
export type DexLogoMotion = 'rotate' | 'oscillate';

export interface IDexLogoProps {
  /** Kantenlänge in px. Ohne Angabe füllt die Komponente die Breite des Containers. */
  size?: number;
  /** Bewegungsmodus, Standard 'oscillate' */
  motion?: DexLogoMotion;
  /** Geschwindigkeit. 1 = ein voller Zyklus in ca. 12 s, 2 = doppelt so schnell */
  speed?: number;
  /** Anzahl der Höhenlinien, Standard 42 */
  contours?: number;
  /** Anzahl der Partikel auf der Oberfläche */
  particles?: number;
  /** Farben überschreiben (teilweise möglich) */
  palette?: Partial<IDexLogoPalette>;
  /** Animation anhalten, Standbild bleibt sichtbar */
  paused?: boolean;
  /** Seed für die Form – gleicher Seed erzeugt immer dasselbe Muster */
  seed?: number;
  /**
   * v28.34: Maus-Antrieb. Faehrt der Zeiger ueber die Kugel, dreht sie sich in
   * die Bewegungsrichtung — je schneller die Maus, desto kraeftiger der
   * Anschub. Danach laeuft sie mit Reibung aus und driftet langsam in die
   * Ruhelage zurueck, damit das Linienzentrum sichtbar bleibt. Standard: an.
   */
  pointerSpin?: boolean;
  className?: string;
  style?: React.CSSProperties;
  /** Wird als aria-label gesetzt. Ohne Angabe gilt die Grafik als dekorativ. */
  title?: string;
}

const TAU = Math.PI * 2;

/** Werte aus dem Original-Logo ausgemessen, nicht geschätzt. */
const DEFAULT_PALETTE: IDexLogoPalette = {
  core: '#0D1603',
  mid: '#5D8918',
  line: '#86BC25',
  edge: '#C2F240',
  particle: '#C8E39A'
};

/* ------------------------------------------------------------------ */
/* Hilfsfunktionen                                                     */
/* ------------------------------------------------------------------ */

/** Park–Miller PRNG, bewusst ohne Math.imul (fehlt in manchen SPFx-libs). */
function createRandom(seed: number): () => number {
  let s = seed % 2147483647;
  if (s <= 0) {
    s += 2147483646;
  }
  return function (): number {
    s = (s * 16807) % 2147483647;
    return (s - 1) / 2147483646;
  };
}

function toRgba(hex: string, alpha: number): string {
  let h = hex.charAt(0) === '#' ? hex.substring(1) : hex;
  if (h.length === 3) {
    h = h.charAt(0) + h.charAt(0) + h.charAt(1) + h.charAt(1) + h.charAt(2) + h.charAt(2);
  }
  const num = parseInt(h, 16);
  const r = (num >> 16) & 255;
  const g = (num >> 8) & 255;
  const b = num & 255;
  return 'rgba(' + r + ',' + g + ',' + b + ',' + alpha + ')';
}

function mergePalette(overrides?: Partial<IDexLogoPalette>): IDexLogoPalette {
  const p: IDexLogoPalette = {
    core: DEFAULT_PALETTE.core,
    mid: DEFAULT_PALETTE.mid,
    line: DEFAULT_PALETTE.line,
    edge: DEFAULT_PALETTE.edge,
    particle: DEFAULT_PALETTE.particle
  };
  if (overrides) {
    if (overrides.core) { p.core = overrides.core; }
    if (overrides.mid) { p.mid = overrides.mid; }
    if (overrides.line) { p.line = overrides.line; }
    if (overrides.edge) { p.edge = overrides.edge; }
    if (overrides.particle) { p.particle = overrides.particle; }
  }
  return p;
}

/* ------------------------------------------------------------------ */
/* Geometrie                                                           */
/* ------------------------------------------------------------------ */

interface IGeometry {
  /** Pro Höhenlinie ein Float32Array aus xyz-Tripeln, letzter Punkt = erster Punkt */
  contours: Float32Array[];
  /** Partikel als xyzr-Quadrupel */
  particles: Float32Array;
}

function buildGeometry(contourCount: number, particleCount: number, seed: number): IGeometry {
  const rnd = createRandom(seed);
  const segments = 190;

  // Mittelpunkt der Höhenlinien, aus dem Original-Logo ausgemessen
  const lat = -0.676;
  const lon = 0.207;
  const cX = Math.cos(lat) * Math.sin(lon);
  const cY = Math.sin(lat);
  const cZ = Math.cos(lat) * Math.cos(lon);

  // Orthonormalbasis um die Achse c aufspannen
  let e1x = -cZ;
  let e1y = 0;
  let e1z = cX;
  const e1len = Math.sqrt(e1x * e1x + e1y * e1y + e1z * e1z) || 1;
  e1x /= e1len; e1y /= e1len; e1z /= e1len;

  const e2x = cY * e1z - cZ * e1y;
  const e2y = cZ * e1x - cX * e1z;
  const e2z = cX * e1y - cY * e1x;

  // Startradius, Endradius und Staffelung sind am Original ausgemessen:
  // innen weite Abstände, zum Rand hin starke Verdichtung. maxAlpha reicht
  // bewusst über den in Ruhelage sichtbaren Bereich hinaus, sonst entsteht
  // bei ausgelenkter Kugel ein unbedeckter Streifen am Rand.
  const startAlpha = 0.17;
  const maxAlpha = 2.62;
  const falloff = 0.70;
  const spread = maxAlpha - startAlpha;
  const spacing = spread / contourCount;
  const steps = contourCount > 1 ? contourCount - 1 : 1;

  // Feste Phasen, damit benachbarte Linien sich ähneln und nie kreuzen
  const ph1 = rnd() * TAU;
  const ph2 = rnd() * TAU;
  const ph3 = rnd() * TAU;

  const contours: Float32Array[] = [];

  for (let i = 0; i < contourCount; i++) {
    const base = startAlpha + spread * Math.pow(i / steps, falloff);
    const pts = new Float32Array((segments + 1) * 3);

    for (let s = 0; s <= segments; s++) {
      const t = (s % segments) / segments * TAU;

      // Normierte Wellenform, Summe der Amplituden = 1
      const wobble =
        0.50 * Math.sin(2 * t + ph1 + i * 0.11) +
        0.30 * Math.sin(3 * t + ph2 - i * 0.08) +
        0.20 * Math.sin(5 * t + ph3 + i * 0.05);

      // Auslenkung bleibt unter der halben Linienabstand -> keine Überschneidungen
      let a = base + spacing * 0.34 * wobble;
      if (a < 0.02) { a = 0.02; }
      if (a > Math.PI - 0.02) { a = Math.PI - 0.02; }

      const sa = Math.sin(a);
      const ca = Math.cos(a);
      const ct = Math.cos(t);
      const st = Math.sin(t);

      const o = s * 3;
      pts[o] = cX * ca + (e1x * ct + e2x * st) * sa;
      pts[o + 1] = cY * ca + (e1y * ct + e2y * st) * sa;
      pts[o + 2] = cZ * ca + (e1z * ct + e2z * st) * sa;
    }

    contours.push(pts);
  }

  // Partikel gleichverteilt auf der Kugeloberfläche
  const particles = new Float32Array(particleCount * 4);
  for (let p = 0; p < particleCount; p++) {
    const z = 1 - 2 * rnd();
    const r = Math.sqrt(Math.max(0, 1 - z * z));
    const phi = rnd() * TAU;
    const o = p * 4;
    particles[o] = r * Math.cos(phi);
    particles[o + 1] = r * Math.sin(phi);
    particles[o + 2] = z;
    particles[o + 3] = 0.35 + rnd() * rnd() * 1.15;
  }

  return { contours: contours, particles: particles };
}

/* ------------------------------------------------------------------ */
/* Zeichnen                                                            */
/* ------------------------------------------------------------------ */

function drawFrame(
  ctx: CanvasRenderingContext2D,
  geo: IGeometry,
  palette: IDexLogoPalette,
  width: number,
  height: number,
  elapsed: number,
  speed: number,
  motion: DexLogoMotion,
  extraRot: number
): void {
  ctx.clearRect(0, 0, width, height);

  const cx = width / 2;
  const cy = height / 2;
  const R = Math.min(width, height) * 0.44;
  if (R <= 0) { return; }

  // Bei speed = 1 dauert ein voller Zyklus rund 12 s
  const rot = (motion === 'oscillate'
    ? Math.sin(elapsed * 0.00052 * speed) * 0.40
    : elapsed * 0.00052 * speed) + extraRot;
  const cosR = Math.cos(rot);
  const sinR = Math.sin(rot);

  /* 1 — Kugelkörper: flache Fläche, im Original gibt es hier keinen Verlauf */
  ctx.save();
  ctx.beginPath();
  ctx.arc(cx, cy, R, 0, TAU);
  ctx.clip();

  ctx.fillStyle = toRgba(palette.core, 1);
  ctx.fill();

  /* 2 — Höhenlinien: ein Pfad, ein Strich, Helligkeit über Radialverlauf */
  ctx.beginPath();
  for (let c = 0; c < geo.contours.length; c++) {
    const pts = geo.contours[c];
    let drawing = false;

    for (let o = 0; o < pts.length; o += 3) {
      const x = pts[o];
      const y = pts[o + 1];
      const z = pts[o + 2];

      const rz = z * cosR - x * sinR;
      if (rz <= 0) {
        drawing = false;
        continue;
      }
      const rx = x * cosR + z * sinR;

      const sx = cx + rx * R;
      const sy = cy - y * R;

      if (drawing) {
        ctx.lineTo(sx, sy);
      } else {
        ctx.moveTo(sx, sy);
        drawing = true;
      }
    }
  }

  const lineGrad = ctx.createRadialGradient(cx, cy, 0, cx, cy, R);
  lineGrad.addColorStop(0, toRgba(palette.mid, 0.29));
  lineGrad.addColorStop(0.5, toRgba(palette.mid, 0.42));
  lineGrad.addColorStop(0.8, toRgba(palette.line, 0.98));
  lineGrad.addColorStop(0.93, toRgba(palette.edge, 0.95));
  lineGrad.addColorStop(1, toRgba(palette.edge, 0.33));

  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  ctx.strokeStyle = lineGrad;

  // Sehr dezenter Streulichtsaum, entspricht der weichen Zeichnung im Original
  ctx.globalAlpha = 0.065;
  ctx.lineWidth = Math.max(1.2, R * 0.022);
  ctx.stroke();

  // Kontur, gemessene Breite im Original: 0.84 bis 1.1 % des Radius
  ctx.globalAlpha = 1;
  ctx.lineWidth = Math.max(0.65, R * 0.0105);
  ctx.stroke();

  /* 3 — Partikel auf der Oberfläche */
  const pts = geo.particles;
  const scale = R / 150;
  ctx.beginPath();
  for (let o = 0; o < pts.length; o += 4) {
    const x = pts[o];
    const y = pts[o + 1];
    const z = pts[o + 2];

    const rz = z * cosR - x * sinR;
    if (rz <= 0.02) { continue; }
    const rx = x * cosR + z * sinR;

    const sx = cx + rx * R;
    const sy = cy - y * R;
    const rad = pts[o + 3] * scale * (0.5 + 0.5 * rz);
    if (rad <= 0.05) { continue; }

    ctx.moveTo(sx + rad, sy);
    ctx.arc(sx, sy, rad, 0, TAU);
  }
  ctx.fillStyle = toRgba(palette.particle, 0.8);
  ctx.fill();

  ctx.restore();
}

/* ------------------------------------------------------------------ */
/* Komponente                                                          */
/* ------------------------------------------------------------------ */

export const DexLogo: React.FC<IDexLogoProps> = (props: IDexLogoProps) => {
  const {
    size,
    motion = 'oscillate',
    speed = 1,
    contours = 42,
    particles = 340,
    palette,
    paused = false,
    seed = 7391,
    pointerSpin = true,
    className,
    style,
    title
  } = props;

  const hostRef = React.useRef<HTMLDivElement | null>(null);
  const canvasRef = React.useRef<HTMLCanvasElement | null>(null);
  const rafRef = React.useRef<number>(0);
  const elapsedRef = React.useRef<number>(0);
  const lastTsRef = React.useRef<number>(0);
  /** Vom Zeiger erzeugter Zusatz-Winkel und dessen Geschwindigkeit (rad/ms). */
  const userRotRef = React.useRef<number>(0);
  const userVelRef = React.useRef<number>(0);

  const [measured, setMeasured] = React.useState<number>(size || 0);
  const [reducedMotion, setReducedMotion] = React.useState<boolean>(false);
  const [visible, setVisible] = React.useState<boolean>(true);

  const geometry = React.useMemo<IGeometry>(
    () => buildGeometry(contours, particles, seed),
    [contours, particles, seed]
  );

  const colors = React.useMemo<IDexLogoPalette>(() => mergePalette(palette), [palette]);

  /* Breite messen, wenn keine feste Größe gesetzt ist */
  React.useEffect(() => {
    if (size) {
      setMeasured(size);
      return undefined;
    }

    const measure = (): void => {
      const host = hostRef.current;
      if (host && host.clientWidth > 0) {
        setMeasured(host.clientWidth);
      }
    };

    measure();

    const RO: any = (window as any).ResizeObserver;
    if (RO && hostRef.current) {
      const ro = new RO(measure);
      ro.observe(hostRef.current);
      return () => ro.disconnect();
    }

    window.addEventListener('resize', measure);
    return () => window.removeEventListener('resize', measure);
  }, [size]);

  /* Systemeinstellung "Bewegung reduzieren" respektieren */
  React.useEffect(() => {
    if (!window.matchMedia) { return undefined; }
    const mq = window.matchMedia('(prefers-reduced-motion: reduce)');
    const apply = (): void => setReducedMotion(mq.matches);
    apply();

    if (mq.addEventListener) {
      mq.addEventListener('change', apply);
      return () => mq.removeEventListener('change', apply);
    }
    (mq as any).addListener(apply);
    return () => (mq as any).removeListener(apply);
  }, []);

  /* Nur animieren, solange die Komponente im Viewport und der Tab aktiv ist */
  React.useEffect(() => {
    const onVisibility = (): void => setVisible(!document.hidden);
    document.addEventListener('visibilitychange', onVisibility);

    const IO: any = (window as any).IntersectionObserver;
    let io: any = null;
    if (IO && hostRef.current) {
      io = new IO(
        (entries: any[]) => setVisible(!document.hidden && entries[0].isIntersecting),
        { threshold: 0 }
      );
      io.observe(hostRef.current);
    }

    return () => {
      document.removeEventListener('visibilitychange', onVisibility);
      if (io) { io.disconnect(); }
    };
  }, []);

  /* Renderloop */
  React.useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || measured <= 0) { return undefined; }

    const ctx = canvas.getContext('2d');
    if (!ctx) { return undefined; }

    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    canvas.width = Math.round(measured * dpr);
    canvas.height = Math.round(measured * dpr);
    canvas.style.width = measured + 'px';
    canvas.style.height = measured + 'px';
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    const animate = !paused && visible && !reducedMotion;

    if (!animate) {
      drawFrame(ctx, geometry, colors, measured, measured, elapsedRef.current, speed, motion, userRotRef.current);
      return undefined;
    }

    lastTsRef.current = 0;

    const loop = (ts: number): void => {
      if (lastTsRef.current === 0) { lastTsRef.current = ts; }
      const delta = Math.min(ts - lastTsRef.current, 64);
      lastTsRef.current = ts;
      elapsedRef.current += delta;

      // v28.34: Maus-Schwung ausrollen lassen. Reibung bremst die
      // Geschwindigkeit, eine sehr weiche Rueckstellung zieht die Kugel
      // ueber ein paar Sekunden in die Ruhelage — sonst koennte das
      // Linienzentrum dauerhaft auf der Rueckseite stehenbleiben.
      const frames = delta / 16.67;
      userRotRef.current += userVelRef.current * delta;
      userVelRef.current *= Math.pow(0.93, frames);
      if (Math.abs(userVelRef.current) < 0.000002) { userVelRef.current = 0; }
      userRotRef.current *= Math.pow(0.9975, frames);

      drawFrame(ctx, geometry, colors, measured, measured, elapsedRef.current, speed, motion, userRotRef.current);
      rafRef.current = window.requestAnimationFrame(loop);
    };

    rafRef.current = window.requestAnimationFrame(loop);
    return () => window.cancelAnimationFrame(rafRef.current);
  }, [measured, geometry, colors, speed, motion, paused, visible, reducedMotion]);

  const hostStyle: React.CSSProperties = {
    display: 'inline-block',
    lineHeight: 0,
    width: size ? size : '100%',
    // v28.34: Zeiger-Hinweis, dass sich die Kugel anschubsen laesst.
    cursor: (pointerSpin && !paused && !reducedMotion) ? 'grab' : undefined
  };
  if (style) {
    for (const key in style) {
      if (Object.prototype.hasOwnProperty.call(style, key)) {
        (hostStyle as any)[key] = (style as any)[key];
      }
    }
  }

  /**
   * v28.34: Jede Zeigerbewegung gibt einen Impuls in Bewegungsrichtung. Die
   * Staerke haengt an der zurueckgelegten Strecke pro Event — also an der
   * Maus-Geschwindigkeit. Normiert auf die Kugelbreite, damit sich das Logo
   * in jeder Groesse gleich anfuehlt; gedeckelt, damit ein Ruck nicht
   * mehrere Umdrehungen ausloest.
   */
  const lastPointerXRef = React.useRef<number | null>(null);
  const onPointerMove = (e: React.PointerEvent<HTMLDivElement>): void => {
    if (!pointerSpin || paused || reducedMotion || measured <= 0) { return; }
    const last = lastPointerXRef.current;
    lastPointerXRef.current = e.clientX;
    if (last === null) { return; }
    const dx = e.clientX - last;
    if (dx === 0) { return; }
    let impulse = (dx / measured) * 0.0016;
    if (impulse > 0.0022) { impulse = 0.0022; }
    if (impulse < -0.0022) { impulse = -0.0022; }
    userVelRef.current += impulse;
    if (userVelRef.current > 0.006) { userVelRef.current = 0.006; }
    if (userVelRef.current < -0.006) { userVelRef.current = -0.006; }
  };
  const onPointerLeave = (): void => { lastPointerXRef.current = null; };

  return (
    <div
      ref={hostRef}
      className={className}
      onPointerMove={onPointerMove}
      onPointerLeave={onPointerLeave}
      style={hostStyle}
      role={title ? 'img' : undefined}
      aria-label={title}
      aria-hidden={title ? undefined : true}
    >
      <canvas ref={canvasRef} style={{ display: 'block' }} />
    </div>
  );
};

export default DexLogo;
