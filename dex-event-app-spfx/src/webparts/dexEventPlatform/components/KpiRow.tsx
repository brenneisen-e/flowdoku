/* v32.0.10: Aus LandingPage herausgelöst — genutzt auf dem Ladebildschirm
 * (DexEventPlatform) und im Info-Fenster (LandingInfoModal). */
import * as React from 'react';

// v11.47: KPI-Box-Reihe über dem "Entwickelt von ..."-Block. Drei Boxen
// nebeneinander: gehostete Events, Teilnehmer, App-Aufrufe. Jede Box mit
// einer AnimatedCounter-Komponente, die beim ersten Verfügbarwerden des
// Werts von 0 zum Zielwert ease-out hochzählt (~1.6s). Solange Daten noch
// laden, steht ein dezenter Skeleton-Punkt im Wert-Feld.
export function KpiRow(props: {
  locale: string;
  eventsLoading: boolean;
  participantsLoading: boolean;
  events: number;
  participants: number;
}): React.ReactElement {
  const isDe = props.locale === 'de';
  const labels = isDe
    ? { ev: 'Events', pa: 'Teilnehmer' }
    : { ev: 'Events', pa: 'Attendees' };
  return (
    <div style={{
      display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 10,
    }}>
      <KpiBox label={labels.ev} value={props.events} loading={props.eventsLoading} />
      <KpiBox label={labels.pa} value={props.participants} loading={props.participantsLoading} />
    </div>
  );
}

function KpiBox(props: { label: string; value: number; loading: boolean }): React.ReactElement {
  return (
    <div style={{
      display: 'flex', flexDirection: 'column', alignItems: 'center',
      padding: '12px 8px',
      background: 'linear-gradient(135deg, rgba(134,188,37,0.08), rgba(0,118,168,0.04))',
      border: '1px solid var(--dex-gray-200)',
      borderRadius: 12,
      minWidth: 0,
    }}>
      <div style={{
        fontSize: 'clamp(1.4rem, 4vw, 1.9rem)', fontWeight: 800,
        color: 'var(--dex-green-dark, #4a7c1f)',
        lineHeight: 1.1, letterSpacing: '-0.02em',
        fontVariantNumeric: 'tabular-nums',
      }}>
        {props.loading
          ? <SkeletonDots />
          : <AnimatedCounter value={props.value} />}
      </div>
      <div style={{
        marginTop: 6,
        fontSize: '0.72rem', color: 'var(--dex-gray-500)',
        fontWeight: 500, textAlign: 'center', lineHeight: 1.2,
      }}>
        {props.label}
      </div>
    </div>
  );
}

function SkeletonDots(): React.ReactElement {
  return (
    <span aria-hidden="true" style={{ display: 'inline-flex', gap: 4, opacity: 0.55 }}>
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: 'currentColor', animation: 'dexKpiPulse 1.2s ease-in-out infinite' }} />
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: 'currentColor', animation: 'dexKpiPulse 1.2s ease-in-out 0.2s infinite' }} />
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: 'currentColor', animation: 'dexKpiPulse 1.2s ease-in-out 0.4s infinite' }} />
      <style>{`
        @keyframes dexKpiPulse {
          0%, 100% { transform: scale(0.6); opacity: 0.3; }
          50%      { transform: scale(1);   opacity: 1; }
        }
      `}</style>
    </span>
  );
}

/** Ease-out-Cubic-Animation von 0 (bzw. dem letzten geanimierten Wert) auf
 *  `value`. Wenn `value` sich ändert, startet eine neue Animation; bei
 *  schnellen Änderungen wird die laufende Animation sauber durch eine neue
 *  ersetzt (requestAnimationFrame + AbortFlag). */
function AnimatedCounter(props: { value: number; durationMs?: number }): React.ReactElement {
  const target = Math.max(0, Math.floor(props.value || 0));
  // v11.79: Default-Dauer von 1600 ms → 600 ms reduziert. Seit der App-Boot
  // unter ~1.6 s liegt, wirkte das längere Count-Up-Tempo träge; die Zahl
  // tickt jetzt knackiger hoch ohne hektisch zu wirken.
  const duration = props.durationMs ?? 600;
  const [shown, setShown] = React.useState<number>(0);
  const startRef = React.useRef<number>(0);
  const fromRef = React.useRef<number>(0);
  const rafRef = React.useRef<number | null>(null);
  React.useEffect(() => {
    if (rafRef.current !== null) cancelAnimationFrame(rafRef.current);
    fromRef.current = shown;
    startRef.current = 0;
    const step = (ts: number): void => {
      if (startRef.current === 0) startRef.current = ts;
      const elapsed = ts - startRef.current;
      const t = Math.min(1, elapsed / duration);
      // ease-out-cubic
      const eased = 1 - Math.pow(1 - t, 3);
      const value = fromRef.current + (target - fromRef.current) * eased;
      setShown(Math.round(value));
      if (t < 1) {
        rafRef.current = requestAnimationFrame(step);
      } else {
        rafRef.current = null;
      }
    };
    rafRef.current = requestAnimationFrame(step);
    return () => {
      if (rafRef.current !== null) cancelAnimationFrame(rafRef.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [target, duration]);
  return <span>{shown.toLocaleString('de-DE')}</span>;
}
