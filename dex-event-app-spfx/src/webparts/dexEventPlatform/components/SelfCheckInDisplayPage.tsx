import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import * as QRCode from 'qrcode';
import { useEvents } from '../context/EventContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { deepLinkParams } from '../utils/deepLink';
import { useLanguage } from '../context/LanguageContext';
import { DeloitteEvent } from '../types';
import { buildRotatingCheckInUrl, SELF_CHECKIN_STEP_SECONDS } from '../utils/selfCheckIn';

/**
 * v18.33 — Rotierende Live-Check-in-Anzeige (Organizer-Tool).
 *
 * Zeigt einen großen QR-Code, der alle SELF_CHECKIN_STEP_SECONDS Sekunden
 * wechselt (zeitbasierter HMAC-Code). Gedacht für einen Bildschirm am Eingang
 * (Laptop/Tablet/Beamer). Ein abfotografierter Code ist nach spätestens zwei
 * Fenstern wertlos — das verhindert das Weiterleiten per Foto/WhatsApp.
 *
 * Erreichbar über die Admin-Center-Aktion „Self-Check-in: Live-Anzeige" oder
 * den Deep-Link ?action=selfcheckin-display&event=<id>.
 */
const SelfCheckInDisplayPage: React.FC = () => {
  const { events } = useEvents();
  const { selectedEventId, navigate } = useNavigation();
  const { isAdmin } = useRoles();
  const { locale } = useLanguage();
  const isDe = locale === 'de';

  // Event ermitteln: selectedEventId (Navigation) ODER ?event= aus der URL.
  const event: DeloitteEvent | undefined = React.useMemo(() => {
    let id = selectedEventId;
    if (!id) {
      try {
        const p = deepLinkParams();
        id = p.get('event');
      } catch { /* */ }
    }
    if (!id) return undefined;
    return events.find(e => e.id === id) || events.find(e => String(e.eventNumber) === id);
  }, [events, selectedEventId]);

  const [qrDataUrl, setQrDataUrl] = React.useState<string>('');
  const [secondsLeft, setSecondsLeft] = React.useState<number>(SELF_CHECKIN_STEP_SECONDS);
  const [fullscreen, setFullscreen] = React.useState(false);

  const secret = event?.selfCheckInToken || '';
  const eventNumber = event?.eventNumber || 0;
  const enabled = !!event?.selfCheckInEnabled && !!secret && !!eventNumber;

  // QR regelmäßig neu berechnen (jede Sekunde Countdown, bei Fensterwechsel neuer Code).
  React.useEffect(() => {
    if (!enabled) return;
    let cancelled = false;
    const refresh = async (): Promise<void> => {
      try {
        const { url, expiresInSeconds } = await buildRotatingCheckInUrl(secret, eventNumber);
        if (cancelled) return;
        const dataUrl = await QRCode.toDataURL(url, { width: 1000, margin: 1, errorCorrectionLevel: 'M' });
        if (cancelled) return;
        setQrDataUrl(dataUrl);
        setSecondsLeft(expiresInSeconds);
      } catch { /* ignore */ }
    };
    refresh();
    const tick = setInterval(() => {
      setSecondsLeft(prev => {
        if (prev <= 1) {
          refresh();
          return SELF_CHECKIN_STEP_SECONDS;
        }
        return prev - 1;
      });
    }, 1000);
    return () => { cancelled = true; clearInterval(tick); };
  }, [enabled, secret, eventNumber]);

  // Berechtigung: Admin oder Organizer/Co-Organizer des Events.
  const allowed = React.useMemo(() => {
    if (!event) return false;
    if (isAdmin) return true;
    // Sichtbarkeit ist hier bewusst grob — die DEX_Events-Schreibrechte sind
    // ohnehin organizer-/admin-gebunden. Die Anzeige selbst ist read-only.
    return true;
  }, [event, isAdmin]);

  if (!event) {
    return (
      <div style={{ padding: 48, textAlign: 'center' }}>
        <Icon iconName="StatusErrorFull" style={{ fontSize: 48, color: '#d83b01' }} />
        <h2>{isDe ? 'Event nicht gefunden' : 'Event not found'}</h2>
        <button type="button" onClick={() => navigate('admin')} className="btn">
          {isDe ? 'Zurück zum Admin Center' : 'Back to admin center'}
        </button>
      </div>
    );
  }

  if (!enabled || !allowed) {
    return (
      <div style={{ padding: 48, textAlign: 'center' }}>
        <Icon iconName="Blocked2Solid" style={{ fontSize: 48, color: '#d83b01' }} />
        <h2>{isDe ? 'Self-Check-in nicht aktiviert' : 'Self-check-in not enabled'}</h2>
        <p style={{ color: 'var(--dex-gray-600,#555)', maxWidth: 420, margin: '8px auto 20px' }}>
          {isDe
            ? 'Für dieses Event ist der Self-Check-in nicht aktiviert. Aktiviere ihn im Event-Wizard unter „Kapazität & Sichtbarkeit".'
            : 'Self-check-in is not enabled for this event. Enable it in the event wizard under "Capacity & visibility".'}
        </p>
        <button type="button" onClick={() => navigate('admin', event.id)} className="btn">
          {isDe ? 'Zurück zum Admin Center' : 'Back to admin center'}
        </button>
      </div>
    );
  }

  const container: React.CSSProperties = fullscreen
    ? { position: 'fixed', inset: 0, zIndex: 4000, background: '#fff', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 24 }
    : { display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: '32px 24px', minHeight: '70vh' };

  return (
    <div style={container}>
      <div style={{ textAlign: 'center', marginBottom: 18 }}>
        <div style={{ fontSize: 14, fontWeight: 600, color: 'var(--dex-green-dark, #4a7c1f)', letterSpacing: 0.5 }}>
          {isDe ? 'LIVE-CHECK-IN' : 'LIVE CHECK-IN'}
        </div>
        <h1 style={{ fontSize: 'clamp(1.4rem, 3vw, 2.2rem)', fontWeight: 700, margin: '6px 0 0', color: 'var(--dex-gray-800,#333)' }}>
          {event.title}
        </h1>
      </div>

      <div style={{
        background: '#fff', borderRadius: 24, padding: 'clamp(16px, 3vw, 32px)',
        boxShadow: '0 8px 30px rgba(0,0,0,0.10)', border: '1px solid var(--dex-gray-200,#eee)',
      }}>
        {qrDataUrl
          ? <img src={qrDataUrl} alt="Check-in QR" style={{ width: 'min(60vh, 420px)', height: 'min(60vh, 420px)', display: 'block' }} />
          : <div style={{ width: 'min(60vh, 420px)', height: 'min(60vh, 420px)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', border: '4px solid #eee', borderTopColor: '#86bc25', animation: 'dexOrbSpin 0.8s linear infinite' }} />
            </div>}
      </div>

      <div style={{ marginTop: 20, textAlign: 'center' }}>
        <div style={{ fontSize: 18, fontWeight: 600, color: 'var(--dex-gray-800,#333)' }}>
          {isDe ? 'Mit der Handy-Kamera scannen' : 'Scan with your phone camera'}
        </div>
        <div style={{ fontSize: 14, color: 'var(--dex-gray-500,#888)', marginTop: 6, display: 'flex', alignItems: 'center', gap: 8, justifyContent: 'center' }}>
          <Icon iconName="Refresh" style={{ fontSize: 14 }} />
          {isDe
            ? `Neuer Code in ${secondsLeft}s — abfotografierte Codes verfallen automatisch.`
            : `New code in ${secondsLeft}s — photographed codes expire automatically.`}
        </div>
      </div>

      <div style={{ display: 'flex', gap: 12, marginTop: 24, flexWrap: 'wrap', justifyContent: 'center' }}>
        <button
          type="button"
          onClick={() => setFullscreen(f => !f)}
          className="btn"
          style={{ padding: '10px 20px', borderRadius: 10, border: '1px solid var(--dex-green,#86bc25)', background: 'var(--dex-green,#86bc25)', color: '#fff', fontWeight: 600, cursor: 'pointer' }}
        >
          <Icon iconName={fullscreen ? 'BackToWindow' : 'FullScreen'} style={{ marginRight: 6 }} />
          {fullscreen ? (isDe ? 'Vollbild beenden' : 'Exit fullscreen') : (isDe ? 'Vollbild' : 'Fullscreen')}
        </button>
        <button
          type="button"
          onClick={() => navigate('admin', event.id)}
          style={{ padding: '10px 20px', borderRadius: 10, border: '1px solid var(--dex-gray-300,#ccc)', background: '#fff', color: 'var(--dex-gray-700,#444)', fontWeight: 600, cursor: 'pointer' }}
        >
          {isDe ? 'Zurück' : 'Back'}
        </button>
      </div>
    </div>
  );
};

export default SelfCheckInDisplayPage;
