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
import { sortAgenda, agendaGroups, groupLabel, groupDateLabel } from '../utils/agendaGroups';
import { suggestCurrentAgendaItem } from '../utils/agendaCheckIns';

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

  // v30.95: Programmpunkte-Modus — der Bildschirm zeigt den Live-QR für EINEN
  // Punkt (Raum-Bildschirm); ein Scan erfasst die Anwesenheit dort, nicht den
  // Event-Status. Vorschlag: der laufende Punkt; Wahl bleibt je Event im Gerät.
  const agendaItems = React.useMemo(() => (event && event.agendaCheckIn ? sortAgenda(event.agenda || []) : []), [event]);
  const agendaMode = agendaItems.length > 0;
  const termS = (event && event.agendaTermSingular) || (isDe ? 'Programmpunkt' : 'Agenda item');
  const [pointId, setPointId] = React.useState<string>('');
  React.useEffect(() => {
    if (!agendaMode || !event) { setPointId(''); return; }
    let stored = '';
    try { stored = window.localStorage.getItem(`dex_selfcheckin_point_${event.id}`) || ''; } catch { /* */ }
    if (agendaItems.some(a => a.id === stored)) { setPointId(stored); return; }
    const s = suggestCurrentAgendaItem(agendaItems);
    setPointId(s ? s.id : '');
  }, [agendaMode, agendaItems, event]);
  const choosePoint = (id: string): void => {
    setPointId(id);
    if (event) { try { window.localStorage.setItem(`dex_selfcheckin_point_${event.id}`, id); } catch { /* */ } }
  };
  const point = agendaItems.find(a => a.id === pointId) || null;
  const [pickerOpen, setPickerOpen] = React.useState(false);

  // QR regelmäßig neu berechnen (jede Sekunde Countdown, bei Fensterwechsel neuer Code).
  React.useEffect(() => {
    if (!enabled) return;
    let cancelled = false;
    const refresh = async (): Promise<void> => {
      try {
        const { url, expiresInSeconds } = await buildRotatingCheckInUrl(secret, eventNumber, Date.now(), agendaMode ? (pointId || undefined) : undefined);
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
  }, [enabled, secret, eventNumber, agendaMode, pointId]);

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
        {/* v30.95: Der Punkt, für den dieser Bildschirm steht — groß, damit die
            Person am Raum sieht, was der Scan erfasst. Klick öffnet die Auswahl. */}
        {agendaMode && (
          <div style={{ marginTop: 10 }}>
            <button type="button" onClick={() => setPickerOpen(o => !o)} title={isDe ? `${termS} wechseln` : `Change ${termS.toLowerCase()}`} style={{
              display: 'inline-flex', alignItems: 'center', gap: 10, padding: '8px 16px', borderRadius: 999, cursor: 'pointer',
              border: '1px solid var(--dex-green, #86bc25)', background: 'rgba(134,188,37,0.12)', color: 'var(--dex-gray-800,#333)', fontSize: 'clamp(0.95rem, 1.8vw, 1.25rem)', fontWeight: 600,
            }}>
              <span style={{ fontSize: '0.72em', fontWeight: 700, color: 'var(--dex-green-dark,#4a7c1f)', letterSpacing: 0.4 }}>{termS.toUpperCase()}</span>
              {point ? <span>{point.time ? `${point.time} · ` : ''}{point.title || (isDe ? '(ohne Titel)' : '(untitled)')}{point.location ? <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}> · {point.location}</span> : null}</span> : <span style={{ color: 'var(--dex-red,#da291c)' }}>{isDe ? 'Bitte wählen' : 'Please choose'}</span>}
              <Icon iconName={pickerOpen ? 'ChevronUp' : 'ChevronDown'} style={{ fontSize: 12 }} />
            </button>
            {pickerOpen && (
              <div style={{ marginTop: 10, textAlign: 'left', maxWidth: 720, marginLeft: 'auto', marginRight: 'auto', background: '#fff', border: '1px solid var(--dex-gray-200,#eee)', borderRadius: 12, padding: 12, boxShadow: '0 6px 20px rgba(0,0,0,0.08)', maxHeight: '40vh', overflow: 'auto' }}>
                {agendaGroups(agendaItems).map((g, gi) => (
                  <div key={g.key} style={{ marginBottom: 8 }}>
                    <div style={{ fontSize: '0.74rem', fontWeight: 700, color: 'var(--dex-green-dark,#4a7c1f)', marginBottom: 4 }}>{groupLabel(g, gi, isDe)} <span style={{ fontWeight: 500, color: 'var(--dex-gray-500)' }}>· {groupDateLabel(g, isDe, false)}</span></div>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                      {g.items.map(it => {
                        const active = it.id === pointId;
                        return (
                          <button key={it.id} type="button" onClick={() => { choosePoint(it.id); setPickerOpen(false); }} style={{
                            display: 'inline-flex', alignItems: 'center', gap: 6, padding: '6px 10px', borderRadius: 999, cursor: 'pointer', fontSize: '0.82rem', fontWeight: active ? 700 : 500,
                            border: `1px solid ${active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`, background: active ? 'var(--dex-green, #86bc25)' : '#fff', color: active ? '#fff' : 'var(--dex-gray-800)',
                          }}>
                            <span style={{ fontVariantNumeric: 'tabular-nums', opacity: 0.85 }}>{it.time}</span>
                            <span>{it.title || (isDe ? '(ohne Titel)' : '(untitled)')}</span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </div>

      <div style={{
        background: '#fff', borderRadius: 24, padding: 'clamp(16px, 3vw, 32px)',
        boxShadow: '0 8px 30px rgba(0,0,0,0.10)', border: '1px solid var(--dex-gray-200,#eee)',
      }}>
        {agendaMode && !point ? (
          <div style={{ width: 'min(60vh, 420px)', height: 'min(60vh, 420px)', display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', color: 'var(--dex-gray-500)', padding: 24 }}>
            {isDe ? `Wähle oben den ${termS}, für den dieser Bildschirm steht.` : `Pick the ${termS.toLowerCase()} this screen stands for above.`}
          </div>
        ) : qrDataUrl
          ? <img src={qrDataUrl} alt="Check-in QR" style={{ width: 'min(60vh, 420px)', height: 'min(60vh, 420px)', display: 'block' }} />
          : <div style={{ width: 'min(60vh, 420px)', height: 'min(60vh, 420px)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', border: '4px solid #eee', borderTopColor: '#86bc25', animation: 'dexOrbSpin 0.8s linear infinite' }} />
            </div>}
      </div>

      <div style={{ marginTop: 20, textAlign: 'center' }}>
        <div style={{ fontSize: 18, fontWeight: 600, color: 'var(--dex-gray-800,#333)' }}>
          {isDe ? 'Mit der Handy-Kamera scannen' : 'Scan with your phone camera'}
          {agendaMode && point && (
            <span style={{ display: 'block', fontSize: 14, fontWeight: 500, color: 'var(--dex-gray-600,#555)', marginTop: 4 }}>
              {isDe ? `Erfasst die Anwesenheit bei „${point.title}“ — der Event-Status bleibt unverändert.` : `Records attendance at “${point.title}” — the event status stays unchanged.`}
            </span>
          )}
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
