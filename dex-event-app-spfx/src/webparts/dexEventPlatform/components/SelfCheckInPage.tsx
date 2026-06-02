import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { useEvents, SelfCheckInResult } from '../context/EventContext';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';

/**
 * v18.33 — Self-Check-in-Ergebnisseite.
 *
 * Wird über den Deep-Link ?action=selfcheckin&… erreicht (per Handy-Kamera
 * gescannter QR-Code). Führt den Check-in einmalig aus und zeigt ein klares,
 * vollflächiges Ergebnis. Kein In-App-Scanner nötig — der native Kamera-Scan
 * umgeht die fehlende Kamera-Freigabe der SharePoint-Mobile-App.
 */
const SelfCheckInPage: React.FC = () => {
  const { selfCheckIn } = useEvents();
  const { navigate } = useNavigation();
  const { locale } = useLanguage();
  const isDe = locale === 'de';

  const [result, setResult] = React.useState<SelfCheckInResult | null>(null);
  const [loading, setLoading] = React.useState(true);
  const didRun = React.useRef(false);

  React.useEffect(() => {
    if (didRun.current) return;
    didRun.current = true;
    (async () => {
      try {
        const params = new URLSearchParams(window.location.search);
        const token = params.get('token') || undefined;
        const eventParam = params.get('event');
        const code = params.get('code') || undefined;
        const t = params.get('t');
        const res = await selfCheckIn({
          token,
          eventNumber: eventParam ? parseInt(eventParam, 10) : undefined,
          code,
          windowIndex: t ? parseInt(t, 10) : undefined,
        });
        setResult(res);
      } catch {
        setResult({ status: 'error' });
      } finally {
        setLoading(false);
      }
    })();
  }, [selfCheckIn]);

  const fmtDateTime = (iso?: string): string => {
    if (!iso) return '';
    try {
      return new Date(iso).toLocaleString(isDe ? 'de-DE' : 'en-GB', {
        day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit',
      });
    } catch { return iso; }
  };

  // Mapping Status → Darstellung
  const view = React.useMemo(() => {
    const s = result?.status;
    const title = result?.eventTitle ? `„${result.eventTitle}“` : (isDe ? 'das Event' : 'the event');
    switch (s) {
      case 'success':
        return {
          color: '#86bc25', icon: 'SkypeCircleCheck',
          head: isDe ? 'Eingecheckt!' : 'Checked in!',
          body: isDe
            ? `Du bist erfolgreich für ${title} eingecheckt. Viel Spaß auf der Veranstaltung!`
            : `You are successfully checked in for ${title}. Enjoy the event!`,
        };
      case 'already':
        return {
          color: '#86bc25', icon: 'CompletedSolid',
          head: isDe ? 'Bereits eingecheckt' : 'Already checked in',
          body: isDe
            ? `Du warst für ${title} bereits eingecheckt — alles in Ordnung, du musst nichts weiter tun.`
            : `You were already checked in for ${title} — all good, nothing more to do.`,
        };
      case 'not-registered':
        return {
          color: '#d83b01', icon: 'BlockedSiteSolid12',
          head: isDe ? 'Keine Anmeldung gefunden' : 'No registration found',
          body: isDe
            ? `Für ${title} ist unter deinem Konto keine aktive Anmeldung hinterlegt. Bitte melde dich zuerst an oder wende dich an den Veranstalter.`
            : `There is no active registration for ${title} under your account. Please register first or contact the organizer.`,
        };
      case 'on-waitlist':
        return {
          color: '#d29200', icon: 'Clock',
          head: isDe ? 'Du stehst auf der Warteliste' : 'You are on the waitlist',
          body: isDe
            ? `Für ${title} stehst du aktuell auf der Warteliste. Ein Check-in ist erst möglich, sobald du einen festen Platz bekommen hast.`
            : `You are currently on the waitlist for ${title}. Check-in is only possible once you have a confirmed seat.`,
        };
      case 'closed':
        return {
          color: '#d29200', icon: 'Clock',
          head: isDe ? 'Check-in noch nicht offen' : 'Check-in not open yet',
          body: isDe
            ? `Der Self-Check-in für ${title} ist aktuell nicht möglich.${result?.opensAt ? ` Er ist ab ${fmtDateTime(result.opensAt)} möglich.` : ''}`
            : `Self-check-in for ${title} is currently not available.${result?.opensAt ? ` It opens at ${fmtDateTime(result.opensAt)}.` : ''}`,
        };
      case 'expired':
        return {
          color: '#d29200', icon: 'Refresh',
          head: isDe ? 'Code abgelaufen' : 'Code expired',
          body: isDe
            ? 'Dieser QR-Code ist nicht mehr aktuell. Bitte scanne den aktuell angezeigten Code am Eingang erneut.'
            : 'This QR code is no longer current. Please scan the code currently shown at the entrance again.',
        };
      case 'disabled':
        return {
          color: '#d83b01', icon: 'Blocked2Solid',
          head: isDe ? 'Self-Check-in nicht aktiv' : 'Self-check-in not active',
          body: isDe
            ? `Für ${title} ist der Self-Check-in nicht aktiviert. Bitte wende dich an das Check-in-Team vor Ort.`
            : `Self-check-in is not enabled for ${title}. Please contact the on-site check-in team.`,
        };
      case 'not-found':
        return {
          color: '#d83b01', icon: 'StatusErrorFull',
          head: isDe ? 'Event nicht gefunden' : 'Event not found',
          body: isDe
            ? 'Zu diesem QR-Code konnte kein Event gefunden werden. Möglicherweise ist der Link veraltet.'
            : 'No event could be found for this QR code. The link may be outdated.',
        };
      default:
        return {
          color: '#d83b01', icon: 'StatusErrorFull',
          head: isDe ? 'Etwas ist schiefgelaufen' : 'Something went wrong',
          body: isDe
            ? 'Der Check-in konnte nicht abgeschlossen werden. Bitte versuche es erneut oder wende dich an das Check-in-Team.'
            : 'Check-in could not be completed. Please try again or contact the check-in team.',
        };
    }
  }, [result, isDe]);

  return (
    <div style={{
      minHeight: '70vh', display: 'flex', alignItems: 'center', justifyContent: 'center',
      padding: 24,
    }}>
      <div style={{
        background: '#fff', borderRadius: 20, maxWidth: 480, width: '100%',
        boxShadow: '0 6px 24px rgba(0,0,0,0.08), 0 1px 3px rgba(0,0,0,0.05)',
        padding: '40px 32px', textAlign: 'center',
      }}>
        {loading ? (
          <>
            <div style={{
              width: 56, height: 56, borderRadius: '50%', margin: '0 auto 22px',
              border: '4px solid var(--dex-gray-200, #e5e5e5)',
              borderTopColor: 'var(--dex-green, #86bc25)',
              animation: 'dexOrbSpin 0.8s linear infinite',
            }} />
            <div style={{ fontSize: 17, fontWeight: 600, color: 'var(--dex-gray-800, #333)' }}>
              {isDe ? 'Check-in läuft…' : 'Checking you in…'}
            </div>
            <div style={{ fontSize: 13, color: 'var(--dex-gray-500, #888)', marginTop: 6 }}>
              {isDe ? 'Einen Moment bitte.' : 'One moment please.'}
            </div>
          </>
        ) : (
          <>
            <Icon iconName={view.icon} style={{ fontSize: 64, color: view.color }} />
            <h1 style={{ fontSize: 24, fontWeight: 700, color: 'var(--dex-gray-800, #333)', margin: '18px 0 10px' }}>
              {view.head}
            </h1>
            <p style={{ fontSize: 15, color: 'var(--dex-gray-600, #555)', lineHeight: 1.5, margin: 0 }}>
              {view.body}
            </p>
            <div style={{ display: 'flex', gap: 12, justifyContent: 'center', marginTop: 28, flexWrap: 'wrap' }}>
              <button
                type="button"
                className="btn"
                onClick={() => navigate('my-events')}
                style={{
                  padding: '11px 22px', borderRadius: 10, border: '1px solid var(--dex-green, #86bc25)',
                  background: 'var(--dex-green, #86bc25)', color: '#fff', fontWeight: 600, cursor: 'pointer',
                }}
              >
                {isDe ? 'Meine Events' : 'My events'}
              </button>
              <button
                type="button"
                onClick={() => navigate('start')}
                style={{
                  padding: '11px 22px', borderRadius: 10, border: '1px solid var(--dex-gray-300, #ccc)',
                  background: '#fff', color: 'var(--dex-gray-700, #444)', fontWeight: 600, cursor: 'pointer',
                }}
              >
                {isDe ? 'Zur Startseite' : 'Go to start'}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

export default SelfCheckInPage;
