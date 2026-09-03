/* Banner und Overlay der Anmeldeseite — aus RegistrationPage.tsx ausgelagert
 * (v30.66). Vier kurze, voneinander unabhaengige Bloecke ueber dem Formular:
 * Standort-Hinweis, Fristen-Hinweis fuer Organizer, das Submit-Overlay und der
 * Demo-Hinweis. Inhalt zeichengleich uebernommen; die Anzeige-Bedingungen sind
 * beim Aufrufer geblieben. */
import * as React from 'react';
import { Locale } from '../../context/LanguageContext';
import { DeloitteEvent } from '../../types';

/** Hinweis, dass das Event fuer den eigenen Standort nicht ausgeschrieben ist. */
export interface LocationBannerProps {
  currentUser: import("../../types/index").User;
  event: DeloitteEvent;
  t: (key: string) => string;
}
export const LocationBanner: React.FC<LocationBannerProps> = (p) => {
  const { currentUser, event, t } = p;
  return (
        <div style={{
          padding: '10px 16px', marginBottom: 16, borderRadius: 'var(--dex-radius-md)',
          background: 'rgba(237,139,0,0.1)', border: '1px solid var(--dex-orange)',
          color: 'var(--dex-orange)', fontSize: '0.85rem',
        }}>
          {t('reg.locationnotice')}
          {event && event.locationAudience.length > 0 && <> {t('reg.locationfilter')}: <strong>{event.locationAudience.join(', ')}</strong>.</>}
          {/* v9.17: bei Einzel-E-Mail-Whitelists in audienceFilter würden bei
              größeren Verteilern (50+ Adressen) der Banner zugekleistert.
              Statt alle Mails auflisten: nur die Anzahl + die ersten 3
              Adressen zeigen, der Rest als "+N weitere". Gruppen-/Group-Namen
              (ohne "@") werden weiterhin alle aufgeführt — die sind kurz. */}
          {event && event.audienceFilter && event.audienceFilter.length > 0 && (() => {
            const items = event.audienceFilter;
            const emails = items.filter(s => s.includes('@'));
            const groups = items.filter(s => !s.includes('@'));
            const showLabel = (() => {
              if (emails.length === 0) return groups.join(', ');
              if (emails.length <= 3) return [...groups, ...emails].join(', ');
              const head = emails.slice(0, 3).join(', ');
              const more = emails.length - 3;
              const tail = `${head} (+${more} ${t('reg.audience.more') || 'weitere E-Mail-Adressen'})`;
              return groups.length > 0 ? `${groups.join(', ')}, ${tail}` : tail;
            })();
            return <> {t('reg.audience')}: <strong>{showLabel}</strong>.</>;
          })()}
          {event && event.filterMode === 'AND' && <> ({t('reg.andmode')})</>}
          {' '}{t('reg.yourlocation')}: {currentUser.location || t('reg.unknown')}.
        </div>
  );
};

/** Fristen-Hinweis, den nur Organizer und Admins sehen. */
export interface DeadlineBannerProps {
  event: DeloitteEvent;
  isFullyClosed: boolean;
  locale: Locale;
  t: (key: string) => string;
}
export const DeadlineBanner: React.FC<DeadlineBannerProps> = (p) => {
  const { event, isFullyClosed, locale, t } = p;
  return (
        <div style={{
          padding: '10px 16px', marginBottom: 16, borderRadius: 'var(--dex-radius-md)',
          background: 'rgba(237,139,0,0.1)', border: '1px solid var(--dex-orange)',
          color: 'var(--dex-orange)', fontSize: '0.85rem',
        }}>
          {/* v22.55: Nur wenn ALLES zu ist ("kein User käme mehr rein") den
              harten Hinweis zeigen. Sind Sub-Events noch offen, kann sich ein
              normaler User weiterhin für diese anmelden — dann ein zutreffender
              Hinweis statt der irreführenden "keine Anmeldung mehr"-Meldung. */}
          {isFullyClosed ? (
            <>
              {t('reg.deadlinepassed.adminnotice')}
              {event && (event.klammerDeadline || event.registrationDeadline) && (
                <> {t('reg.deadlinepassed.date')}: <strong>{new Date(event.klammerDeadline || event.registrationDeadline).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}</strong>.</>
              )}
            </>
          ) : (
            locale === 'de' ? (
              <>Hinweis: Die Anmeldefrist des Hauptevents ist abgelaufen{event.registrationDeadline ? <> (war <strong>{new Date(event.registrationDeadline).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}</strong>)</> : ''} — die noch offenen Sub-Events sind aber weiterhin buchbar, auch für reguläre User.</>
            ) : (
              <>Note: The main event’s registration deadline has passed{event.registrationDeadline ? <> (was <strong>{new Date(event.registrationDeadline).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}</strong>)</> : ''} — but the still-open sub-events remain bookable, also for regular users.</>
            )
          )}
        </div>
  );
};

/** Submit-Overlay mit Spinner, Prozentanzeige und Live-Label (v11.33). */
export interface SubmitOverlayProps {
  displayProgress: number;
  locale: Locale;
  submitProgressLabel: string;
  /** v30.73: Überschrift — der Wizard nutzt dasselbe Overlay fürs Speichern. */
  title?: string;
}
export const SubmitOverlay: React.FC<SubmitOverlayProps> = (p) => {
  const { displayProgress, locale, submitProgressLabel } = p;
  const title = p.title || (locale === 'de' ? 'Anmeldung läuft …' : 'Submitting registration …');
  return (
        <div
          role="dialog"
          aria-modal="true"
          aria-label={title}
          style={{
            position: 'fixed', inset: 0, zIndex: 2000,
            background: 'rgba(0,0,0,0.55)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            padding: 16,
          }}
        >
          <div style={{
            background: '#fff', borderRadius: 12, padding: '28px 32px',
            maxWidth: 460, width: '100%',
            boxShadow: '0 16px 48px rgba(0,0,0,0.35)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16,
          }}>
            {/* v15.13: Doppel-Ladebalken entfernt — die deterministische
                Progress-Bar weiter unten (0-100% mit Label) reicht aus.
                Die zusätzliche indeterministische „Pulse"-Bar war für den
                User verwirrend. */}
            <div style={{ fontWeight: 700, fontSize: '1rem', color: 'var(--dex-gray-800)' }}>
              {title}
            </div>
            <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', textAlign: 'center', minHeight: 18 }}>
              {submitProgressLabel}
            </div>
            <div style={{ width: '100%', height: 8, borderRadius: 4, background: 'var(--dex-gray-200)', overflow: 'hidden' }}>
              <div style={{
                width: `${Math.min(100, Math.max(0, displayProgress))}%`,
                height: '100%',
                background: 'var(--dex-green, #86bc25)',
                // v29.29: kurze Übergangszeit — die Anzeige wird alle 60 ms
                // nachgezogen, eine lange Transition würde hinterherhinken.
                transition: 'width 0.15s linear',
              }} />
            </div>
            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontVariantNumeric: 'tabular-nums' }}>
              {displayProgress}%
            </div>
            {/* v30.19: deutlicher, pulsierender Warnhinweis — der Hintergrund
                ist durch das Overlay ohnehin nicht klickbar, aber „bitte
                warten" muss man SEHEN. Zusätzlich warnt ein
                beforeunload-Guard (s. Hook oben) vor dem Schließen. */}
            <div style={{
              display: 'flex', alignItems: 'center', gap: 8,
              padding: '8px 16px', borderRadius: 999,
              background: 'rgba(237,139,0,0.12)', border: '1px solid var(--dex-orange, #ed8b00)',
              color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 700, fontSize: '0.85rem',
              textAlign: 'center',
              animation: 'dexWaitPulse 1.5s ease-in-out infinite',
            }}>
              {locale === 'de'
                ? 'Bitte warten — Fenster nicht schließen'
                : 'Please wait — do not close this window'}
            </div>
          </div>
          <style>{`@keyframes dexProgressSlide { 0% { left: -40%; } 100% { left: 100%; } }
@keyframes dexWaitPulse { 0%, 100% { opacity: 1; transform: scale(1); } 50% { opacity: 0.55; transform: scale(1.04); } }`}</style>
        </div>
  );
};

/** Demo-Hinweis-Banner (v18). */
export interface DemoBannerProps {
  locale: Locale;
}
export const DemoBanner: React.FC<DemoBannerProps> = (p) => {
  const { locale } = p;
  return (
        <div style={{
          padding: '10px 16px', marginBottom: 16, borderRadius: 'var(--dex-radius, 12px)',
          background: 'rgba(0,118,168,0.08)', border: '1px solid var(--dex-blue, #0076a8)',
          color: 'var(--dex-gray-800)', fontSize: '0.85rem', lineHeight: 1.55,
          display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap',
        }}>
          <span style={{
            padding: '2px 8px', borderRadius: 999, background: 'var(--dex-blue, #0076a8)',
            color: '#fff', fontSize: '0.66rem', fontWeight: 700, letterSpacing: 1,
          }}>DEMO</span>
          {locale === 'de'
            ? <span>Dies ist ein <strong>Demo-Event</strong> — es wird genau so angezeigt wie ein echtes Event. Du kannst die Anmeldemaske ansehen, aber <strong>keine echte Anmeldung</strong> absenden.</span>
            : <span>This is a <strong>demo event</strong> — shown exactly like a real one. You can explore the registration form, but <strong>cannot submit a real registration</strong>.</span>}
        </div>
  );
};
