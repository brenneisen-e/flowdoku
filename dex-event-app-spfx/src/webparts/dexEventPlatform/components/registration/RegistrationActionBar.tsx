/* Fuss der Anmeldeseite — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Die Aktionsleiste (Platz-Badge, Anmelden/Absagen, v24.57) und darunter der
 * Datenschutz-Hinweis. Inhalt zeichengleich uebernommen. */
import * as React from 'react';
import { DeloitteEvent } from '../../types';
import { Icon } from '@fluentui/react/lib/Icon';
import { Send, X } from '../Icons';
import { Locale } from '../../context/LanguageContext';

/** Platz-Badge und die Aktions-Buttons unter dem Formular (v24.57). */
export interface RegistrationActionBarProps {
  childEvents: DeloitteEvent[];
  childOneDe: string;
  childTermPlural: string;
  childTermSingular: string;
  email: string;
  event: DeloitteEvent;
  handleDecline: () => Promise<void>;
  handleSubmit: () => Promise<void>;
  isDeclining: boolean;
  isSubmitting: boolean;
  isTeamMode: boolean;
  liveStats: { active: number; waitlist: number; };
  locale: Locale;
  nothingToSubmit: boolean;
  otherConsentConfirmed: boolean;
  parentAlreadyRegistered: boolean;
  pendingJoinTeam: { teamId: string; teamName: string; };
  registerForOther: boolean;
  resolveMainEventLabel: (defaultLabel: string) => string | null;
  selectedSessions: Set<string>;
  t: (key: string) => string;
  teamMembersParsed: { displayName: string; email: string; }[];
  teamValidation: { ok: boolean; reason?: string; };
  thirdPartyCheck: { alreadyRegistered: boolean; notInAudience: boolean; registeredName?: string; registeredDate?: string; };
  willRegisterParent: boolean;
}
export const RegistrationActionBar: React.FC<RegistrationActionBarProps> = (p) => {
  const { childEvents, childOneDe, childTermPlural, childTermSingular, email, event, handleDecline, handleSubmit, isDeclining, isSubmitting, isTeamMode, liveStats, locale, nothingToSubmit, otherConsentConfirmed, parentAlreadyRegistered, pendingJoinTeam, registerForOther, resolveMainEventLabel, selectedSessions, t, teamMembersParsed, teamValidation, thirdPartyCheck, willRegisterParent } = p;
  return (
      <div style={{ maxWidth: 1100, margin: '24px auto 0', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
        {/* v24.57: Badge mit Icon — freie Plätze ODER (wenn voll + Warteliste
            aktiv) „Warteliste". Bei unbegrenzter Teilnehmerzahl gar nichts. */}
        {event.maxParticipants > 0 && (() => {
          // v24.72: Wartelisten-Anzahl von den freien Plätzen abziehen. Ein frei
          // gewordener Platz geht IMMER zuerst an die Warteliste — er ist also
          // nicht „frei" für neue Anmeldungen. Das verhindert auch das kurze,
          // fälschliche „1 freier Platz" während des Nachrückens.
          // v24.73: Aktiv-/Warteliste-Zahl bevorzugt aus dem für alle lesbaren
          // Counter (liveStats) — sonst sieht ein normaler Teilnehmer wegen der
          // Item-Level-Security der Teilnehmerliste keine korrekte Zahl.
          const effActive = liveStats ? liveStats.active : (event.currentParticipants || 0);
          const effWaitlist = (liveStats && liveStats.waitlist >= 0) ? liveStats.waitlist : (event.waitlistCount || 0);
          const free = Math.max(0, event.maxParticipants - effActive - effWaitlist);
          const isFull = free <= 0;
          // v27.11: voll + Warteliste deaktiviert → NICHT mehr stumm bleiben,
          // sondern rote Badge zeigen (vorher: return null; die Anmeldung lief
          // dann trotzdem still auf die abgeschaltete Warteliste).
          if (isFull && !event.waitlistEnabled) {
            return (
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 7, background: 'rgba(204,0,0,0.08)', color: 'var(--dex-red, #c00)', border: '1px solid var(--dex-red, #c00)', borderRadius: 999, padding: '5px 14px', fontSize: '0.82rem', fontWeight: 700 }}>
                <Icon iconName="People" style={{ fontSize: 15 }} />
                {locale === 'de' ? 'Alle Plätze belegt — keine Warteliste' : 'All seats taken — no waitlist'}
              </span>
            );
          }
          const waitlist = isFull && !!event.waitlistEnabled;
          const nearlyFull = !isFull && free <= Math.max(1, Math.round(event.maxParticipants * 0.1));
          const isTeamEvent = !!(event.teamRegistrationEnabled && event.teamSize && event.teamSize > 1);
          const teamsFree = isTeamEvent ? Math.floor(free / (event.teamSize || 1)) : 0;
          const orange = waitlist || nearlyFull;
          const bg = orange ? 'rgba(237,139,0,0.12)' : 'rgba(134,188,37,0.14)';
          const fg = orange ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)';
          const bd = orange ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)';
          return (
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 7, background: bg, color: fg, border: `1px solid ${bd}`, borderRadius: 999, padding: '5px 14px', fontSize: '0.82rem', fontWeight: 700 }}>
              <Icon iconName={waitlist ? 'Clock' : 'People'} style={{ fontSize: 15 }} />
              {waitlist
                ? (() => {
                    const wc = effWaitlist;
                    return locale === 'de'
                      ? `Alle Plätze belegt | Warteliste aktuell ${wc} ${wc === 1 ? 'Person' : 'Personen'}`
                      : `All places taken | Waitlist currently ${wc} ${wc === 1 ? 'person' : 'people'}`;
                  })()
                : (
                  <span>
                    {free} / {event.maxParticipants} {locale === 'de' ? 'freie Plätze' : 'available'}
                    {isTeamEvent && <span style={{ fontWeight: 500, marginLeft: 6 }}>({teamsFree} {teamsFree === 1 ? (locale === 'de' ? 'Team' : 'team') : (locale === 'de' ? 'Teams' : 'teams')} {locale === 'de' ? 'frei' : 'available'})</span>}
                  </span>
                )}
            </span>
          );
        })()}
        <div className="registration-actions" style={{ alignItems: 'center' }}>
        {(() => {
          // v15.11: im subEventsOnlyMode (Hauptevent nicht anmeldbar) muss
          // mindestens ein Sub-Event ausgewählt sein, sonst Button ausgrauen
          // + Hinweis statt „Registrieren (Haupt-Event)" zeigen.
          const isSubOnly = !!(event && event.subEventsOnlyMode) && !registerForOther;
          const nothingPicked = isSubOnly && selectedSessions.size === 0;
          // v15.16: Consent-Pflicht bei „Für andere registrieren".
          const needsOtherConsent = registerForOther && !!email.trim() && !otherConsentConfirmed;
          // v19.8: Bei stellvertretender Anmeldung den Button sperren, wenn die
          // ausgewählte Person bereits angemeldet ist — vorher konnte man
          // trotz Hinweis auf „Registrieren" klicken (und es kam danach noch
          // die CC-Frage). Jetzt klare Blockade direkt am Button.
          const targetAlreadyRegistered = registerForOther && !!(thirdPartyCheck && thirdPartyCheck.alreadyRegistered);
          // v18: Demo-Event — Register-Button ist bewusst NICHT auswählbar
          // (keine echte Anmeldung; reine Showcase-Ansicht).
          const isDemo = !!(event && event.isDemoShowcase);
          // v28.88: Bereits angemeldet und nichts (mehr) auszuwählen → der
          // Klick konnte ohnehin nichts bewirken und endete in einer
          // Fehlermeldung. Jetzt sagt der Button selbst, dass die Anmeldung
          // schon steht. (nothingToSubmit deckt auch Abwahl-Änderungen ab —
          // wer Sub-Events abmeldet, kommt weiterhin durch.)
          const alreadyDone = parentAlreadyRegistered && nothingToSubmit;
          const isDisabled = isDemo || isSubmitting || (isTeamMode && !teamValidation.ok) || nothingPicked || needsOtherConsent || targetAlreadyRegistered || alreadyDone;
          const titleAttr = isDemo
            ? (locale === 'de' ? 'Demo-Event — eine echte Anmeldung ist nicht möglich.' : 'Demo event — real registration is not possible.')
            : (alreadyDone
            ? (locale === 'de' ? 'Du bist für dieses Event bereits angemeldet. Abmelden kannst du dich über „Meine Events".' : 'You are already registered for this event. You can cancel via „My events".')
            : (targetAlreadyRegistered
            ? (locale === 'de' ? 'Diese Person ist bereits für das Event angemeldet.' : 'This person is already registered for this event.')
            : (isTeamMode && !teamValidation.ok
            ? (teamValidation.reason || '')
            : (nothingPicked
                ? (locale === 'de'
                    ? `Bitte mindestens ${childOneDe} auswählen.`
                    : `Please pick at least one ${childTermSingular || 'sub-event'}.`)
                : (needsOtherConsent
                    ? (locale === 'de'
                        ? 'Bitte bestätige die Zustimmung der Person.'
                        : 'Please confirm the person\'s consent.')
                    : '')))));
          return (
            <button
              className="btn btn-primary"
              onClick={handleSubmit}
              disabled={isDisabled}
              title={titleAttr}
            >
              {/* v24.94: Label in EINEN Span wickeln. Sonst werden „Register" und
                  das „(Warteliste)"-Suffix-Span zu separaten Flex-Items des
                  Buttons (.btn ist inline-flex mit gap:8px) → der Flex-Gap PLUS
                  das Leerzeichen ergaben einen doppelten Abstand. */}
              <Send size={16} /> <span>{(() => {
                if (isSubmitting) return t('reg.submitting');
                // v28.88: Bestehende Anmeldung, nichts zu ändern — der Button
                // sagt das jetzt selbst, statt „Registrieren" anzubieten und
                // beim Klick zu meckern.
                if (alreadyDone) return locale === 'de' ? 'Bereits angemeldet' : 'Already registered';
                // v24.62: Wenn das Hauptevent voll ist und eine Warteliste hat,
                // landet die Anmeldung auf der Warteliste — im Button steht das als
                // kurzer, NICHT fetter Zusatz „(Warteliste)" (die aktuelle Anzahl
                // steht im Badge über dem Button).
                const mfActive = liveStats ? liveStats.active : (event.currentParticipants || 0);
                const mfWaitlist = (liveStats && liveStats.waitlist >= 0) ? liveStats.waitlist : (event.waitlistCount || 0);
                const mainFull = event.maxParticipants > 0
                  && Math.max(0, event.maxParticipants - mfActive - mfWaitlist) <= 0
                  && !!event.waitlistEnabled;
                const waitlistSuffixNode: React.ReactNode = mainFull
                  ? <span style={{ fontWeight: 400 }}> ({locale === 'de' ? 'Warteliste' : 'waitlist'})</span>
                  : null;
                // v18.73: Vorgemerkter Team-Beitritt — eigener Button-Text.
                if (pendingJoinTeam) {
                  return event?.teamJoinRequiresApproval
                    ? (locale === 'de' ? 'Beitritt anfragen' : 'Request to join')
                    : (locale === 'de' ? 'Team beitreten & anmelden' : 'Join team & register');
                }
                // v11.82: Team-Modus — eigener Button-Text mit Personen-Zahl.
                if (isTeamMode) {
                  const n = 1 + teamMembersParsed.filter(Boolean).length;
                  return locale === 'de'
                    ? `Team anmelden (${n} ${n === 1 ? 'Person' : 'Personen'})`
                    : `Register team (${n} ${n === 1 ? 'person' : 'people'})`;
                }
                if (nothingPicked) {
                  return locale === 'de'
                    ? `Bitte mindestens ${childOneDe} auswählen`
                    : `Please pick at least one ${childTermSingular || 'sub-event'}`;
                }
                if (registerForOther) return <>{t('reg.register')}{waitlistSuffixNode}</>;
                // v7.3: Kein Selection-Block → einfacher "Registrieren"-Text ohne
                // Parantheses-Info. Erst wenn Sub-Events existieren, zeigen wir
                // detailliert an, was gerade submittet wird.
                if (childEvents.length === 0) return <>{t('reg.register')}{waitlistSuffixNode}</>;
                const parts: string[] = [];
                if (willRegisterParent) parts.push(resolveMainEventLabel(t('reg.selection.mainevent') || 'Haupt-Event') || event.title);
                if (selectedSessions.size > 0) {
                  parts.push(`${selectedSessions.size} ${selectedSessions.size === 1 ? (childTermSingular || t('reg.selection.sessioncount.one') || 'Session') : (childTermPlural || t('reg.selection.sessioncount.many') || 'Sessions')}`);
                }
                if (parts.length === 0) return <>{t('reg.register')}{waitlistSuffixNode}</>;
                // Bei gleichzeitiger Hauptevent-Anmeldung den Warteliste-Hinweis anhängen.
                return <>{t('reg.register')} ({parts.join(' + ')}){willRegisterParent ? waitlistSuffixNode : null}</>;
              })()}</span>
            </button>
          );
        })()}
        {/* v18.11: „Ich nehme nicht teil" — proaktive Absage. Nur bei
            Selbst-Anmeldung (nicht „für andere", nicht Team-Modus, kein
            Demo-Event). Braucht keine Pflichtfelder. */}
        {!registerForOther && !isTeamMode && !pendingJoinTeam && !(event && event.isDemoShowcase) && (
          <button
            type="button"
            className="btn btn-secondary"
            onClick={handleDecline}
            disabled={isDeclining || isSubmitting}
            title={locale === 'de'
              ? (childEvents.length > 0
                ? `Melde zurück, dass du nicht teilnehmen wirst — gilt für das gesamte Event inklusive aller ${childTermPlural || 'Sub-Events'}. Eine Auswahl ist dafür nicht nötig.`
                : 'Melde zurück, dass du nicht teilnehmen wirst (keine Anmeldung).')
              : (childEvents.length > 0
                ? 'Let us know you will not attend — applies to the whole event including all sub-events. No selection needed.'
                : 'Let us know you will not attend (no registration).')}
            style={{ color: 'var(--dex-gray-700, #444)' }}
          >
            <X size={16} /> {isDeclining
              ? (locale === 'de' ? 'Wird gesendet…' : 'Submitting…')
              : (locale === 'de' ? 'Ich nehme nicht teil' : 'I will not attend')}
          </button>
        )}
        </div>
      </div>
  );
};

/** Datenschutz-Hinweis als Fussnote ganz unten. */
export interface PrivacyNoteProps {
  event: DeloitteEvent;
  t: (key: string) => string;
}
export const PrivacyNote: React.FC<PrivacyNoteProps> = (p) => {
  const { event, t } = p;
  return (
      <div
        className="footer-disclaimer mt-24"
        style={{ borderRadius: 'var(--dex-radius-lg)', maxWidth: 1100, margin: '24px auto 0' }}
      >
        <p>
          {/* Datenverarbeitungs-Einwilligung. „{link}" wird als Anchor auf die
              Deloitte-Datenschutzhinweise gerendert; der Rest ist reiner Text
              mit {title}-Ersetzung. */}
          {(() => {
            const raw = t('reg.privacy.data').replace('{title}', event.title);
            const parts = raw.split('{link}');
            const linkLabel = t('reg.privacy.data.link');
            return (
              <>
                {parts[0]}
                <a
                  href="https://www.deloitte.com/de/de/legal/privacy.html"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  {linkLabel}
                </a>
                {parts[1] || ''}
              </>
            );
          })()}
        </p>
        <p>
          {/* Bild-/Video-Einwilligung. privacy@deloitte.de wird als
              mailto-Link gerendert (gleiche Adresse in DE und EN). */}
          {(() => {
            const raw = t('reg.privacy').replace('{title}', event.title);
            const mail = 'privacy@deloitte.de';
            const parts = raw.split(mail);
            return (
              <>
                {parts[0]}
                {parts.length > 1 && (
                  <>
                    <a href={`mailto:${mail}`}>{mail}</a>
                    {parts.slice(1).join(mail)}
                  </>
                )}
              </>
            );
          })()}
        </p>
      </div>
  );
};
