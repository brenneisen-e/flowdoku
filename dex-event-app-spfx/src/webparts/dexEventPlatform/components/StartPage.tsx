// Start-Seite - Navigation zu Registration, My Events und Admin

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useLanguage } from '../context/LanguageContext';
import { Calendar, Pin, Settings, QrCode } from './Icons';
import InquiryModal from './InquiryModal';

export default function StartPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { canCreateEvents, isAdmin } = useRoles();
  const { events } = useEvents();
  const { currentUser } = useCurrentUser();
  const { t, locale } = useLanguage();
  // v12.5: Organizer-Kachel wird jetzt IMMER gerendert. Wer keine Organizer-
  // Rechte hat sieht sie ausgegraut mit Overlay-Button „Want to become an
  // organizer?" — Klick öffnet das gleiche Inquiry-Modal wie das
  // Bubble-CTA auf der Landing-Page.
  const [showInquiry, setShowInquiry] = React.useState(false);

  // v11.38/v11.46: Co-Organizer pro Event (per-Event-Rolle, ohne globale
  // Organizer-Rolle in DEX_Roles) sehen die Organizer-Kachel ebenfalls —
  // AdminPage gewährt ihnen ohnehin Zugriff auf "ihre" Events (siehe
  // isOrganizerFor dort), aber ohne Kachel im Startmenü gab es bisher
  // keinen Einstieg.
  //
  // v11.46-Fix: seit v9.20 hat der Wizard nur EINEN Organizer-Picker und
  // schreibt alle Personen (Haupt-Organizer wie Co-Organizer) in
  // event.organizerEmails — das alte Feld event.coOrganizerEmails wird
  // garnicht mehr befüllt (siehe EventCreationPage const-State ohne Setter).
  // Der bisherige Check auf nur coOrganizerEmails fand also für alle nach
  // v9.20 angelegten Events nichts. Prüfung jetzt analog zu
  // AdminPage.isOrganizerFor: organizerEmails ODER coOrganizerEmails
  // (Backward-Compat für alte Events).
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isOrganizerOfAnyEvent = !!currentEmailLc && (events || []).some(e => {
    const inOrg = (e.organizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
    if (inOrg) return true;
    return (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
  });
  // v23.36: Im Demo-Modus sieht der Demo-User die Start-Kacheln EXAKT wie ein
  // normaler User — d.h. die Organizer-Kachel ist ausgegraut (mit „Organizer
  // werden?"-Hinweis) und die Check-In-Kachel fehlt. Früher (v18.3) war die
  // Organizer-Kachel im Demo-Modus klickbar (`|| isImpersonating`) — das
  // widersprach dem Zweck „so sieht es ein echter User".
  const isOrganizer = canCreateEvents || isOrganizerOfAnyEvent;
  // v13.12: Check-In-Team-Mitgliedschaft = User ist in qrScannerEmails
  // mindestens eines AKTIVEN Events. Wenn ja: eine eigene 4. Kachel
  // „Check-In" erscheint mit QrCode-Icon + Pulse-Animation. Die
  // Organizer-Kachel bleibt für reine Check-In-Helfer trotzdem ausgegraut
  // (keine Verwirrung über die tatsächlichen Berechtigungen — Check-In ist
  // ein eigener, klar abgegrenzter Bereich).
  const isCheckInTeamOfActive = !!currentEmailLc && (events || []).some(e => {
    if (e.status !== 'Active') return false;
    if ((e.qrScannerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc)) return true;
    if ((e.organizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc)) return true;
    return (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
  });
  // v12.5: Grid bleibt immer 3-spaltig — auch User sehen die Organizer-
  // Kachel (ausgegraut + CTA-Overlay).
  // v13.12 / v14.3: Wenn der User Check-In-Team-Mitgliedschaft hat ODER
  // Admin / Organizer ist, kommt die 4. Check-In-Kachel dazu → 4-spaltiges
  // Grid. Admins und Organizer können ohnehin überall einchecken — sie
  // sollen den Direkteinstieg also auch immer prominent auf der Startseite
  // sehen, nicht erst über den Header-Button suchen müssen.
  const showAdminTile = true;
  const showCheckInTile = isCheckInTeamOfActive || isAdmin || isOrganizer;

  return (
    <div className="page-container">
      {/* v9.36: Keyframes inline injizieren — SPFx hasht sonst die Names im
          .module.scss und die Animation findet sie nicht. */}
      <style>{`
        @keyframes dexStartIconBounce {
          0%   { transform: translateY(0) scale(1); }
          30%  { transform: translateY(-8px) scale(1.1); }
          60%  { transform: translateY(0) scale(1); }
          80%  { transform: translateY(-3px) scale(1.04); }
          100% { transform: translateY(0) scale(1); }
        }
        @keyframes dexStartIconWiggle {
          0%   { transform: rotate(0deg) scale(1); }
          20%  { transform: rotate(-10deg) scale(1.08); }
          40%  { transform: rotate(8deg) scale(1.08); }
          60%  { transform: rotate(-6deg) scale(1.05); }
          80%  { transform: rotate(4deg) scale(1.03); }
          100% { transform: rotate(0deg) scale(1); }
        }
        @keyframes dexStartIconSpin {
          0%   { transform: rotate(0deg) scale(1); }
          100% { transform: rotate(360deg) scale(1); }
        }
        @keyframes dexStartIconScanPulse {
          0%   { transform: scale(1); filter: drop-shadow(0 0 0 rgba(134,188,37,0.0)); }
          40%  { transform: scale(1.12); filter: drop-shadow(0 0 6px rgba(134,188,37,0.55)); }
          70%  { transform: scale(0.96); filter: drop-shadow(0 0 0 rgba(134,188,37,0.0)); }
          100% { transform: scale(1); filter: drop-shadow(0 0 0 rgba(134,188,37,0.0)); }
        }
        .start-card--checkin:hover .start-card__icon svg { animation: dexStartIconScanPulse 0.9s ease; }
        /* v13.14: 4-Spalten-Variante. Grid bekommt mehr Maximalbreite,
           damit alle vier Kacheln gleich breit UND nicht zu schmal sind. */
        .start-grid--with-checkin {
          grid-template-columns: 1fr 1fr 1fr 1fr !important;
          max-width: 1200px !important;
        }
        .start-grid--with-checkin .start-card { width: 100%; }
        @media (max-width: 900px) {
          .start-grid--with-checkin { grid-template-columns: 1fr 1fr !important; }
          .start-grid--with-checkin .start-card--checkin { grid-column: 1 / -1; }
        }
      `}</style>
      <div className={`start-grid${showAdminTile ? ' start-grid--with-admin' : ''}${showCheckInTile ? ' start-grid--with-checkin' : ''}`}>
        {/* v22.21: data-tour-Anker — Spotlight-Ziele des geführten Tutorials. */}
        <div className="card card-clickable start-card" data-tour="tile-register" onClick={() => navigate('register')}>
          <div className="start-card__icon">
            <Calendar size={64} strokeWidth={1} />
          </div>
          <h2>{t('start.register')}</h2>
          <p style={{ whiteSpace: 'nowrap' }}>{t('start.register.desc')}</p>
        </div>
        <div className="card card-clickable start-card" data-tour="tile-myevents" onClick={() => navigate('my-events')}>
          <div className="start-card__icon">
            <Pin size={64} strokeWidth={1} />
          </div>
          <h2>{t('start.myevents')}</h2>
          <p style={{ whiteSpace: 'nowrap' }}>{t('start.myevents.desc')}</p>
        </div>
        {isOrganizer ? (
          <div className="card card-clickable start-card start-card--admin" data-tour="tile-admin" onClick={() => navigate('admin')}>
            <div className="start-card__icon">
              <Settings size={64} strokeWidth={1} />
            </div>
            <h2>{t('start.admin')}</h2>
            <p style={{ whiteSpace: 'nowrap' }}>{t('start.admin.desc')}</p>
          </div>
        ) : (
          // v12.5: Ausgegraute Organizer-Kachel mit CTA-Overlay-Button
          // — User können so direkt anfragen Organizer zu werden. Bleibt
          // ausgegraut auch für reine Check-In-Helfer (v13.12), die haben
          // ja keine Organizer-Rechte — Check-In ist eine eigene Kachel.
          <div
            className="card start-card start-card--admin"
            style={{ position: 'relative', cursor: 'default', opacity: 0.55 }}
          >
            <div className="start-card__icon">
              <Settings size={64} strokeWidth={1} />
            </div>
            <h2>{t('start.admin')}</h2>
            <p style={{ whiteSpace: 'nowrap' }}>{t('start.admin.desc')}</p>
            <div
              style={{
                position: 'absolute', inset: 0,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                background: 'rgba(255,255,255,0.0)',
                opacity: 1,
              }}
            >
              <button
                type="button"
                onClick={(e) => { e.stopPropagation(); setShowInquiry(true); }}
                style={{
                  background: 'var(--dex-green, #86bc25)', color: '#fff',
                  border: 'none', borderRadius: 999,
                  padding: '10px 18px', fontWeight: 700,
                  cursor: 'pointer', boxShadow: '0 4px 14px rgba(0,0,0,0.18)',
                  fontSize: '0.85rem', whiteSpace: 'nowrap',
                }}
              >
                {locale === 'de' ? 'Organizer werden?' : 'Want to become an organizer?'}
              </button>
            </div>
          </div>
        )}
        {/* v13.12: Vierte Kachel „Check-In" — nur sichtbar wenn der User
            in mindestens einem aktiven Event als QR-Scanner / Organizer /
            Co-Organizer eingetragen ist. Eigenes QR-Code-Icon und eigene
            Scan-Pulse-Animation, klar abgegrenzt von der Admin-Kachel. */}
        {showCheckInTile && (
          <div
            className="card card-clickable start-card start-card--checkin"
            data-tour="tile-checkin"
            onClick={() => navigate('check-in')}
          >
            <div className="start-card__icon">
              <QrCode size={64} strokeWidth={1} />
            </div>
            <h2>{locale === 'de' ? 'Check-In' : 'Check-in'}</h2>
            <p style={{ whiteSpace: 'nowrap' }}>
              {locale === 'de'
                ? 'Teilnehmer einchecken'
                : 'Check in attendees'}
            </p>
          </div>
        )}
      </div>
      <InquiryModal open={showInquiry} onClose={() => setShowInquiry(false)} />
    </div>
  );
}
