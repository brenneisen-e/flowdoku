// Start-Seite - Navigation zu Registration, My Events und Admin

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useLanguage } from '../context/LanguageContext';
import { Calendar, Pin, Settings } from './Icons';
import InquiryModal from './InquiryModal';

export default function StartPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { canCreateEvents } = useRoles();
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
  // AdminPage gewaehrt ihnen ohnehin Zugriff auf "ihre" Events (siehe
  // isOrganizerFor dort), aber ohne Kachel im Startmenue gab es bisher
  // keinen Einstieg.
  //
  // v11.46-Fix: seit v9.20 hat der Wizard nur EINEN Organizer-Picker und
  // schreibt alle Personen (Haupt-Organizer wie Co-Organizer) in
  // event.organizerEmails — das alte Feld event.coOrganizerEmails wird
  // garnicht mehr befuellt (siehe EventCreationPage const-State ohne Setter).
  // Der bisherige Check auf nur coOrganizerEmails fand also fuer alle nach
  // v9.20 angelegten Events nichts. Pruefung jetzt analog zu
  // AdminPage.isOrganizerFor: organizerEmails ODER coOrganizerEmails
  // (Backward-Compat fuer alte Events).
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isOrganizerOfAnyEvent = !!currentEmailLc && (events || []).some(e => {
    const inOrg = (e.organizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
    if (inOrg) return true;
    return (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
  });
  const isOrganizer = canCreateEvents || isOrganizerOfAnyEvent;
  // v13.11: Reines Check-In-Team — User ist per qrScannerEmails Mitglied
  // mindestens eines AKTIVEN Events, hat aber sonst keine Organizer-Rechte.
  // Die Organizer-Kachel wird dann klickbar und navigiert auf die
  // Check-In-Seite. Ein kleines „Check-In"-Badge oben rechts macht klar,
  // dass die Berechtigung auf Check-In beschränkt ist (kein Event-Edit,
  // keine Massenmails, keine Rollen-Verwaltung).
  const isCheckInTeamOfActive = !isOrganizer && !!currentEmailLc && (events || []).some(e => {
    if (e.status !== 'Active') return false;
    return (e.qrScannerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
  });
  // v12.5: Grid bleibt immer 3-spaltig — auch User sehen die Organizer-
  // Kachel (ausgegraut + CTA-Overlay).
  const showAdminTile = true;

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
      `}</style>
      <div className={`start-grid${showAdminTile ? ' start-grid--with-admin' : ''}`}>
        <div className="card card-clickable start-card" onClick={() => navigate('register')}>
          <div className="start-card__icon">
            <Calendar size={64} strokeWidth={1} />
          </div>
          <h2>{t('start.register')}</h2>
          <p style={{ whiteSpace: 'nowrap' }}>{t('start.register.desc')}</p>
        </div>
        <div className="card card-clickable start-card" onClick={() => navigate('my-events')}>
          <div className="start-card__icon">
            <Pin size={64} strokeWidth={1} />
          </div>
          <h2>{t('start.myevents')}</h2>
          <p style={{ whiteSpace: 'nowrap' }}>{t('start.myevents.desc')}</p>
        </div>
        {isOrganizer ? (
          <div className="card card-clickable start-card start-card--admin" onClick={() => navigate('admin')}>
            <div className="start-card__icon">
              <Settings size={64} strokeWidth={1} />
            </div>
            <h2>{t('start.admin')}</h2>
            <p style={{ whiteSpace: 'nowrap' }}>{t('start.admin.desc')}</p>
          </div>
        ) : isCheckInTeamOfActive ? (
          // v13.11: Check-In-Team-Variante — klickbar, leitet auf die
          // Check-In-Seite. Rechts oben sitzt ein Badge „Check-In" als
          // Hinweis, dass nur die Check-In-Funktion freigeschaltet ist.
          <div
            className="card card-clickable start-card start-card--admin"
            onClick={() => navigate('check-in')}
            style={{ position: 'relative' }}
          >
            <div
              style={{
                position: 'absolute', top: 12, right: 12,
                background: 'var(--dex-green, #86bc25)', color: '#fff',
                padding: '4px 10px', borderRadius: 999,
                fontSize: '0.72rem', fontWeight: 700,
                boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
                letterSpacing: 0.3, textTransform: 'uppercase',
              }}
            >
              {locale === 'de' ? 'Check-In' : 'Check-in'}
            </div>
            <div className="start-card__icon">
              <Settings size={64} strokeWidth={1} />
            </div>
            <h2>{t('start.admin')}</h2>
            <p style={{ whiteSpace: 'nowrap' }}>
              {locale === 'de'
                ? 'Teilnehmer einchecken (Scan + manuell)'
                : 'Check in attendees (scan + manual)'}
            </p>
          </div>
        ) : (
          // v12.5: Ausgegraute Organizer-Kachel mit CTA-Overlay-Button
          // — User können so direkt anfragen Organizer zu werden.
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
      </div>
      <InquiryModal open={showInquiry} onClose={() => setShowInquiry(false)} />
    </div>
  );
}
