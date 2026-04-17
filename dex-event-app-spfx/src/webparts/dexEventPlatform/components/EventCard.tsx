/**
 * Event-Karte fuer die Uebersichtsseite
 *
 * Zeigt Gradient-Hintergrund, Event-Infos und freie Plaetze.
 * Die Gradient-Farben rotieren basierend auf dem Index.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';
import { useRoles } from '../context/RoleContext';
import { DeloitteEvent } from '../types';

// Deutsches Datumsformat
function formatDate(iso: string): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('de-DE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }) + ' ' + d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
}

// Nur Datum (ohne Uhrzeit) - wird z.B. fuer die Registration Deadline verwendet,
// die im Formular ausschliesslich als Datum gepflegt wird.
function formatDateOnly(iso: string): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('de-DE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  });
}

// 4 verschiedene Farbverlaeufe, rotieren durch
// type wird aktuell nicht genutzt, koennte man spaeter fuer typspezifische Farben verwenden
function getEventGradient(_type: string, index: number): string {
  const gradients = [
    'linear-gradient(135deg, #0a2e1a 0%, #1a6b3c 40%, #00ff88 100%)',
    'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
    'linear-gradient(135deg, #2d5016 0%, #86bc25 50%, #c5e63c 100%)',
    'linear-gradient(135deg, #1a0a2e 0%, #4a1a6b 40%, #8800ff 100%)',
  ];
  return gradients[index % gradients.length];
}

interface Props {
  event: DeloitteEvent;
  index: number;
  isRegistered?: boolean;
  isWaitlisted?: boolean;
}

export default function EventCard({ event, index, isRegistered, isWaitlisted }: Props): React.ReactElement {
  const { navigate } = useNavigation();
  const { t } = useLanguage();
  const { canCreateEvents } = useRoles();
  const isUnlimited = !event.maxParticipants || event.maxParticipants === 0;
  const freePlaces = isUnlimited ? Infinity : event.maxParticipants - event.currentParticipants;
  const isFull = !isUnlimited && freePlaces <= 0;
  const alreadySignedUp = isRegistered || isWaitlisted;
  const isDeadlinePassed = !!event.registrationDeadline && new Date(event.registrationDeadline) < new Date();
  // Nur normale User bekommen den Deadline-Overlay. Organizer/Admins duerfen
  // trotzdem reinklicken, um ggf. manuell zu registrieren.
  const showDeadlineOverlay = isDeadlinePassed && !canCreateEvents && !alreadySignedUp;

  return (
    <div className="event-card" style={{ position: 'relative', cursor: showDeadlineOverlay ? 'not-allowed' : 'pointer' }} onClick={() => (!alreadySignedUp && !showDeadlineOverlay) ? navigate('registration', event.id) : undefined}>
      {event.isFictive && (
        <div style={{
          position: 'absolute', top: 10, right: 10, zIndex: 5,
          padding: '3px 10px', borderRadius: 999,
          background: 'var(--dex-orange, #ed8b00)', color: '#fff',
          fontSize: '0.68rem', fontWeight: 700, letterSpacing: 0.5,
          boxShadow: '0 1px 4px rgba(0,0,0,0.15)',
        }}>
          {t('create.fictive.badge')}
        </div>
      )}
      {showDeadlineOverlay && (
        <div style={{
          position: 'absolute', inset: 0, zIndex: 10, borderRadius: 'var(--dex-radius)',
          background: 'rgba(0,0,0,0.65)', display: 'flex', flexDirection: 'column',
          alignItems: 'center', justifyContent: 'center', padding: 24, textAlign: 'center',
        }}>
          <div style={{ color: '#fff', fontWeight: 700, fontSize: '1rem', marginBottom: 4 }}>
            {t('events.deadlinepassed')}
          </div>
          <div style={{ color: 'rgba(255,255,255,0.85)', fontSize: '0.85rem', marginBottom: 4, maxWidth: 320, lineHeight: 1.4 }}>
            {t('events.deadlinepassed.hint')}
          </div>
          <div style={{ color: 'rgba(255,255,255,0.6)', fontSize: '0.8rem', marginTop: 4 }}>
            {event.title}
          </div>
        </div>
      )}
      {alreadySignedUp && (
        <div style={{
          position: 'absolute', inset: 0, zIndex: 10, borderRadius: 'var(--dex-radius)',
          background: 'rgba(0,0,0,0.65)', display: 'flex', flexDirection: 'column',
          alignItems: 'center', justifyContent: 'center', padding: 24, textAlign: 'center',
        }}>
          <div style={{ color: '#fff', fontWeight: 700, fontSize: '1rem', marginBottom: 4 }}>
            {isWaitlisted ? t('status.waitlist') : t('status.registered')}
          </div>
          <div style={{ color: 'rgba(255,255,255,0.7)', fontSize: '0.8rem', marginBottom: 16 }}>
            {event.title}
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8, alignItems: 'center' }}>
            <button
              className="btn btn-primary"
              style={{ fontSize: '0.85rem', padding: '8px 20px' }}
              onClick={(e: React.MouseEvent) => { e.stopPropagation(); navigate('my-events'); }}
            >
              {t('myevents.title')}
            </button>
            {canCreateEvents && (
              <button
                className="btn btn-secondary"
                style={{ fontSize: '0.8rem', padding: '6px 16px' }}
                onClick={(e: React.MouseEvent) => {
                  e.stopPropagation();
                  navigate('registration', event.id, 'register-other');
                }}
              >
                {t('reg.registerother')}
              </button>
            )}
          </div>
        </div>
      )}
      <div className="event-card__image" style={{
        background: event.imageUrl
          ? `url(${event.imageUrl}) center/cover no-repeat`
          : getEventGradient(event.type, index),
      }}>
        <div className="event-card__overlay">
          <h3 className="event-card__title">{event.title}</h3>
          <div className="event-card__meta">
            <span>{event.location}</span>
            <span className="event-card__places">
              {isUnlimited ? '' : isFull ? '' : `${freePlaces} free`}
            </span>
          </div>
        </div>
      </div>
      <div className="event-card__body" style={{ display: 'flex', flexDirection: 'column', flex: 1 }}>
        {/* Datum links + Freie-Plaetze-Badge rechts in einer Zeile */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 12, marginBottom: 8 }}>
          <div className="event-card__dates" style={{ flex: 1, minWidth: 0 }}>
            {formatDate(event.startDate)} {t('events.until')}
            <br />
            {formatDate(event.endDate)}
          </div>
          <span style={{
            padding: '3px 10px', borderRadius: 12, fontSize: '0.78rem', fontWeight: 600, whiteSpace: 'nowrap', flexShrink: 0,
            background: isFull ? 'rgba(218,41,28,0.12)' : 'rgba(134,188,37,0.12)',
            color: isFull ? 'var(--dex-red)' : 'var(--dex-green-dark)',
          }}>
            {isFull ? t('status.waitlist') : (isUnlimited ? t('reg.unlimited') : `${freePlaces} ${t('reg.free')}`)}
          </span>
        </div>
        {event.registrationDeadline && formatDateOnly(event.registrationDeadline) && (
          <div className="event-card__deadline">
            {t('events.regopen')} {formatDateOnly(event.registrationDeadline)}
          </div>
        )}
        {event.registrationDeadline && new Date(event.registrationDeadline) < new Date() && (
          <div style={{
            marginTop: 6,
            padding: '6px 10px',
            background: 'rgba(218,41,28,0.10)',
            color: 'var(--dex-red)',
            borderRadius: 6,
            fontSize: '0.78rem',
            fontWeight: 600,
            lineHeight: 1.35,
          }}>
            {t('events.deadlinepassed')}
          </div>
        )}
        <div style={{ marginTop: 'auto', paddingTop: 12 }}>
          <button
            className="btn btn-primary event-card__register-btn"
            onClick={(e: React.MouseEvent) => {
              e.stopPropagation();
              navigate('registration', event.id);
            }}
          >
            {t('reg.register')}
          </button>
        </div>
      </div>
    </div>
  );
}
