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
import { Info } from './Icons';
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

  return (
    <div className="event-card" style={{ position: 'relative' }} onClick={() => !alreadySignedUp ? navigate('registration', event.id) : undefined}>
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
        <button
          className="event-card__info-btn"
          onClick={(e: React.MouseEvent) => {
            e.stopPropagation();
            navigate('registration', event.id);
          }}
          aria-label="Event info"
        >
          <Info size={18} />
        </button>
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
        {event.registrationDeadline && formatDate(event.registrationDeadline) && (
          <div className="event-card__deadline">
            {t('events.regopen')} {formatDate(event.registrationDeadline)}
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
