/**
 * Event-Karte fuer die Uebersichtsseite
 *
 * Zeigt Gradient-Hintergrund, Event-Infos und freie Plaetze.
 * Die Gradient-Farben rotieren basierend auf dem Index.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
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
}

export default function EventCard({ event, index }: Props): React.ReactElement {
  const { navigate } = useNavigation();
  const freePlaces = event.maxParticipants - event.currentParticipants;
  const isFull = freePlaces <= 0;

  return (
    <div className="event-card" onClick={() => navigate('registration', event.id)}>
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
              {isFull ? (
                <>
                  No free places.
                  {event.waitlistCount > 0 && (
                    <> Current position on waiting list: {event.waitlistCount}</>
                  )}
                </>
              ) : (
                <>Free places: {freePlaces}</>
              )}
            </span>
          </div>
        </div>
      </div>
      <div className="event-card__body">
        <div className="event-card__dates">
          {formatDate(event.startDate)} until
          <br />
          {formatDate(event.endDate)}
        </div>
        {event.registrationDeadline && formatDate(event.registrationDeadline) && (
          <div className="event-card__deadline">
            Registration open until: {formatDate(event.registrationDeadline)}
          </div>
        )}
        <button
          className="btn btn-primary event-card__register-btn"
          onClick={(e: React.MouseEvent) => {
            e.stopPropagation();
            navigate('registration', event.id);
          }}
        >
          Registrate
        </button>
      </div>
    </div>
  );
}
