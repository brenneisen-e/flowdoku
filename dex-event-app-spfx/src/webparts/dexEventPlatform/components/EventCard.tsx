/**
 * Event-Karte für die Uebersichtsseite
 *
 * Zeigt Gradient-Hintergrund, Event-Infos und freie Plätze.
 * Die Gradient-Farben rotieren basierend auf dem Index.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { DeloitteEvent } from '../types';
import { useCachedImage } from '../utils/imageCache';
import { isRegistrationFullyClosed } from '../utils/eventFormat';

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

// Nur Datum (ohne Uhrzeit) - wird z.B. für die Registration Deadline verwendet,
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

// 4 verschiedene Farbverläufe, rotieren durch
// type wird aktuell nicht genutzt, könnte man später für typspezifische Farben verwenden
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
  /** v11.90: eingeloggter User ist Haupt- oder Co-Organizer dieses Events. */
  isOwnOrganizer?: boolean;
}

export default function EventCard({ event, index, isRegistered, isWaitlisted, isOwnOrganizer }: Props): React.ReactElement {
  const { navigate } = useNavigation();
  const { t } = useLanguage();
  const { canCreateEvents } = useRoles();
  const { childEventsOf } = useEvents();
  // v19.22: Event-Bild über den IndexedDB-Cache — beim zweiten App-Aufruf sofort
  // da, ohne SharePoint-Roundtrip.
  const cachedImage = useCachedImage(event.imageUrl);
  // v9.8: B2Run-Events haben maxParticipants=0, weil die Kapazität auf
  // Durchstarter + Funstarter aufgeteilt ist. Die Summe gilt als
  // Gesamtkapazität — sonst zeigt die Karte fälschlich "Unbegrenzt", obwohl
  // es z.B. 140 Plätze gibt.
  const splitCapacity = (event.durchstarterCapacity || 0) + (event.funstarterCapacity || 0);
  const effectiveMax = event.maxParticipants && event.maxParticipants > 0
    ? event.maxParticipants
    : splitCapacity;
  const isUnlimited = !effectiveMax || effectiveMax === 0;
  const freePlaces = isUnlimited ? Infinity : effectiveMax - event.currentParticipants;
  const isFull = !isUnlimited && freePlaces <= 0;
  const alreadySignedUp = isRegistered || isWaitlisted;
  // v22.54: Die Anmeldung bleibt offen, solange das Hauptevent ODER mindestens
  // ein Sub-Event noch offen ist — eine abgelaufene Klammer-/Hauptevent-Frist
  // sperrt nicht mehr das ganze Event, wenn die Sub-Events noch laufen.
  const isDeadlinePassed = isRegistrationFullyClosed(event, childEventsOf(event.id));
  // Nur normale User bekommen den Deadline-Overlay. Organizer/Admins dürfen
  // trotzdem reinklicken, um ggf. manuell zu registrieren.
  const showDeadlineOverlay = isDeadlinePassed && !canCreateEvents && !alreadySignedUp;
  // v23.14: Vorschau vor Aktivierung („Aktiv ab" in der Zukunft + previewBeforeActive).
  // Reguläre User sehen die Karte, dürfen aber NICHT in die Anmeldeseite —
  // blockierender Overlay „Anmeldung ab …". Organizer/Admins dürfen weiterhin
  // rein (zum Vorbereiten) und bekommen stattdessen nur einen Hinweis-Badge.
  const activeFromTs = event.activeFrom ? new Date(event.activeFrom).getTime() : 0;
  const notYetActive = activeFromTs > 0 && activeFromTs > Date.now();
  const showPreviewOverlay = notYetActive && !canCreateEvents && !isOwnOrganizer && !alreadySignedUp;
  const showOrganizerActiveBadge = notYetActive && (canCreateEvents || isOwnOrganizer);
  const blockClick = showDeadlineOverlay || showPreviewOverlay;

  return (
    <div className="event-card" style={{ position: 'relative', cursor: blockClick ? 'not-allowed' : 'pointer', ...(event.isDemoShowcase ? { outline: '2px dashed var(--dex-blue, #0076a8)', outlineOffset: 2 } : {}) }} onClick={() => (!alreadySignedUp && !blockClick) ? navigate('registration', event.id) : undefined}>
      {/* v17.25: Demo-Showcase-Event deutlich markieren (blaues DEMO-Badge
          oben rechts, gestrichelter Rahmen). Nur im Demo-Impersonation-Modus
          ueberhaupt in der Liste. */}
      {event.isDemoShowcase && (
        <div style={{
          position: 'absolute', top: 10, right: 10, zIndex: 6,
          padding: '3px 10px', borderRadius: 999,
          background: 'var(--dex-blue, #0076a8)', color: '#fff',
          fontSize: '0.68rem', fontWeight: 700, letterSpacing: 1,
          boxShadow: '0 1px 4px rgba(0,0,0,0.15)',
        }}>
          DEMO
        </div>
      )}
      {event.isFictive && !event.isDemoShowcase && (
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
      {/* v11.90: Eigene Events (User selbst Haupt- oder Co-Organizer)
          bekommen ein grünes „Organizer"-Badge oben links — auch in
          Kombination mit dem orangen „Entwurf"-Badge oben rechts. */}
      {isOwnOrganizer && (
        <div style={{
          position: 'absolute', top: 10, left: 10, zIndex: 5,
          padding: '3px 10px', borderRadius: 999,
          background: 'var(--dex-green, #86bc25)', color: '#fff',
          fontSize: '0.68rem', fontWeight: 700, letterSpacing: 0.5,
          boxShadow: '0 1px 4px rgba(0,0,0,0.15)',
        }}>
          Organizer
        </div>
      )}
      {/* v23.14: Organizer/Admin sehen bei „Aktiv ab" in der Zukunft einen
          Hinweis-Banner „Anmeldung ab …" (nicht blockierend — sie dürfen rein,
          um das Event vorzubereiten). Regulären Usern wird stattdessen der
          blockierende Vorschau-Overlay gezeigt. */}
      {showOrganizerActiveBadge && (
        <div style={{
          position: 'absolute', top: 0, left: 0, right: 0, zIndex: 6,
          padding: '5px 10px',
          background: 'rgba(74,124,31,0.95)', color: '#fff',
          fontSize: '0.72rem', fontWeight: 700, textAlign: 'center',
          borderTopLeftRadius: 'var(--dex-radius)', borderTopRightRadius: 'var(--dex-radius)',
        }}>
          {t('events.regfrom')} {formatDate(event.activeFrom || '')}
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
      {/* v23.14: Vorschau-Overlay für reguläre User — sichtbar, aber Anmeldung
          erst ab dem Aktivierungszeitpunkt (Anmeldeseite nicht öffenbar). */}
      {showPreviewOverlay && (
        <div style={{
          position: 'absolute', inset: 0, zIndex: 10, borderRadius: 'var(--dex-radius)',
          background: 'rgba(0,0,0,0.65)', display: 'flex', flexDirection: 'column',
          alignItems: 'center', justifyContent: 'center', padding: 24, textAlign: 'center',
        }}>
          <div style={{ color: '#fff', fontWeight: 700, fontSize: '1rem', marginBottom: 4 }}>
            {t('events.previewsoon')}
          </div>
          <div style={{ color: 'rgba(255,255,255,0.9)', fontSize: '0.9rem', marginBottom: 4 }}>
            {t('events.regfrom')} {formatDate(event.activeFrom || '')}
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
          ? `url(${cachedImage}) center/cover no-repeat`
          : getEventGradient(event.type, index),
        position: 'relative',
      }}>
        {/* v23.19: Optionale Pro-Ansicht-Darstellung (Karte) — überlagert das
            Hintergrund-Bild mit individuellem Zoom/Position. Nur wenn gesetzt. */}
        {event.imageUrl && event.imageDisplay?.card && (
          <img
            src={cachedImage}
            alt=""
            style={{
              position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'cover',
              objectPosition: `center ${event.imageDisplay.card.posY}%`,
              transform: `scale(${event.imageDisplay.card.zoom})`,
              transformOrigin: `center ${event.imageDisplay.card.posY}%`,
            }}
          />
        )}
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
        {/* Datum links + Freie-Plätze-Badge rechts in einer Zeile */}
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
        {isDeadlinePassed && (
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
