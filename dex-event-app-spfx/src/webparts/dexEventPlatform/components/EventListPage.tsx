/**
 * Event-Uebersicht
 *
 * Zeigt Events als Karten an.
 * Standort + Zielgruppen-Filter: User sehen nur passende Events.
 * Admin/Organizer sehen alle Events.
 */

import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { DeloitteEvent } from '../types';
import EventCard from './EventCard';

/**
 * Prüft ob ein User-Standort zu einem LocationFilter passt.
 * "DE - Koeln" matcht "Köln" oder "Koeln".
 */
function matchesLocation(userLocation: string, locationFilters: string[]): boolean {
  if (locationFilters.length === 0) return true;
  const filters = locationFilters.map(s => s.trim().toLowerCase());
  if (filters.indexOf('all') >= 0) return true;
  if (!userLocation) return false;
  const loc = userLocation.toLowerCase();
  return filters.some(f => {
    const normalized = f.replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ä/g, 'ae').replace(/ß/g, 'ss');
    return loc.indexOf(f) >= 0 || loc.indexOf(normalized) >= 0;
  });
}

/**
 * Prüft ob ein User zur Zielgruppe passt.
 * Zielgruppe kann Verteilergruppen (DEALL, SAPALL) oder Einzelmails sein.
 * Einzelmails: exakter Match auf User-Email.
 * Gruppen: werden gegen User-Email-Prefix gematcht (z.B. DEALL matcht alle @deloitte.de).
 */
function matchesAudience(userEmail: string, userLocation: string, audienceFilters: string[]): boolean {
  if (audienceFilters.length === 0) return true;
  const email = userEmail.toLowerCase();
  const loc = (userLocation || '').toLowerCase();

  return audienceFilters.some(filter => {
    const f = filter.trim().toLowerCase();
    if (!f) return false;

    // Einzelne E-Mail-Adresse
    if (f.indexOf('@') >= 0) {
      return email === f;
    }

    // Gruppen-Patterns
    if (f === 'deall') return true; // Alle DE-Mitarbeiter
    // Standort-Gruppen: DEKOELN, DEHAMBURG, etc.
    if (f.startsWith('de')) {
      const city = f.substring(2);
      const normalized = city.replace(/oe/g, 'ö').replace(/ue/g, 'ü').replace(/ae/g, 'ä');
      return loc.indexOf(city) >= 0 || loc.indexOf(normalized) >= 0;
    }
    // Abteilungs-Gruppen: SAPALL, TECHALL, etc. (Zukunft: Graph API Integration)
    return false;
  });
}

/**
 * Prüft ob ein Event fuer den User sichtbar ist.
 */
function isEventVisibleForUser(
  event: DeloitteEvent,
  userEmail: string,
  userLocation: string
): boolean {
  const hasLocationFilter = event.locationAudience.length > 0;
  const hasAudienceFilter = event.audienceFilter.length > 0;

  if (!hasLocationFilter && !hasAudienceFilter) return true;

  const locMatch = matchesLocation(userLocation, event.locationAudience);
  const audMatch = matchesAudience(userEmail, userLocation, event.audienceFilter);

  if (event.filterMode === 'AND') {
    // Beide müssen passen (sofern gesetzt)
    if (hasLocationFilter && hasAudienceFilter) return locMatch && audMatch;
    if (hasLocationFilter) return locMatch;
    return audMatch;
  }
  // OR: mindestens einer muss passen
  if (hasLocationFilter && hasAudienceFilter) return locMatch || audMatch;
  if (hasLocationFilter) return locMatch;
  return audMatch;
}

export default function EventListPage(): React.ReactElement {
  const { events, isEventsLoading } = useEvents();
  const { currentUser } = useCurrentUser();
  const { canCreateEvents, currentUserRole } = useRoles();
  const [onlyActive, setOnlyActive] = React.useState(true);
  const [showDebug, setShowDebug] = React.useState(false);

  const statusFiltered = onlyActive
    ? events.filter((e) => e.status === 'Active')
    : events;

  // Admin/Organizer sehen alle, normale User nur passende
  const filteredEvents = canCreateEvents
    ? statusFiltered
    : statusFiltered.filter((e: DeloitteEvent) =>
        isEventVisibleForUser(e, currentUser.email, currentUser.location)
      );

  return (
    <div className="page-container">
      {/* Debug-Panel fuer mobiles Testen */}
      <div style={{ marginBottom: 8, textAlign: 'right' }}>
        <button onClick={() => setShowDebug(!showDebug)} style={{ fontSize: '0.7rem', padding: '2px 8px', opacity: 0.5 }}>Debug</button>
      </div>
      {showDebug && (
        <div style={{ background: '#1e1e1e', color: '#0f0', padding: 12, borderRadius: 8, fontSize: '0.7rem', fontFamily: 'monospace', marginBottom: 16, maxHeight: 300, overflowY: 'auto', whiteSpace: 'pre-wrap' }}>
          {`User: ${currentUser.email}
Location: "${currentUser.location}"
Role: ${currentUserRole}
canCreateEvents: ${canCreateEvents}
isEventsLoading: ${isEventsLoading}
events.length: ${events.length}
statusFiltered.length: ${statusFiltered.length}
filteredEvents.length: ${filteredEvents.length}
onlyActive: ${onlyActive}

Events Detail:
${events.map(e => `  #${e.eventNumber} "${e.title}" status=${e.status} loc=[${e.locationAudience.join(',')}] aud=[${e.audienceFilter.join(',')}] filterMode=${e.filterMode} visible=${isEventVisibleForUser(e, currentUser.email, currentUser.location)}`).join('\n') || '(keine Events geladen)'}
`}
        </div>
      )}
      <div className="flex-between mb-16">
        <div />
        <div className="toggle-wrapper">
          <label className="toggle">
            <input
              type="checkbox"
              checked={onlyActive}
              onChange={(e) => setOnlyActive(e.target.checked)}
            />
            <span className="toggle-slider" />
          </label>
          <span>Only Active Events</span>
        </div>
      </div>
      <div className="event-grid">
        {filteredEvents.map((event, i) => (
          <EventCard key={event.id} event={event} index={i} />
        ))}
      </div>
      {filteredEvents.length === 0 && (
        <p className="text-center mt-24" style={{ color: 'var(--dex-gray-400)' }}>
          Keine Events für dich gefunden.
        </p>
      )}
    </div>
  );
}
