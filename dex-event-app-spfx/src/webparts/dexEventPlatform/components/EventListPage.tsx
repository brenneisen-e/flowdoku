/**
 * Event-Uebersicht
 *
 * Zeigt Events als Karten an.
 * Standort-Filter: User sehen nur Events fuer ihren Standort.
 * Admin/Organizer sehen alle Events.
 */

import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { DeloitteEvent } from '../types';
import EventCard from './EventCard';

/**
 * Prüft ob ein User-Standort (z.B. "DE - Koeln") zu einem LocationFilter passt.
 * LocationFilter kann "Köln, Düsseldorf, All" sein.
 */
function matchesLocation(userLocation: string, locationFilter: string): boolean {
  if (!locationFilter || !locationFilter.trim()) return true; // Kein Filter = alle
  const filters = locationFilter.split(',').map(s => s.trim().toLowerCase());
  if (filters.indexOf('all') >= 0) return true;
  if (!userLocation) return false;
  const loc = userLocation.toLowerCase();
  // Matche z.B. "DE - Koeln" gegen "Köln" oder "Koeln"
  return filters.some(f => {
    const normalized = f.replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ä/g, 'ae').replace(/ß/g, 'ss');
    return loc.indexOf(f) >= 0 || loc.indexOf(normalized) >= 0;
  });
}

export default function EventListPage(): React.ReactElement {
  const { events } = useEvents();
  const { currentUser } = useCurrentUser();
  const { canCreateEvents } = useRoles();
  const [onlyActive, setOnlyActive] = React.useState(true);

  const statusFiltered = onlyActive
    ? events.filter((e) => e.status === 'Active')
    : events;

  // Admin/Organizer sehen alle, normale User nur passende Standorte
  const filteredEvents = canCreateEvents
    ? statusFiltered
    : statusFiltered.filter((e: DeloitteEvent) =>
        matchesLocation(currentUser.location, e.locationAudience.join(', '))
      );

  return (
    <div className="page-container">
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
          Keine Events für deinen Standort gefunden.
        </p>
      )}
    </div>
  );
}
