/**
 * Event-Uebersicht
 *
 * Zeigt alle verfuegbaren Events als Karten an.
 * Mit Toggle kann man zwischen "nur aktive" und "alle" Events wechseln.
 */

import * as React from 'react';
import { useEvents } from '../context/EventContext';
import EventCard from './EventCard';

export default function EventListPage(): React.ReactElement {
  const { events } = useEvents();
  const [onlyActive, setOnlyActive] = React.useState(true);

  const filteredEvents = onlyActive
    ? events.filter((e) => e.status === 'Active')
    : events;

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
          No events found.
        </p>
      )}
    </div>
  );
}
