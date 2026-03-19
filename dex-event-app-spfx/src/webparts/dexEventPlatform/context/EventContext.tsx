/**
 * Event Context - zentraler State fuer alle Events
 *
 * Mock-Events + vom User erstellte Events werden zusammengefuehrt.
 * Neue Events werden im localStorage gespeichert, damit sie
 * nach einem Reload nicht verloren gehen.
 *
 * - Eike, Maerz 2026
 */

import * as React from 'react';
import { DeloitteEvent } from '../types';
import { events as mockEvents } from '../data/mockData';

interface EventContextType {
  events: DeloitteEvent[];
  addEvent: (event: DeloitteEvent) => void;
}

const EventContext = React.createContext<EventContextType | undefined>(undefined);

const STORAGE_KEY = 'dex-custom-events';

function loadCustomEvents(): DeloitteEvent[] {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    return stored ? JSON.parse(stored) : [];
  } catch {
    return [];
  }
}

export function EventProvider(props: { children: React.ReactNode }): React.ReactElement {
  const [customEvents, setCustomEvents] = React.useState<DeloitteEvent[]>(loadCustomEvents);

  // Bei jeder Aenderung in localStorage schreiben
  React.useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(customEvents));
  }, [customEvents]);

  const allEvents = [...mockEvents, ...customEvents];

  const addEvent = (event: DeloitteEvent): void => {
    setCustomEvents(prev => [...prev, event]);
  };

  return React.createElement(
    EventContext.Provider,
    { value: { events: allEvents, addEvent } },
    props.children
  );
}

export function useEvents(): EventContextType {
  const ctx = React.useContext(EventContext);
  if (!ctx) throw new Error('useEvents must be used within EventProvider');
  return ctx;
}
