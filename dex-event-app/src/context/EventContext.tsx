import { createContext, useContext, useState, useEffect } from 'react';
import type { DeloitteEvent } from '../types';
import { events as mockEvents } from '../data/mockData';

interface EventContextType {
  events: DeloitteEvent[];
  addEvent: (event: DeloitteEvent) => void;
}

const EventContext = createContext<EventContextType | null>(null);

const STORAGE_KEY = 'dex-custom-events';

function loadCustomEvents(): DeloitteEvent[] {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    return stored ? JSON.parse(stored) : [];
  } catch {
    return [];
  }
}

export function EventProvider({ children }: { children: React.ReactNode }) {
  const [customEvents, setCustomEvents] = useState<DeloitteEvent[]>(loadCustomEvents);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(customEvents));
  }, [customEvents]);

  const allEvents = [...mockEvents, ...customEvents];

  const addEvent = (event: DeloitteEvent) => {
    setCustomEvents((prev) => [...prev, event]);
  };

  return (
    <EventContext.Provider value={{ events: allEvents, addEvent }}>
      {children}
    </EventContext.Provider>
  );
}

export function useEvents() {
  const ctx = useContext(EventContext);
  if (!ctx) throw new Error('useEvents must be used within EventProvider');
  return ctx;
}
