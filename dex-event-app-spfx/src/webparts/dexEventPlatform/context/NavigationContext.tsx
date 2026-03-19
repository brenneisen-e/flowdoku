/**
 * Navigation Context - ersetzt react-router fuer SPFx
 *
 * SPFx WebParts koennen kein Browser-Routing verwenden,
 * deshalb wird hier alles ueber State gesteuert.
 * Der History-Stack ermoeglicht die Zurueck-Navigation.
 */

import * as React from 'react';

export type Page = 'landing' | 'start' | 'register' | 'registration' | 'my-events' | 'create-event' | 'settings';

interface NavigationContextType {
  currentPage: Page;
  selectedEventId: string | null;
  navigate: (page: Page, eventId?: string) => void;
  goBack: () => void;
}

const NavigationContext = React.createContext<NavigationContextType | undefined>(undefined);

interface HistoryEntry {
  page: Page;
  eventId: string | null;
}

export function NavigationProvider(props: { children: React.ReactNode }): React.ReactElement {
  const [currentPage, setCurrentPage] = React.useState<Page>('landing');
  const [selectedEventId, setSelectedEventId] = React.useState<string | null>(null);
  const [history, setHistory] = React.useState<HistoryEntry[]>([]);

  const navigate = (page: Page, eventId?: string): void => {
    // Aktuelle Seite in den History-Stack pushen
    setHistory(prev => [...prev, { page: currentPage, eventId: selectedEventId }]);
    setCurrentPage(page);
    setSelectedEventId(eventId || null);
  };

  const goBack = (): void => {
    if (history.length > 0) {
      const prev = history[history.length - 1];
      setHistory(h => h.slice(0, -1));
      setCurrentPage(prev.page);
      setSelectedEventId(prev.eventId);
    }
  };

  return React.createElement(
    NavigationContext.Provider,
    { value: { currentPage, selectedEventId, navigate, goBack } },
    props.children
  );
}

export function useNavigation(): NavigationContextType {
  const ctx = React.useContext(NavigationContext);
  if (!ctx) throw new Error('useNavigation must be used within NavigationProvider');
  return ctx;
}
