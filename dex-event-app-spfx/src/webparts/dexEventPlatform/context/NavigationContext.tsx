/**
 * Navigation Context - ersetzt react-router fuer SPFx
 *
 * SPFx WebParts koennen kein Browser-Routing verwenden,
 * deshalb wird hier alles ueber State gesteuert.
 * Der History-Stack ermoeglicht die Zurueck-Navigation.
 */

import * as React from 'react';

export type Page = 'landing' | 'start' | 'register' | 'registration' | 'my-events' | 'create-event' | 'edit-event' | 'settings' | 'profile' | 'admin' | 'role-matrix' | 'participants' | 'flowcharts' | 'check-in' | 'help' | 'manual';

// Optionale Absicht beim Navigieren (z.B. Registration-Seite direkt im "Fuer andere"-Modus oeffnen)
export type NavIntent = 'register-other' | 'auto-cancel' | undefined;

interface NavigationContextType {
  currentPage: Page;
  selectedEventId: string | null;
  navIntent: NavIntent;
  navigate: (page: Page, eventId?: string, intent?: NavIntent) => void;
  goBack: () => void;
  clearIntent: () => void;
}

// Exportiert, damit Preview-Wrapper im Handbuch den Context mit Demo-Daten
// überschreiben können (v6.27 App-Screenshots).
export const NavigationContext = React.createContext<NavigationContextType | undefined>(undefined);

interface HistoryEntry {
  page: Page;
  eventId: string | null;
  intent: NavIntent;
}

export function NavigationProvider(props: { children: React.ReactNode }): React.ReactElement {
  const [currentPage, setCurrentPage] = React.useState<Page>('landing');
  const [selectedEventId, setSelectedEventId] = React.useState<string | null>(null);
  const [navIntent, setNavIntent] = React.useState<NavIntent>(undefined);
  const [history, setHistory] = React.useState<HistoryEntry[]>([]);

  const navigate = (page: Page, eventId?: string, intent?: NavIntent): void => {
    setHistory(prev => [...prev, { page: currentPage, eventId: selectedEventId, intent: navIntent }]);
    setCurrentPage(page);
    setSelectedEventId(eventId || null);
    setNavIntent(intent);
  };

  const goBack = (): void => {
    if (history.length > 0) {
      const prev = history[history.length - 1];
      setHistory(h => h.slice(0, -1));
      setCurrentPage(prev.page);
      setSelectedEventId(prev.eventId);
      setNavIntent(prev.intent);
    }
  };

  const clearIntent = (): void => setNavIntent(undefined);

  return React.createElement(
    NavigationContext.Provider,
    { value: { currentPage, selectedEventId, navIntent, navigate, goBack, clearIntent } },
    props.children
  );
}

export function useNavigation(): NavigationContextType {
  const ctx = React.useContext(NavigationContext);
  if (!ctx) throw new Error('useNavigation must be used within NavigationProvider');
  return ctx;
}
