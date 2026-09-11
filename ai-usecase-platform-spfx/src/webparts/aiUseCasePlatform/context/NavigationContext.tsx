/**
 * Seitenwechsel und Zurueck.
 *
 * Aufbau wie DEX: EINE Stelle kennt die aktuelle Seite, alles andere ruft
 * `navigate`. Der Verlauf ist ein eigener Stapel und nicht die Browser-
 * Historie — in SharePoint gehoert die Adresszeile der Host-Seite, ein
 * `pushState` daran ist unzuverlaessig.
 *
 * Deep-Links laufen ueber Abfrageparameter der Host-Seite
 * (`?uc=<id>`), gelesen EINMAL beim Start. Muster wie `?action=` in DEX.
 */

import * as React from 'react';

// v1.1: `landing` ist die erste Seite — der Startbildschirm mit dem Orb.
// `start` ist die Kachelwand dahinter.
export type Page = 'landing' | 'start' | 'detail' | 'verwaltung' | 'rollen';

interface NavEntry { page: Page; useCaseId?: number }

interface NavigationContextType {
  currentPage: Page;
  currentUseCaseId?: number;
  navigate: (page: Page, useCaseId?: number) => void;
  goBack: () => void;
  canGoBack: boolean;
}

const NavigationContext = React.createContext<NavigationContextType | undefined>(undefined);

export function NavigationProvider(props: { children: React.ReactNode }): React.ReactElement {
  const [entry, setEntry] = React.useState<NavEntry>({ page: 'landing' });
  const [stack, setStack] = React.useState<NavEntry[]>([]);

  const navigate = React.useCallback((page: Page, useCaseId?: number): void => {
    setStack(prev => [...prev, entry]);
    setEntry({ page, useCaseId });
    // Nach oben — sonst steht die neue Seite mitten im Text der alten.
    try { window.scrollTo({ top: 0, behavior: 'smooth' }); } catch { /* aeltere Browser */ }
  }, [entry]);

  const goBack = React.useCallback((): void => {
    setStack(prev => {
      if (prev.length === 0) return prev;
      const last = prev[prev.length - 1];
      setEntry(last);
      return prev.slice(0, -1);
    });
  }, []);

  const value = React.useMemo<NavigationContextType>(() => ({
    currentPage: entry.page,
    currentUseCaseId: entry.useCaseId,
    navigate,
    goBack,
    canGoBack: stack.length > 0,
  }), [entry, navigate, goBack, stack.length]);

  return React.createElement(NavigationContext.Provider, { value }, props.children);
}

export function useNavigation(): NavigationContextType {
  const ctx = React.useContext(NavigationContext);
  if (!ctx) throw new Error('useNavigation muss innerhalb des NavigationProvider stehen');
  return ctx;
}
