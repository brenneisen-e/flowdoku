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
import { DEEPLINK_PARAM } from '../constants';

// v1.1: `landing` ist die erste Seite — der Startbildschirm mit dem Orb.
// `start` ist die Kachelwand dahinter.
// v1.3: `protokoll` — das Änderungsprotokoll (Organizer). Die `useCaseId`
// engt es auf einen Use Case ein.
export type Page = 'landing' | 'start' | 'usecases' | 'detail' | 'studio' | 'rollen' | 'protokoll';

interface NavEntry { page: Page; useCaseId?: number }

/**
 * v1.3: Die Use-Case-Id aus der Adresse der Host-Seite — oder `undefined`.
 *
 * Nur eine ganze, positive Zahl zählt. Alles andere (`uc=abc`, `uc=-1`,
 * `uc=7x`) ist kein Deep-Link und fällt auf die Startseite zurück: Eine
 * halb lesbare Id auf eine Detailseite zu schicken hieße, „nicht gefunden"
 * zu behaupten, obwohl der Link nur kaputt getippt war.
 *
 * Bewusst ein Regex und kein `URLSearchParams`: Die Adresse gehört der
 * SharePoint-Seite, und wir wollen genau EINEN Parameter daraus, ohne uns auf
 * die Verfügbarkeit einer weiteren Browser-API zu verlassen.
 */
function ucIdAusAdresse(): number | undefined {
  try {
    const m = new RegExp('[?&]' + DEEPLINK_PARAM + '=(\\d{1,9})(?:&|#|$)').exec(window.location.search || '');
    if (!m) return undefined;
    const id = parseInt(m[1], 10);
    return id > 0 ? id : undefined;
  } catch {
    return undefined;
  }
}

interface NavigationContextType {
  currentPage: Page;
  currentUseCaseId?: number;
  navigate: (page: Page, useCaseId?: number) => void;
  goBack: () => void;
  canGoBack: boolean;
}

const NavigationContext = React.createContext<NavigationContextType | undefined>(undefined);

export function NavigationProvider(props: { children: React.ReactNode }): React.ReactElement {
  // Die Adresse wird EINMAL beim Start gelesen (lazy Initialisierer), nicht
  // bei jedem Render: Nach dem ersten Seitenwechsel gehört die Navigation
  // der App, und ein Deep-Link darf sie nicht zurückholen.
  const [entry, setEntry] = React.useState<NavEntry>(() => {
    const id = ucIdAusAdresse();
    return id ? { page: 'detail', useCaseId: id } : { page: 'landing' };
  });
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
