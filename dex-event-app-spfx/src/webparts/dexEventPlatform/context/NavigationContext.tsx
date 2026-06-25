/**
 * Navigation Context - ersetzt react-router für SPFx
 *
 * SPFx WebParts können kein Browser-Routing verwenden,
 * deshalb wird hier alles ueber State gesteuert.
 * Der History-Stack ermöglicht die Zurück-Navigation.
 */

import * as React from 'react';

export type Page = 'landing' | 'start' | 'register' | 'registration' | 'my-events' | 'assistant' | 'create-event' | 'edit-event' | 'settings' | 'profile' | 'admin' | 'admin-hub' | 'role-matrix' | 'participants' | 'flowcharts' | 'check-in' | 'self-checkin-display' | 'help' | 'manual' | 'email-templates';

// Optionale Absicht beim Navigieren (z.B. Registration-Seite direkt im "Für andere"-Modus oeffnen)
export type NavIntent = 'register-other' | 'auto-cancel' | undefined;

interface NavigationContextType {
  currentPage: Page;
  selectedEventId: string | null;
  navIntent: NavIntent;
  navigate: (page: Page, eventId?: string, intent?: NavIntent) => void;
  goBack: () => void;
  clearIntent: () => void;
  /** v17.3: Page registriert einen Confirm-Hook für ungespeicherte
   *  Aenderungen. Wird VOR jeder Navigation aufgerufen — wenn er false
   *  zurückliefert, blockiert das Navigation. Null = keine Aenderungen,
   *  durchnavigieren. */
  setNavigationGuard: (guard: (() => Promise<boolean>) | null) => void;
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
  const guardRef = React.useRef<(() => Promise<boolean>) | null>(null);

  // v20.0 (Audit): Context-Value memoizen — die Funktionen schließen über
  // currentPage/selectedEventId/navIntent/history, daher sind genau diese
  // vier die Memo-Abhängigkeiten. Verhindert App-weite Re-Renders aller
  // useNavigation()-Consumer bei Parent-Re-Renders ohne Navigations-Änderung.
  const value = React.useMemo<NavigationContextType>(() => {
    const setNavigationGuard = (guard: (() => Promise<boolean>) | null): void => {
      guardRef.current = guard;
    };

    const navigate = (page: Page, eventId?: string, intent?: NavIntent): void => {
      // v17.3: Wenn eine Page einen Guard registriert hat (z.B.
      // EventCreationPage bei unsaved-changes), erst dort bestätigen lassen.
      const proceed = (): void => {
        setHistory(prev => [...prev, { page: currentPage, eventId: selectedEventId, intent: navIntent }]);
        setCurrentPage(page);
        setSelectedEventId(eventId || null);
        setNavIntent(intent);
        guardRef.current = null; // Guard nach erfolgreichem Wegnavigieren räumen.
      };
      if (guardRef.current) {
        guardRef.current().then(ok => { if (ok) proceed(); }).catch(() => { /* abort */ });
      } else {
        proceed();
      }
    };

    const goBack = (): void => {
      const proceed = (): void => {
        if (history.length > 0) {
          const prev = history[history.length - 1];
          setHistory(h => h.slice(0, -1));
          setCurrentPage(prev.page);
          setSelectedEventId(prev.eventId);
          setNavIntent(prev.intent);
          guardRef.current = null;
        }
      };
      if (guardRef.current) {
        guardRef.current().then(ok => { if (ok) proceed(); }).catch(() => { /* abort */ });
      } else {
        proceed();
      }
    };

    const clearIntent = (): void => setNavIntent(undefined);

    return { currentPage, selectedEventId, navIntent, navigate, goBack, clearIntent, setNavigationGuard };
  }, [currentPage, selectedEventId, navIntent, history]);

  return React.createElement(
    NavigationContext.Provider,
    { value },
    props.children
  );
}

export function useNavigation(): NavigationContextType {
  const ctx = React.useContext(NavigationContext);
  if (!ctx) throw new Error('useNavigation must be used within NavigationProvider');
  return ctx;
}
