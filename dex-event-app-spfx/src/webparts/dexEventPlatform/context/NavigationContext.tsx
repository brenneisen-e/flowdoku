/**
 * Navigation Context - ersetzt react-router für SPFx
 *
 * SPFx WebParts können kein Browser-Routing verwenden,
 * deshalb wird hier alles ueber State gesteuert.
 * Der History-Stack ermöglicht die Zurück-Navigation.
 */

import * as React from 'react';
import { deepLinkParams } from '../utils/deepLink';

export type Page = 'landing' | 'start' | 'register' | 'registration' | 'my-events' | 'assistant' | 'create-event' | 'edit-event' | 'settings' | 'profile' | 'admin' | 'admin-hub' | 'role-matrix' | 'participants' | 'flowcharts' | 'check-in' | 'self-checkin-display' | 'help' | 'manual' | 'email-templates' | 'tickets' | 'architecture' | 'stats-archive';

// v27.12 (Feedback Datenschutz-Review): Beim Seiten-Refresh landete man immer
// wieder auf der Startseite — die Navigation lebt nur im React-State. Jetzt
// wird die zuletzt besuchte Seite (+ Event) pro Browser-Tab in sessionStorage
// gemerkt und beim Boot wiederhergestellt. Deep-Links (#action=…) behalten
// Vorrang: steht ein action-Parameter in der URL, wird NICHT restauriert
// (der Deep-Link-Effekt in DexEventPlatform navigiert ohnehin gleich weiter).
const NAV_STORAGE_KEY = 'dex-nav-state';
const ALL_PAGES: Page[] = ['landing', 'start', 'register', 'registration', 'my-events', 'assistant', 'create-event', 'edit-event', 'settings', 'profile', 'admin', 'admin-hub', 'role-matrix', 'participants', 'flowcharts', 'check-in', 'self-checkin-display', 'help', 'manual', 'email-templates', 'tickets', 'architecture', 'stats-archive'];

function readStoredNav(): { page: Page; eventId: string | null } | null {
  try {
    if (deepLinkParams().get('action')) return null;
    const raw = window.sessionStorage.getItem(NAV_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as { page?: string; eventId?: string | null };
    if (parsed && typeof parsed.page === 'string' && ALL_PAGES.indexOf(parsed.page as Page) >= 0) {
      return { page: parsed.page as Page, eventId: parsed.eventId || null };
    }
  } catch { /* sessionStorage nicht verfügbar / kaputter Eintrag */ }
  return null;
}

function storeNav(page: Page, eventId: string | null): void {
  try {
    window.sessionStorage.setItem(NAV_STORAGE_KEY, JSON.stringify({ page, eventId }));
  } catch { /* best-effort */ }
}

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
  // v27.12: zuletzt besuchte Seite pro Tab wiederherstellen (s. readStoredNav).
  const restoredNav = React.useRef(readStoredNav()).current;
  const [currentPage, setCurrentPage] = React.useState<Page>(restoredNav ? restoredNav.page : 'landing');
  const [selectedEventId, setSelectedEventId] = React.useState<string | null>(restoredNav ? restoredNav.eventId : null);
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
        storeNav(page, eventId || null); // v27.12: Refresh-Restore
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
          storeNav(prev.page, prev.eventId); // v27.12: Refresh-Restore
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
