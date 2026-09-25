/**
 * Navigation Context - ersetzt react-router für SPFx
 *
 * SPFx WebParts können kein Browser-Routing verwenden,
 * deshalb wird hier alles über State gesteuert.
 * Der History-Stack ermöglicht die Zurück-Navigation.
 *
 * v31.59: „Zurück" heißt jetzt wirklich zurück (Nutzer-Ansage 15.09.2026:
 * „ich habe das Gefühl, dass es nicht immer sinnvoll ist, wohin ich zurück-
 * komme"). Der Befund des Audits: Der Stack existierte seit jeher, aber der
 * Zurück-Knopf im Kopf ignorierte ihn und sprang IMMER auf die Startseite —
 * von der Rollenmatrix, vom Check-in, aus dem Feedback, überall. Dazu füllten
 * Rechte-Guards und die Tour den Stack mit Seiten, die nie jemand gesehen
 * hat. Jetzt: (1) `goBack()` fällt bei leerem Stack (Deep-Link, F5) auf ein
 * Ziel je Seite zurück statt nichts zu tun, (2) `navigate(..., { replace })`
 * für Umleitungen, die kein Rücksprungziel sein dürfen, (3) kein Eintrag bei
 * Navigation auf dieselbe Seite, (4) Stack auf 20 begrenzt und pro Tab in
 * sessionStorage, damit F5 den Rückweg nicht kappt.
 */

import * as React from 'react';
import { deepLinkParams } from '../utils/deepLink';

export type Page = 'landing' | 'start' | 'register' | 'registration' | 'my-events' | 'assistant' | 'create-event' | 'edit-event' | 'settings' | 'profile' | 'admin' | 'admin-hub' | 'role-matrix' | 'participants' | 'flowcharts' | 'check-in' | 'self-checkin-display' | 'help' | 'manual' | 'email-templates' | 'tickets' | 'architecture' | 'stats-archive' | 'feedback-overview' | 'intro-onepager' | 'fa-center';

// v27.12 (Feedback Datenschutz-Review): Beim Seiten-Refresh landete man immer
// wieder auf der Startseite — die Navigation lebt nur im React-State. Jetzt
// wird die zuletzt besuchte Seite (+ Event) pro Browser-Tab in sessionStorage
// gemerkt und beim Boot wiederhergestellt. Deep-Links (#action=…) behalten
// Vorrang: steht ein action-Parameter in der URL, wird NICHT restauriert
// (der Deep-Link-Effekt in DexEventPlatform navigiert ohnehin gleich weiter).
const NAV_STORAGE_KEY = 'dex-nav-state';
// v31.59: 'fa-center' fehlte — ein Refresh im F&A Center warf auf die Landing.
const ALL_PAGES: Page[] = ['landing', 'start', 'register', 'registration', 'my-events', 'assistant', 'create-event', 'edit-event', 'settings', 'profile', 'admin', 'admin-hub', 'role-matrix', 'participants', 'flowcharts', 'check-in', 'self-checkin-display', 'help', 'manual', 'email-templates', 'tickets', 'architecture', 'stats-archive', 'feedback-overview', 'intro-onepager', 'fa-center'];
const HISTORY_MAX = 20;

interface HistoryEntry {
  page: Page;
  eventId: string | null;
  intent: NavIntent;
}

function readStoredNav(): { page: Page; eventId: string | null; history: HistoryEntry[] } | null {
  try {
    if (deepLinkParams().get('action')) return null;
    const raw = window.sessionStorage.getItem(NAV_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as { page?: string; eventId?: string | null; history?: Array<{ page?: string; eventId?: string | null }> };
    if (parsed && typeof parsed.page === 'string' && ALL_PAGES.indexOf(parsed.page as Page) >= 0) {
      const history: HistoryEntry[] = Array.isArray(parsed.history)
        ? parsed.history
          .filter(h => h && typeof h.page === 'string' && ALL_PAGES.indexOf(h.page as Page) >= 0)
          .map(h => ({ page: h.page as Page, eventId: h.eventId || null, intent: undefined }))
        : [];
      return { page: parsed.page as Page, eventId: parsed.eventId || null, history };
    }
  } catch { /* sessionStorage nicht verfügbar / kaputter Eintrag */ }
  return null;
}

function storeNav(page: Page, eventId: string | null, history: HistoryEntry[]): void {
  try {
    // Absichten (intent) werden nicht gespeichert — sie gehören zum Klick,
    // nicht zur Seite; nach F5 wäre „für andere anmelden" ein Fehlstart.
    window.sessionStorage.setItem(NAV_STORAGE_KEY, JSON.stringify({
      page, eventId, history: history.slice(-HISTORY_MAX).map(h => ({ page: h.page, eventId: h.eventId })),
    }));
  } catch { /* best-effort */ }
}

/**
 * v31.59: Wohin „Zurück" führt, wenn es keinen Rückweg gibt (Deep-Link aus
 * einer Mail, F5, erster Aufruf). Die Regel: dorthin, von wo man diese Seite
 * normalerweise öffnet — Werkzeuge eines Events ins Organizer Center DIESES
 * Events, Admin-Seiten in den Admin-Hub, die Anmeldeseite in die Übersicht.
 */
export function fallbackFor(page: Page, eventId: string | null): { page: Page; eventId?: string } {
  switch (page) {
    case 'landing': return { page: 'landing' };
    case 'start': return { page: 'landing' };
    case 'check-in':
    case 'participants':
    case 'self-checkin-display':
    case 'edit-event':
      return eventId ? { page: 'admin', eventId } : { page: 'admin' };
    case 'create-event':
      return { page: 'admin' };
    case 'settings':
    case 'role-matrix':
    case 'email-templates':
    case 'architecture':
    case 'stats-archive':
    case 'feedback-overview':
    case 'intro-onepager':
    case 'flowcharts':
      return { page: 'admin-hub' };
    case 'registration':
      return { page: 'register' };
    default:
      return { page: 'start' };
  }
}

// Optionale Absicht beim Navigieren (z.B. Registration-Seite direkt im "Für andere"-Modus öffnen)
// v31.9.3: 'open-comms' kommt aus dem Deep-Link `?action=comms&event=<Nr>`
// im Hinweis „Bereits versendete Infos zu diesem Event" und oeffnet in
// „Meine Events" direkt die Nachrichten dieses Events.
// v31.60: 'resume-draft' öffnet die Event-Erstellung und wendet den
// gespeicherten Entwurf sofort an (Knopf „Entwurf weiter bearbeiten").
// v31.98: 'open-teilnehmer' / 'open-concur' / 'open-fa' kommen aus den drei
// Listen-Links der Danke-Mail (`#action=admin&event=<Id>&open=…`) und öffnen
// im Organizer Center den passenden Dialog.
export type NavIntent = 'register-other' | 'auto-cancel' | 'open-comms' | 'resume-draft' | 'open-teilnehmer' | 'open-concur' | 'open-fa' | undefined;

export interface NavigateOptions {
  /** v31.59: Seite ersetzen statt anhängen — für Rechte-Umleitungen und die
   *  Tour. Die verlassene Seite wird dann NICHT zum Rücksprungziel. */
  replace?: boolean;
}

interface NavigationContextType {
  currentPage: Page;
  selectedEventId: string | null;
  navIntent: NavIntent;
  navigate: (page: Page, eventId?: string, intent?: NavIntent, opts?: NavigateOptions) => void;
  /** Eine Seite zurück; ohne Rückweg zum Fallback (Parameter oder `fallbackFor`). */
  goBack: (fallback?: { page: Page; eventId?: string }) => void;
  /** v31.59: Gibt es einen echten Rückweg (Stack nicht leer)? */
  canGoBack: boolean;
  clearIntent: () => void;
  /** v17.3: Page registriert einen Confirm-Hook für ungespeicherte
   *  Änderungen. Wird VOR jeder Navigation aufgerufen — wenn er false
   *  zurückliefert, blockiert das Navigation. Null = keine Änderungen,
   *  durchnavigieren. */
  setNavigationGuard: (guard: (() => Promise<boolean>) | null) => void;
}

// Exportiert, damit Preview-Wrapper im Handbuch den Context mit Demo-Daten
// überschreiben können (v6.27 App-Screenshots).
export const NavigationContext = React.createContext<NavigationContextType | undefined>(undefined);

export function NavigationProvider(props: { children: React.ReactNode }): React.ReactElement {
  // v27.12: zuletzt besuchte Seite pro Tab wiederherstellen (s. readStoredNav).
  const restoredNav = React.useRef(readStoredNav()).current;
  const [currentPage, setCurrentPage] = React.useState<Page>(restoredNav ? restoredNav.page : 'landing');
  const [selectedEventId, setSelectedEventId] = React.useState<string | null>(restoredNav ? restoredNav.eventId : null);
  const [navIntent, setNavIntent] = React.useState<NavIntent>(undefined);
  const [history, setHistory] = React.useState<HistoryEntry[]>(restoredNav ? restoredNav.history : []);
  const guardRef = React.useRef<(() => Promise<boolean>) | null>(null);

  // v20.0 (Audit): Context-Value memoizen — die Funktionen schließen über
  // currentPage/selectedEventId/navIntent/history, daher sind genau diese
  // vier die Memo-Abhängigkeiten. Verhindert App-weite Re-Renders aller
  // useNavigation()-Consumer bei Parent-Re-Renders ohne Navigations-Änderung.
  const value = React.useMemo<NavigationContextType>(() => {
    const setNavigationGuard = (guard: (() => Promise<boolean>) | null): void => {
      guardRef.current = guard;
    };

    const navigate = (page: Page, eventId?: string, intent?: NavIntent, opts?: NavigateOptions): void => {
      // v17.3: Wenn eine Page einen Guard registriert hat (z.B.
      // EventCreationPage bei unsaved-changes), erst dort bestätigen lassen.
      const proceed = (): void => {
        const same = page === currentPage && (eventId || null) === selectedEventId;
        let next = history;
        if (!opts?.replace && !same) {
          // Dieselbe Seite nicht doppelt stapeln (Kachel → Kachel → Kachel
          // hieß sonst dreimal „Zurück" auf denselben Hub).
          next = [...history, { page: currentPage, eventId: selectedEventId, intent: navIntent }].slice(-HISTORY_MAX);
          setHistory(next);
        }
        setCurrentPage(page);
        setSelectedEventId(eventId || null);
        setNavIntent(intent);
        storeNav(page, eventId || null, next); // v27.12: Refresh-Restore
        guardRef.current = null; // Guard nach erfolgreichem Wegnavigieren räumen.
      };
      if (guardRef.current) {
        guardRef.current().then(ok => { if (ok) proceed(); }).catch(() => { /* abort */ });
      } else {
        proceed();
      }
    };

    const goBack = (fallback?: { page: Page; eventId?: string }): void => {
      const proceed = (): void => {
        if (history.length > 0) {
          const prev = history[history.length - 1];
          const next = history.slice(0, -1);
          setHistory(next);
          setCurrentPage(prev.page);
          setSelectedEventId(prev.eventId);
          setNavIntent(prev.intent);
          storeNav(prev.page, prev.eventId, next); // v27.12: Refresh-Restore
          guardRef.current = null;
          return;
        }
        // v31.59: Kein Rückweg (Deep-Link, F5) — bis hierher passierte NICHTS,
        // der Knopf war tot. Jetzt das Ziel, von dem aus man diese Seite
        // normalerweise öffnet; ohne Eintrag im Stack.
        const fb = fallback || fallbackFor(currentPage, selectedEventId);
        setCurrentPage(fb.page);
        setSelectedEventId(fb.eventId || null);
        setNavIntent(undefined);
        storeNav(fb.page, fb.eventId || null, []);
        guardRef.current = null;
      };
      if (guardRef.current) {
        guardRef.current().then(ok => { if (ok) proceed(); }).catch(() => { /* abort */ });
      } else {
        proceed();
      }
    };

    const clearIntent = (): void => setNavIntent(undefined);

    return { currentPage, selectedEventId, navIntent, navigate, goBack, canGoBack: history.length > 0, clearIntent, setNavigationGuard };
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
