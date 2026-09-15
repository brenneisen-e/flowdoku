/**
 * v31.60: Der Entwurf der Event-Erstellung (localStorage) — lesbar auch
 * außerhalb des Assistenten. Nutzer-Ansage 15.09.2026: „wenn ich einen
 * Entwurf habe, dann soll es einen Button geben ‚Entwurf weiter bearbeiten',
 * und wenn ich draufklicke, kommt ein Modal, wo ich den Entwurf anklicken
 * kann." Bis dahin sah man den Entwurf erst als Kachel in Schritt 1, nachdem
 * man „Neues Event erstellen" geklickt hatte — wer nicht wusste, dass ein
 * Entwurf liegt, fand ihn nicht.
 *
 * Schlüssel und Regeln (Substanz, 14 Tage) sind dieselben wie in
 * EventCreationPage — wer sie dort ändert, ändert sie hier mit.
 */
export const EVENT_DRAFT_KEY = 'dex_event_creation_draft_v1';
export const EVENT_DRAFT_MAX_AGE_MS = 14 * 86400000;

export interface EventDraftInfo {
  savedAt: number;
  title: string;
  /** 0-basierter Wizard-Schritt, wie gespeichert. */
  step: number;
  subEventCount: number;
  location: string;
  startDate: string;
  /** v31.61: Nutzungsbedingungen für diesen Entwurf schon bestätigt. */
  tcAccepted: boolean;
}

/** Liest den Entwurf; null = keiner, ohne Substanz oder älter als 14 Tage. */
export function readEventDraft(): EventDraftInfo | null {
  try {
    const raw = localStorage.getItem(EVENT_DRAFT_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as { savedAt?: number; data?: Record<string, unknown> };
    const data = parsed?.data;
    if (!data) return null;
    const title = typeof data.title === 'string' ? data.title.trim() : '';
    const subs = Array.isArray(data.subEvents) ? data.subEvents.length : 0;
    const age = Date.now() - (parsed?.savedAt || 0);
    if ((!title && subs === 0) || !(age >= 0) || age > EVENT_DRAFT_MAX_AGE_MS) return null;
    return {
      savedAt: parsed?.savedAt || 0,
      title,
      step: typeof data.currentStep === 'number' ? data.currentStep : 0,
      subEventCount: subs,
      location: typeof data.location === 'string' ? data.location : '',
      startDate: typeof data.startDate === 'string' ? data.startDate : '',
      tcAccepted: data.tcAccepted === true,
    };
  } catch {
    return null;
  }
}

export function clearEventDraft(): void {
  try { localStorage.removeItem(EVENT_DRAFT_KEY); } catch { /* best-effort */ }
}
