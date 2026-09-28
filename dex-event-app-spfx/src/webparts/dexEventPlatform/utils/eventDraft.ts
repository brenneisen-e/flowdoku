/**
 * v31.60: Der Entwurf der Event-Erstellung (localStorage) — lesbar auch
 * außerhalb des Assistenten. Nutzer-Ansage 15.09.2026: „wenn ich einen
 * Entwurf habe, dann soll es einen Button geben ‚Entwurf weiter bearbeiten',
 * und wenn ich draufklicke, kommt ein Modal, wo ich den Entwurf anklicken
 * kann." Bis dahin sah man den Entwurf erst als Kachel in Schritt 1, nachdem
 * man „Neues Event erstellen" geklickt hatte — wer nicht wusste, dass ein
 * Entwurf liegt, fand ihn nicht.
 *
 * v32.1.4: MEHRERE Entwürfe — je Assistenten-Sitzung ein eigener Platz.
 * Nutzer-Befund 28.09.2026: Entwurf A liegt, neues Event B angefangen,
 * zurück, „Event verwerfen" — und A war auch weg. Es gab genau EINEN Platz:
 * Der Autosave von B überschrieb A schon beim ersten Tippen, das Verwerfen
 * löschte danach den Platz. Seither schreibt jede Sitzung unter
 * `EVENT_DRAFT_PREFIX + id`, und Verwerfen/Anlegen räumt nur den eigenen
 * Platz. Der alte Einzel-Schlüssel wird weiter gelesen (Id `legacy`).
 *
 * Regeln (Substanz, 14 Tage) sind dieselben wie in EventCreationPage — wer
 * sie dort ändert, ändert sie hier mit.
 */
export const EVENT_DRAFT_KEY = 'dex_event_creation_draft_v1';
export const EVENT_DRAFT_PREFIX = 'dex_event_creation_draft_v1__';
export const EVENT_DRAFT_MAX_AGE_MS = 14 * 86400000;
const LEGACY_ID = 'legacy';
const RESUME_KEY = 'dex_resume_draft_id';

export interface EventDraftInfo {
  id: string;
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

/** localStorage-Schlüssel eines Entwurfs. */
export function eventDraftKey(id: string): string {
  return id === LEGACY_ID ? EVENT_DRAFT_KEY : EVENT_DRAFT_PREFIX + id;
}

/** Neue Id für eine frische Assistenten-Sitzung. */
export function newEventDraftId(): string {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

/** Roh-Inhalt eines Entwurfs (für das Anwenden im Assistenten). */
export function readEventDraftRaw(id: string): { savedAt: number; data: Record<string, unknown> } | null {
  try {
    const raw = localStorage.getItem(eventDraftKey(id));
    if (!raw) return null;
    const parsed = JSON.parse(raw) as { savedAt?: number; data?: Record<string, unknown> };
    if (!parsed || !parsed.data) return null;
    return { savedAt: parsed.savedAt || 0, data: parsed.data };
  } catch { return null; }
}

function infoOf(id: string, parsed: { savedAt: number; data: Record<string, unknown> } | null): EventDraftInfo | null {
  if (!parsed) return null;
  const data = parsed.data;
  const title = typeof data.title === 'string' ? data.title.trim() : '';
  const subs = Array.isArray(data.subEvents) ? data.subEvents.length : 0;
  const age = Date.now() - (parsed.savedAt || 0);
  if ((!title && subs === 0) || !(age >= 0) || age > EVENT_DRAFT_MAX_AGE_MS) return null;
  return {
    id,
    savedAt: parsed.savedAt || 0,
    title,
    step: typeof data.currentStep === 'number' ? data.currentStep : 0,
    subEventCount: subs,
    location: typeof data.location === 'string' ? data.location : '',
    startDate: typeof data.startDate === 'string' ? data.startDate : '',
    tcAccepted: data.tcAccepted === true,
  };
}

/** Ein Entwurf nach Id; null = keiner, ohne Substanz oder älter als 14 Tage. */
export function readEventDraftById(id: string | null | undefined): EventDraftInfo | null {
  if (!id) return null;
  return infoOf(id, readEventDraftRaw(id));
}

/** Alle gültigen Entwürfe, neuester zuerst. Abgelaufene werden dabei entfernt. */
export function listEventDrafts(): EventDraftInfo[] {
  const out: EventDraftInfo[] = [];
  try {
    const ids: string[] = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i) || '';
      if (k === EVENT_DRAFT_KEY) ids.push(LEGACY_ID);
      else if (k.indexOf(EVENT_DRAFT_PREFIX) === 0) ids.push(k.slice(EVENT_DRAFT_PREFIX.length));
    }
    for (const id of ids) {
      const raw = readEventDraftRaw(id);
      const info = infoOf(id, raw);
      if (info) out.push(info);
      else if (raw && Date.now() - raw.savedAt > EVENT_DRAFT_MAX_AGE_MS) {
        try { localStorage.removeItem(eventDraftKey(id)); } catch { /* */ }
      }
    }
  } catch { /* localStorage gesperrt — keine Entwürfe */ }
  return out.sort((a, b) => b.savedAt - a.savedAt);
}

/** Neuester Entwurf (Kompatibilität für Aufrufer mit nur einem Entwurf). */
export function readEventDraft(): EventDraftInfo | null {
  return listEventDrafts()[0] || null;
}

/** Einen Entwurf löschen — nie alle. */
export function clearEventDraft(id: string): void {
  try { localStorage.removeItem(eventDraftKey(id)); } catch { /* best-effort */ }
}

/** Übergabe „diesen Entwurf fortsetzen" an den Assistenten (mit `resume-draft`). */
export function setResumeDraftId(id: string): void {
  try { sessionStorage.setItem(RESUME_KEY, id); } catch { /* */ }
}
export function peekResumeDraftId(): string | null {
  try { return sessionStorage.getItem(RESUME_KEY); } catch { return null; }
}
export function clearResumeDraftId(): void {
  try { sessionStorage.removeItem(RESUME_KEY); } catch { /* */ }
}
