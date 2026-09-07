/**
 * v31.2: „Wer bearbeitet dieses Event gerade auch?" — Herzschlag und Abfrage
 * für den Wizard (Liste DEX_EditPresence, s. services/events/editPresence).
 *
 * Läuft nur im Edit-Modus eines bestehenden Events: Beim Anlegen gibt es
 * noch nichts, was zwei Personen gleichzeitig bearbeiten könnten. Alle 20 s
 * ein Takt: eigene Zeile stempeln, fremde frische Zeilen lesen. Beim
 * Verlassen (Unmount, Event-Wechsel, Tab-Wechsel per pagehide) wird die
 * eigene Zeile gelöscht — best-effort; bleibt sie stehen, fällt sie nach
 * PRESENCE_FRESH_MS ohnehin aus der Anzeige.
 *
 * Ein Lesefehler ist keine Null (CLAUDE.md): Liefert die Liste nichts
 * Lesbares, bleibt der letzte bekannte Stand stehen, statt „niemand da" zu
 * behaupten.
 */
import * as React from 'react';
import { EventService } from '../../../services/EventService';
import type { EditPresenceEntry } from '../../../services/events/editPresence';

const TICK_MS = 20 * 1000;

export interface UseEditPresenceArgs {
  enabled: boolean;
  eventId: string;
  email: string;
  name: string;
}

export function useEditPresence({ enabled, eventId, email, name }: UseEditPresenceArgs): EditPresenceEntry[] {
  const [others, setOthers] = React.useState<EditPresenceEntry[]>([]);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const context = (window as any).__dexSpfxContext;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const svc = React.useMemo(() => context ? new EventService(context) : null, []);
  const ownIdRef = React.useRef<number | null>(null);
  const nameRef = React.useRef(name);
  nameRef.current = name;

  React.useEffect(() => {
    if (!enabled || !svc || !eventId || !email) { setOthers([]); return undefined; }
    let alive = true;
    ownIdRef.current = null;
    const tick = async (): Promise<void> => {
      const id = await svc.heartbeatEditPresence(eventId, email, nameRef.current, ownIdRef.current);
      if (!alive) return;
      if (id !== null) ownIdRef.current = id;
      const list = await svc.readEditPresence(eventId, email);
      if (!alive || list === null) return;
      setOthers(list);
    };
    void tick();
    const timer = window.setInterval(() => { void tick(); }, TICK_MS);
    const clear = (): void => {
      const id = ownIdRef.current;
      if (id !== null) { ownIdRef.current = null; void svc.clearEditPresence(id); }
    };
    window.addEventListener('pagehide', clear);
    return () => {
      alive = false;
      window.clearInterval(timer);
      window.removeEventListener('pagehide', clear);
      clear();
    };
  }, [enabled, svc, eventId, email]);

  return others;
}
