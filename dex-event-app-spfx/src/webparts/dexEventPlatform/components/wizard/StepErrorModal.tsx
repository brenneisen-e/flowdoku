/* v32.0.6: Hinweis-Dialog, wenn „Weiter", ein späterer Schritt oben oder
 * „Event erstellen" an einem leeren Pflichtfeld scheitert.
 *
 * Nutzer-Befund 28.09.2026: Wer unten gescrollt auf „Weiter" klickte, sah
 * nichts passieren — die rote Markierung stand oben außerhalb des Bildes.
 * Der Dialog nennt, WAS fehlt, und „Zum Feld" scrollt zur ersten roten
 * Markierung (errorBorderStyle setzt inline `--dex-red`). Die Fehler-
 * Schlüssel kommen aus getStepErrorsForImpl (wizardMisc.ts) — wer dort einen
 * neuen ergänzt, trägt hier den Text nach, sonst erscheint der Schlüssel. */
import * as React from 'react';
import Modal from '../Modal';
import { AlertCircle } from '../Icons';

export interface StepErrorState { step: number; errs: string[] }

const TEXTE: Record<string, { de: string; en: string }> = {
  title: { de: 'Titel des Events („Wie heißt das Event?")', en: 'Event title' },
  startDate: { de: 'Beginn des Events', en: 'Event start' },
  description: { de: 'Beschreibung — schreib eine oder schalte „Beschreibung anzeigen" aus', en: 'Description — write one or switch off “Show a description”' },
  endDate: { de: 'Ende des Events', en: 'Event end' },
  endBeforeStart: { de: 'Das Ende liegt vor dem Beginn', en: 'The end is before the start' },
  subEventEndBeforeStart: { de: 'Bei einem Sub-Event liegt das Ende vor dem Beginn', en: 'A sub-event ends before it starts' },
  organizer: { de: 'Mindestens ein Organizer', en: 'At least one organizer' },
  deadlineAfterStart: { de: 'Die Anmeldefrist liegt nach dem Event-Beginn', en: 'The registration deadline is after the event start' },
  deregAfterStart: { de: 'Die Abmeldefrist liegt nach dem Event-Beginn', en: 'The cancellation deadline is after the event start' },
  maxParticipants: { de: 'Anzahl der Plätze', en: 'Number of seats' },
};

/** Zur ersten roten Markierung im Wizard scrollen und das Feld fokussieren. */
export function springeZumFehlerfeld(root: HTMLElement | null): void {
  window.setTimeout(() => {
    const scope: ParentNode = root || document;
    const el = scope.querySelector('[style*="--dex-red"]') as HTMLElement | null;
    if (!el) return;
    try { el.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch { el.scrollIntoView(); }
    const feld = (el.matches('input, textarea, select') ? el : el.querySelector('input, textarea, select')) as HTMLElement | null;
    if (feld) { try { feld.focus({ preventScroll: true }); } catch { feld.focus(); } }
  }, 120);
}

export function StepErrorModal(p: {
  state: StepErrorState | null;
  stepLabel: string;
  isDe: boolean;
  onClose: () => void;
  onJump: () => void;
}): React.ReactElement | null {
  const { state, stepLabel, isDe, onClose, onJump } = p;
  if (!state) return null;
  return (
    <Modal
      open
      onClose={onClose}
      maxWidth={460}
      icon={<AlertCircle size={18} />}
      title={isDe ? 'Da fehlt noch etwas' : 'Something is missing'}
      subtitle={isDe ? `Schritt ${state.step + 1} · ${stepLabel}` : `Step ${state.step + 1} · ${stepLabel}`}
      footer={
        <>
          <button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Schließen' : 'Close'}</button>
          <button type="button" className="btn btn-primary" onClick={onJump}>{isDe ? 'Zum Feld' : 'Go to field'}</button>
        </>
      }
    >
      <p style={{ margin: '0 0 8px' }}>
        {isDe ? 'Bevor es weitergeht, bitte noch ausfüllen bzw. korrigieren:' : 'Please fill in or correct before continuing:'}
      </p>
      <ul style={{ margin: 0, paddingLeft: 20 }}>
        {state.errs.map(k => (
          <li key={k} style={{ marginBottom: 4 }}>{TEXTE[k] ? (isDe ? TEXTE[k].de : TEXTE[k].en) : k}</li>
        ))}
      </ul>
    </Modal>
  );
}
