/* IdGapHintBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10553-10630 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { formatDate } from '../../../utils/eventStatus';
import { Hash, RefreshCw } from '../../Icons';
import { DeloitteEvent } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';

export interface IdGapHintBoxProps {
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  eventServiceRef: EventService;
  idRecheckBusy: boolean;
  isAdmin: boolean;
  isDe: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  isReorderingIDs: boolean;
  recentCancellation: (regs: SPRegistration[]) => {    recent: boolean;    whenIso: string;    detail: string;};
  registrations: SPRegistration[];
  reloadRegistrationsForIdCheck: () => Promise<void>;
  runIdReorder: () => Promise<void>;
  selectedEvent: DeloitteEvent;
}

export const IdGapHintBox: React.FC<IdGapHintBoxProps> = (p) => {
  const { confirmDialog, eventServiceRef, idRecheckBusy, isAdmin, isDe, isOrganizerFor, isReorderingIDs, recentCancellation, registrations, reloadRegistrationsForIdCheck, runIdReorder, selectedEvent } = p;
          if (!selectedEvent) return null;
          // v22.67: Im Klammer-Modus („Nur Sub-Events") greift die
          // TeilnehmerID-Durchgängigkeits-Prüfung NICHT — die geprüfte Liste
          // sind die Schatten-Zeilen der Klammer (ohne fortlaufende Nummern);
          // die echten TeilnehmerIDs leben pro Sub-Event. Die Warnung war hier
          // ein Fehlalarm.
          if (selectedEvent.subEventsOnlyMode) return null;
          const info = recentCancellation(registrations);
          if (!info.recent) return null;
          const whenStr = info.whenIso ? formatDate(info.whenIso) : '';
          // v22.12: zweiphasig — innerhalb von ~10 Min nach der letzten
          // Abmeldung läuft die automatische Korrektur evtl. noch (warten);
          // danach ist die Lücke ECHT stehengeblieben (typisch: die höchste
          // Nummer wurde abgemeldet, während gleichzeitig neue Anmeldungen
          // bereits höhere Nummern gezogen haben — ein späterer automatischer
          // Lauf kommt nicht, weil nur Abmeldungen die Korrektur anstoßen).
          const minutesSinceCancel = info.whenIso
            ? Math.floor((Date.now() - new Date(info.whenIso).getTime()) / 60000)
            : 999;
          const probablyStillRunning = minutesSinceCancel >= 0 && minutesSinceCancel < 10;
          // v30.74: In den ersten Minuten nach einer Abmeldung GAR NICHTS
          // zeigen. Die App nummeriert bewusst nicht selbst (seit v6.7 macht
          // das der Flow DEX_IDReorder_TeilnehmerIDs); der Trigger fragt die
          // Queue etwa minütlich ab, dazu die Laufzeit. Die Lücke ist in
          // dieser Phase also der Normalfall — und die orange Box las sich
          // nach jeder Organizer-Abmeldung wie ein Fehler (Nutzer-Ansage
          // 03.09.2026: „erst, wenn ein paar Minuten nichts passiert ist").
          // Die 30-Sekunden-Nachlade-Schleife (useTeamActions) läuft
          // unabhängig von dieser Box weiter; ist der Flow schneller, sieht
          // niemand je eine Box. Ein Gruppenwechsel setzt kein Abmeldedatum
          // → minutesSinceCancel 999 → die Box erscheint wie bisher sofort.
          const ID_GAP_GRACE_MINUTES = 3;
          if (minutesSinceCancel >= 0 && minutesSinceCancel < ID_GAP_GRACE_MINUTES) return null;
          return (
            <div style={{
              margin: '0 0 16px',
              padding: '14px 16px',
              borderRadius: 8,
              background: 'rgba(237,139,0,0.10)',
              border: '1px solid var(--dex-orange, #ed8b00)',
              color: 'var(--dex-orange-dark, #b35a00)',
            }}>
              <div style={{ fontWeight: 700, marginBottom: 6, fontSize: '0.9rem' }}>
                {probablyStillRunning
                  ? 'TeilnehmerIDs sind gerade nicht durchgängig — automatische Korrektur läuft vermutlich noch'
                  : 'TeilnehmerIDs sind nicht durchgängig — bitte einmal korrigieren'}
              </div>
              <div style={{ fontSize: '0.82rem', lineHeight: 1.5 }}>
                <strong>Geprüft an der geladenen Teilnehmerliste:</strong> {info.detail}.{whenStr ? <> Letzte Abmeldung: <strong>{whenStr}</strong>.</> : ''}{' '}
                {probablyStillRunning ? (
                  <>Die automatische Korrektur — <strong>Nachrücken von der Warteliste</strong> und <strong>Neu-Nummerierung</strong> — übernimmt der Nachrück-Flow und braucht nach einer Abmeldung typischerweise 1–5 Minuten; seit der Abmeldung sind schon {minutesSinceCancel} Minuten vergangen. Die Liste wird hier <strong>automatisch alle 30 Sekunden neu geladen</strong>; diese Box verschwindet von selbst, sobald alles stimmt. Bitte in dieser Phase NICHT manuell korrigieren (sonst laufen zwei Korrekturen ineinander).</>
                ) : (
                  <>Die letzte Abmeldung liegt länger zurück. Die Lücke kann trotzdem frisch sein — ein <strong>Gruppenwechsel</strong> auf die Warteliste vergibt eine neue Nummer und lässt die alte leer, setzt aber kein Abmeldedatum; und wenn der Nachrück-Flow <strong>gestaut oder ausgefallen</strong> ist, kommt die automatische Korrektur verspätet oder gar nicht (Admin: <strong>Run history</strong> des Flows prüfen). Die Lücke ist rein kosmetisch (Nachrücken/Check-in funktionieren trotzdem) und <strong>gefahrlos per Klick zu beheben</strong>:</>
                )}
              </div>
              {/* v22.12: manueller Sofort-Check (lädt die Liste neu). */}
              <button
                type="button"
                className="btn btn-secondary"
                disabled={idRecheckBusy}
                style={{ marginTop: 12, marginRight: 10, fontSize: '0.82rem', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                onClick={() => { reloadRegistrationsForIdCheck().catch(() => { /* */ }); }}
              >
                <RefreshCw size={14} /> {idRecheckBusy ? 'Prüft…' : 'Jetzt neu prüfen'}
              </button>
              {/* v18.60: Direkter Korrektur-Button in der Box — Admin ODER
                  Organizer des Events. Use-Case: die automatische Batch-
                  Korrektur ist offensichtlich NICHT gelaufen (IDs seit längerem
                  falsch). Vorher musste der Organizer den Button im Aktionen-
                  Dropdown suchen. */}
              {(isAdmin || isOrganizerFor(selectedEvent)) && !!selectedEvent.subsiteUrl && (
                <button
                  type="button"
                  className="btn btn-primary"
                  disabled={isReorderingIDs}
                  style={{ marginTop: 12, fontSize: '0.82rem', opacity: isReorderingIDs ? 0.6 : 1, display: 'inline-flex', alignItems: 'center', gap: 6 }}
                  onClick={async () => {
                    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                    if (!(await confirmDialog(isDe
                      ? 'TeilnehmerIDs jetzt neu vergeben (1, 2, 3, …)?\n\nNur klicken, wenn die automatische Korrektur offensichtlich nicht gelaufen ist (IDs schon länger falsch) — NICHT mitten in einer Anmeldewelle.'
                      : 'Reassign participant IDs now (1, 2, 3, …)?\n\nOnly click if the automatic correction clearly did not run (IDs wrong for a while) — NOT in the middle of a registration wave.'))) return;
                    await runIdReorder();
                  }}
                >
                  <Hash size={14} /> {isReorderingIDs ? (isDe ? 'IDs werden korrigiert…' : 'Fixing IDs…') : (isDe ? 'IDs jetzt korrigieren' : 'Fix IDs now')}
                </button>
              )}
            </div>
          );
};

