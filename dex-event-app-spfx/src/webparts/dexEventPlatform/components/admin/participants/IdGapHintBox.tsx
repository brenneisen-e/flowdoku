/* IdGapHintBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10553-10630 des
 * Stands vor dem Schnitt). Die Rechenlogik (Phasen, Karenzzeit) ist unverändert.
 *
 * v31.3: Nach docs/ui-leitfaden.md 5a (Punkt 2) umgebaut — ein
 * `dex-ui-callout--warn` mit der Folge im ersten Satz und beiden Knöpfen IN
 * der Zeile (Leitfaden 2a′). Die lange Erklärung (Flow-Laufzeit,
 * Gruppenwechsel, Run history) steckt hinter einem Aufklapper: Sie wird nur
 * beim Nachforschen gebraucht, stand aber als sechs Zeilen Fließtext über
 * der Teilnehmerliste. Titel, Erklärung und der Prüf-Knopf gibt es jetzt
 * auch auf Englisch — sie waren als einzige Stelle hier nur deutsch.
 */
import * as React from 'react';
import { formatDate } from '../../../utils/eventStatus';
import { ChevronDown, Hash, RefreshCw } from '../../Icons';
import { cx, ensureDexUiStyles } from '../../dexUi';
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
          // v31.3: Hook vor jedem frühen Return (rules-of-hooks ist `error`);
          // ensureDexUiStyles ist idempotent und nötig, weil der Kasten ohne
          // Modal/WizardFormShell auf der Seite steht.
          const [detailsOpen, setDetailsOpen] = React.useState(false);
          ensureDexUiStyles();
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
            <div className="dex-ui-callout dex-ui-callout--warn" style={{ margin: '0 0 16px' }}>
              <span className="dex-ui-callout-icon" aria-hidden="true"><Hash size={16} /></span>
              <div style={{ minWidth: 0, flex: 1 }}>
                <div style={{ fontWeight: 700 }}>
                  {probablyStillRunning
                    ? (isDe ? 'TeilnehmerIDs haben Lücken — die automatische Korrektur läuft vermutlich noch' : 'Participant IDs have gaps — the automatic fix is probably still running')
                    : (isDe ? 'TeilnehmerIDs haben Lücken — bitte einmal korrigieren' : 'Participant IDs have gaps — please fix them once')}
                </div>
                {/* v31.3: Zwei sichtbare Zeilen mit der Folge, alles Weitere im
                    Aufklapper (Leitfaden 1.2 und 2c). */}
                <div style={{ marginTop: 2 }}>
                  {probablyStillRunning
                    ? (isDe
                      ? <>Die Liste lädt sich hier alle 30 Sekunden neu, der Hinweis verschwindet dann von selbst. Bitte jetzt <strong>nicht</strong> zusätzlich von Hand korrigieren.</>
                      : <>This list reloads every 30 seconds, so the hint disappears on its own. Please do <strong>not</strong> fix anything manually right now.</>)
                    : (isDe
                      ? <>Die Lücke ist rein kosmetisch — Nachrücken und Check-in laufen weiter. <strong>IDs jetzt korrigieren</strong> vergibt 1, 2, 3, … neu.</>
                      : <>The gap is cosmetic only — promotion and check-in keep working. <strong>Fix IDs now</strong> reassigns 1, 2, 3, ….</>)}
                </div>
                {/* v31.3: Beide Knöpfe stehen in der Zeile beim Text (Leitfaden
                    2a′) — vorher hingen sie mit eigenen Rändern darunter.
                    v22.12: „Jetzt neu prüfen" lädt die Liste neu. */}
                <div className="dex-ui-inline" style={{ marginTop: 10 }}>
                  <button type="button" className="btn btn-secondary dex-ui-btn-sm" disabled={idRecheckBusy}
                    onClick={() => { reloadRegistrationsForIdCheck().catch(() => { /* */ }); }}>
                    <RefreshCw size={14} /> {idRecheckBusy ? (isDe ? 'Prüft…' : 'Checking…') : (isDe ? 'Jetzt neu prüfen' : 'Check again now')}
                  </button>
                  {/* v18.60: Direkter Korrektur-Button in der Box — Admin ODER
                      Organizer des Events. Use-Case: die automatische Batch-
                      Korrektur ist offensichtlich NICHT gelaufen (IDs seit längerem
                      falsch). Vorher musste der Organizer den Button im Aktionen-
                      Dropdown suchen. */}
                  {(isAdmin || isOrganizerFor(selectedEvent)) && !!selectedEvent.subsiteUrl && (
                    <button type="button" className="btn btn-primary dex-ui-btn-sm" disabled={isReorderingIDs}
                      onClick={async () => {
                        if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                        if (!(await confirmDialog(isDe
                          ? 'TeilnehmerIDs jetzt neu vergeben (1, 2, 3, …)?\n\nNur klicken, wenn die automatische Korrektur offensichtlich nicht gelaufen ist (IDs schon länger falsch) — NICHT mitten in einer Anmeldewelle.'
                          : 'Reassign participant IDs now (1, 2, 3, …)?\n\nOnly click if the automatic correction clearly did not run (IDs wrong for a while) — NOT in the middle of a registration wave.'))) return;
                        await runIdReorder();
                      }}>
                      <Hash size={14} /> {isReorderingIDs ? (isDe ? 'IDs werden korrigiert…' : 'Fixing IDs…') : (isDe ? 'IDs jetzt korrigieren' : 'Fix IDs now')}
                    </button>
                  )}
                </div>
                <button type="button" className={cx('dex-ui-disclosure', detailsOpen && 'is-open')} aria-expanded={detailsOpen} onClick={() => setDetailsOpen(o => !o)}>
                  <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                  {isDe ? 'Was genau fehlt — und warum' : 'What exactly is missing — and why'}
                </button>
                {detailsOpen && (
                  <div className="dex-ui-disclosure-body">
                    <div>
                      <strong>{isDe ? 'Geprüft an der geladenen Teilnehmerliste:' : 'Checked against the loaded attendee list:'}</strong>{' '}
                      {info.detail}.{whenStr ? <> {isDe ? 'Letzte Abmeldung:' : 'Last cancellation:'} <strong>{whenStr}</strong>.</> : ''}
                    </div>
                    <div style={{ marginTop: 6 }}>
                      {probablyStillRunning
                        ? (isDe
                          ? <>Die automatische Korrektur — <strong>Nachrücken von der Warteliste</strong> und <strong>Neu-Nummerierung</strong> — übernimmt der Nachrück-Flow und braucht nach einer Abmeldung typischerweise 1–5 Minuten; seit der Abmeldung sind schon {minutesSinceCancel} Minuten vergangen. Zwei Korrekturen gleichzeitig laufen ineinander — deshalb hier abwarten.</>
                          : <>The automatic fix — <strong>promoting from the waitlist</strong> and <strong>renumbering</strong> — is done by the promotion flow and usually takes 1–5 minutes after a cancellation; {minutesSinceCancel} minutes have passed already. Two corrections at once collide — so wait here.</>)
                        : (isDe
                          ? <>Die letzte Abmeldung liegt länger zurück. Die Lücke kann trotzdem frisch sein: Ein <strong>Gruppenwechsel</strong> auf die Warteliste vergibt eine neue Nummer und lässt die alte leer, setzt aber kein Abmeldedatum. Und wenn der Nachrück-Flow <strong>gestaut oder ausgefallen</strong> ist, kommt die automatische Korrektur verspätet oder gar nicht (Admin: <strong>Run history</strong> des Flows prüfen).</>
                          : <>The last cancellation was a while ago. The gap can still be fresh: a <strong>group change</strong> onto the waitlist assigns a new number and leaves the old one empty, but sets no cancellation date. And if the promotion flow is <strong>backed up or has failed</strong>, the automatic fix arrives late or never (admin: check the flow&apos;s <strong>Run history</strong>).</>)}
                    </div>
                  </div>
                )}
              </div>
            </div>
          );
};

