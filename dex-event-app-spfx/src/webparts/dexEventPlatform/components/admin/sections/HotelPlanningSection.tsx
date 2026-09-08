/* HotelPlanningSection — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 9570-9621 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich übernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3: Nur der Rahmen um `HotelPlanningPanel` — Kopfzeile nach
 * docs/ui-leitfaden.md 5a Punkt 8 (Auswertungen zuletzt, eingeklappt, Kopf mit
 * Zähler und Chevron rechts). Das Panel selbst ist eine eigene Datei und
 * bleibt unangetastet.
 */
import * as React from 'react';
import HotelPlanningPanel from '../../HotelPlanningPanel';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';
import { ChevronDown, Pin } from '../../Icons';
import { cx, ensureDexUiStyles } from '../../dexUi';

export interface HotelPlanningSectionProps {
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
  hotelPanelOpen: boolean;
  isDe: boolean;
  refreshEvents: () => Promise<void>;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setHotelPanelOpen: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
}

export const HotelPlanningSection: React.FC<HotelPlanningSectionProps> = (p) => {
  const { childEventsOf, confirmDialog, hotelPanelOpen, isDe, refreshEvents, registrations, reloadRegistrations, selectedEvent, setHotelPanelOpen, showAlert, subEventRegsByEventId } = p;
        // v31.3: Die gemeinsamen Klassen sicherstellen — das Organizer Center
        // rendert nicht zwangsläufig vorher ein Modal, das sie injiziert.
        ensureDexUiStyles();
        const HOTEL_LABEL = /hotel|unterkunft|übernacht|uebernacht|accommodation|lodging/i;
        const asksForHotel = (ev: { eventSpecificFields?: Array<{ type?: string; label?: string; labelEn?: string }> }): boolean =>
          (ev.eventSpecificFields || []).some(f =>
            f.type === 'daterange' || HOTEL_LABEL.test(`${f.label || ''} ${f.labelEn || ''}`));
        const planningStarted = (selectedEvent.hotels || []).length > 0
          || registrations.some(r => (r.Hotel || '').trim());
        if (!planningStarted && !asksForHotel(selectedEvent) && !childEventsOf(selectedEvent.id).some(asksForHotel)) return null;
        const hotelCount = (selectedEvent.hotels || []).length;
        const assignedCount = registrations.filter(r => (r.Hotel || '').trim()).length;
        return (
        <div className="dex-ui-card" style={{ marginBottom: 16, padding: 12 }}>
          {/* v31.3: Die ganze Kopfzeile ist der Aufklapper (Hover über
              `dex-ui-row`) — vorher lag der Klick auf einem Knopf ohne Hover,
              und das Dreieck stand links vor dem Titel statt rechts.
              span statt h3: Überschriften gehören nicht in einen <button>. */}
          <button type="button" onClick={() => setHotelPanelOpen(o => !o)} aria-expanded={hotelPanelOpen}
            className="dex-ui-rowbtn dex-ui-row dex-ui-card-head">
            <span className="dex-ui-card-head-title">
              <Pin size={16} strokeWidth={2} />
              {isDe ? 'Hotels & Übernachtungen' : 'Hotels & accommodation'}
            </span>
            {hotelCount > 0 && (
              <span className="dex-ui-card-head-meta">
                {hotelCount} {isDe ? 'Hotels' : 'hotels'} · {assignedCount} {isDe ? 'Personen zugeordnet' : 'people assigned'}
              </span>
            )}
            {selectedEvent.hotelVisibleToAttendees && (
              <span className="dex-ui-pill dex-ui-pill--green">
                {isDe ? 'für Teilnehmer sichtbar' : 'visible to attendees'}
              </span>
            )}
            <span className={cx('dex-ui-disclosure-chevron', hotelPanelOpen && 'is-open')} style={{ marginLeft: 'auto' }}>
              <ChevronDown size={18} />
            </span>
          </button>
          {hotelPanelOpen && (
            <div style={{ marginTop: 12, padding: '0 2px 2px' }}>
              <HotelPlanningPanel
                event={selectedEvent}
                registrations={registrations}
                isDe={isDe}
                childEvents={childEventsOf(selectedEvent.id)}
                subEventRegsByEventId={subEventRegsByEventId}
                onReloadRegistrations={async () => {
                  // v30.67 (Review): gemeinsamer Nachlade-Pfad statt `[]` bei 429.
                  await reloadRegistrations();
                }}
                onReloadEvents={async () => { await refreshEvents(); }}
                showAlert={showAlert}
                confirmDialog={confirmDialog}
              />
            </div>
          )}
        </div>
        );
};

