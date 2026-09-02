/* HotelPlanningSection — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 9570-9621 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import HotelPlanningPanel from '../../HotelPlanningPanel';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

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
        const HOTEL_LABEL = /hotel|unterkunft|übernacht|uebernacht|accommodation|lodging/i;
        const asksForHotel = (ev: { eventSpecificFields?: Array<{ type?: string; label?: string; labelEn?: string }> }): boolean =>
          (ev.eventSpecificFields || []).some(f =>
            f.type === 'daterange' || HOTEL_LABEL.test(`${f.label || ''} ${f.labelEn || ''}`));
        const planningStarted = (selectedEvent.hotels || []).length > 0
          || registrations.some(r => (r.Hotel || '').trim());
        if (!planningStarted && !asksForHotel(selectedEvent) && !childEventsOf(selectedEvent.id).some(asksForHotel)) return null;
        return (
        <div className="card" style={{ marginBottom: 16 }}>
          <button
            type="button"
            onClick={() => setHotelPanelOpen(o => !o)}
            style={{ width: '100%', textAlign: 'left', background: 'none', border: 'none', cursor: 'pointer', padding: '4px 0', display: 'flex', alignItems: 'center', gap: 10 }}
          >
            <span style={{ transform: hotelPanelOpen ? 'rotate(90deg)' : 'none', transition: 'transform .15s', color: 'var(--dex-gray-400)' }}>▶</span>
            <span style={{ fontWeight: 700, fontSize: '1rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Hotels & Übernachtungen' : 'Hotels & accommodation'}
            </span>
            {(selectedEvent.hotels || []).length > 0 && (
              <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                {(selectedEvent.hotels || []).length} {isDe ? 'Hotels' : 'hotels'} ·{' '}
                {registrations.filter(r => (r.Hotel || '').trim()).length} {isDe ? 'zugeordnet' : 'assigned'}
              </span>
            )}
            {selectedEvent.hotelVisibleToAttendees && (
              <span style={{ fontSize: '0.68rem', fontWeight: 700, padding: '2px 8px', borderRadius: 999, background: 'rgba(134,188,37,0.15)', color: 'var(--dex-green-dark, #4a7c1f)' }}>
                {isDe ? 'für Teilnehmer sichtbar' : 'visible to attendees'}
              </span>
            )}
          </button>
          {hotelPanelOpen && (
            <div style={{ marginTop: 12 }}>
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

