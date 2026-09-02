/* KpiTiles — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 9403-9512 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { SPRegistration } from '../../../services/EventService';
import { DeloitteEvent } from '../../../types';
import { countConsolidatedActive } from '../logic/parentRegs';

export interface KpiTilesProps {
  isConsolidatedMode: boolean;
  isSplitCapacity: boolean;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
  t: (key: string) => string;
}

export const KpiTiles: React.FC<KpiTilesProps> = (p) => {
  const { isConsolidatedMode, isSplitCapacity, registrations, selectedEvent, subEventRegsByEventId, t } = p;
        // v30.67: Bei geteilten Kapazitäten ist `maxParticipants` per
        // Definition 0 (die Kapazität steht in durchstarter-/funstarterCapacity).
        // `> 0` sollte nur „Unbegrenzt" ausschließen, schloss aber jedes
        // Split-Event mit aus — 25 Wartende, keine Kachel.
        const hasWaitlistKPI = !!(selectedEvent?.waitlistEnabled && ((selectedEvent?.maxParticipants || 0) > 0 || isSplitCapacity));
        // Fraktionen pro Spalte — Angemeldet bekommt 2fr wenn Split aktiv ist.
        const angeFr = isSplitCapacity ? '2fr' : '1fr';
        const tail = `1fr 1fr${hasWaitlistKPI ? ' 1fr' : ''} 1fr`; // QR / Eingecheckt / [Warteliste] / Abgemeldet
        const gridCols = `${angeFr} ${tail}`;
        // v15.14: Im subEventsOnlyMode (Hauptevent ohne eigene Anmeldungen)
        // beziehen sich die Stat-Cards auf die konsolidierten Teilnehmer
        // über alle Sub-Events. Die Hauptevent-Liste selbst hat hier nur
        // Alt-Daten und würde das echte Bild verfälschen.
        const consolidatedRegs: SPRegistration[] = isConsolidatedMode
          // v22.63: Klammer-eigene Abmeldungen (Absagen auf der Klammer-Subsite)
          // in die KPI „Abgemeldet" mitzählen, damit KPI und Abmeldungs-Liste
          // übereinstimmen. Nur Abgemeldet-Zeilen, um die Aktiv-Zahlen nicht zu
          // verfälschen.
          ? ([] as SPRegistration[]).concat(
              registrations.filter(r => r.Status === 'Abgemeldet'),
              ...Object.values(subEventRegsByEventId),
            )
          : [];
        const consolidatedQRByEmail = new Set<string>();
        const consolidatedCheckedByEmail = new Set<string>();
        const consolidatedWaitlistByEmail = new Set<string>();
        const consolidatedCancelledByEmail = new Set<string>();
        const consolidatedAnyByEmail = new Set<string>();
        for (const r of consolidatedRegs) {
          // v23.3: emaillose Zeile zaehlt als eigener Kopf (Zeilen-Id-Fallback),
          // statt aus den KPIs zu verschwinden — sonst KPI < Tabelle.
          const key = (r.ParticipantEmail || '').toLowerCase().trim() || `__noemail#${r.Id}`;
          consolidatedAnyByEmail.add(key);
          if (r.Status === 'QR versendet') consolidatedQRByEmail.add(key);
          if (r.Status === 'Eingecheckt') consolidatedCheckedByEmail.add(key);
          if (r.Status === 'Warteliste') consolidatedWaitlistByEmail.add(key);
          if (r.Status === 'Abgemeldet') consolidatedCancelledByEmail.add(key);
        }
        const active = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
        // v30.67: Über den gemeinsamen Helfer, damit die Detail-Karte
        // („Aktuell registriert") dieselbe Zahl zeigt — vorher zählte sie die
        // Matrix-Zeilen inklusive Warteliste.
        const totalActive = isConsolidatedMode ? countConsolidatedActive(subEventRegsByEventId) : active.length;
        // v19.12: nach EFFEKTIVER Gruppe zählen (StarterType ODER, falls leer,
        // PreferredStarterType). Sonst fehlen angemeldete Nachrücker, deren
        // StarterType der Flow nicht gesetzt hat, in der Gruppen-Zahl — dann ist
        // „Angemeldet" gesamt ≠ Summe der beiden Gruppen (Beobachtung: 142 vs.
        // 130+11). Eine angemeldete Person mit Wunsch „Funstarter" belegt real
        // einen Funstarter-Platz und muss dort mitgezählt werden.
        const durchActive = active.filter(r => (r.StarterType || r.PreferredStarterType) === 'Durchstarter').length;
        const funActive = active.filter(r => (r.StarterType || r.PreferredStarterType) === 'Funstarter').length;
        const durchCap = selectedEvent?.durchstarterCapacity || 0;
        const funCap = selectedEvent?.funstarterCapacity || 0;
        const labelA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
        const labelB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
        const reversed = !!selectedEvent?.splitDisplayOrderReversed;
        const grpA = (
          <div key="grpA" style={{ display: 'flex', alignItems: 'baseline', gap: 8 }}>
            <span style={{ color: 'var(--dex-green-dark, #6b9a1e)', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={labelA}>● {labelA}</span>
            <strong style={{ whiteSpace: 'nowrap' }}>{durchActive}<span style={{ color: 'var(--dex-gray-400)' }}>/{durchCap}</span></strong>
          </div>
        );
        const grpB = (
          <div key="grpB" style={{ display: 'flex', alignItems: 'baseline', gap: 8 }}>
            <span style={{ color: 'var(--dex-orange, #ff8c00)', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={labelB}>● {labelB}</span>
            <strong style={{ whiteSpace: 'nowrap' }}>{funActive}<span style={{ color: 'var(--dex-gray-400)' }}>/{funCap}</span></strong>
          </div>
        );
        return (
          <div className="admin-counters" style={{ display: 'grid', gridTemplateColumns: gridCols, gap: 12, marginBottom: 24 }}>
            <div className="card" style={{ padding: 16, textAlign: 'center' }}>
              <div style={{ fontSize: '1.8rem', fontWeight: 700, color: '#1565c0' }}>{totalActive}</div>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.registered')}</div>
              {isSplitCapacity && (
                <div style={{
                  marginTop: 10, paddingTop: 10,
                  borderTop: '1px solid var(--dex-gray-200)',
                  fontSize: '0.82rem', textAlign: 'left',
                  display: 'flex', flexDirection: 'column', gap: 4,
                }}>
                  {reversed ? <>{grpB}{grpA}</> : <>{grpA}{grpB}</>}
                </div>
              )}
            </div>
            {/* v23.30: Warteliste direkt rechts neben „Angemeldet". */}
            {hasWaitlistKPI && (
              <div className="card" style={{ padding: 16, textAlign: 'center' }}>
                <div style={{ fontSize: '1.8rem', fontWeight: 700, color: 'var(--dex-orange)' }}>
                  {isConsolidatedMode ? consolidatedWaitlistByEmail.size : registrations.filter(r => r.Status === 'Warteliste').length}
                </div>
                <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.waitlist')}</div>
              </div>
            )}
            <div className="card" style={{ padding: 16, textAlign: 'center' }}>
              <div style={{ fontSize: '1.8rem', fontWeight: 700, color: '#6a1b9a' }}>
                {isConsolidatedMode ? consolidatedQRByEmail.size : registrations.filter(r => r.Status === 'QR versendet').length}
              </div>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.qrsent')}</div>
            </div>
            <div className="card" style={{ padding: 16, textAlign: 'center' }}>
              <div style={{ fontSize: '1.8rem', fontWeight: 700, color: 'var(--dex-green)' }}>
                {isConsolidatedMode ? consolidatedCheckedByEmail.size : registrations.filter(r => r.Status === 'Eingecheckt').length}
              </div>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.checkedin')}</div>
            </div>
            <div className="card" style={{ padding: 16, textAlign: 'center' }}>
              <div style={{ fontSize: '1.8rem', fontWeight: 700, color: 'var(--dex-gray-400)' }}>
                {isConsolidatedMode ? consolidatedCancelledByEmail.size : registrations.filter(r => r.Status === 'Abgemeldet').length}
              </div>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.cancelled')}</div>
            </div>
          </div>
        );
};

