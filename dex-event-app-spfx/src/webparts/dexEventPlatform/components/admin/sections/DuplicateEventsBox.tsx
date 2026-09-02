/* DuplicateEventsBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 6933-6986 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { AlertCircle, Trash2 } from '../../Icons';
import { DeloitteEvent } from '../../../types';

export interface DuplicateEventsBoxProps {
  duplicateEvents: DeloitteEvent[];
  isDe: boolean;
  selectedEvent: DeloitteEvent;
  setConfirmDeleteEvent: React.Dispatch<React.SetStateAction<DeloitteEvent>>;
}

export const DuplicateEventsBox: React.FC<DuplicateEventsBoxProps> = (p) => {
  const { duplicateEvents, isDe, selectedEvent, setConfirmDeleteEvent } = p;
        // Alle Versionen (geöffnetes Event + Duplikate), älteste zuerst — so ist
        // die „alte" Version oben gut erkennbar.
        const allVersions = [selectedEvent, ...duplicateEvents].slice().sort((a, b) => {
          const ta = a.created ? new Date(a.created).getTime() : 0;
          const tb = b.created ? new Date(b.created).getTime() : 0;
          return ta - tb;
        });
        const fmt = (d?: string): string => { if (!d) return ''; try { return new Date(d).toLocaleString(isDe ? 'de-DE' : 'en-GB'); } catch { return ''; } };
        return (
          <div style={{
            display: 'flex', alignItems: 'flex-start', gap: 12,
            padding: '14px 16px', marginBottom: 20,
            background: '#fff3e0', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 12,
            color: 'var(--dex-gray-800)',
          }}>
            <AlertCircle size={20} />
            <div style={{ minWidth: 0, flex: 1 }}>
              <div style={{ fontWeight: 700, marginBottom: 4, color: 'var(--dex-orange-dark, #b35a00)' }}>
                {isDe
                  ? `Achtung: Dieses Event existiert in ${allVersions.length} Versionen mit gleichem oder ähnlichem Namen am gleichen Tag`
                  : `Heads up: this event exists in ${allVersions.length} versions with the same or similar name on the same day`}
              </div>
              <div style={{ fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, marginBottom: 8 }}>
                {isDe
                  ? 'Sieht so aus, als wäre dieses Event versehentlich mehrfach angelegt worden. Bitte die nicht mehr benötigte(n) Version(en) löschen, damit Teilnehmer sich nicht auf der falschen Version anmelden.'
                  : 'It looks like this event was created more than once. Please delete the version(s) you no longer need so attendees do not register on the wrong one.'}
              </div>
              <ul style={{ margin: 0, paddingLeft: 0, listStyle: 'none', display: 'flex', flexDirection: 'column', gap: 6 }}>
                {allVersions.map(ver => {
                  const created = fmt(ver.created);
                  const statusLabel = ver.isFictive ? (isDe ? 'Entwurf' : 'Draft') : (ver.status || '');
                  const isOpen = ver.id === selectedEvent.id;
                  return (
                    <li key={ver.id} style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', fontSize: '0.85rem' }}>
                      <strong>{ver.title}</strong>
                      {statusLabel && <span style={{ padding: '1px 8px', borderRadius: 10, background: 'rgba(0,0,0,0.06)', fontSize: 12 }}>{statusLabel}</span>}
                      <span style={{ color: 'var(--dex-gray-500)' }}>{created ? (isDe ? `erstellt am ${created}` : `created ${created}`) : (isDe ? 'Erstelldatum unbekannt' : 'creation date unknown')}</span>
                      {isOpen && <span style={{ color: 'var(--dex-orange-dark, #b35a00)', fontSize: 12, fontWeight: 600 }}>{isDe ? '· gerade geöffnet' : '· currently open'}</span>}
                      <button
                        type="button"
                        onClick={() => setConfirmDeleteEvent(ver)}
                        style={{ display: 'inline-flex', alignItems: 'center', gap: 5, background: 'transparent', border: '1px solid var(--dex-red, #c00)', color: 'var(--dex-red, #c00)', borderRadius: 6, padding: '3px 10px', cursor: 'pointer', fontSize: '0.78rem', fontWeight: 600, fontFamily: 'inherit' }}
                      >
                        <Trash2 size={12} /> {isDe ? 'Diese Version löschen' : 'Delete this version'}
                      </button>
                    </li>
                  );
                })}
              </ul>
            </div>
          </div>
        );
};

