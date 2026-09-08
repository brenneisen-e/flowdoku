/* DuplicateEventsBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 6933-6986 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich übernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3: Nach docs/ui-leitfaden.md 5a umgebaut — `dex-ui-callout--warn` statt
 * eigener Orange-Fläche, jede Version eine `dex-ui-row` mit Statuspunkt,
 * Status-Pille und dem Löschen-Knopf IN der Zeile. Das Erstelldatum steht als
 * Unterzeile, damit die Titelzeile die Versionen unterscheidbar hält.
 */
import * as React from 'react';
import { AlertCircle, Trash2 } from '../../Icons';
import { cx, ensureDexUiStyles } from '../../dexUi';
import { DeloitteEvent } from '../../../types';

export interface DuplicateEventsBoxProps {
  duplicateEvents: DeloitteEvent[];
  isDe: boolean;
  selectedEvent: DeloitteEvent;
  setConfirmDeleteEvent: React.Dispatch<React.SetStateAction<DeloitteEvent>>;
}

export const DuplicateEventsBox: React.FC<DuplicateEventsBoxProps> = (p) => {
  const { duplicateEvents, isDe, selectedEvent, setConfirmDeleteEvent } = p;
  // Idempotent — der Kasten steht ohne Modal/WizardFormShell in der Event-Seite.
  ensureDexUiStyles();
        // Alle Versionen (geöffnetes Event + Duplikate), älteste zuerst — so ist
        // die „alte" Version oben gut erkennbar.
        const allVersions = [selectedEvent, ...duplicateEvents].slice().sort((a, b) => {
          const ta = a.created ? new Date(a.created).getTime() : 0;
          const tb = b.created ? new Date(b.created).getTime() : 0;
          return ta - tb;
        });
        const fmt = (d?: string): string => { if (!d) return ''; try { return new Date(d).toLocaleString(isDe ? 'de-DE' : 'en-GB'); } catch { return ''; } };
        return (
          <div className="dex-ui-callout dex-ui-callout--warn" style={{ flexDirection: 'column', marginBottom: 20 }}>
            <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10 }}>
              <span className="dex-ui-callout-icon"><AlertCircle size={18} /></span>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontWeight: 700 }}>
                  {isDe
                    ? `Dieses Event gibt es ${allVersions.length}-mal — gleicher Tag, gleicher oder ähnlicher Name`
                    : `This event exists ${allVersions.length} times — same day, same or similar name`}
                </div>
                <div style={{ marginTop: 3 }}>
                  {isDe
                    ? 'Sieht so aus, als wäre es versehentlich mehrfach angelegt worden. Lösch die Versionen, die du nicht mehr brauchst — sonst melden sich Teilnehmer auf der falschen an.'
                    : 'It looks like it was created more than once. Delete the versions you no longer need — otherwise attendees register on the wrong one.'}
                </div>
              </div>
            </div>
            <ul className="dex-ui-card dex-ui-card--list" style={{ margin: 0, listStyle: 'none', background: '#fff' }}>
              {allVersions.map(ver => {
                const created = fmt(ver.created);
                const statusLabel = ver.isFictive ? (isDe ? 'Entwurf' : 'Draft') : (ver.status || '');
                const isOpen = ver.id === selectedEvent.id;
                return (
                  <li key={ver.id} className="dex-ui-row dex-ui-row--bordered">
                    {/* Grüner Punkt = die Version, die gerade offen ist. */}
                    <span className={cx('dex-ui-dot', isOpen && 'dex-ui-dot--green')} aria-hidden="true" />
                    <div className="dex-ui-row-main">
                      <div className="dex-ui-row-title" style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap', whiteSpace: 'normal' }}>
                        {ver.title}
                        {statusLabel && <span className="dex-ui-pill dex-ui-pill--gray">{statusLabel}</span>}
                        {isOpen && <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'gerade geöffnet' : 'currently open'}</span>}
                      </div>
                      <div className="dex-ui-row-sub">
                        {created ? (isDe ? `erstellt am ${created}` : `created ${created}`) : (isDe ? 'Erstelldatum unbekannt' : 'creation date unknown')}
                      </div>
                    </div>
                    <div className="dex-ui-row-actions">
                      <button type="button" className="dex-ui-textbtn dex-ui-textbtn--danger" onClick={() => setConfirmDeleteEvent(ver)}>
                        <Trash2 size={13} /> {isDe ? 'Diese Version löschen' : 'Delete this version'}
                      </button>
                    </div>
                  </li>
                );
              })}
            </ul>
          </div>
        );
};

