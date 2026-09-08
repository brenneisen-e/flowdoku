/* DuplicateInSubEventHintBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10737-10772 des
 * Stands vor dem Schnitt). Wer hier steht, entscheidet der Aufrufer.
 *
 * v31.3: Nach docs/ui-leitfaden.md 5a (Punkt 2) umgebaut — `dex-ui-callout--danger`
 * mit der Folge im ersten Satz, Zahl als Pille, Personenliste hinter einem
 * Aufklapper. Die Überschrift nennt jetzt den Ort („im selben Sub-Event"): Sie
 * war wortgleich mit der des Kastens darüber, und zwei identische
 * Ueberschriften untereinander liest niemand als zwei verschiedene Fälle.
 */
import * as React from 'react';
import { AlertCircle, ChevronDown } from '../../Icons';
import { cx, ensureDexUiStyles } from '../../dexUi';

export interface DuplicateInSubEventHintBoxProps {
  isDe: boolean;
  subEventDupGroups: { sectionTitle: string; email: string; name: string; count: number; rowsInfo: string; }[];
}

/** v31.3: Eine Person je Zeile im Aufklapper — ruhig, ohne Hover (nichts klickbar). */
const LINE: React.CSSProperties = { display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'baseline', padding: '8px 0', fontSize: '0.84rem' };

export const DuplicateInSubEventHintBox: React.FC<DuplicateInSubEventHintBoxProps> = (p) => {
  const { isDe, subEventDupGroups } = p;
          // v31.3: Hook vor dem frühen Return (rules-of-hooks ist `error`);
          // ensureDexUiStyles ist idempotent und nötig, weil der Kasten ohne
          // Modal/WizardFormShell auf der Seite steht.
          const [listOpen, setListOpen] = React.useState(false);
          ensureDexUiStyles();
          // v28.21: ECHTE Doppel-Anmeldung im Klammer-Modus — dieselbe Person
          // zweimal aktiv IM SELBEN Sub-Event. Nur das belegt zwei Plätze und
          // gehört rot gemeldet; die doppelten Klammer-Schatten-Zeilen laufen
          // über den neutralen Hinweis oben.
          if (subEventDupGroups.length === 0) return null;
          return (
            <div className="dex-ui-callout dex-ui-callout--danger" style={{ marginBottom: 20 }}>
              <span className="dex-ui-callout-icon" aria-hidden="true"><AlertCircle size={16} /></span>
              <div style={{ minWidth: 0, flex: 1 }}>
                <div style={{ fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                  {isDe ? 'Doppel-Anmeldungen im selben Sub-Event' : 'Duplicate registrations in the same sub-event'}
                  <span className="dex-ui-pill dex-ui-pill--red">{isDe ? `${subEventDupGroups.length} Fälle` : `${subEventDupGroups.length} cases`}</span>
                </div>
                <div style={{ marginTop: 2 }}>
                  {isDe
                    ? 'Dieselbe Person belegt dort zwei Plätze. Im jeweiligen Sub-Event-Reiter entfernst du die doppelte Zeile über „Abmelden“ bzw. „Entfernen“ (Warteliste) still.'
                    : 'The same person occupies two seats there. Remove the duplicate row silently via „Cancel“ or „Remove“ (waitlist) in that sub-event tab.'}
                </div>
                <button type="button" className={cx('dex-ui-disclosure', listOpen && 'is-open')} aria-expanded={listOpen} onClick={() => setListOpen(o => !o)}>
                  <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                  {isDe ? `Wen es betrifft (${subEventDupGroups.length})` : `Who is affected (${subEventDupGroups.length})`}
                </button>
                {listOpen && (
                  <div className="dex-ui-disclosure-body">
                    <div className="dex-ui-muted" style={{ marginBottom: 6 }}>
                      {isDe
                        ? 'Achtung: Die doppelten Zeilen können auch auf der WARTELISTE des Sub-Events stehen — dann fehlen sie in der Teilnehmer-Tabelle. Status und Zeitpunkt stehen hinter jedem Eintrag.'
                        : 'Note: the duplicate rows may sit on the sub-event’s WAITLIST — then they are missing from the attendee table. Status and time are shown behind each entry.'}
                    </div>
                    <div className="dex-ui-card" style={{ padding: '2px 12px' }}>
                      {subEventDupGroups.map(g => (
                        <div key={`${g.sectionTitle}::${g.email}`} className="dex-ui-row--bordered" style={LINE}>
                          <strong>{g.name}</strong>
                          <span className="dex-ui-muted">{g.email}</span>
                          <span className="dex-ui-pill dex-ui-pill--red">{isDe ? `${g.count}× angemeldet` : `${g.count}× registered`}</span>
                          <span className="dex-ui-muted">— {g.sectionTitle}</span>
                          {/* v30.14: Status + Zeitpunkt je Zeile — Warteliste-Duplikate sind sonst nicht auffindbar. */}
                          <span className="dex-ui-muted">({g.rowsInfo})</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            </div>
          );
};

