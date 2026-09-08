/* MissingEmailHintBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10774-10808 des
 * Stands vor dem Schnitt). Die Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3: Nach docs/ui-leitfaden.md 5a (Punkt 2) umgebaut — ein
 * `dex-ui-callout--warn` mit der Folge im Satz und der Zahl als Pille; die
 * Personenliste liegt hinter einem Aufklapper, weil sie sonst die
 * Teilnehmerliste nach unten schiebt, obwohl der Satz darüber schon alles sagt.
 */
import * as React from 'react';
import { ChevronDown, Mail } from '../../Icons';
import { cx, ensureDexUiStyles } from '../../dexUi';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

export interface MissingEmailHintBoxProps {
  isDe: boolean;
  missingEmailRegs: SPRegistration[];
  selectedEvent: DeloitteEvent;
}

/** v31.3: Eine Person je Zeile im Aufklapper — ruhig, ohne Hover (nichts klickbar). */
const LINE: React.CSSProperties = { display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'baseline', padding: '8px 0', fontSize: '0.84rem' };

export const MissingEmailHintBox: React.FC<MissingEmailHintBoxProps> = (p) => {
  const { isDe, missingEmailRegs, selectedEvent } = p;
          // v31.3: Hook vor dem frühen Return (rules-of-hooks ist `error`);
          // ensureDexUiStyles ist idempotent und nötig, weil der Kasten ohne
          // Modal/WizardFormShell auf der Seite steht.
          const [listOpen, setListOpen] = React.useState(false);
          ensureDexUiStyles();
          // v23.3: Hinweis auf aktive Anmeldungen ohne gültige E-Mail. Diese
          // belegen einen Platz (zählen also in „Aktuell registriert"/Tabelle),
          // bekommen aber KEINE Bestätigung/QR/Outlook und tauchen in den
          // entdoppelten Zahlen sonst nicht auf (E-Mail = Dedup-Schlüssel).
          if (missingEmailRegs.length === 0 || !selectedEvent) return null;
          return (
            <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 20 }}>
              <span className="dex-ui-callout-icon" aria-hidden="true"><Mail size={16} /></span>
              <div style={{ minWidth: 0, flex: 1 }}>
                <div style={{ fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                  {isDe ? 'Anmeldungen ohne gültige E-Mail' : 'Registrations without a valid email'}
                  <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? `${missingEmailRegs.length} Personen` : `${missingEmailRegs.length} people`}</span>
                </div>
                <div style={{ marginTop: 2 }}>
                  {isDe
                    ? 'Diese Personen belegen einen Platz, bekommen aber keine Bestätigung, keinen QR-Code und keinen Outlook-Termin. Trage die Adresse in der Teilnehmerzeile über „Felder“ bzw. „Details“ nach.'
                    : 'These people occupy a seat but receive no confirmation, no QR code and no calendar invite. Add the address in the attendee row via „Fields“ or „Details“.'}
                </div>
                <button type="button" className={cx('dex-ui-disclosure', listOpen && 'is-open')} aria-expanded={listOpen} onClick={() => setListOpen(o => !o)}>
                  <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                  {isDe ? `Wen es betrifft (${missingEmailRegs.length})` : `Who is affected (${missingEmailRegs.length})`}
                </button>
                {listOpen && (
                  <div className="dex-ui-disclosure-body dex-ui-card" style={{ padding: '2px 12px' }}>
                    {missingEmailRegs.map(r => (
                      <div key={r.Id} className="dex-ui-row--bordered" style={LINE}>
                        <strong>{(r.Vorname && r.Nachname) ? `${r.Vorname} ${r.Nachname}` : (r.ParticipantName || (isDe ? '(ohne Namen)' : '(no name)'))}</strong>
                        {typeof r.TeilnehmerID === 'number' && <span className="dex-ui-pill dex-ui-pill--gray">#{r.TeilnehmerID}</span>}
                        {r.TeamName && <span className="dex-ui-muted">— „{r.TeamName}“</span>}
                        <span className="dex-ui-muted">{r.ParticipantEmail ? `(${r.ParticipantEmail})` : (isDe ? '(keine E-Mail)' : '(no email)')}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          );
};

