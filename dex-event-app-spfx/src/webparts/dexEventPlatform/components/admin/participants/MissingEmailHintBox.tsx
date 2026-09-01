/* MissingEmailHintBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10774-10808 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

export interface MissingEmailHintBoxProps {
  isDe: boolean;
  missingEmailRegs: SPRegistration[];
  selectedEvent: DeloitteEvent;
}

export const MissingEmailHintBox: React.FC<MissingEmailHintBoxProps> = (p) => {
  const { isDe, missingEmailRegs, selectedEvent } = p;
          // v23.3: Hinweis auf aktive Anmeldungen ohne gültige E-Mail. Diese
          // belegen einen Platz (zählen also in „Aktuell registriert"/Tabelle),
          // bekommen aber KEINE Bestätigung/QR/Outlook und tauchen in den
          // entdoppelten Zahlen sonst nicht auf (E-Mail = Dedup-Schlüssel).
          if (missingEmailRegs.length === 0 || !selectedEvent) return null;
          return (
            <div style={{ marginBottom: 20, padding: 16, borderRadius: 12, border: '1px solid var(--dex-orange, #ed8b00)', background: 'rgba(237,139,0,0.07)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 10 }}>
                <Icon iconName="Mail" style={{ fontSize: 18, color: 'var(--dex-orange-dark, #b35a00)' }} />
                <strong style={{ color: 'var(--dex-orange-dark, #b35a00)', fontSize: '0.95rem' }}>
                  {isDe ? `Anmeldungen ohne gültige E-Mail (${missingEmailRegs.length})` : `Registrations without a valid email (${missingEmailRegs.length})`}
                </strong>
                <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Diese Personen belegen einen Platz, bekommen aber keine Bestätigung/QR/Outlook. Über „Felder" bzw. „Details" die E-Mail-Adresse nachtragen.'
                    : 'These people occupy a seat but receive no confirmation/QR/Outlook. Add their email address via „Fields" or „Details".'}
                </span>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {missingEmailRegs.map(r => {
                  const nm = (r.Vorname && r.Nachname) ? `${r.Vorname} ${r.Nachname}` : (r.ParticipantName || (isDe ? '(ohne Namen)' : '(no name)'));
                  return (
                    <div key={r.Id} style={{ fontSize: '0.84rem', color: 'var(--dex-gray-800)', display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'baseline' }}>
                      <strong>{nm}</strong>
                      {typeof r.TeilnehmerID === 'number' && <span style={{ color: 'var(--dex-gray-500)' }}>#{r.TeilnehmerID}</span>}
                      {r.TeamName && <span style={{ color: 'var(--dex-gray-600)' }}>— „{r.TeamName}“</span>}
                      <span style={{ color: 'var(--dex-gray-400)', fontSize: '0.78rem' }}>{r.ParticipantEmail ? `(${r.ParticipantEmail})` : (isDe ? '(keine E-Mail)' : '(no email)')}</span>
                    </div>
                  );
                })}
              </div>
            </div>
          );
};

