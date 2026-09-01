/* MassmailPasteModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 15021-15103 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

export interface MassmailPasteModalProps {
  massmailMode: "closed" | "pick" | "paste" | "editor";
  massmailPasteRaw: string;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setMassmailMode: React.Dispatch<React.SetStateAction<"closed" | "pick" | "paste" | "editor">>;
  setMassmailPasteRaw: React.Dispatch<React.SetStateAction<string>>;
  setShowEmailModal: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export const MassmailPasteModal: React.FC<MassmailPasteModalProps> = (p) => {
  const { massmailMode, massmailPasteRaw, registrations, selectedEvent, setMassmailMode, setMassmailPasteRaw, setShowEmailModal, showAlert } = p;
        const closeAll = (): void => { setMassmailMode('closed'); setMassmailPasteRaw(''); };
        const back = (): void => { setMassmailMode('pick'); };
        // E-Mail-Adressen aus dem Rohtext extrahieren — robust gegen Vorname
        // Nachname <mail@…> / mail@…; sep / Outlook-Verteiler-Dumps.
        const extractEmails = (raw: string): string[] => {
          if (!raw) return [];
          const matches = raw.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g) || [];
          const seen = new Set<string>();
          const out: string[] = [];
          for (const m of matches) {
            const e = m.toLowerCase();
            if (!seen.has(e)) { seen.add(e); out.push(e); }
          }
          return out;
        };
        const pasted = extractEmails(massmailPasteRaw);
        const active = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
        const pastedSet = new Set(pasted);
        const missing = active.filter(r => !pastedSet.has((r.ParticipantEmail || '').toLowerCase()));
        const continueAction = (): void => {
          if (missing.length === 0) { showAlert('Alle aktiven Teilnehmer stehen bereits in deiner Liste — niemand zum Anschreiben uebrig.'); return; }
          setShowEmailModal(true);
          setMassmailMode('editor');
        };
        return (
          <Modal open={true} onClose={closeAll} maxWidth={680} padding={24} ariaLabel="Nachrücker — Liste einfügen">
            <h3 style={{ margin: '0 0 8px', fontSize: '1.1rem' }}>Nachrücker — bestehende Empfänger-Liste einfügen</h3>
            <p style={{ margin: '0 0 14px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              Hau alles rein, was du hast — Verteiler-Export, Outlook-To-Liste, Vorname Nachname &lt;mail@deloitte.de&gt;-Format, kommagetrennt, semikolongetrennt, Zeilenumbruch — die App pickt die E-Mail-Adressen automatisch raus.
            </p>
            <textarea
              value={massmailPasteRaw}
              onChange={e => setMassmailPasteRaw(e.target.value)}
              placeholder={'Max Mustermann <mmustermann@deloitte.de>; anna.schmidt@deloitte.de; ...'}
              style={{ width: '100%', minHeight: 160, fontFamily: 'monospace', fontSize: '0.82rem', padding: 8, border: '1px solid var(--dex-gray-300)', borderRadius: 6, resize: 'vertical' }}
            />
            <div style={{ marginTop: 10, padding: 10, borderRadius: 6, background: 'var(--dex-gray-50, #fafafa)', fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
              <strong>{pasted.length}</strong> Adressen aus dem Text extrahiert.<br />
              <strong>{active.length}</strong> aktive Teilnehmer im Event.<br />
              <strong style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>{missing.length}</strong> Teilnehmer NICHT in deiner Liste — die werden angeschrieben.
            </div>
            {missing.length > 0 && (
              <details style={{ marginTop: 8, padding: 0, borderRadius: 6, background: 'rgba(237,139,0,0.06)', border: '1px solid var(--dex-orange, #ed8b00)', fontSize: '0.82rem' }}>
                <summary style={{ padding: '8px 12px', cursor: 'pointer', fontWeight: 600, color: 'var(--dex-orange-dark, #b35a00)' }}>
                  Empfänger anzeigen ({missing.length})
                </summary>
                <div style={{ maxHeight: 320, overflowY: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.8rem' }}>
                    <thead>
                      <tr style={{ borderBottom: '1px solid var(--dex-orange, #ed8b00)', background: 'rgba(255,255,255,0.6)' }}>
                        <th style={{ textAlign: 'left', padding: 6 }}>Vorname</th>
                        <th style={{ textAlign: 'left', padding: 6 }}>Nachname</th>
                        <th style={{ textAlign: 'left', padding: 6 }}>Position</th>
                        <th style={{ textAlign: 'left', padding: 6 }}>Email</th>
                      </tr>
                    </thead>
                    <tbody>
                      {missing.map(r => {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const anyR = r as any;
                        return (
                          <tr key={r.Id} style={{ borderBottom: '1px solid rgba(237,139,0,0.15)' }}>
                            <td style={{ padding: 6 }}>{r.Vorname || '-'}</td>
                            <td style={{ padding: 6 }}>{r.Nachname || '-'}</td>
                            <td style={{ padding: 6, color: 'var(--dex-gray-600)' }}>{anyR.JobTitle || '-'}</td>
                            <td style={{ padding: 6, color: 'var(--dex-gray-600)' }}>{r.ParticipantEmail}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </details>
            )}
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 14 }}>
              <button type="button" className="btn btn-secondary" onClick={back} style={{ fontSize: '0.85rem' }}>Zurück</button>
              <button type="button" className="btn btn-secondary" onClick={closeAll} style={{ fontSize: '0.85rem' }}>Abbrechen</button>
              <button type="button" className="btn btn-primary" disabled={missing.length === 0} onClick={continueAction} style={{ fontSize: '0.85rem' }}>Weiter zum Mail-Editor</button>
            </div>
          </Modal>
        );
};

