/* MassmailPickModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14939-15018 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { MassmailAudience } from '../adminTypes';
import Modal from '../../Modal';
import { SPRegistration } from '../../../services/EventService';

export interface MassmailPickModalProps {
  massmailAudience: MassmailAudience;
  massmailStatuses: Set<string>;
  registrations: SPRegistration[];
  setMassmailAudience: React.Dispatch<React.SetStateAction<MassmailAudience>>;
  setMassmailMode: React.Dispatch<React.SetStateAction<"closed" | "pick" | "paste" | "editor">>;
  setMassmailPasteRaw: React.Dispatch<React.SetStateAction<string>>;
  setMassmailStatuses: React.Dispatch<React.SetStateAction<Set<string>>>;
  setShowEmailModal: React.Dispatch<React.SetStateAction<boolean>>;
}

export const MassmailPickModal: React.FC<MassmailPickModalProps> = (p) => {
  const { massmailAudience, massmailStatuses, registrations, setMassmailAudience, setMassmailMode, setMassmailPasteRaw, setMassmailStatuses, setShowEmailModal } = p;
        const closeAll = (): void => { setMassmailMode('closed'); setMassmailPasteRaw(''); };
        const proceed = (): void => {
          if (massmailAudience === 'custom' && massmailStatuses.size === 0) return;
          if (massmailAudience === 'nachruecker') setMassmailMode('paste');
          else { setShowEmailModal(true); setMassmailMode('editor'); }
        };
        const STATUS_OPTIONS = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
        const toggleStatus = (st: string): void => {
          setMassmailStatuses(prev => {
            const next = new Set(prev);
            if (next.has(st)) next.delete(st); else next.add(st);
            return next;
          });
        };
        const Row = (props: { value: MassmailAudience; label: string; desc: string }): React.ReactElement => (
          <label style={{
            display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
            borderRadius: 8, border: `1px solid ${massmailAudience === props.value ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
            background: massmailAudience === props.value ? 'rgba(134,188,37,0.08)' : '#fff',
            cursor: 'pointer', marginBottom: 8,
          }}>
            <input
              type="radio"
              name="massmail-target"
              checked={massmailAudience === props.value}
              onChange={() => setMassmailAudience(props.value)}
              style={{ marginTop: 3 }}
            />
            <div>
              <div style={{ fontWeight: 600, fontSize: '0.92rem' }}>{props.label}</div>
              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>{props.desc}</div>
            </div>
          </label>
        );
        return (
          <Modal open={true} onClose={closeAll} maxWidth={560} padding={24} ariaLabel="Empfänger wählen">
            <h3 style={{ margin: '0 0 14px', fontSize: '1.1rem' }}>An wen soll die Mail gehen?</h3>
            <Row value="active" label="Teilnehmer (alle aktiven)" desc="Status: Angemeldet, QR versendet, Eingecheckt — Default für die ueblichen Info-Mails." />
            <Row value="activePlusWait" label="Teilnehmer + Warteliste" desc="Alle aktiven UND Wartelistler — z.B. wenn sich noch Plätze frei machen und du auch die Warteliste vorwarnen willst." />
            <Row value="waitOnly" label="Nur Warteliste" desc={'Nur Wartelistler — z.B. Info „Es wird wahrscheinlich keinen Platz mehr geben".'} />
            <Row value="nachruecker" label="Nachrücker (Manueller Abgleich)" desc="Du fügst im nächsten Schritt eine Liste von E-Mails ein (Verteiler, Vorname Nachname Email, beliebig formatiert) — die App extrahiert die Adressen und schickt die Mail an alle aktiven Teilnehmer, die NICHT in deiner Liste stehen." />
            {/* v22.9: Eigene Status-Auswahl — einzelne Status getrennt anhaken. */}
            <div style={{
              borderRadius: 8, border: `1px solid ${massmailAudience === 'custom' ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
              background: massmailAudience === 'custom' ? 'rgba(134,188,37,0.08)' : '#fff',
              marginBottom: 8, padding: 10,
            }}>
              <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                <input type="radio" name="massmail-target" checked={massmailAudience === 'custom'} onChange={() => setMassmailAudience('custom')} style={{ marginTop: 3 }} />
                <div>
                  <div style={{ fontWeight: 600, fontSize: '0.92rem' }}>Eigene Auswahl (nach Status)</div>
                  <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>Häkchen setzen, welche Status die Mail bekommen sollen — z.B. nur „QR versendet“.</div>
                </div>
              </label>
              {massmailAudience === 'custom' && (
                <div style={{ marginTop: 10, paddingLeft: 28, display: 'flex', flexDirection: 'column', gap: 7 }}>
                  {STATUS_OPTIONS.map(st => {
                    const count = registrations.filter(r => r.Status === st).length;
                    return (
                      <label key={st} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                        <input type="checkbox" checked={massmailStatuses.has(st)} onChange={() => toggleStatus(st)} />
                        <span style={{ fontWeight: 500 }}>{st}</span>
                        <span style={{ color: 'var(--dex-gray-500)' }}>({count})</span>
                      </label>
                    );
                  })}
                  {massmailStatuses.size === 0 && (
                    <span style={{ fontSize: '0.78rem', color: 'var(--dex-orange-dark, #b35a00)' }}>Bitte mindestens einen Status anhaken.</span>
                  )}
                </div>
              )}
            </div>
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 14 }}>
              <button type="button" className="btn btn-secondary" onClick={closeAll} style={{ fontSize: '0.85rem' }}>Abbrechen</button>
              <button type="button" className="btn btn-primary" onClick={proceed} disabled={massmailAudience === 'custom' && massmailStatuses.size === 0} style={{ fontSize: '0.85rem' }}>Weiter</button>
            </div>
          </Modal>
        );
};

