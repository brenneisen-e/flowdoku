/* MassmailPasteModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 15021-15103 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { cx } from '../../dexUi';
import { Check, ChevronDown, Mail } from '../../Icons';
import { useLocaleSafe } from '../../../context/LanguageContext';
import { SPRegistration } from '../../../services/EventService';

export interface MassmailPasteModalProps {
  massmailPasteRaw: string;
  registrations: SPRegistration[];
  setMassmailMode: React.Dispatch<React.SetStateAction<"closed" | "pick" | "paste" | "editor">>;
  setMassmailPasteRaw: React.Dispatch<React.SetStateAction<string>>;
  setShowEmailModal: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export const MassmailPasteModal: React.FC<MassmailPasteModalProps> = (p) => {
  const { massmailPasteRaw, registrations, setMassmailMode, setMassmailPasteRaw, setShowEmailModal, showAlert } = p;
        // v31.2: Die Props kennen kein isDe (Schnittstelle bleibt) — die Sprache
        // kommt wie in Modal.tsx aus dem Kontext.
        const isDe = useLocaleSafe() === 'de';
        const [showMissing, setShowMissing] = React.useState(false);
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
          if (missing.length === 0) { showAlert(isDe ? 'Alle aktiven Teilnehmer stehen bereits in deiner Liste — niemand zum Anschreiben übrig.' : 'All active participants are already on your list — nobody left to write to.'); return; }
          setShowEmailModal(true);
          setMassmailMode('editor');
        };
        // v31.2: Frage statt Feldname im Kopf, drei Kennzahlen statt drei Fließtext-
        // Zeilen, die Empfängerliste hinter einem Aufklapper; der Primär-Knopf nennt
        // die Zahl. backdropClose aus: Ein Klick neben die Karte warf den
        // eingefügten Verteiler weg (Modal.tsx, v30.51).
        return (
          <Modal open={true} onClose={closeAll} maxWidth={680} backdropClose={false}
            ariaLabel={isDe ? 'Nachrücker — Liste einfügen' : 'Late joiners — paste list'}
            title={isDe ? 'Wer hat die Mail schon bekommen?' : 'Who already received the mail?'}
            subtitle={isDe ? 'Schritt 2 von 2 — füge deine bisherige Empfänger-Liste ein. Angeschrieben wird, wer im Event aktiv ist, aber dort nicht steht.' : 'Step 2 of 2 — paste your existing recipient list. The mail goes to everyone active in the event who is not on it.'}
            icon={<Mail size={20} />}
            footer={<>
              <span className="dex-ui-modal-foot-left"><button type="button" className="btn btn-secondary" onClick={back}>{isDe ? 'Zurück' : 'Back'}</button></span>
              <button type="button" className="btn btn-secondary" onClick={closeAll}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
              <button type="button" className="btn btn-primary" disabled={missing.length === 0} onClick={continueAction}>{isDe ? `Weiter zum Mail-Editor (${missing.length})` : `Continue to mail editor (${missing.length})`}</button>
            </>}>
            <div className="dex-ui-field">
              <label className="dex-ui-label" htmlFor="massmail-paste-raw">{isDe ? 'Deine bisherige Empfänger-Liste' : 'Your existing recipient list'}</label>
              <textarea id="massmail-paste-raw" className="dex-ui-textarea" value={massmailPasteRaw} onChange={e => setMassmailPasteRaw(e.target.value)}
                placeholder={'Max Mustermann <mmustermann@deloitte.de>; anna.schmidt@deloitte.de; ...'}
                style={{ minHeight: 160, fontFamily: 'monospace', fontSize: '0.82rem' }} />
              <div className="dex-ui-help">
                {isDe ? 'Hau alles rein, was du hast: Verteiler-Export, Outlook-To-Zeile, „Vorname Nachname <mail@deloitte.de>“, komma-, semikolon- oder zeilengetrennt — die App pickt die E-Mail-Adressen automatisch heraus.' : 'Paste whatever you have: distribution list export, Outlook To line, "First Last <mail@deloitte.de>", separated by comma, semicolon or line break — the app picks out the email addresses automatically.'}
              </div>
            </div>
            <div className="dex-ui-grid-3">
              {([
                [pasted.length, isDe ? 'Adressen erkannt' : 'addresses found', ''],
                [active.length, isDe ? 'aktive Teilnehmer' : 'active participants', ''],
                [missing.length, isDe ? 'nicht in deiner Liste — werden angeschrieben' : 'not on your list — will be mailed', 'dex-ui-kpi--orange'],
              ] as [number, string, string][]).map(([v, label, mod]) => (
                <div key={label} className={cx('dex-ui-kpi', mod)}><div className="dex-ui-kpi-value">{v}</div><div className="dex-ui-kpi-label">{label}</div></div>
              ))}
            </div>
            {missing.length > 0 ? (
              <div>
                <button type="button" className={cx('dex-ui-disclosure', showMissing && 'is-open')} aria-expanded={showMissing} onClick={() => setShowMissing(v => !v)}>
                  <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                  {isDe ? 'Empfänger anzeigen' : 'Show recipients'}
                  <span className="dex-ui-disclosure-count">{missing.length}</span>
                </button>
                {showMissing && (
                  <div className="dex-ui-disclosure-body">
                    <div className="dex-ui-table-wrap" style={{ maxHeight: 320, overflowY: 'auto' }}>
                      <table className="dex-ui-table">
                        <thead><tr><th>{isDe ? 'Vorname' : 'First name'}</th><th>{isDe ? 'Nachname' : 'Last name'}</th><th>{isDe ? 'Position' : 'Job title'}</th><th>{isDe ? 'E-Mail' : 'Email'}</th></tr></thead>
                        <tbody>
                          {missing.map(r => (
                            <tr key={r.Id}>
                              <td>{r.Vorname || '-'}</td><td>{r.Nachname || '-'}</td>
                              <td style={{ color: 'var(--dex-gray-600)' }}>{(r as SPRegistration & { JobTitle?: string }).JobTitle || '-'}</td>
                              <td style={{ color: 'var(--dex-gray-600)' }}>{r.ParticipantEmail}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            ) : (
              <div className="dex-ui-callout dex-ui-callout--success"><span className="dex-ui-callout-icon"><Check size={16} /></span>
                <span>{isDe ? 'Alle aktiven Teilnehmer stehen schon in deiner Liste — es bleibt niemand zum Anschreiben.' : 'All active participants are already on your list — nobody left to write to.'}</span></div>
            )}
          </Modal>
        );
};

