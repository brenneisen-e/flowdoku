/* OverbookDecisionModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16437-16560 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { getCachedLogoBase64, getCachedOrbBase64 } from '../../../services/EmailTemplates';
import { SPRegistration } from '../../../services/EventService';

export interface OverbookDecisionModalProps {
  obBusy: boolean;
  obKeepVariant: "active" | "firstWaitlist";
  obMailBody: string;
  obMailLang: "DE" | "EN";
  obMailSubject: string;
  obRemoveCalendar: boolean;
  obWithMail: boolean;
  overbookModal: { mode: "confirm" | "keep"; targets: SPRegistration[]; };
  runOverbookResolution: () => Promise<void>;
  setObKeepVariant: React.Dispatch<React.SetStateAction<"active" | "firstWaitlist">>;
  setObMailBody: React.Dispatch<React.SetStateAction<string>>;
  setObMailLang: React.Dispatch<React.SetStateAction<"DE" | "EN">>;
  setObMailSubject: React.Dispatch<React.SetStateAction<string>>;
  setObRemoveCalendar: React.Dispatch<React.SetStateAction<boolean>>;
  setObWithMail: React.Dispatch<React.SetStateAction<boolean>>;
  setOverbookModal: React.Dispatch<React.SetStateAction<{ mode: "confirm" | "keep"; targets: SPRegistration[]; }>>;
}

export const OverbookDecisionModal: React.FC<OverbookDecisionModalProps> = (p) => {
  const { obBusy, obKeepVariant, obMailBody, obMailLang, obMailSubject, obRemoveCalendar, obWithMail, overbookModal, runOverbookResolution, setObKeepVariant, setObMailBody, setObMailLang, setObMailSubject, setObRemoveCalendar, setObWithMail, setOverbookModal } = p;
  return (
        <Modal
          open={true}
          onClose={() => setOverbookModal(null)}
          dismissable={!obBusy}
          maxWidth={560}
          padding={24}
          ariaLabel="Überbuchung"
        >
            {overbookModal.mode === 'confirm' ? (
              <>
                <h3 style={{ marginTop: 0 }}>
                  Auf Warteliste bestätigen ({overbookModal.targets.length})
                </h3>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                  {overbookModal.targets.length === 1
                    ? `${(overbookModal.targets[0].Vorname && overbookModal.targets[0].Nachname) ? `${overbookModal.targets[0].Vorname} ${overbookModal.targets[0].Nachname}` : overbookModal.targets[0].ParticipantName} wird auf die Warteliste der Gruppe zurückgesetzt. Im Audit-Log wird vermerkt, dass die Person fälschlich angemeldet war.`
                    : `${overbookModal.targets.length} Personen werden auf die Warteliste zurückgesetzt. Im Audit-Log jeder Person wird der Vorgang vermerkt.`}
                </p>
                <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.88rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="checkbox" checked={obWithMail} onChange={e => setObWithMail(e.target.checked)} disabled={obBusy} />
                  Mit Entschuldigungs-Mail (Deloitte-Layout, in die Mail-Queue)
                </label>
                {obWithMail && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, margin: '0 0 10px', fontSize: '0.82rem' }}>
                    <span style={{ color: 'var(--dex-gray-600)' }}>Sprache:</span>
                    {(['DE', 'EN'] as const).map(lng => (
                      <button
                        key={lng}
                        type="button"
                        className="btn btn-secondary"
                        disabled={obBusy}
                        onClick={() => setObMailLang(lng)}
                        style={{
                          fontSize: '0.75rem', padding: '3px 12px',
                          ...(obMailLang === lng ? { background: 'var(--dex-green, #86bc25)', color: '#fff', fontWeight: 600 } : {}),
                        }}
                      >
                        {lng === 'DE' ? 'Deutsch' : 'English'}
                      </button>
                    ))}
                  </div>
                )}
                {obWithMail && overbookModal.targets.length === 1 && (
                  <div style={{ marginBottom: 10 }}>
                    <input
                      className="form-input"
                      value={obMailSubject}
                      onChange={e => setObMailSubject(e.target.value)}
                      disabled={obBusy}
                      style={{ width: '100%', marginBottom: 6, padding: '6px 10px', fontSize: '0.82rem' }}
                    />
                    <textarea
                      value={obMailBody}
                      onChange={e => setObMailBody(e.target.value)}
                      disabled={obBusy}
                      rows={5}
                      style={{ width: '100%', fontFamily: 'monospace', fontSize: '0.72rem', padding: 8 }}
                    />
                    <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', margin: '4px 0 8px' }}>
                      Vorschlagstext — editierbar. Wird in die Mail-Queue gelegt, nicht direkt versendet.
                    </p>
                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)', marginBottom: 4 }}>Vorschau (echte Deloitte-Mail):</div>
                    <div
                      style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 6, maxHeight: 280, overflow: 'auto', background: '#fff' }}
                      dangerouslySetInnerHTML={{
                        __html: obMailBody
                          .replace(/\{\{LOGO_URL\}\}/g, getCachedLogoBase64() || '')
                          .replace(/\{\{ORB_URL\}\}/g, getCachedOrbBase64() || ''),
                      }}
                    />
                  </div>
                )}
                {obWithMail && overbookModal.targets.length > 1 && (
                  // v13.0: Preview teilt sich den obMailBody-State mit der
                  // Modal-Open-useEffect — beide rendern den Body der
                  // ersten Person. Vorher wurde buildOverbookApologyEmail
                  // synchron im Render aufgerufen; seit der Template-DB-
                  // Lookup async ist, geht das nicht mehr direkt im JSX.
                  <div style={{ marginBottom: 10 }}>
                    <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', margin: '0 0 6px' }}>
                      Bei &bdquo;Alle&ldquo; wird der Standardtext je Person personalisiert versendet (eigene Wartelisten-Position). Vorschau am Beispiel der ersten Person:
                    </p>
                    <div
                      style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 6, maxHeight: 260, overflow: 'auto', background: '#fff' }}
                      dangerouslySetInnerHTML={{
                        __html: obMailBody
                          .replace(/\{\{LOGO_URL\}\}/g, getCachedLogoBase64() || '')
                          .replace(/\{\{ORB_URL\}\}/g, getCachedOrbBase64() || ''),
                      }}
                    />
                  </div>
                )}
                <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.88rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="checkbox" checked={obRemoveCalendar} onChange={e => setObRemoveCalendar(e.target.checked)} disabled={obBusy} />
                  Vom Kalendereintrag abmelden (falls vorhanden)
                </label>
              </>
            ) : (
              <>
                <h3 style={{ marginTop: 0 }}>Platz behalten ({overbookModal.targets.length})</h3>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                  Wie soll die Person ihren Platz behalten?
                </p>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: '0.85rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="radio" name="obkeep" checked={obKeepVariant === 'firstWaitlist'} onChange={() => setObKeepVariant('firstWaitlist')} disabled={obBusy} style={{ marginTop: 3 }} />
                  <span><strong>Erste(r) auf der Warteliste</strong> — rückt beim nächsten frei werdenden Platz der Gruppe garantiert als Erste(r) nach (risikoarm).</span>
                </label>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: '0.85rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="radio" name="obkeep" checked={obKeepVariant === 'active'} onChange={() => setObKeepVariant('active')} disabled={obBusy} style={{ marginTop: 3 }} />
                  <span><strong>Bleibt angemeldet</strong> (als Letzte(r)) — Gruppe bleibt +1, der nächste frei werdende Platz wird einmal nicht nachgerückt, bis die Überzahl absorbiert ist.</span>
                </label>
              </>
            )}
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 18 }}>
              <button className="btn btn-secondary" onClick={() => setOverbookModal(null)} disabled={obBusy}>
                Abbrechen
              </button>
              <button className="btn btn-primary" onClick={() => { runOverbookResolution().catch(() => { /* */ }); }} disabled={obBusy}>
                {obBusy ? 'Wird ausgeführt…' : (overbookModal.mode === 'confirm' ? 'Bestätigen & IDs neu vergeben' : 'Übernehmen & IDs neu vergeben')}
              </button>
            </div>
        </Modal>
  );
};

