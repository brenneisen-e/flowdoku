/* MassImportModal — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Massenimport von Teilnehmern (v18.13). Inhalt zeichengleich uebernommen;
 * die Anzeige-Bedingung (`massImportOpen`) ist beim Aufrufer geblieben. */
import * as React from 'react';
import Modal from '../Modal';
import { X } from '../Icons';
import { Locale } from '../../context/LanguageContext';

/** Massenimport von Teilnehmern (v18.13). */
export interface MassImportModalProps {
  locale: Locale;
  massImportBusy: boolean;
  massImportMode: "mail" | "nomail" | "silent";
  massImportOpen: boolean;
  massImportProgress: string;
  massImportResolving: boolean;
  massImportResult: { ok: number; failed: string[]; };
  massImportRows: { email: string; firstName: string; lastName: string; jobTitle: string; location: string; status: "ok" | "duplicate" | "notfound"; raw: string; }[];
  massImportStep: "input" | "preview";
  massImportText: string;
  resolveMassImport: () => Promise<void>;
  runMassImport: () => Promise<void>;
  setMassImportMode: React.Dispatch<React.SetStateAction<"mail" | "nomail" | "silent">>;
  setMassImportOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setMassImportResult: React.Dispatch<React.SetStateAction<{ ok: number; failed: string[]; }>>;
  setMassImportRows: React.Dispatch<React.SetStateAction<{ email: string; firstName: string; lastName: string; jobTitle: string; location: string; status: "ok" | "duplicate" | "notfound"; raw: string; }[]>>;
  setMassImportStep: React.Dispatch<React.SetStateAction<"input" | "preview">>;
  setMassImportText: React.Dispatch<React.SetStateAction<string>>;
}
export const MassImportModal: React.FC<MassImportModalProps> = (p) => {
  const { locale, massImportBusy, massImportMode, massImportOpen, massImportProgress, massImportResolving, massImportResult, massImportRows, massImportStep, massImportText, resolveMassImport, runMassImport, setMassImportMode, setMassImportOpen, setMassImportResult, setMassImportRows, setMassImportStep, setMassImportText } = p;
  return (
        <Modal
          open={massImportOpen}
          onClose={() => { if (!massImportBusy) setMassImportOpen(false); }}
          maxWidth={600}
          padding={24}
          dismissable={!massImportBusy}
          ariaLabel={locale === 'de' ? 'Teilnehmer-Massenimport' : 'Bulk participant import'}
        >
          <h3 style={{ margin: '0 0 8px', fontSize: '1.1rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
            {locale === 'de' ? 'Teilnehmer-Massenimport' : 'Bulk participant import'}
          </h3>

          {massImportStep === 'input' && (
            <>
              <p style={{ margin: '0 0 12px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {locale === 'de'
                  ? <>Namen und/oder E-Mail-Adressen einfügen — <strong>eine Person pro Zeile</strong>. Das Tool gleicht jede Zeile mit dem Deloitte-Verzeichnis ab und zeigt dir danach eine <strong>Vorschau-Tabelle</strong> (Vorname, Nachname, Position, Standort, E-Mail) zum Prüfen, bevor angemeldet wird.</>
                  : <>Paste names and/or email addresses — <strong>one person per line</strong>. The tool matches each line against the Deloitte directory and then shows a <strong>preview table</strong> (first name, last name, position, location, email) to review before registering.</>}
              </p>
              <textarea
                className="form-input"
                value={massImportText}
                onChange={e => setMassImportText(e.target.value)}
                disabled={massImportResolving}
                rows={8}
                placeholder={'Mustermann, Max\nerika.musterfrau@deloitte.de\nMax Mustermann; max.mustermann@deloitte.de'}
                style={{ width: '100%', fontFamily: 'monospace', fontSize: '0.82rem', resize: 'vertical' }}
              />
              {massImportResolving && (
                <div style={{ marginTop: 10, fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                  {locale === 'de' ? 'Verzeichnis-Abgleich läuft…' : 'Matching against the directory…'} {massImportProgress}
                </div>
              )}
              <div style={{ marginTop: 18, display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
                <button className="btn btn-secondary" onClick={() => setMassImportOpen(false)} disabled={massImportResolving}>
                  {locale === 'de' ? 'Abbrechen' : 'Cancel'}
                </button>
                <button className="btn btn-primary" onClick={resolveMassImport} disabled={massImportResolving || !massImportText.trim()}>
                  {massImportResolving ? (locale === 'de' ? 'Wird abgeglichen…' : 'Matching…') : (locale === 'de' ? 'Abgleichen & Vorschau' : 'Match & preview')}
                </button>
              </div>
            </>
          )}

          {massImportStep === 'preview' && (() => {
            const okCount = massImportRows.filter(r => r.status === 'ok').length;
            const dupCount = massImportRows.filter(r => r.status === 'duplicate').length;
            const nfCount = massImportRows.filter(r => r.status === 'notfound').length;
            const removeRow = (idx: number): void => setMassImportRows(prev => prev.filter((_, i) => i !== idx));
            const tdStyle: React.CSSProperties = { padding: '6px 8px', fontSize: '0.8rem', borderBottom: '1px solid var(--dex-gray-100)' };
            const thStyle: React.CSSProperties = { padding: '6px 8px', fontSize: '0.75rem', textAlign: 'left', color: 'var(--dex-gray-500)', borderBottom: '2px solid var(--dex-gray-200)', textTransform: 'uppercase', letterSpacing: 0.4 };
            return (
              <>
                <p style={{ margin: '0 0 10px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {locale === 'de'
                    ? <><strong>{okCount}</strong> bereit zum Anmelden{dupCount > 0 ? `, ${dupCount} Duplikat(e)` : ''}{nfCount > 0 ? `, ${nfCount} nicht gefunden` : ''}. Prüfe die Tabelle — nicht passende Zeilen kannst du entfernen.</>
                    : <><strong>{okCount}</strong> ready to register{dupCount > 0 ? `, ${dupCount} duplicate(s)` : ''}{nfCount > 0 ? `, ${nfCount} not found` : ''}. Review the table — remove rows that don&apos;t fit.</>}
                </p>
                <div style={{ maxHeight: 320, overflowY: 'auto', overflowX: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 8 }}>
                  <table style={{ width: '100%', minWidth: 560, borderCollapse: 'collapse' }}>
                    <thead>
                      <tr>
                        <th style={thStyle}>{locale === 'de' ? 'Vorname' : 'First name'}</th>
                        <th style={thStyle}>{locale === 'de' ? 'Nachname' : 'Last name'}</th>
                        <th style={thStyle}>{locale === 'de' ? 'Position' : 'Position'}</th>
                        <th style={thStyle}>{locale === 'de' ? 'Standort' : 'Location'}</th>
                        <th style={thStyle}>E-Mail</th>
                        <th style={thStyle}>{locale === 'de' ? 'Status' : 'Status'}</th>
                        <th style={thStyle} />
                      </tr>
                    </thead>
                    <tbody>
                      {massImportRows.map((r, idx) => (
                        <tr key={`${r.email || r.raw}-${idx}`} style={{ opacity: r.status === 'ok' ? 1 : 0.6 }}>
                          <td style={tdStyle}>{r.firstName || '–'}</td>
                          <td style={tdStyle}>{r.lastName || '–'}</td>
                          <td style={{ ...tdStyle, color: 'var(--dex-gray-600)' }}>{r.jobTitle || '–'}</td>
                          <td style={{ ...tdStyle, color: 'var(--dex-gray-600)' }}>{r.location || '–'}</td>
                          <td style={{ ...tdStyle, color: 'var(--dex-gray-600)' }}>{r.email || <span style={{ color: 'var(--dex-red, #c00)' }}>{r.raw}</span>}</td>
                          <td style={tdStyle}>
                            {r.status === 'ok' && <span style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-green-dark, #4a7c1f)' }}>{locale === 'de' ? 'OK' : 'OK'}</span>}
                            {r.status === 'duplicate' && <span style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-orange-dark, #b35a00)' }}>{locale === 'de' ? 'Duplikat' : 'Duplicate'}</span>}
                            {r.status === 'notfound' && <span style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-red, #c00)' }}>{locale === 'de' ? 'Nicht gefunden' : 'Not found'}</span>}
                          </td>
                          <td style={tdStyle}>
                            <button type="button" onClick={() => removeRow(idx)} disabled={massImportBusy} title={locale === 'de' ? 'Zeile entfernen' : 'Remove row'} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-400)', padding: 2 }}>
                              <X size={14} />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div style={{ marginTop: 14, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
                    {locale === 'de' ? 'Benachrichtigung' : 'Notification'}
                  </div>
                  {([
                    { v: 'mail', de: 'Bestätigungsmail senden (+ Outlook-Termin)', en: 'Send confirmation email (+ Outlook invite)' },
                    { v: 'nomail', de: 'Ohne Bestätigungsmail (aber Outlook-Termin)', en: 'No confirmation email (but Outlook invite)' },
                    { v: 'silent', de: 'Stille Anmeldung (keine Mail, kein Kalendereintrag)', en: 'Silent registration (no email, no calendar invite)' },
                  ] as const).map(opt => (
                    <label key={opt.v} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                      <input type="radio" name="massImportMode" checked={massImportMode === opt.v} onChange={() => setMassImportMode(opt.v)} disabled={massImportBusy} />
                      {locale === 'de' ? opt.de : opt.en}
                    </label>
                  ))}
                </div>

                {massImportBusy && (
                  <div style={{ marginTop: 12, fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                    {locale === 'de' ? 'Anmeldung läuft…' : 'Registering…'} {massImportProgress}
                  </div>
                )}
                {massImportResult && (
                  <div style={{
                    marginTop: 12, padding: '10px 12px', borderRadius: 8, fontSize: '0.82rem',
                    background: massImportResult.failed.length > 0 ? 'rgba(237,139,0,0.08)' : 'rgba(134,188,37,0.10)',
                    border: `1px solid ${massImportResult.failed.length > 0 ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)'}`,
                    color: 'var(--dex-gray-700)', lineHeight: 1.5,
                  }}>
                    {locale === 'de' ? <><strong>{massImportResult.ok}</strong> Person(en) angemeldet.</> : <><strong>{massImportResult.ok}</strong> person(s) registered.</>}
                    {massImportResult.failed.length > 0 && (
                      <div style={{ marginTop: 4 }}>
                        {locale === 'de' ? 'Nicht angemeldet (bereits angemeldet / Fehler): ' : 'Not registered (already registered / error): '}
                        {massImportResult.failed.join(', ')}
                      </div>
                    )}
                  </div>
                )}

                <div style={{ marginTop: 18, display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
                  {massImportResult ? (
                    <button className="btn btn-primary" onClick={() => setMassImportOpen(false)}>
                      {locale === 'de' ? 'Schließen' : 'Close'}
                    </button>
                  ) : (
                    <>
                      <button className="btn btn-secondary" onClick={() => { setMassImportStep('input'); setMassImportResult(null); }} disabled={massImportBusy}>
                        {locale === 'de' ? 'Zurück' : 'Back'}
                      </button>
                      <button className="btn btn-primary" onClick={runMassImport} disabled={massImportBusy || okCount === 0}>
                        {massImportBusy ? (locale === 'de' ? 'Läuft…' : 'Running…') : (locale === 'de' ? `${okCount} anmelden` : `Register ${okCount}`)}
                      </button>
                    </>
                  )}
                </div>
              </>
            );
          })()}
        </Modal>
  );
};
