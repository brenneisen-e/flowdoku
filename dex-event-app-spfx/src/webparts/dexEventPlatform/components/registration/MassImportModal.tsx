/* MassImportModal — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Massenimport von Teilnehmern (v18.13). Inhalt zeichengleich uebernommen;
 * die Anzeige-Bedingung (`massImportOpen`) ist beim Aufrufer geblieben.
 *
 * v31.2: Optik auf die dex-ui-Klassen umgestellt (docs/ui-leitfaden.md). Der
 * Ablauf Einfügen → Prüfen → Übernehmen stand vorher nur in der Knopf-Beschriftung;
 * jetzt als Schritt-Leiste oben. Benachrichtigung als Auswahl mit Folge-Zeile
 * statt Radio-Zeilen, Kopf/Fuß aus `Modal`. Handler und Bindungen unverändert. */
import * as React from 'react';
import Modal from '../Modal';
import { cx } from '../dexUi';
import { AlertCircle, Check, Search, Users, X } from '../Icons';
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

// v31.2: Die drei Ablauf-Schritte der Leiste oben. Reihenfolge = Bedienfolge.
const STEP_META = [
  { de: 'Einfügen', en: 'Paste', hintDe: 'Eine Person je Zeile', hintEn: 'One person per line' },
  { de: 'Prüfen', en: 'Review', hintDe: 'Vorschau aus dem Verzeichnis', hintEn: 'Preview from the directory' },
  { de: 'Übernehmen', en: 'Register', hintDe: 'Mit oder ohne Benachrichtigung', hintEn: 'With or without notification' },
] as const;

// v31.2: Die Benachrichtigung ist eine Entscheidung mit Folgen — deshalb je
// Option ein Titel UND eine Zeile, was die Personen dann bekommen.
const MODE_OPTIONS = [
  { v: 'mail', de: 'Bestätigungsmail und Outlook-Termin', en: 'Confirmation email and Outlook invite',
    descDe: 'Jede Person bekommt die Bestätigung per Mail und den Termin in ihren Kalender.', descEn: 'Everyone gets the confirmation by email and the invite in their calendar.' },
  { v: 'nomail', de: 'Nur Outlook-Termin', en: 'Outlook invite only',
    descDe: 'Keine Bestätigungsmail — der Termin landet trotzdem im Kalender.', descEn: 'No confirmation email — the invite still lands in the calendar.' },
  { v: 'silent', de: 'Still anmelden', en: 'Register silently',
    descDe: 'Keine Mail, kein Kalendereintrag — die Personen erfahren nichts davon.', descEn: 'No email, no calendar entry — the people are not notified.' },
] as const;

export const MassImportModal: React.FC<MassImportModalProps> = (p) => {
  const { locale, massImportBusy, massImportMode, massImportOpen, massImportProgress, massImportResolving, massImportResult, massImportRows, massImportStep, massImportText, resolveMassImport, runMassImport, setMassImportMode, setMassImportOpen, setMassImportResult, setMassImportRows, setMassImportStep, setMassImportText } = p;
  const isDe = locale === 'de';
  // v31.2: Der Aufrufer-State ist nullable (`useState<… | null>`), der Prop-Typ
  // nicht — die Wahrheitsprüfung ist deshalb dieselbe wie vor dem Umbau.
  const hasResult = !!massImportResult;
  const okCount = massImportRows.filter(r => r.status === 'ok').length;
  const dupCount = massImportRows.filter(r => r.status === 'duplicate').length;
  const nfCount = massImportRows.filter(r => r.status === 'notfound').length;
  const lineCount = massImportText.split('\n').filter(l => l.trim()).length;
  const removeRow = (idx: number): void => setMassImportRows(prev => prev.filter((_, i) => i !== idx));
  const stepStates: Array<'done' | 'active' | 'pending'> = massImportStep === 'input'
    ? ['active', 'pending', 'pending']
    : ['done', hasResult ? 'done' : 'active', hasResult ? 'done' : (massImportBusy ? 'active' : 'pending')];
  // v31.2: Tabellenkopf bleibt beim Scrollen stehen — bei 50 Zeilen weiß man sonst nicht mehr, welche Spalte was ist.
  const thSticky: React.CSSProperties = { position: 'sticky', top: 0, zIndex: 1 };
  const dim: React.CSSProperties = { color: 'var(--dex-gray-600)' };

  // v31.2: Fuß über `Modal.footer` — Primär-Knopf rechts außen, und die Zahl im
  // Knopf nennt die Folge („12 Personen anmelden"), nicht nur die Aktion.
  const footer = massImportStep === 'input' ? (
    <>
      <button type="button" className="btn btn-secondary" onClick={() => setMassImportOpen(false)} disabled={massImportResolving}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
      <button type="button" className="btn btn-primary" onClick={resolveMassImport} disabled={massImportResolving || !massImportText.trim()}>
        {massImportResolving ? (isDe ? 'Wird abgeglichen…' : 'Matching…') : (isDe ? 'Abgleichen und prüfen' : 'Match and review')}
      </button>
    </>
  ) : hasResult ? (
    <button type="button" className="btn btn-primary" onClick={() => setMassImportOpen(false)}>{isDe ? 'Schließen' : 'Close'}</button>
  ) : (
    <>
      <button type="button" className="btn btn-secondary" onClick={() => { setMassImportStep('input'); setMassImportResult(null); }} disabled={massImportBusy}>{isDe ? 'Zurück' : 'Back'}</button>
      <button type="button" className="btn btn-primary" onClick={runMassImport} disabled={massImportBusy || okCount === 0}>
        {massImportBusy ? (isDe ? 'Läuft…' : 'Running…') : (isDe ? (okCount === 1 ? '1 Person anmelden' : `${okCount} Personen anmelden`) : (okCount === 1 ? 'Register 1 person' : `Register ${okCount} people`))}
      </button>
    </>
  );

  return (
    <Modal
      open={massImportOpen} onClose={() => { if (!massImportBusy) setMassImportOpen(false); }} maxWidth={640} dismissable={!massImportBusy}
      ariaLabel={isDe ? 'Teilnehmer-Massenimport' : 'Bulk participant import'}
      title={isDe ? 'Mehrere Personen anmelden' : 'Register several people'}
      subtitle={isDe ? 'Liste einfügen, Treffer prüfen, anmelden — in drei Schritten.' : 'Paste a list, review the matches, register — in three steps.'}
      icon={<Users size={20} />} footer={footer}
    >
      <div className="dex-ui-grid-3" style={{ gap: 8 }} aria-label={isDe ? 'Ablauf' : 'Progress'}>
        {STEP_META.map((s, i) => { const st = stepStates[i]; return (
          <div key={s.en} className={cx('dex-ui-step', st === 'done' && 'is-done', st === 'pending' && 'is-pending')} style={{ padding: '8px 10px', gap: 10 }} aria-current={st === 'active' ? 'step' : undefined}>
            <span className="dex-ui-step-num" aria-hidden="true">{st === 'done' ? <Check size={14} /> : i + 1}</span>
            <div className="dex-ui-step-body">
              <div className="dex-ui-step-title">{isDe ? s.de : s.en}</div>
              <div className="dex-ui-step-hint">{isDe ? s.hintDe : s.hintEn}</div>
            </div>
          </div>
        ); })}
      </div>

      {massImportStep === 'input' && (
        <div className="dex-ui-field">
          <label className="dex-ui-label" htmlFor="dex-mass-import-text">
            {isDe ? 'Wen möchtest du anmelden?' : 'Who do you want to register?'}
            {lineCount > 0 && <span className="dex-ui-pill dex-ui-pill--gray" style={{ marginLeft: 'auto' }}>{isDe ? `${lineCount} Zeile${lineCount === 1 ? '' : 'n'}` : `${lineCount} line${lineCount === 1 ? '' : 's'}`}</span>}
          </label>
          <textarea
            id="dex-mass-import-text" className="dex-ui-textarea" value={massImportText} onChange={e => setMassImportText(e.target.value)}
            disabled={massImportResolving} rows={8} style={{ fontFamily: 'monospace', fontSize: '0.82rem' }}
            placeholder={'Mustermann, Max\nerika.musterfrau@deloitte.de\nMax Mustermann; max.mustermann@deloitte.de'}
          />
          <div className="dex-ui-help">
            {isDe
              ? <>Namen und/oder E-Mail-Adressen, <strong>eine Person je Zeile</strong>. Jede Zeile wird mit dem Deloitte-Verzeichnis abgeglichen; du prüfst danach Vorname, Nachname, Position, Standort und E-Mail, bevor jemand angemeldet wird.</>
              : <>Names and/or email addresses, <strong>one person per line</strong>. Each line is matched against the Deloitte directory; you then review first name, last name, position, location and email before anyone is registered.</>}
          </div>
          {massImportResolving && (
            <div className="dex-ui-callout dex-ui-callout--neutral" role="status" style={{ marginTop: 10 }}>
              <span className="dex-ui-callout-icon"><Search size={16} /></span><span>{isDe ? 'Verzeichnis-Abgleich läuft…' : 'Matching against the directory…'} {massImportProgress}</span>
            </div>
          )}
        </div>
      )}

      {massImportStep === 'preview' && (
        <>
          <div className="dex-ui-inline" role="status">
            <span className="dex-ui-pill dex-ui-pill--green"><Check size={12} /> {okCount} {isDe ? 'bereit zum Anmelden' : 'ready to register'}</span>
            {dupCount > 0 && <span className="dex-ui-pill dex-ui-pill--orange">{dupCount} {isDe ? 'doppelt in der Liste' : 'duplicate in the list'}</span>}
            {nfCount > 0 && <span className="dex-ui-pill dex-ui-pill--red">{nfCount} {isDe ? 'nicht gefunden' : 'not found'}</span>}
            <span className="dex-ui-muted">{isDe ? 'Zeilen, die nicht passen, entfernst du mit dem ×.' : 'Remove rows that don’t fit with the ×.'}</span>
          </div>
          <div className="dex-ui-table-wrap" style={{ maxHeight: 320, overflowY: 'auto' }}>
            <table className="dex-ui-table" style={{ minWidth: 560 }}>
              <thead>
                <tr>
                  <th style={thSticky}>{isDe ? 'Vorname' : 'First name'}</th>
                  <th style={thSticky}>{isDe ? 'Nachname' : 'Last name'}</th>
                  <th style={thSticky}>{isDe ? 'Position' : 'Position'}</th>
                  <th style={thSticky}>{isDe ? 'Standort' : 'Location'}</th>
                  <th style={thSticky}>E-Mail</th>
                  <th style={thSticky}>{isDe ? 'Status' : 'Status'}</th>
                  <th style={thSticky} aria-label={isDe ? 'Entfernen' : 'Remove'} />
                </tr>
              </thead>
              <tbody>
                {massImportRows.map((r, idx) => (
                  <tr key={`${r.email || r.raw}-${idx}`} style={{ opacity: r.status === 'ok' ? 1 : 0.6 }}>
                    <td>{r.firstName || '–'}</td>
                    <td>{r.lastName || '–'}</td>
                    <td style={dim}>{r.jobTitle || '–'}</td>
                    <td style={dim}>{r.location || '–'}</td>
                    <td style={dim}>{r.email || <span style={{ color: 'var(--dex-red, #c00)' }}>{r.raw}</span>}</td>
                    <td>
                      {r.status === 'ok' && <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'OK' : 'OK'}</span>}
                      {r.status === 'duplicate' && <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'Doppelt' : 'Duplicate'}</span>}
                      {r.status === 'notfound' && <span className="dex-ui-pill dex-ui-pill--red">{isDe ? 'Nicht gefunden' : 'Not found'}</span>}
                    </td>
                    <td style={{ padding: '4px 6px' }}>
                      <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => removeRow(idx)} disabled={massImportBusy} title={isDe ? 'Zeile entfernen' : 'Remove row'} aria-label={isDe ? 'Zeile entfernen' : 'Remove row'} style={{ width: 28, height: 28 }}>
                        <X size={14} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="dex-ui-section" style={{ margin: 0 }}>
            <div className="dex-ui-section-title">{isDe ? 'Was erfahren die Personen?' : 'What do the people receive?'}</div>
            <div className="dex-ui-stack" role="radiogroup" aria-label={isDe ? 'Benachrichtigung' : 'Notification'}>
              {MODE_OPTIONS.map(opt => { const on = massImportMode === opt.v; return (
                <button key={opt.v} type="button" role="radio" aria-checked={on} className={cx('dex-ui-choice', on && 'is-active')} onClick={() => setMassImportMode(opt.v)} disabled={massImportBusy} style={{ padding: '10px 14px', alignItems: 'center' }}>
                  <span className="dex-ui-choice-body">
                    <span className="dex-ui-choice-title" style={{ display: 'block' }}>{isDe ? opt.de : opt.en}</span>
                    <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{isDe ? opt.descDe : opt.descEn}</span>
                  </span>
                  <span className="dex-ui-choice-check" aria-hidden="true">{on && <Check size={12} />}</span>
                </button>
              ); })}
            </div>
          </div>

          {massImportBusy && (
            <div className="dex-ui-callout dex-ui-callout--neutral" role="status"><span>{isDe ? 'Anmeldung läuft…' : 'Registering…'} {massImportProgress}</span></div>
          )}
          {massImportResult && (
            <div className={cx('dex-ui-callout', massImportResult.failed.length > 0 ? 'dex-ui-callout--warn' : 'dex-ui-callout--success')} role="status">
              <span className="dex-ui-callout-icon">{massImportResult.failed.length > 0 ? <AlertCircle size={16} /> : <Check size={16} />}</span>
              <div>
                {isDe ? <><strong>{massImportResult.ok}</strong> Person(en) angemeldet.</> : <><strong>{massImportResult.ok}</strong> person(s) registered.</>}
                {massImportResult.failed.length > 0 && (
                  <div style={{ marginTop: 4 }}>{isDe ? 'Nicht angemeldet (bereits angemeldet / Fehler): ' : 'Not registered (already registered / error): '}{massImportResult.failed.join(', ')}</div>
                )}
              </div>
            </div>
          )}
        </>
      )}
    </Modal>
  );
};
