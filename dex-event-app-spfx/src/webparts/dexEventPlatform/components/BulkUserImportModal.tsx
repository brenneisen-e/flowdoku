import * as React from 'react';
import { X } from './Icons';

/**
 * Generischer Massenimport-Dialog für Team-Felder im Event-Wizard.
 *
 * Wird aktuell genutzt für:
 *   - Co-Organizer (Schritt 1, Grundlagen)
 *   - Test-Team (Schritt 1, Grundlagen)
 *   - Check-In Team / QR-Scanner (Schritt 1, Grundlagen)
 *
 * Pattern eng angelehnt an den Audience-Massenimport (Sichtbarkeits-Reiter,
 * EventCreationPage `runBulkImport`/`resolveAmbiguous`). Wesentlicher Unterschied:
 * Audience speichert eine `,`-separierte Email-Liste in einem String-Feld, Teams
 * speichern parallele `Names[]` + `Emails[]`-Arrays. Daher liefert dieser Modal
 * keine fertig formatierte Liste zurück sondern ruft `onAdd(item)` für jeden
 * erfolgreich aufgelösten Eintrag auf — der Caller hängt das an seine Arrays an.
 *
 * Eager-Add-Verhalten: jeder Treffer löst sofort `onAdd` aus (statt erst beim
 * Schließen) — analog zum Audience-Modal, das auch eagerly `setAudience()` ruft.
 * So kann der User parallel zum Modal die Chips wachsen sehen und bei Mehrdeutigkeit
 * eine Auswahl-Entscheidung treffen, die direkt sichtbar wird.
 */

export interface BulkImportItem {
  email: string;
  displayName: string;
}

interface SearchHit {
  email: string;
  displayName: string;
  location?: string;
}

interface BulkImportReport {
  added: Array<{ lastname: string; firstname: string; email: string; originalInput?: string }>;
  alreadyIn: Array<{ lastname: string; firstname: string; email: string }>;
  notFound: string[];
  ambiguous: Array<{ input: string; matches: Array<{ email: string; displayName: string }> }>;
}

interface Props {
  open: boolean;
  onClose: () => void;
  /** Anzeigetitel im Modal-Header, z.B. "Massenimport — Co-Organizer". */
  title: string;
  /** Kurze Beschreibung was die Liste später bewirkt (rolle-spezifisch). */
  description?: React.ReactNode;
  /** Bereits in der Liste vorhandene Emails (Lowercase-Vergleich, für Doppel-Check). */
  existingEmails: string[];
  /** Tenant-User-Lookup (gleicher Service wie der Single-Picker). */
  searchUsers: (q: string) => Promise<SearchHit[]>;
  /** Wird pro erfolgreich gefundenem User aufgerufen — Caller appendet an seine Arrays. */
  onAdd: (item: BulkImportItem) => void;
}

const parseDisplayName = (displayName: string, fallbackEmail: string): { lastname: string; firstname: string } => {
  const dn = (displayName || '').trim();
  if (!dn) return { lastname: fallbackEmail || '', firstname: '' };
  if (dn.indexOf(',') >= 0) {
    const parts = dn.split(',').map(s => s.trim());
    return { lastname: parts[0] || dn, firstname: parts[1] || '' };
  }
  const words = dn.split(/\s+/);
  if (words.length >= 2) {
    return { lastname: words[words.length - 1], firstname: words.slice(0, -1).join(' ') };
  }
  return { lastname: dn, firstname: '' };
};

const sortByLastFirst = (a: { lastname: string; firstname: string }, b: { lastname: string; firstname: string }): number => {
  const ln = a.lastname.localeCompare(b.lastname, 'de', { sensitivity: 'base' });
  if (ln !== 0) return ln;
  return a.firstname.localeCompare(b.firstname, 'de', { sensitivity: 'base' });
};

const BulkUserImportModal: React.FC<Props> = ({ open, onClose, title, description, existingEmails, searchUsers, onAdd }) => {
  const [text, setText] = React.useState('');
  const [running, setRunning] = React.useState(false);
  const [report, setReport] = React.useState<BulkImportReport | null>(null);

  // Die Lokal-Snapshot-Menge der „bereits drin"-Emails. Wird inkrementell
  // erweitert während des Runs, damit Doppel-Eingaben innerhalb desselben
  // Massen-Pastes auch als alreadyIn erkannt werden.
  const buildExistingSet = React.useCallback((): Set<string> => {
    return new Set((existingEmails || []).map(e => (e || '').toLowerCase()));
  }, [existingEmails]);

  React.useEffect(() => {
    // Bei Öffnen jedes Mal frisch starten — sonst zeigt der Modal noch das
    // Ergebnis des letzten Imports beim Wiederöffnen.
    if (open) {
      setText('');
      setReport(null);
      setRunning(false);
    }
  }, [open]);

  const runImport = async (): Promise<void> => {
    if (!text.trim()) return;
    setRunning(true);
    setReport(null);
    const existingLc = buildExistingSet();

    // Zwei-Pass-Parser für Excel-/GAL-Pastes wie „Lastname, Firstname email@x.de;":
    //
    // 1. **Email-Pass** — Zeile-für-Zeile nach Email-SUBSTRINGS suchen (nicht
    //    Full-String-Match). Wenn eine Zeile eine oder mehrere Emails enthält,
    //    werden DIESE Emails als Identifier genommen — der Namens-Schnipsel
    //    drumherum (z.B. "Lastname, Firstname ") wird ignoriert, weil
    //    `searchUsers(email)` den korrekten Display-Namen ohnehin liefert.
    //    Vorher wurde die Zeile durch Komma-Split zerhackt: aus "Pagano,
    //    Giovanna gpagano@deloitte.de" wurden ["Pagano", "Giovanna gpagano@..."]
    //    — der erste Teil landete als Lastname-only-Suche in „mehrdeutig", der
    //    zweite scheiterte am Full-String-Email-Match.
    //
    // 2. **Name-Pass** — nur für Zeilen OHNE Email. Die ganze Zeile geht als
    //    Suchquery zu searchUsers. Falls 0 Treffer und Zeile enthält Komma:
    //    Retry mit Komma → Space (für „Lastname, Firstname"-Format).
    const SUBSTRING_EMAIL_RE = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
    const lines = text.split(/[\n\r]/).map(s => s.trim()).filter(Boolean);

    interface QueueItem { kind: 'email' | 'name'; value: string; }
    const queue: QueueItem[] = [];
    for (const line of lines) {
      const emails = line.match(SUBSTRING_EMAIL_RE);
      if (emails && emails.length > 0) {
        // Eine oder mehrere Emails auf der Zeile gefunden → 1 Eintrag pro Email.
        // Lowercase-Normalisierung beim Vergleich, aber Original-Casing fürs Display
        // bleibt erhalten (manche Tenants pushen ALLCAPS-Domains aus historischen
        // Gründen — das stört nicht und sieht im Report mit Original-Casing besser aus).
        for (const em of emails) queue.push({ kind: 'email', value: em.trim() });
      } else {
        // Keine Email → Zeile als Namens-Suchquery behandeln. Falls die Zeile mehrere
        // durch `;` oder `,` getrennte Namen enthält (z.B. „Mustermann, Max; Schmitz,
        // Anna") splitten wir konservativ — ABER nur wenn die Zeile MEHR als ein
        // Komma-getrenntes Paar hat. „Mustermann, Max" allein bleibt als Ganzes
        // erhalten, sonst würde Graph einzeln nach „Mustermann" und „Max" suchen.
        const semiSplit = line.split(';').map(s => s.trim()).filter(Boolean);
        if (semiSplit.length > 1) {
          for (const part of semiSplit) queue.push({ kind: 'name', value: part });
        } else {
          queue.push({ kind: 'name', value: line });
        }
      }
    }

    const added: BulkImportReport['added'] = [];
    const alreadyIn: BulkImportReport['alreadyIn'] = [];
    const notFound: string[] = [];
    const ambiguous: BulkImportReport['ambiguous'] = [];

    for (const item of queue) {
      const f = item.value;
      const fLc = f.toLowerCase();

      // Schon drin (Email-Identität, case-insensitive)?
      if (item.kind === 'email' && existingLc.has(fLc)) {
        try {
          const hits = await searchUsers(f);
          if (hits.length > 0) {
            const { lastname, firstname } = parseDisplayName(hits[0].displayName, f);
            alreadyIn.push({ lastname, firstname, email: f });
          } else {
            alreadyIn.push({ lastname: f, firstname: '', email: f });
          }
        } catch {
          alreadyIn.push({ lastname: f, firstname: '', email: f });
        }
        continue;
      }

      if (item.kind === 'email') {
        try {
          const hits = await searchUsers(f);
          const exact = hits.find(h => h.email && h.email.toLowerCase() === fLc);
          if (exact) {
            const { lastname, firstname } = parseDisplayName(exact.displayName, f);
            onAdd({ email: exact.email, displayName: exact.displayName });
            existingLc.add(exact.email.toLowerCase());
            added.push({ lastname, firstname, email: exact.email });
          } else {
            // Email passt syntaktisch aber Tenant kennt sie nicht — trotzdem übernehmen,
            // weil externe oder neu-aktivierte Accounts manuell legitim sein können.
            onAdd({ email: f, displayName: f });
            existingLc.add(fLc);
            added.push({ lastname: f, firstname: '', email: f });
          }
        } catch {
          onAdd({ email: f, displayName: f });
          existingLc.add(fLc);
          added.push({ lastname: f, firstname: '', email: f });
        }
        await new Promise(res => setTimeout(res, 80));
        continue;
      }

      // Name-Pass.
      const tryName = async (query: string): Promise<SearchHit[]> => {
        try { return await searchUsers(query); } catch { return []; }
      };
      let hits = await tryName(f);
      if ((!hits || hits.length === 0) && f.indexOf(',') >= 0) {
        // Retry: „Lastname, Firstname" → „Firstname Lastname" (Graph mag oft
        // diese Reihenfolge besser).
        const parts = f.split(',').map(s => s.trim()).filter(Boolean);
        if (parts.length >= 2) {
          const reordered = parts.slice(1).concat(parts[0]).join(' ');
          hits = await tryName(reordered);
        }
      }

      if (!hits || hits.length === 0) {
        notFound.push(f);
        await new Promise(res => setTimeout(res, 120));
        continue;
      }
      if (hits.length === 1) {
        const em = hits[0].email;
        if (em && !existingLc.has(em.toLowerCase())) {
          onAdd({ email: em, displayName: hits[0].displayName });
          existingLc.add(em.toLowerCase());
          const { lastname, firstname } = parseDisplayName(hits[0].displayName, em);
          added.push({ lastname, firstname, email: em, originalInput: f });
        } else {
          const { lastname, firstname } = parseDisplayName(hits[0].displayName, em || f);
          alreadyIn.push({ lastname, firstname, email: em || f });
        }
        await new Promise(res => setTimeout(res, 120));
        continue;
      }
      ambiguous.push({
        input: f,
        matches: hits.slice(0, 5).map(h => ({ email: h.email, displayName: h.displayName })),
      });
      await new Promise(res => setTimeout(res, 120));
    }

    added.sort(sortByLastFirst);
    alreadyIn.sort(sortByLastFirst);
    notFound.sort();

    setReport({ added, alreadyIn, notFound, ambiguous });
    setRunning(false);
  };

  const resolveAmbiguous = (input: string, email: string, displayName: string): void => {
    if (!report) return;
    if (existingEmails.map(e => e.toLowerCase()).indexOf(email.toLowerCase()) >= 0) {
      // Sollte nicht vorkommen weil Filter vor dem Klick schon greifen sollte —
      // aber zur Sicherheit no-op.
      setReport({ ...report, ambiguous: report.ambiguous.filter(a => a.input !== input) });
      return;
    }
    onAdd({ email, displayName });
    const { lastname, firstname } = parseDisplayName(displayName, email);
    const newAdded = [...report.added, { lastname, firstname, email, originalInput: input }];
    newAdded.sort(sortByLastFirst);
    setReport({
      ...report,
      ambiguous: report.ambiguous.filter(a => a.input !== input),
      added: newAdded,
    });
  };

  if (!open) return null;

  return (
    <div
      style={{
        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 1100,
        display: 'flex', alignItems: 'center', justifyContent: 'center',
      }}
      onClick={() => { if (!running) onClose(); }}
    >
      <div className="card" style={{ width: '90%', maxWidth: 720, maxHeight: '85vh', overflow: 'auto', padding: 24 }} onClick={e => e.stopPropagation()}>
        <div className="flex-between mb-16">
          <h3 style={{ margin: 0 }}>{title}</h3>
          <button
            type="button"
            onClick={() => { if (!running) onClose(); }}
            style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4 }}
            aria-label="Schließen"
          >
            <X size={18} />
          </button>
        </div>
        {description && (
          <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)', marginTop: 0, lineHeight: 1.55, marginBottom: 12 }}>
            {description}
          </div>
        )}
        <ul style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginTop: 0, marginBottom: 12, paddingLeft: 18, lineHeight: 1.5 }}>
          <li><strong>Email-Adressen</strong> (z.B. <code>vorname.nachname@deloitte.de</code>) werden direkt übernommen.</li>
          <li><strong>Namen</strong> werden im Deloitte-Tenant gesucht — eindeutige Treffer werden automatisch hinzugefügt, mehrdeutige Namen musst du unten manuell auflösen.</li>
          <li>Personen die <strong>schon in der Liste sind</strong>, werden übersprungen (keine Doppel-Einträge).</li>
          <li>Trennzeichen: <code>,</code>&nbsp; <code>;</code>&nbsp; <code>Tab</code> oder <code>Zeilenumbruch</code>.</li>
        </ul>
        <textarea
          className="form-input"
          style={{ width: '100%', minHeight: 160, fontFamily: 'monospace', fontSize: '0.8rem' }}
          placeholder="z.B.:&#10;max.mustermann@deloitte.de; erika.mustermann@deloitte.de&#10;Schmitz, Alexander, Kraus, Annika&#10;oder aus Excel kopiert (Tab-getrennt)"
          value={text}
          onChange={e => setText(e.target.value)}
          disabled={running}
        />

        {report && (
          <div style={{ marginTop: 16, padding: 12, background: 'var(--dex-gray-50)', borderRadius: 'var(--dex-radius)', fontSize: '0.8rem' }}>
            {report.added.length > 0 && (
              <div style={{ marginBottom: 10 }}>
                <strong style={{ color: 'var(--dex-green-dark)' }}>✓ Hinzugefügt ({report.added.length}):</strong>
                <ul style={{ margin: '4px 0 0 16px', padding: 0 }}>
                  {report.added.map((a, i) => (
                    <li key={`a-${i}`}>
                      <strong>{a.lastname}</strong>{a.firstname ? `, ${a.firstname}` : ''} <span style={{ color: 'var(--dex-gray-400)' }}>— {a.email}</span>
                    </li>
                  ))}
                </ul>
              </div>
            )}
            {report.alreadyIn.length > 0 && (
              <div style={{ marginBottom: 10 }}>
                <strong style={{ color: 'var(--dex-gray-500)' }}>— Bereits in der Liste ({report.alreadyIn.length}):</strong>
                <ul style={{ margin: '4px 0 0 16px', padding: 0, color: 'var(--dex-gray-500)' }}>
                  {report.alreadyIn.map((a, i) => (
                    <li key={`w-${i}`}>
                      <strong>{a.lastname}</strong>{a.firstname ? `, ${a.firstname}` : ''} <span>— {a.email}</span>
                    </li>
                  ))}
                </ul>
              </div>
            )}
            {report.notFound.length > 0 && (
              <div style={{ marginBottom: 10 }}>
                <strong style={{ color: 'var(--dex-red)' }}>✗ Nicht gefunden ({report.notFound.length}):</strong>
                <ul style={{ margin: '4px 0 0 16px', padding: 0, color: 'var(--dex-red)' }}>
                  {report.notFound.map((a, i) => <li key={`n-${i}`}>{a} — bitte manuell suchen oder als Email eintragen</li>)}
                </ul>
              </div>
            )}
            {report.ambiguous.length > 0 && (
              <div style={{ marginBottom: 10 }}>
                <strong style={{ color: 'var(--dex-orange, #ed8b00)' }}>? Mehrdeutig — bitte auswählen ({report.ambiguous.length}):</strong>
                {report.ambiguous.map((a, i) => (
                  <div key={`m-${i}`} style={{ marginTop: 6, padding: 8, background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 4 }}>
                    <div style={{ fontWeight: 600, marginBottom: 4 }}>&bdquo;{a.input}&ldquo;</div>
                    {a.matches.map((m, j) => (
                      <button
                        key={`mm-${i}-${j}`}
                        type="button"
                        onClick={() => resolveAmbiguous(a.input, m.email, m.displayName)}
                        style={{
                          display: 'block', width: '100%', textAlign: 'left', padding: '6px 8px',
                          marginTop: 4, background: '#fff', border: '1px solid var(--dex-gray-200)',
                          borderRadius: 4, cursor: 'pointer', fontSize: '0.8rem',
                        }}
                      >
                        <strong>{m.displayName}</strong> <span style={{ color: 'var(--dex-gray-400)' }}>{m.email}</span>
                      </button>
                    ))}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        <div style={{ display: 'flex', gap: 8, marginTop: 16, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
          <button
            type="button"
            className="btn btn-secondary"
            onClick={() => { if (!running) onClose(); }}
            disabled={running}
          >
            {report ? 'Verwerfen' : 'Schließen'}
          </button>
          {!report && (
            <button
              type="button"
              className="btn btn-primary"
              onClick={runImport}
              disabled={running || !text.trim()}
            >
              {running ? 'Suche läuft...' : 'Verarbeiten'}
            </button>
          )}
          {report && (
            <>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={runImport}
                disabled={running || !text.trim()}
              >
                Erneut verarbeiten
              </button>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => { setText(''); setReport(null); }}
                disabled={running}
              >
                Weitere Liste hinzufügen
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => { if (!running) onClose(); }}
                disabled={running}
              >
                Speichern und schließen
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
};

export default BulkUserImportModal;
