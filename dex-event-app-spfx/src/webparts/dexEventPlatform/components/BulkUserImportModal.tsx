import * as React from 'react';
import { AlertCircle, Check, Users } from './Icons';
import InternationalSearchToggle from './InternationalSearchToggle';
import Modal from './Modal';
import { InfoTooltip } from './InfoTooltip';
import { cx } from './dexUi';
import { useLanguage } from '../context/LanguageContext';
import { useRoles } from '../context/RoleContext';

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
  searchUsers: (q: string, includeIntl?: boolean) => Promise<SearchHit[]>;
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
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  // v30.82: E-Mail-Einträge vorab in EINEM Graph-Batch auflösen — vorher je
  // Adresse eine Personensuche plus Profil-Nachladen (bis zu fünf Aufrufe),
  // bei einem Paste mit 100 Adressen also Hunderte sequentielle Requests.
  // In der Handbuch-Vorschau (Stub-Context) fehlt die Funktion → Einzelweg.
  const { getBasicProfiles } = useRoles();
  const [text, setText] = React.useState('');
  const [running, setRunning] = React.useState(false);
  const [report, setReport] = React.useState<BulkImportReport | null>(null);
  const [includeIntl, setIncludeIntl] = React.useState(false);

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

    // v30.82: Vorab-Batch für alle E-Mail-Einträge (s. Kommentar oben).
    let pre: Record<string, { displayName: string; jobTitle: string; location: string }> = {};
    try {
      const mails = queue.filter(q => q.kind === 'email').map(q => q.value);
      if (mails.length > 0 && typeof getBasicProfiles === 'function') pre = await getBasicProfiles(mails);
    } catch { pre = {}; }

    for (const item of queue) {
      const f = item.value;
      const fLc = f.toLowerCase();

      // Schon drin (Email-Identität, case-insensitive)?
      if (item.kind === 'email' && existingLc.has(fLc)) {
        const preHit = pre[fLc];
        if (preHit && preHit.displayName) {
          const { lastname, firstname } = parseDisplayName(preHit.displayName, f);
          alreadyIn.push({ lastname, firstname, email: f });
          continue;
        }
        try {
          const hits = await searchUsers(f, includeIntl);
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
        const preHit = pre[fLc];
        if (preHit && preHit.displayName) {
          const { lastname, firstname } = parseDisplayName(preHit.displayName, f);
          onAdd({ email: f, displayName: preHit.displayName });
          existingLc.add(fLc);
          added.push({ lastname, firstname, email: f });
          continue;
        }
        try {
          const hits = await searchUsers(f, includeIntl);
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
        try { return await searchUsers(query, includeIntl); } catch { return []; }
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

  // v31.2: Der Dialog erzählt seinen Ablauf als drei nummerierte Schritte
  // (Einfügen → Prüfen → Ergebnis), der Knopf sitzt IN der Schritt-Zeile.
  // Vorher standen Erklärliste, Textfeld, Ergebnis-Absätze und fünf Knöpfe
  // untereinander, und „Verwerfen" versprach ein Rückgängig, das es nie gab —
  // die Treffer sind beim Prüfen bereits übernommen (Eager-Add, s. oben).
  const t = (de: string, en: string): string => (isDe ? de : en);
  const hasText = !!text.trim();
  const close = (): void => { if (!running) onClose(); };
  // Ein Ergebnis gilt erst als „erledigt", wenn nichts mehr zu klären ist —
  // mehrdeutige Namen brauchen einen Klick, nicht gefundene eine Korrektur.
  const openIssues = report ? report.ambiguous.length + report.notFound.length : 0;

  type RowKind = 'added' | 'alreadyIn' | 'notFound';
  const personCell = (a: { lastname: string; firstname: string }): React.ReactNode => (
    <><strong>{a.lastname}</strong>{a.firstname ? `, ${a.firstname}` : ''}</>
  );
  // Eine Tabelle statt drei Listen: gleiche Spalten, Status als Pill —
  // Reihenfolge wie bisher (hinzugefügt, schon drin, nicht gefunden).
  const rows: Array<{ kind: RowKind; name: React.ReactNode; email: string; from?: string }> = report ? [
    ...report.added.map(a => ({ kind: 'added' as RowKind, name: personCell(a), email: a.email, from: a.originalInput })),
    ...report.alreadyIn.map(a => ({ kind: 'alreadyIn' as RowKind, name: personCell(a), email: a.email })),
    ...report.notFound.map(n => ({ kind: 'notFound' as RowKind, name: <strong>{n}</strong>, email: '' })),
  ] : [];
  const PILL: Record<RowKind, { cls: string; de: string; en: string }> = {
    added: { cls: 'dex-ui-pill--green', de: 'Hinzugefügt', en: 'Added' },
    alreadyIn: { cls: 'dex-ui-pill--gray', de: 'Schon in der Liste', en: 'Already listed' },
    notFound: { cls: 'dex-ui-pill--red', de: 'Nicht gefunden', en: 'Not found' },
  };
  // Aufrufer geben die Beschreibung mal als String, mal als <p> — ein <p> im
  // <p>-Untertitel des Modals wäre ungültiges HTML, deshalb landet JSX im Körper.
  const descIsText = typeof description === 'string';

  return (
    <Modal
      open={open}
      onClose={close}
      maxWidth={720}
      dismissable={!running}
      ariaLabel={title}
      title={title}
      subtitle={descIsText ? description : undefined}
      icon={<Users size={20} />}
      footer={report ? (
        <>
          <button type="button" className="btn btn-secondary" onClick={() => { setText(''); setReport(null); }} disabled={running}>
            {t('Weitere Liste einfügen', 'Paste another list')}
          </button>
          <button type="button" className="btn btn-primary" onClick={close} disabled={running}>
            {t('Fertig', 'Done')}
          </button>
        </>
      ) : (
        <button type="button" className="btn btn-secondary" onClick={close} disabled={running}>
          {t('Abbrechen', 'Cancel')}
        </button>
      )}
    >
      {description && !descIsText && (
        <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>{description}</div>
      )}

      {/* Schritt 1 — Einfügen */}
      <div className={cx('dex-ui-step', hasText && 'is-done')} style={{ alignItems: 'flex-start' }}>
        <span className="dex-ui-step-num">1</span>
        <div className="dex-ui-step-body">
          <div className="dex-ui-step-title" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            {t('Namen oder E-Mail-Adressen einfügen', 'Paste names or email addresses')}
            <InfoTooltip text={isDe ? (
              <ul style={{ margin: 0, paddingLeft: 16 }}>
                <li><strong>E-Mail-Adressen</strong> (z.B. vorname.nachname@deloitte.de) werden direkt übernommen.</li>
                <li><strong>Namen</strong> werden im Deloitte-Verzeichnis gesucht — eindeutige Treffer kommen automatisch dazu, mehrdeutige klärst du in Schritt 3.</li>
                <li>Wer <strong>schon in der Liste</strong> steht, wird übersprungen — keine Doppel-Einträge.</li>
                <li>Trennzeichen: Komma, Semikolon, Tab oder Zeilenumbruch.</li>
              </ul>
            ) : (
              <ul style={{ margin: 0, paddingLeft: 16 }}>
                <li><strong>Email addresses</strong> (e.g. first.last@deloitte.de) are taken over directly.</li>
                <li><strong>Names</strong> are looked up in the Deloitte directory — unique hits are added automatically, ambiguous ones you resolve in step 3.</li>
                <li>People <strong>already on the list</strong> are skipped — no duplicates.</li>
                <li>Separators: comma, semicolon, tab or line break.</li>
              </ul>
            )} />
          </div>
          <div className="dex-ui-step-hint">
            {t('Aus Outlook oder Excel kopiert — eine Person je Zeile oder getrennt durch Komma, Semikolon oder Tab.',
              'Copied from Outlook or Excel — one person per line or separated by comma, semicolon or tab.')}
          </div>
          <textarea
            className="dex-ui-textarea"
            style={{ marginTop: 10, minHeight: 140, fontFamily: 'monospace', fontSize: '0.8rem' }}
            placeholder={isDe
              ? 'z.B.:\nmax.mustermann@deloitte.de; erika.mustermann@deloitte.de\nSchmitz, Alexander; Kraus, Annika\noder aus Excel kopiert (Tab-getrennt)'
              : 'e.g.:\nmax.mustermann@deloitte.de; erika.mustermann@deloitte.de\nSchmitz, Alexander; Kraus, Annika\nor copied from Excel (tab-separated)'}
            value={text}
            onChange={e => setText(e.target.value)}
            disabled={running}
          />
          <InternationalSearchToggle query={text} checked={includeIntl} onChange={setIncludeIntl} isDe={isDe} />
        </div>
      </div>

      {/* Schritt 2 — Prüfen und übernehmen (der Knopf ist der Schritt) */}
      <div className={cx('dex-ui-step', report ? 'is-done' : !hasText && 'is-pending')}>
        <span className="dex-ui-step-num">2</span>
        <div className="dex-ui-step-body">
          <div className="dex-ui-step-title">{t('Prüfen und übernehmen', 'Check and add')}</div>
          <div className="dex-ui-step-hint">
            {t('E-Mail-Adressen kommen direkt dazu, Namen werden im Verzeichnis gesucht. Wer schon in der Liste steht, wird übersprungen.',
              'Email addresses are added directly, names are looked up in the directory. People already on the list are skipped.')}
          </div>
        </div>
        <div className="dex-ui-step-action">
          <button
            type="button"
            className={cx('btn', report ? 'btn-secondary' : 'btn-primary', 'dex-ui-btn-sm')}
            onClick={runImport}
            disabled={running || !hasText}
          >
            {running ? t('Suche läuft…', 'Searching…') : report ? t('Erneut prüfen', 'Check again') : t('Prüfen und übernehmen', 'Check and add')}
          </button>
        </div>
      </div>

      {/* Schritt 3 — Ergebnis: erst offene Entscheidungen, dann die Tabelle */}
      <div className={cx('dex-ui-step', report ? (openIssues === 0 && 'is-done') : 'is-pending')} style={{ alignItems: 'flex-start' }}>
        <span className="dex-ui-step-num">3</span>
        <div className="dex-ui-step-body">
          <div className="dex-ui-step-title">{t('Ergebnis', 'Result')}</div>
          {!report ? (
            <div className="dex-ui-step-hint">
              {running
                ? t('Die Liste wird gerade aufgelöst …', 'Resolving the list …')
                : t('Erscheint nach der Prüfung — alles Eindeutige ist dann schon übernommen.', 'Appears after the check — everything unambiguous is already added by then.')}
            </div>
          ) : (
            <>
              <div className="dex-ui-inline" style={{ marginTop: 6 }}>
                <span className="dex-ui-pill dex-ui-pill--green">{report.added.length} {t('hinzugefügt', 'added')}</span>
                {report.alreadyIn.length > 0 && <span className="dex-ui-pill dex-ui-pill--gray">{report.alreadyIn.length} {t('schon in der Liste', 'already listed')}</span>}
                {report.notFound.length > 0 && <span className="dex-ui-pill dex-ui-pill--red">{report.notFound.length} {t('nicht gefunden', 'not found')}</span>}
                {report.ambiguous.length > 0 && <span className="dex-ui-pill dex-ui-pill--orange">{report.ambiguous.length} {t('zu klären', 'to resolve')}</span>}
              </div>
              <div className="dex-ui-step-hint" style={{ marginTop: 6 }}>
                {t('Hinzugefügte Personen stehen bereits in der Liste hinter dem Dialog — „Fertig“ schließt nur das Fenster.',
                  'Added people are already on the list behind this dialog — “Done” only closes the window.')}
              </div>

              {report.ambiguous.length > 0 && (
                <div className="dex-ui-callout dex-ui-callout--warn" style={{ flexDirection: 'column', gap: 10, marginTop: 10 }}>
                  <div><strong>{t('Mehrdeutig — wähle die richtige Person', 'Ambiguous — pick the right person')}</strong> ({report.ambiguous.length})</div>
                  {report.ambiguous.map((a, i) => (
                    <div key={`m-${i}`}>
                      <div style={{ fontWeight: 600, marginBottom: 6 }}>&bdquo;{a.input}&ldquo;</div>
                      <div className="dex-ui-stack" style={{ gap: 6 }}>
                        {a.matches.map((m, j) => (
                          <button
                            key={`mm-${i}-${j}`}
                            type="button"
                            className="dex-ui-choice"
                            style={{ padding: '8px 12px', alignItems: 'center' }}
                            onClick={() => resolveAmbiguous(a.input, m.email, m.displayName)}
                          >
                            <span className="dex-ui-choice-body">
                              <span className="dex-ui-choice-title">{m.displayName}</span>
                              <span className="dex-ui-choice-desc" style={{ overflowWrap: 'anywhere' }}>{m.email}</span>
                            </span>
                            <span className="dex-ui-choice-check"><Check size={12} /></span>
                          </button>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {report.notFound.length > 0 && (
                <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 10 }}>
                  <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                  <span>
                    {t('Nicht gefunden: Prüf die Schreibweise oder trag die E-Mail-Adresse ein und lass erneut prüfen — oder such die Person einzeln über das Suchfeld.',
                      'Not found: check the spelling or enter the email address and check again — or look the person up individually via the search field.')}
                  </span>
                </div>
              )}

              {rows.length > 0 && (
                <div className="dex-ui-table-wrap" style={{ marginTop: 10 }}>
                  <table className="dex-ui-table">
                    <thead>
                      <tr><th>{t('Person', 'Person')}</th><th>{t('E-Mail', 'Email')}</th><th>{t('Status', 'Status')}</th></tr>
                    </thead>
                    <tbody>
                      {rows.map((r, i) => (
                        <tr key={`${r.kind}-${i}`} style={r.kind === 'alreadyIn' ? { opacity: 0.7 } : undefined}>
                          <td>
                            {r.name}
                            {r.from && <div className="dex-ui-muted" style={{ fontSize: '0.72rem' }}>{t('gesucht als', 'searched as')} &bdquo;{r.from}&ldquo;</div>}
                          </td>
                          <td style={{ overflowWrap: 'anywhere', color: 'var(--dex-gray-500)' }}>{r.email || '—'}</td>
                          <td><span className={cx('dex-ui-pill', PILL[r.kind].cls)}>{isDe ? PILL[r.kind].de : PILL[r.kind].en}</span></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </Modal>
  );
};

export default BulkUserImportModal;
