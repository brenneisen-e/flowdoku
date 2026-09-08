/**
 * v31.4: „Startnummern zuteilen" — freie Nummern an Personen ohne Nummer.
 *
 * ## Was diese Aktion tut — und was ausdrücklich nicht
 *
 * Sie **erfindet keine Nummer**. Frei ist ausschließlich, was DEX gesehen hat:
 * eine Nummer auf einer abgemeldeten Zeile, oder eine Nummer aus einer
 * festgehaltenen „abmelden"-Aufgabe des Imports. Es wird kein Nummernkreis
 * gespeichert und keiner aus Lücken abgeleitet (Nutzer-Entscheidung
 * 08.09.2026) — Nummern, die der Import als „unknown" gesehen hat, laufen beim
 * Veranstalter auf reale Personen und sind in DEX unsichtbar.
 *
 * Daraus folgt die ehrliche Erwartung, mit der der Dialog aufmacht: Sind beim
 * Veranstalter genauso viele Nummern gemeldet wie Leute angemeldet, entstehen
 * freie Nummern NUR durch Abmeldungen. Für alle anderen lautet die richtige
 * Antwort „beim Veranstalter nachmelden", nicht „Nummer zuteilen". Der Satz
 * steht oben, damit niemand eine leere Liste für einen Fehler hält.
 *
 * ## Die Regel, die über allem steht
 *
 * Nutzer-Ansage (08.09.2026, Vorabend des Laufs): „Ich habe ja heute
 * importiert und nun bereits 20 Nummern vergeben bzw. eingecheckt. Die dürfen
 * nun nicht geändert werden." Kandidat ist deshalb ausschließlich eine AKTIVE
 * Zeile mit LEERER Startnummer; eine Zeile, die eine Nummer trägt, wird nicht
 * angefasst — nicht zum Korrigieren, nicht zum Lückenschließen. Die Rechnung
 * dazu steht in `utils/b2runBibPool`; dieser Dialog darf sie nicht umgehen und
 * baut sie deshalb VOR dem Schreiben aus frisch gelesenen Daten neu auf.
 *
 * ## Warum die Verpflichtung vor den Daten geschrieben wird
 *
 * Eine Nummer in DEX ist nur die halbe Zuteilung: Beim Veranstalter läuft sie
 * weiter auf die abgemeldete Person, und niemand außer dem Organizer kann das
 * ummelden. Deshalb wird ZUERST die Aufgabenliste (`_b2runTodo`) geschrieben
 * und erst danach die Teilnehmerzeile. Geht dabei etwas schief, sieht man eine
 * Aufgabe ohne Nummer — das fällt am Tisch auf. Eine Nummer ohne Aufgabe
 * würde niemandem auffallen (der v30.54-Fehler).
 *
 * ## Warum der Entzug zuerst kommt
 *
 * Sonst gilt: Der unumkehrbare Schritt kommt zuletzt (CLAUDE.md). Eine
 * Startnummer ist aber ein AUSSCHLIESSLICHES Recht — zwei Zeilen mit
 * derselben Nummer sind zwei Personen mit demselben Zettel am Einlass.
 * Deshalb wird die abgemeldete Altzeile ZUERST und GEPRÜFT geleert; scheitert
 * das, wird die neue Zeile gar nicht geschrieben. Der schlimmste Ausgang ist
 * dann „Nummer trägt niemand" — sichtbar und behebbar.
 */
import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import Modal from '../Modal';
import { useDialog } from '../../context/DialogContext';
import { useEvents } from '../../context/EventContext';
import { useCurrentUser } from '../../context/UserContext';
import { DeloitteEvent } from '../../types';
import { EventService, REG_LIST_NAME, SPRegistration } from '../../services/EventService';
import { bibKey, buildBibPool, proposeBibAssignments, BibPool, BibProposal, FreeBib } from '../../utils/b2runBibPool';
import { StoredB2RunTodo, b2runNameOf, mergeStoredTodos, cleanDoneKeys, deriveB2RunTodos } from '../../utils/b2runTodos';
import { isThrottled } from '../../utils/spThrottle';
import { cx } from '../dexUi';
import { AlertCircle, Check, Download, Hash } from '../Icons';

/** Status, die als „läuft mit" zählen — dieselbe enge Liste wie im Rechenmodul. */
const ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
const DONE_UNREG = 'unregister|';
const DONE_REG = 'register|';

const lc = (s: string | undefined | null): string => (s || '').toLowerCase().trim();
const groupOf = (r: SPRegistration | undefined): string =>
  r ? String(r.StarterType || r.PreferredStarterType || '').trim() : '';

/**
 * Widersprechen sich zwei Startblock-Angaben? Nur für die BESCHRIFTUNG.
 *
 * Die verbindliche Regel — bei geteilten Kapazitäten wird gruppenfremd nie
 * automatisch vorgeschlagen — steckt in `proposeBibAssignments`
 * (`blockMismatch`). Hier geht es allein um den Warnsatz an einer von Hand
 * gewählten Nummer; verglichen wird auf Enthaltensein, weil in DEX
 * „Funstarter" steht und beim Veranstalter „17:00 Uhr Funstarter (Grün)".
 */
function blockDiffers(a: string, b: string): boolean {
  const x = lc(a), y = lc(b);
  if (!x || !y) return false;
  return x.indexOf(y) < 0 && y.indexOf(x) < 0;
}

// v31.2/v31.4: Dieselben kleinen Bausteine wie im Import-Fenster (Tabelle,
// Zeile, dicktengleiche Nummer, durchgestrichener Name, Ablauf-Zeile,
// Kennzahl). Sie sind dort nicht exportiert; sie hier erneut aufzusetzen ist
// bewusst — am Vorabend des Laufs die funktionierende Import-Datei umzubauen,
// nur um drei Hilfskomponenten zu teilen, wäre das teurere Risiko.
const Tbl = (p: { head: string[]; children: React.ReactNode }): React.ReactElement => (
  <div className="dex-ui-table-wrap">
    <table className="dex-ui-table">
      <thead><tr>{p.head.map(h => <th key={h}>{h}</th>)}</tr></thead>
      <tbody>{p.children}</tbody>
    </table>
  </div>
);
const Row = (p: { children: React.ReactNode }): React.ReactElement => (
  <tr>{React.Children.map(p.children, (c, i) => <td key={i}>{c}</td>)}</tr>
);
const Bib = (p: { children: React.ReactNode }): React.ReactElement => (
  <span style={{ fontFamily: 'monospace', fontWeight: 700 }}>{p.children}</span>
);
const Gone = (p: { children: React.ReactNode }): React.ReactElement => (
  <span style={{ textDecoration: 'line-through', color: 'var(--dex-gray-500)' }}>{p.children}</span>
);
const Step = (p: { n: number; done?: boolean; pending?: boolean; title: string; hint: React.ReactNode; children: React.ReactNode }): React.ReactElement => (
  <div className={cx('dex-ui-step', p.pending && 'is-pending', p.done && 'is-done')}>
    <span className="dex-ui-step-num">{p.done ? <Check size={14} /> : p.n}</span>
    <div className="dex-ui-step-body">
      <div className="dex-ui-step-title">{p.title}</div>
      <div className="dex-ui-step-hint">{p.hint}</div>
    </div>
    <div className="dex-ui-step-action">{p.children}</div>
  </div>
);
const Kpi = (p: { value: number; label: string; tone?: string | false }): React.ReactElement => (
  <div className={cx('dex-ui-kpi', p.tone)}>
    <div className="dex-ui-kpi-value">{p.value}</div>
    <div className="dex-ui-kpi-label">{p.label}</div>
  </div>
);

/** Der gelesene Stand — alles, worauf die Rechnung beruht, an EINER Stelle.
 *  `null` bei den Aufgaben heißt „nicht lesbar" und sperrt (nie `[]`). */
interface Snapshot {
  regs: SPRegistration[];
  /** Ungefilterte Zahl der Listeneinträge; -1 = nicht ermittelbar = Sperre. */
  itemCount: number;
  readFailed: boolean;
  storedTodos: StoredB2RunTodo[] | null;
  doneKeys: string[] | null;
}

/** Ein bestätigtes Paar: diese Nummer an diese Person. */
interface Pair { free: FreeBib; target: SPRegistration; bib: string; key: string }

interface WriteResult {
  ok: number;
  /** Namentliche Fehlschläge und Aussortierungen — nie stillschweigend. */
  failed: string[];
  dropped: string[];
  abortNote: string;
  todoSaved: boolean;
  rows: Array<{ bib: string; toName: string; fromName: string; block: string; toEmail: string; status: string }>;
}

/**
 * Abgehakte „abmelden"-Haken, die der Organizer ausdrücklich freigegeben hat,
 * aus der Haken-Liste nehmen.
 *
 * Nutzer-Entscheidung (08.09.2026): Solche Nummern bleiben angezeigt und
 * gesperrt; ein zweiter Klick mit Warnung gibt sie frei. Technisch ist die
 * Freigabe genau das hier — der Haken wird für die Rechnung ausgeblendet,
 * NICHT gelöscht. Verglichen wird über `bibKey`, damit ein Haken auf „0412"
 * auch „412" freigibt.
 */
function applyUnlock(doneKeys: string[] | null, unlocked: string[]): string[] | null {
  if (doneKeys === null) return null;
  if (unlocked.length === 0) return doneKeys;
  return doneKeys.filter(k =>
    k.indexOf(DONE_UNREG) !== 0 || unlocked.indexOf(bibKey(k.substring(DONE_UNREG.length))) < 0);
}

export default function B2RunAssignBibsModal(props: {
  event: DeloitteEvent;
  service: EventService;
  onClose: () => void;
  onDone: () => void;
}): React.ReactElement {
  const { getAllRegistrations, refreshEvents, events, childEventsOf } = useEvents();
  const { showAlert, confirmDialog } = useDialog();
  const { currentUser } = useCurrentUser();
  const [loading, setLoading] = React.useState(true);
  const [busy, setBusy] = React.useState(false);
  const [progress, setProgress] = React.useState('');
  const [snapshot, setSnapshot] = React.useState<Snapshot | null>(null);
  /** Freigegebene Nummern (als `bibKey`) — s. `applyUnlock`. */
  const [unlocked, setUnlocked] = React.useState<string[]>([]);
  /** Teilnehmer-Item-Id → Schlüssel der gewählten Nummer. Fehlender Eintrag =
   *  noch nicht entschieden, `''` = ausdrücklich „keine". */
  const [choice, setChoice] = React.useState<Record<number, string>>({});
  const [result, setResult] = React.useState<WriteResult | null>(null);
  const [xlsxBusy, setXlsxBusy] = React.useState(false);

  const eventId = props.event.id;
  const subsiteUrl = props.event.subsiteUrl || '';

  /**
   * Alles lesen, was die Rechnung braucht — Zeilen, ungefilterte Zeilenzahl,
   * Aufgaben, Haken.
   *
   * Der `ItemCount`-Abgleich ist der wichtigste Teil: Die Teilnehmerliste läuft
   * mit Element-Sicherheit (ReadSecurity=2). Wer ohne „Manage Lists" liest,
   * bekommt ein FEHLERFREIES HTTP 200 mit nur der eigenen Zeile — `onHttpError`
   * feuert nicht, und dann sieht JEDE Nummer frei aus. Muster: `getActiveCounts`
   * in services/events/seats.ts.
   */
  const loadAll = React.useCallback(async (): Promise<Snapshot> => {
    let readFailed = false;
    const regsP = getAllRegistrations(eventId, () => { readFailed = true; });
    const countP = (async (): Promise<number> => {
      if (!subsiteUrl) return -1;
      try {
        const r = await props.service._sp.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')?$select=ItemCount`,
          SPHttpClient.configurations.v1,
        );
        if (!r.ok) return -1;
        const d = await r.json();
        const raw = (d && d.ItemCount !== undefined) ? d.ItemCount : (d && d.d ? d.d.ItemCount : undefined);
        const n = typeof raw === 'number' ? raw : parseInt(String(raw === undefined || raw === null ? '' : raw), 10);
        return isFinite(n) && n >= 0 ? n : -1;
      } catch { return -1; }
    })();
    const both = await Promise.all([regsP, countP]);
    const regs = both[0];
    const itemCount = both[1];
    // v30.56: IMMER den frischesten Event-Stand nehmen — das über die Props
    // hereingereichte Objekt kann älter sein als der Kontext.
    const liveEv = events.filter(e => e.id === eventId)[0] || props.event;
    let storedTodos: StoredB2RunTodo[] | null = [];
    let doneKeys: string[] | null = [];
    const raw = (liveEv.emailTemplateOverrides || '').trim();
    if (raw) {
      try {
        const o = JSON.parse(raw);
        if (!o || typeof o !== 'object') {
          storedTodos = null; doneKeys = null;
        } else {
          // Vorhanden, aber kein Array = kaputt = unbekannt. Ein fehlender
          // Schlüssel dagegen heißt wirklich „es gibt noch keine Aufgaben".
          if (o._b2runTodo !== undefined) {
            storedTodos = Array.isArray(o._b2runTodo)
              ? o._b2runTodo.filter((x: unknown) => !!x && typeof x === 'object')
              : null;
          }
          if (o._b2runTodoDone !== undefined) {
            doneKeys = Array.isArray(o._b2runTodoDone)
              ? o._b2runTodoDone.filter((x: unknown) => typeof x === 'string')
              : null;
          }
        }
      } catch { storedTodos = null; doneKeys = null; }
    }
    return { regs, itemCount, readFailed, storedTodos, doneKeys };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [eventId, subsiteUrl, events]);

  const hasChildren = childEventsOf(eventId).length > 0;
  const isChild = !!props.event.parentEventId;
  const splitGroups = !!(props.event.durchstarterCapacity || props.event.funstarterCapacity);

  const buildFrom = React.useCallback((s: Snapshot, unl: string[]): BibPool => buildBibPool({
    regs: s.regs,
    itemCount: s.itemCount,
    readFailed: s.readFailed,
    storedTodos: s.storedTodos,
    doneKeys: applyUnlock(s.doneKeys, unl),
    hasChildren,
    isChild,
    splitGroups,
  }), [hasChildren, isChild, splitGroups]);

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const s = await loadAll();
        if (!cancelled) setSnapshot(s);
      } catch {
        if (!cancelled) {
          setSnapshot({ regs: [], itemCount: -1, readFailed: true, storedTodos: null, doneKeys: null });
        }
      } finally { if (!cancelled) setLoading(false); }
    })().catch(() => { /* oben behandelt */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [eventId]);

  const pool = React.useMemo(
    () => (snapshot ? buildFrom(snapshot, unlocked) : null),
    [snapshot, unlocked, buildFrom]);
  const proposals = React.useMemo(() => (pool ? proposeBibAssignments(pool) : []), [pool]);

  /**
   * Vorgehakt ist ausschließlich die aufgezeichnete Nachrück-Kette.
   *
   * Sie steht in den Daten (`ReplacedByParticipantEmail`, seit v17.15) — alles
   * andere ist erschlossen und wird nur VORGESCHLAGEN. Bestehende
   * Entscheidungen werden nie überschrieben, auch nicht das ausdrückliche
   * „keine".
   */
  React.useEffect(() => {
    if (!pool || !pool.ok || result) return;
    setChoice(prev => {
      const next: Record<number, string> = {};
      let changed = false;
      Object.keys(prev).forEach(k => { next[Number(k)] = prev[Number(k)]; });
      const taken: string[] = Object.keys(next).map(k => next[Number(k)]).filter(Boolean);
      for (const p of proposals) {
        if (p.stufe !== 'chain' || !p.target) continue;
        if (next[p.target.Id] !== undefined) continue;
        if (taken.indexOf(p.key) >= 0) continue;
        next[p.target.Id] = p.key;
        taken.push(p.key);
        changed = true;
      }
      return changed ? next : prev;
    });
  }, [pool, proposals, result]);

  const freeOf = React.useCallback((key: string): FreeBib | undefined =>
    (pool ? pool.free.filter(f => f.key === key)[0] : undefined), [pool]);
  const proposalFor = React.useCallback((regId: number): BibProposal | undefined =>
    proposals.filter(p => p.target && p.target.Id === regId)[0], [proposals]);

  /** Die aktuell geplanten Paare — die EINE Quelle für Zähler, Knopf, Excel
   *  und Schreibvorgang. */
  const planned: Pair[] = React.useMemo(() => {
    if (!pool || !pool.ok) return [];
    const out: Pair[] = [];
    const used: string[] = [];
    for (const c of pool.candidates) {
      const key = choice[c.Id];
      if (!key) continue;
      if (used.indexOf(key) >= 0) continue; // kann die Oberfläche nicht erzeugen — Gürtel und Hosenträger
      const f = pool.free.filter(x => x.key === key)[0];
      if (!f) continue;
      used.push(key);
      out.push({ free: f, target: c, bib: f.bib, key });
    }
    return out;
  }, [pool, choice]);

  /** Zeilen, die eine Nummer TRAGEN — sie bleiben unverändert, und das muss
   *  dastehen (Nutzer-Ansage: „Die dürfen nun nicht geändert werden."). */
  const untouched = React.useMemo(() => {
    if (!snapshot) return 0;
    return snapshot.regs.filter(r => ACTIVE_STATI.indexOf(r.Status) >= 0 && String(r.Startnummer || '').trim()).length;
  }, [snapshot]);

  const actor = {
    name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email,
    email: currentUser.email,
  };

  const unlockBib = async (bib: string): Promise<void> => {
    const okGo = await confirmDialog(
      `Startnummer ${bib} ist beim Veranstalter bereits ABGEMELDET — der Haken dazu steht in deiner Aufgabenliste. `
      + 'Nur freigeben, wenn der Zettel noch da ist UND du sie beim Veranstalter erneut anmeldest. '
      + 'Sonst läuft am Lauftag jemand mit einer Nummer, die in der Wertung nicht existiert.',
      { title: 'Abgemeldete Nummer freigeben?', confirmLabel: 'Freigeben', danger: true });
    if (!okGo) return;
    setUnlocked(prev => (prev.indexOf(bibKey(bib)) >= 0 ? prev : prev.concat(bibKey(bib))));
  };

  const write = async (): Promise<void> => {
    if (!snapshot || !pool || !pool.ok || busy || result) return;
    if (!subsiteUrl) {
      showAlert('Dieses Event hat keine Teilnehmerliste — es kann nichts geschrieben werden.', { variant: 'error' });
      return;
    }
    const wanted = planned.slice();
    if (wanted.length === 0) return;
    const okGo = await confirmDialog(
      `${wanted.length} Startnummer${wanted.length === 1 ? '' : 'n'} zuteilen. `
      + `Alle ${wanted.length === 1 ? 'trägt' : 'tragen'} noch einen anderen Namen und ${wanted.length === 1 ? 'muss' : 'müssen'} am Check-in-Tisch überklebt werden. `
      + `Beim Veranstalter musst du danach ${wanted.length} Ummeldung${wanted.length === 1 ? '' : 'en'} vornehmen — DEX legt dir die Aufgaben dafür an.`,
      { title: 'Startnummern zuteilen?', confirmLabel: 'Jetzt zuteilen' });
    if (!okGo) return;

    setBusy(true);
    const failed: string[] = [];
    const dropped: string[] = [];
    const rows: WriteResult['rows'] = [];
    let abortNote = '';
    let todoSaved = false;
    let doneCount = 0;
    try {
      // ---- 0. Alles neu lesen und JEDES Paar erneut prüfen -----------------
      // Zwischen Vorschau und Klick kann sich alles geändert haben: Jemand
      // meldet sich an, ein zweiter Organizer trägt eine Nummer ein, eine
      // Abmeldung wird zurückgenommen. Ein Paar, das der frische Stand nicht
      // mehr hergibt, wird NAMENTLICH verworfen — nicht stillschweigend
      // geschrieben.
      setProgress('Stand wird neu gelesen…');
      const fresh = await loadAll();
      const freshPool = buildFrom(fresh, unlocked);
      if (!freshPool.ok) {
        abortNote = `Abgebrochen, nichts geschrieben: ${freshPool.hardBlock}`;
        setResult({ ok: 0, failed, dropped, abortNote, todoSaved: false, rows });
        return;
      }
      const okPairs: Pair[] = [];
      const usedKeys: string[] = [];
      for (const w of wanted) {
        const f = freshPool.free.filter(x => x.key === w.key)[0];
        const t = freshPool.candidates.filter(c => c.Id === w.target.Id)[0];
        const who = b2runNameOf(w.target);
        if (!f) { dropped.push(`${w.bib} für ${who} — die Nummer ist nicht mehr frei.`); continue; }
        if (!t) { dropped.push(`${w.bib} für ${who} — die Person ist kein Kandidat mehr (trägt inzwischen eine Nummer, ist abgemeldet oder hat eine offene Aufgabe).`); continue; }
        if (f.source !== w.free.source || (f.fromRow ? f.fromRow.Id : 0) !== (w.free.fromRow ? w.free.fromRow.Id : 0)) {
          dropped.push(`${w.bib} für ${who} — die Nummer steht inzwischen auf einer anderen Zeile.`);
          continue;
        }
        if (usedKeys.indexOf(f.key) >= 0) { dropped.push(`${w.bib} für ${who} — die Nummer war schon vergeben.`); continue; }
        usedKeys.push(f.key);
        okPairs.push({ free: f, target: t, bib: f.bib, key: f.key });
      }
      if (okPairs.length === 0) {
        abortNote = 'Es ist kein Paar übrig geblieben — es wurde nichts geschrieben.';
        setResult({ ok: 0, failed, dropped, abortNote, todoSaved: false, rows });
        return;
      }

      // ---- 1. Spalte sicherstellen, BEVOR eine Verpflichtung entsteht ------
      setProgress('Spalte „Startnummer" wird geprüft…');
      const colOk = await props.service.ensureStartNumberColumn(subsiteUrl);
      if (!colOk) {
        abortNote = 'Die Spalte „Startnummer" konnte in der Teilnehmerliste nicht angelegt werden. Ohne sie kann nichts geschrieben werden — es wurde nichts geändert.';
        setResult({ ok: 0, failed, dropped, abortNote, todoSaved: false, rows });
        return;
      }

      // ---- 2. Verpflichtung zuerst ----------------------------------------
      setProgress('Aufgabenliste wird gespeichert…');
      const nowIso = new Date().toISOString();
      const nowDe = new Date().toLocaleDateString('de-DE');
      const added: StoredB2RunTodo[] = okPairs.map(p => {
        const to = b2runNameOf(p.target);
        const chain = !!(p.free.chainTarget && p.free.chainTarget.Id === p.target.Id);
        return {
          key: `${chain ? 'transfer' : 'assign'}|${p.bib}|${lc(p.target.ParticipantEmail)}`,
          kind: chain ? 'transfer' : 'assign',
          bib: p.bib,
          certain: chain,
          ts: nowIso,
          fromName: p.free.fromName,
          fromEmail: p.free.fromEmail,
          toName: to,
          toEmail: p.target.ParticipantEmail,
          action: `Startnummer ${p.bib} beim Veranstalter von ${p.free.fromName} auf ${to} ummelden — von dir zugeteilt am ${nowDe}. Der Zettel trägt noch den alten Namen.`,
        };
      });
      const assignedKeys = okPairs.map(p => p.key);
      const servedEmails = okPairs.map(p => lc(p.target.ParticipantEmail)).filter(Boolean);
      // Was diese Zuteilung erledigt: „Nummer abmelden" (sie bekommt jemanden)
      // und „Person nachmelden" (sie hat jetzt eine Nummer). Dieselbe Regel
      // gilt für die Aufgaben UND für die Haken — sonst bleibt ein Haken als
      // Karteileiche liegen und markiert die nächste Aufgabe zur selben Nummer
      // sofort als erledigt.
      const obsolete = (key: string): boolean => {
        if (key.indexOf(DONE_UNREG) === 0) return assignedKeys.indexOf(bibKey(key.substring(DONE_UNREG.length))) >= 0;
        if (key.indexOf(DONE_REG) === 0) return servedEmails.indexOf(key.substring(DONE_REG.length)) >= 0;
        return false;
      };
      const stored = fresh.storedTodos || [];
      const removeKeys = stored.filter(t => obsolete(t.key)).map(t => t.key);
      const merged = mergeStoredTodos(stored, added, removeKeys);
      const r1 = await props.service.patchEventOverridesValueEx(Number(eventId), '_b2runTodo', merged);
      if (!r1.ok) {
        abortNote = `Die Aufgabenliste konnte nicht gespeichert werden (HTTP ${r1.status}${r1.detail ? `: ${r1.detail}` : ''}). `
          + 'Es wurde KEINE Startnummer geschrieben — eine Nummer ohne Ummelde-Aufgabe würde niemandem auffallen.';
        setResult({ ok: 0, failed, dropped, abortNote, todoSaved: false, rows });
        return;
      }
      // `cleanDoneKeys` braucht die VOLLSTÄNDIGE Aufgabenliste (festgehalten
      // UND abgeleitet) — sonst löscht es die Haken aller abgeleiteten
      // Aufgaben, und der Organizer hakt sie morgen erneut ab.
      const allTodos = merged.concat(deriveB2RunTodos(fresh.regs));
      const nextDone = cleanDoneKeys(fresh.doneKeys || [], allTodos).filter(k => !obsolete(k));
      const r2 = await props.service.patchEventOverridesValueEx(Number(eventId), '_b2runTodoDone', nextDone);
      if (!r2.ok) {
        abortNote = `Die Haken zur Aufgabenliste konnten nicht gespeichert werden (HTTP ${r2.status}${r2.detail ? `: ${r2.detail}` : ''}). `
          + 'Es wurde KEINE Startnummer geschrieben. Die Aufgaben stehen bereits — bitte gleich noch einmal versuchen.';
        // Den lokalen Event-Stand trotzdem nachziehen: `_b2runTodo` IST
        // geschrieben, und ein zweiter Versuch soll auf dem neuen Stand
        // aufsetzen statt auf dem alten (v30.56).
        try { await refreshEvents(); } catch { /* beim nächsten Laden */ }
        setResult({ ok: 0, failed, dropped, abortNote, todoSaved: true, rows });
        return;
      }
      todoSaved = true;

      // ---- 3. Zeilen schreiben — sequentiell, Entzug zuerst ---------------
      let streak = 0;
      for (let i = 0; i < okPairs.length; i++) {
        const p = okPairs[i];
        const to = b2runNameOf(p.target);
        setProgress(`Startnummern werden zugeteilt… ${i + 1} / ${okPairs.length}`);
        if (isThrottled()) {
          abortNote = `SharePoint drosselt gerade — abgebrochen. Bis dahin geschrieben: ${doneCount}.`;
          break;
        }
        if (p.free.source === 'row' && p.free.fromRow) {
          const cleared = await props.service.adminUpdateRegistration(
            subsiteUrl, p.free.fromRow.Id, { Startnummer: '' }, actor);
          if (!cleared) {
            failed.push(`${p.bib} — konnte bei ${p.free.fromName} nicht entfernt werden, deshalb NICHT an ${to} vergeben. Sonst stünde dieselbe Nummer auf zwei Zeilen.`);
            rows.push({ bib: p.bib, toName: to, fromName: p.free.fromName, block: p.free.block, toEmail: p.target.ParticipantEmail || '', status: 'fehlgeschlagen — Altzeile nicht geleert' });
            streak++;
            if (streak >= 3) { abortNote = `Drei Fehlschläge in Folge — abgebrochen. Bis dahin geschrieben: ${doneCount}.`; break; }
            continue;
          }
        }
        const set = await props.service.adminUpdateRegistration(
          subsiteUrl, p.target.Id, { Startnummer: p.bib }, actor);
        if (!set) {
          failed.push(`${p.bib} trägt jetzt niemand — bei ${to} konnte die Nummer nicht gesetzt werden. Die Aufgabe steht; übertrage sie über „Offen beim Veranstalter".`);
          rows.push({ bib: p.bib, toName: to, fromName: p.free.fromName, block: p.free.block, toEmail: p.target.ParticipantEmail || '', status: 'fehlgeschlagen — trägt niemand' });
          streak++;
          if (streak >= 3) { abortNote = `Drei Fehlschläge in Folge — abgebrochen. Bis dahin geschrieben: ${doneCount}.`; break; }
          continue;
        }
        streak = 0;
        doneCount++;
        rows.push({ bib: p.bib, toName: to, fromName: p.free.fromName, block: p.free.block, toEmail: p.target.ParticipantEmail || '', status: 'zugeteilt — Namen überkleben' });
      }

      // ---- 4. Audit --------------------------------------------------------
      try {
        await props.service.writeChangeLog({
          action: 'B2RunBibAssigned',
          targetType: 'Event',
          eventId: String(eventId),
          eventTitle: props.event.title,
          actorName: actor.name,
          actorEmail: actor.email,
          details: {
            vergeben: doneCount,
            fehlgeschlagen: failed.length,
            verworfen: dropped.length,
            nummern: rows.filter(r => r.status.indexOf('zugeteilt') === 0).map(r => `${r.bib} → ${r.toName}`),
          },
        });
      } catch { /* Audit best-effort — der Lauf ist wichtiger als sein Protokoll */ }

      // ---- 5. Nachladen ----------------------------------------------------
      try { await refreshEvents(); } catch { /* beim nächsten Laden */ }
      setResult({ ok: doneCount, failed, dropped, abortNote, todoSaved, rows });
      props.onDone();
    } catch (err) {
      console.warn('[DEX] Startnummern-Zuteilung fehlgeschlagen:', err);
      setResult({
        ok: doneCount, failed, dropped,
        abortNote: abortNote || 'Unerwarteter Fehler — bitte den Stand über „Offen beim Veranstalter" und die Teilnehmerliste prüfen.',
        todoSaved, rows,
      });
    } finally { setBusy(false); setProgress(''); }
  };

  /**
   * Die Überklebe-Liste als Excel.
   *
   * Sie ist auch dann verfügbar, wenn noch nichts geschrieben wurde: Wenn
   * heute nur die Hälfte fertig wird, ist DIESE Liste die Hälfte, die trägt —
   * sie funktioniert am Ausgabetisch auch ohne Browser.
   */
  const downloadXlsx = async (): Promise<void> => {
    if (xlsxBusy) return;
    const src: WriteResult['rows'] = result
      ? result.rows
      : planned.map(p => ({
        bib: p.bib,
        toName: b2runNameOf(p.target),
        fromName: p.free.fromName,
        block: p.free.block || groupOf(p.target),
        toEmail: p.target.ParticipantEmail || '',
        status: 'geplant — noch nicht geschrieben',
      }));
    if (src.length === 0) return;
    setXlsxBusy(true);
    try {
      const head = ['Startnummer', 'Neuer Name', 'Alter Name (überkleben)', 'Startblock', 'E-Mail neu', 'Status'];
      const data: string[][] = [head];
      src.slice()
        .sort((a, b) => a.bib.localeCompare(b.bib, 'de', { numeric: true }))
        .forEach(r => data.push([r.bib, r.toName, r.fromName, r.block, r.toEmail, r.status]));
      // Bundle-Regel: xlsx MUSS dynamisch importiert werden (s. HotelImportModal).
      const XLSX = await import('xlsx');
      const ws = XLSX.utils.aoa_to_sheet(data);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = [{ wch: 13 }, { wch: 26 }, { wch: 26 }, { wch: 24 }, { wch: 32 }, { wch: 34 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Überkleben');
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Startnummern_ueberkleben_${new Date().toISOString().slice(0, 10)}.xlsx`;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
    } catch (err) {
      console.warn('[DEX] Überklebe-Liste fehlgeschlagen:', err);
      showAlert('Die Excel-Datei konnte nicht erzeugt werden — bitte erneut versuchen.', { variant: 'error' });
    } finally { setXlsxBusy(false); }
  };

  const waitLabel = progress || 'Bitte warten…';
  const blockedDone = pool ? pool.blocked.filter(b => b.reason === 'done') : [];
  const noneLeft = pool && pool.ok ? pool.candidates.filter(c => !choice[c.Id]).length : 0;

  return (
    <Modal
      open
      onClose={props.onClose}
      maxWidth={920}
      dismissable={!busy}
      ariaLabel="Startnummern zuteilen"
      title="Startnummern zuteilen"
      icon={<Hash size={20} />}
      subtitle={<>Freie Startnummern gibt es in DEX <strong>nur durch Abmeldungen</strong>. Es sind immer gebrauchte Nummern: Auf jeder steht noch ein anderer Name. Vergeben wird erst, wenn du bestätigst.</>}
      footer={<>
        <div className="dex-ui-modal-foot-left">
          <button type="button" className="btn btn-secondary" disabled={xlsxBusy || busy || (planned.length === 0 && !result)}
            onClick={() => { void downloadXlsx(); }}
          >
            <Download size={16} /> {xlsxBusy ? 'Wird erzeugt…' : 'Überklebe-Liste als Excel'}
          </button>
        </div>
        <button type="button" className="btn btn-secondary" disabled={busy} onClick={props.onClose}>
          {result ? 'Schließen' : 'Abbrechen'}
        </button>
      </>}
    >
      <div className="dex-ui-stack" style={{ fontSize: '0.87rem', lineHeight: 1.55 }}>
        {loading && <p className="dex-ui-muted" style={{ margin: 0 }}>Teilnehmerliste wird gelesen…</p>}

        {/* Die Sperren zuerst — bei einer harten Sperre gibt es weder Tabelle
            noch Knopf, nur den Grund im Klartext. */}
        {!loading && pool && !pool.ok && (
          <>
            <div className="dex-ui-callout dex-ui-callout--warn">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <div><strong>Es wird nichts zugeteilt.</strong><div style={{ marginTop: 4 }}>{pool.hardBlock}</div></div>
            </div>
            {pool.duplicatesActive.length > 0 && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  Achtung: doppelt belegt
                  <span className="dex-ui-pill dex-ui-pill--red">{pool.duplicatesActive.length}</span>
                </div>
                <p className="dex-ui-section-desc">
                  Dieselbe Startnummer steht auf mehreren nicht-abgemeldeten Zeilen. Solange das so ist, teilt DEX gar nichts zu —
                  am Einlass stünden sonst zwei Personen mit demselben Zettel.
                </p>
                <Tbl head={['Startnummer', 'Steht bei', 'Status']}>
                  {pool.duplicatesActive.map(d => (
                    <Row key={d.bib}>
                      <Bib>{d.bib}</Bib>
                      <>{d.regs.map(b2runNameOf).join(' · ')}</>
                      <>{d.regs.map(r => r.Status || 'unbekannt').join(' · ')}</>
                    </Row>
                  ))}
                </Tbl>
              </div>
            )}
          </>
        )}

        {!loading && pool && pool.ok && (
          <>
            <div className="dex-ui-grid-auto">
              <Kpi value={pool.free.length} label="Freie Nummern" tone={pool.free.length > 0 && 'dex-ui-kpi--green'} />
              <Kpi value={pool.candidates.length} label="Ohne Nummer" tone={pool.candidates.length > 0 && 'dex-ui-kpi--orange'} />
              <Kpi value={planned.length} label="Wird zugeteilt" />
              <Kpi value={pool.blocked.length} label="Nicht angeboten" tone="dex-ui-kpi--gray" />
            </div>

            {/* Die ehrliche Erwartung steht VOR der Tabelle — sonst liest man
                eine leere Liste als Fehler. */}
            {pool.free.length === 0 && (
              <div className="dex-ui-callout dex-ui-callout--neutral">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <div>
                  <strong>DEX kennt gerade keine freie Startnummer.</strong>
                  <div style={{ marginTop: 4 }}>
                    Das ist der Normalfall: Sind beim Veranstalter genauso viele Nummern gemeldet wie Leute angemeldet,
                    entstehen freie Nummern nur durch Abmeldungen.
                    {pool.candidates.length > 0
                      ? <> Für die {pool.candidates.length} Person{pool.candidates.length === 1 ? '' : 'en'} ohne Nummer lautet die richtige Antwort deshalb: <strong>beim Veranstalter nachmelden</strong>.</>
                      : <> Es ist auch niemand ohne Nummer — hier ist nichts zu tun.</>}
                  </div>
                </div>
              </div>
            )}

            {pool.free.length > 0 && pool.candidates.length === 0 && (
              <div className="dex-ui-callout dex-ui-callout--neutral">
                <span className="dex-ui-callout-icon"><Check size={16} /></span>
                <div>
                  <strong>Es ist niemand ohne Startnummer.</strong>
                  <div style={{ marginTop: 4 }}>
                    {pool.free.length} Nummer{pool.free.length === 1 ? '' : 'n'} {pool.free.length === 1 ? 'ist' : 'sind'} frei geworden, aber alle angemeldeten Personen haben bereits eine.
                    Melde die frei gewordenen Nummern beim Veranstalter ab — die Aufgaben dazu stehen unter &bdquo;Offen beim Veranstalter&ldquo;.
                  </div>
                </div>
              </div>
            )}

            <Step n={1} done={!!result} title="Zuordnung prüfen"
              hint={<>
                Vorgehakt ist nur, was in den Daten steht (die aufgezeichnete Nachrück-Kette). Alles andere ist ein Vorschlag und muss von dir gewählt werden.
                {untouched > 0 && <> {untouched} Person{untouched === 1 ? '' : 'en'} tragen bereits eine Nummer — die bleiben unverändert.</>}
              </>}
            >
              <span className="dex-ui-muted">{planned.length} von {pool.candidates.length} gewählt</span>
            </Step>

            <Step n={2} pending={planned.length === 0} done={!!result} title="Startnummern zuteilen"
              hint={result
                ? 'Erledigt — die Zusammenfassung steht unten.'
                : planned.length === 0
                  ? (pool.free.length === 0
                    ? 'Gerade nicht möglich: Es gibt keine freie Startnummer, die DEX belegen kann.'
                    : 'Wähle oben mindestens eine Nummer aus.')
                  : <>Schreibt zuerst die Ummelde-Aufgaben für den Veranstalter, dann die Nummern in die Teilnehmerliste. Geht dabei etwas schief, siehst du eine Aufgabe ohne Nummer — das fällt auf. Eine Nummer ohne Aufgabe würde niemandem auffallen.</>}
            >
              <button type="button" className={cx('btn', planned.length > 0 ? 'btn-primary' : 'btn-secondary', 'dex-ui-btn-sm')}
                disabled={planned.length === 0 || busy || !!result}
                onClick={() => { void write(); }}
              >
                {busy ? waitLabel : planned.length === 0 ? 'Nichts ausgewählt' : `${planned.length} Startnummer${planned.length === 1 ? '' : 'n'} zuteilen`}
              </button>
            </Step>

            {result && (
              <div className={cx('dex-ui-callout', (result.failed.length || result.dropped.length || result.abortNote) ? 'dex-ui-callout--danger' : 'dex-ui-callout--success')}>
                <span className="dex-ui-callout-icon">{(result.failed.length || result.dropped.length || result.abortNote) ? <AlertCircle size={16} /> : <Check size={16} />}</span>
                <div>
                  <strong>{result.ok} Startnummer{result.ok === 1 ? '' : 'n'} zugeteilt{result.failed.length ? `, ${result.failed.length} fehlgeschlagen` : ''}.</strong>
                  {result.ok > 0 && (
                    <div style={{ marginTop: 4 }}>
                      Alle {result.ok} tragen noch einen anderen Namen und müssen am Check-in-Tisch überklebt werden.
                      Beim Veranstalter musst du {result.ok} Ummeldung{result.ok === 1 ? '' : 'en'} vornehmen — sie stehen unter <strong>&bdquo;Offen beim Veranstalter&ldquo;</strong>.
                    </div>
                  )}
                  {result.abortNote && <div style={{ marginTop: 4 }}>{result.abortNote}</div>}
                  {result.dropped.length > 0 && (
                    <div style={{ marginTop: 4 }}>
                      Zwischen Vorschau und Klick hat sich etwas geändert — diese Paare wurden verworfen:
                      <ul style={{ margin: '4px 0 0', paddingLeft: 18 }}>{result.dropped.map((d, i) => <li key={i}>{d}</li>)}</ul>
                    </div>
                  )}
                  {result.failed.length > 0 && (
                    <ul style={{ margin: '4px 0 0', paddingLeft: 18 }}>{result.failed.map((f, i) => <li key={i}>{f}</li>)}</ul>
                  )}
                  {result.ok > 0 && !result.todoSaved && (
                    <div style={{ marginTop: 4, color: 'var(--dex-red)' }}>
                      Die Nummern stehen in DEX, die Aufgabenliste konnte nicht gespeichert werden — bitte jetzt die Überklebe-Liste als Excel laden,
                      sonst weiß der Veranstalter von den Ummeldungen nichts.
                    </div>
                  )}
                </div>
              </div>
            )}

            {pool.candidates.length > 0 && pool.free.length > 0 && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  Wer bekommt welche Nummer?
                  <span className="dex-ui-pill dex-ui-pill--orange">{pool.candidates.length} ohne Nummer</span>
                </div>
                <p className="dex-ui-section-desc">
                  Eine Nummer, die du hier wählst, verschwindet aus den anderen Listen — dieselbe Nummer kann in diesem Fenster nicht zweimal vergeben werden.
                </p>
                <Tbl head={['Person', 'Gruppe', 'Vorschlag', 'Startnummer', 'Nummer trägt aktuell den Namen', 'Herkunft']}>
                  {pool.candidates.map(c => {
                    const sel = choice[c.Id] || '';
                    const f = sel ? freeOf(sel) : undefined;
                    const prop = proposalFor(c.Id);
                    const takenElsewhere = Object.keys(choice)
                      .filter(k => Number(k) !== c.Id && !!choice[Number(k)])
                      .map(k => choice[Number(k)]);
                    const mismatch = f ? blockDiffers(f.block, groupOf(c)) : false;
                    return (
                      <Row key={c.Id}>
                        <><strong>{b2runNameOf(c)}</strong><span className="dex-ui-muted"> · {c.ParticipantEmail}</span></>
                        <>{groupOf(c) || '—'}</>
                        <>
                          {prop && prop.stufe === 'chain' && (
                            <span className="dex-ui-pill dex-ui-pill--green" title={prop.reason}>Kette: {prop.bib}</span>
                          )}
                          {prop && prop.stufe === 'suggestion' && (
                            <>
                              <span className="dex-ui-pill dex-ui-pill--gray" title={prop.reason}>Vorschlag: {prop.bib}</span>
                              {sel !== prop.key && !result && (
                                <button type="button" className="dex-ui-textbtn" disabled={busy}
                                  onClick={() => setChoice(prev => ({ ...prev, [c.Id]: prop.key }))}
                                >übernehmen</button>
                              )}
                            </>
                          )}
                          {!prop && <span className="dex-ui-muted">—</span>}
                        </>
                        <>
                          <select
                            className="dex-ui-select dex-ui-input--sm"
                            value={sel}
                            disabled={busy || !!result}
                            style={{ minWidth: 190 }}
                            aria-label={`Startnummer für ${b2runNameOf(c)}`}
                            onChange={e => setChoice(prev => ({ ...prev, [c.Id]: e.target.value }))}
                          >
                            <option value="">— keine, beim Veranstalter nachmelden —</option>
                            {pool.free
                              .filter(x => x.key === sel || takenElsewhere.indexOf(x.key) < 0)
                              .map(x => (
                                <option key={x.key} value={x.key}>
                                  {x.bib}{blockDiffers(x.block, groupOf(c)) ? ' · anderer Startblock' : ''}
                                </option>
                              ))}
                          </select>
                          {mismatch && (
                            <div className="dex-ui-help" style={{ color: '#7a4a00' }}>
                              Anderer Startblock ({f ? f.block : ''}) — bitte nur wählen, wenn du weißt, dass es passt.
                            </div>
                          )}
                        </>
                        <>{f ? <Gone>{f.fromName}</Gone> : <span className="dex-ui-muted">—</span>}</>
                        <>{f
                          ? (f.source === 'row'
                            ? <span title="Die Abmeldung steht in DEX; die Nummer wird dort geleert, bevor sie neu vergeben wird.">Abmeldung in DEX belegt</span>
                            : <span title="Die Nummer stand nie auf einer DEX-Zeile — sie stammt aus einer festgehaltenen Aufgabe des Imports.">nur aus dem Import erinnert</span>)
                          : <span className="dex-ui-muted">—</span>}</>
                      </Row>
                    );
                  })}
                </Tbl>
                {noneLeft > 0 && (
                  <p className="dex-ui-muted" style={{ margin: '6px 0 0' }}>
                    {noneLeft} Person{noneLeft === 1 ? '' : 'en'} ohne Auswahl — für sie hat DEX keine Nummer. Melde sie beim Veranstalter nach.
                  </p>
                )}
              </div>
            )}

            {pool.blocked.length > 0 && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  Nicht angeboten
                  <span className="dex-ui-pill dex-ui-pill--gray">{pool.blocked.length}</span>
                </div>
                <p className="dex-ui-section-desc">
                  Diese Nummern sind in DEX sichtbar, aber gesperrt. Der Grund steht dahinter — er ist immer eine Aussage über die Daten,
                  nie eine Vermutung.
                  {blockedDone.length > 0 && <> Beim Veranstalter abgemeldete Nummern lassen sich mit einem zweiten Klick freigeben; das ist nie automatisch.</>}
                </p>
                <Tbl head={['Startnummer', 'Grund', '']}>
                  {pool.blocked.map(b => (
                    <Row key={`${b.reason}-${b.bib}`}>
                      <Bib>{b.bib}</Bib>
                      <>{b.detail}</>
                      <>{b.reason === 'done' && !result && (
                        <button type="button" className="btn btn-secondary dex-ui-btn-sm" disabled={busy}
                          onClick={() => { void unlockBib(b.bib); }}
                        >Nummer freigeben</button>
                      )}</>
                    </Row>
                  ))}
                </Tbl>
              </div>
            )}

            {pool.skippedPeople.length > 0 && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  Bleibt ohne Nummer
                  <span className="dex-ui-pill dex-ui-pill--gray">{pool.skippedPeople.length}</span>
                </div>
                <p className="dex-ui-section-desc">
                  Diese Personen sind angemeldet, bekommen hier aber bewusst keine Nummer — sie haben schon eine oder eine offene Aufgabe.
                </p>
                <Tbl head={['Person', 'Grund']}>
                  {pool.skippedPeople.map(s => (
                    <Row key={s.reg.Id}>
                      <><strong>{b2runNameOf(s.reg)}</strong><span className="dex-ui-muted"> · {s.reg.ParticipantEmail}</span></>
                      <>{s.reason}</>
                    </Row>
                  ))}
                </Tbl>
              </div>
            )}

            <p className="dex-ui-muted" style={{ margin: 0 }}>
              Halte während der Zuteilung keinen Wizard für dieses Event offen — er würde die frisch entstandenen Aufgaben überschreiben.
            </p>
          </>
        )}
      </div>
    </Modal>
  );
}
