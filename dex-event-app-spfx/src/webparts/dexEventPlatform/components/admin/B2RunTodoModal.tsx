/**
 * v30.54: „Offen beim Veranstalter" — die Aufgabenliste zum B2Run.
 *
 * **v30.55 korrigiert den Ansatz aus v30.54.** Dort wurde die Liste
 * ausschließlich aus den Teilnehmerzeilen abgeleitet — und war direkt nach
 * einem Import leer, obwohl beim Veranstalter neun Ummeldungen warteten. Der
 * Grund steht in `utils/b2runTodos`: Die Startnummer einer abgemeldeten Person
 * existiert nur in der Datei des Veranstalters, nie in DEX. Was der Import
 * erzeugt hat, wird deshalb FESTGEHALTEN (`_b2runTodo`); was seither passiert
 * ist, kommt weiterhin aus der Ableitung. Beides wird hier zusammengeführt.
 *
 * Der Haken liegt getrennt davon in `_b2runTodoDone`.
 */
import * as React from 'react';
import Modal from '../Modal';
import { useDialog } from '../../context/DialogContext';
import { useEvents } from '../../context/EventContext';
import { useCurrentUser } from '../../context/UserContext';
import { DeloitteEvent } from '../../types';
import { EventService, SPRegistration } from '../../services/EventService';
import { mergeB2RunTodos, B2RunTodo, StoredB2RunTodo, B2RUN_TODO_LABELS } from '../../utils/b2runTodos';
// v31.4: „0412" und „412" sind auf dem Zettel dieselbe Nummer — die Sperre
// unten muss beide erkennen (s. utils/b2runBibPool).
import { bibKey } from '../../utils/b2runBibPool';
import { cx } from '../dexUi';
import { Check, ChevronDown, Download, Hash } from '../Icons';

// v31.2: Die Art der Aufgabe als Status-Pill (dex-ui-pill) statt eigener
// Farbwerte — dieselbe Farblogik wie vorher: Ummelden orange, Abmelden rot,
// Nachmelden blau.
const KIND_PILL: Record<B2RunTodo['kind'], string> = {
  transfer: 'dex-ui-pill--orange',
  assign: 'dex-ui-pill--orange',
  unregister: 'dex-ui-pill--red',
  register: 'dex-ui-pill--blue',
};

export default function B2RunTodoModal(props: {
  event: DeloitteEvent;
  service: EventService;
  onClose: () => void;
}): React.ReactElement {
  const { getAllRegistrations, events } = useEvents();
  const { showAlert } = useDialog();
  // v31.4: Wer die Nummer umgetragen hat, gehört in den ChangeLog der Zeile.
  // Vorher stand dort „DEX" — im Nachhinein war nicht mehr feststellbar, wer
  // am Vorabend welche Nummer verschoben hat.
  const { currentUser } = useCurrentUser();
  const [loading, setLoading] = React.useState(true);
  const [todos, setTodos] = React.useState<B2RunTodo[]>([]);
  const [regs, setRegs] = React.useState<SPRegistration[]>([]);
  const [bibBusy, setBibBusy] = React.useState<string>('');
  const [done, setDone] = React.useState<string[]>([]);
  const [saving, setSaving] = React.useState(false);
  const [showDone, setShowDone] = React.useState(false);
  const [xlsxBusy, setXlsxBusy] = React.useState(false);

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        // v30.67 (Review): unvollständig gelesen = unvollständige Aufgaben —
        // das muss die Person wissen, sonst gelten offene Nummern als erledigt.
        let readFailed = false;
        const regs = await getAllRegistrations(props.event.id, () => { readFailed = true; });
        if (cancelled) return;
        setRegs(regs);
        if (readFailed) {
          await showAlert('Die Teilnehmerliste konnte nicht vollständig gelesen werden — die Aufgabenliste ist unvollständig. Bitte den Dialog schließen und erneut öffnen.', { variant: 'error' });
        }
        let storedTodos: StoredB2RunTodo[] = [];
        let storedDone: string[] = [];
        try {
          // v30.56: IMMER den frischesten Stand nehmen. Das über die Props
          // hereingereichte Event kann älter sein als der Kontext (der Import
          // schreibt die Aufgaben und lädt danach nach) — und aus einem
          // veralteten Objekt gelesen, sieht die Liste leer aus.
          const liveEv = events.find(e => e.id === props.event.id) || props.event;
          const o = JSON.parse(liveEv.emailTemplateOverrides || '{}');
          if (Array.isArray(o?._b2runTodo)) storedTodos = o._b2runTodo.filter((x: unknown) => !!x && typeof x === 'object');
          if (Array.isArray(o?._b2runTodoDone)) storedDone = o._b2runTodoDone.filter((x: unknown) => typeof x === 'string');
        } catch { /* kein Piggyback — dann ist nichts festgehalten */ }
        setTodos(mergeB2RunTodos(storedTodos, regs));
        setDone(storedDone);
      } catch {
        if (!cancelled) await showAlert('Die Teilnehmerliste konnte nicht gelesen werden.', { variant: 'error' });
      } finally { if (!cancelled) setLoading(false); }
    })().catch(() => { /* oben behandelt */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.event.id]);

  const toggle = async (key: string): Promise<void> => {
    const next = done.indexOf(key) >= 0 ? done.filter(k => k !== key) : done.concat(key);
    setDone(next);
    setSaving(true);
    // Nur den Haken schreiben — patchEventOverridesValue liest den aktuellen
    // Stand und merged, damit parallele Änderungen an anderen Piggybacks
    // (Abrechnung, Hotels) nicht verloren gehen.
    const ok = await props.service.patchEventOverridesValue(Number(props.event.id), '_b2runTodoDone', next)
      .catch(() => false);
    setSaving(false);
    if (!ok) {
      setDone(done); // zurückrollen — sonst sieht es erledigt aus und ist es nicht
      await showAlert('Der Haken konnte nicht gespeichert werden — bitte erneut versuchen.', { variant: 'error' });
    }
  };

  /**
   * v30.55: Die Startnummer in DEX auf die Person umtragen, die sie beim
   * Veranstalter bekommt.
   *
   * Die Ummeldung beim Veranstalter ist das eine — steht die Nummer in DEX
   * weiter bei der abgemeldeten Person (oder bei niemandem), zeigt der
   * Check-in für die Person, die wirklich läuft, keine Nummer an. Der Knopf
   * setzt sie deshalb um und räumt sie bei der Vorgängerin weg; danach
   * verschwindet die Aufgabe von selbst aus der abgeleiteten Liste.
   *
   * **v31.4: Reihenfolge umgedreht und Rückgabewerte geprüft.** Vorher wurde
   * ERST die neue Zeile gesetzt und DANN die alte geleert, ohne einen der
   * beiden Rückgabewerte anzusehen — und `bibInDex: true` samt grüner Pille
   * kam bedingungslos hinterher. Zwei Fehler in einem:
   *
   *  1. `adminUpdateRegistration` wirft nicht, sie liefert `response.ok`. Ein
   *     403/429 lief also stillschweigend durch und die Anzeige meldete
   *     Erfolg. Genau daraus entsteht eine aktive Zeile ohne Nummer — der
   *     die Aktion „Startnummern zuteilen" später eine ZWEITE gibt.
   *  2. Eine Startnummer ist ein ausschließliches Recht. Scheitert das Leeren
   *     der Altzeile, tragen zwei Zeilen dieselbe Nummer. Deshalb kommt hier
   *     der Entzug ZUERST (anders als beim Austausch von Datensätzen, wo der
   *     unumkehrbare Schritt zuletzt kommt): Bleibt die neue Zeile leer, ist
   *     das sichtbar und behebbar; eine doppelt vergebene Nummer fällt erst
   *     am Einlass auf.
   */
  const moveBibInDex = async (t: B2RunTodo): Promise<void> => {
    if (bibBusy || !t.toReg || !t.bib || !props.event.subsiteUrl) return;
    // v31.4: Trägt die Nummer schon jemand, der nicht abgemeldet ist? Dann
    // wird nichts geschrieben. Warteliste und No-Show zählen als BELEGT — die
    // Person kann zurückkommen (dieselbe Regel wie utils/b2runBibPool).
    const holder = regs.filter(r => r.Status !== 'Abgemeldet'
      && bibKey(String(r.Startnummer || '')) === bibKey(t.bib)
      && r.Id !== t.toReg!.Id)[0];
    if (holder) {
      await showAlert(
        `Startnummer ${t.bib} steht bereits bei ${holder.Vorname || ''} ${holder.Nachname || ''}`.trim()
        + ` (Status ${holder.Status || 'unbekannt'}). Es wurde nichts geändert — bitte erst klären, wem die Nummer gehört.`,
        { variant: 'error' });
      return;
    }
    setBibBusy(t.key);
    try {
      const actor = {
        name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email,
        email: currentUser.email,
      };
      // 1) Bei der Vorgängerin leeren — GEPRÜFT, und vor dem Setzen.
      const from = regs.filter(r => (r.ParticipantEmail || '').toLowerCase().trim() === (t.fromEmail || '').toLowerCase().trim()
        && bibKey(String(r.Startnummer || '')) === bibKey(t.bib))[0];
      if (from) {
        const cleared = await props.service.adminUpdateRegistration(props.event.subsiteUrl, from.Id, { Startnummer: '' }, actor);
        if (!cleared) {
          await showAlert(
            `Die Startnummer konnte bei ${t.fromName || 'der bisherigen Person'} nicht entfernt werden — deshalb wurde sie auch NICHT auf ${t.toName} übertragen. Sonst stünde dieselbe Nummer auf zwei Zeilen. Bitte erneut versuchen.`,
            { variant: 'error' });
          return;
        }
      }
      // 2) Erst jetzt die neue Zeile setzen — ebenfalls geprüft.
      const set = await props.service.adminUpdateRegistration(props.event.subsiteUrl, t.toReg.Id, { Startnummer: t.bib }, actor);
      if (!set) {
        await showAlert(
          `Startnummer ${t.bib} trägt jetzt niemand: Bei ${t.fromName || 'der bisherigen Person'} ist sie entfernt, bei ${t.toName} konnte sie nicht gesetzt werden. Die Aufgabe bleibt offen — bitte gleich noch einmal übertragen.`,
          { variant: 'error' });
        return;
      }
      // v30.67 (Review): Die Übertragung ist durch — scheitert nur das
      // Nachlesen, bleibt der alte Stand stehen und die Aufgabe gilt als
      // erledigt; sonst stünde die Nummer wieder als „offen" da.
      let freshFailed = false;
      const fresh = await getAllRegistrations(props.event.id, () => { freshFailed = true; });
      if (!freshFailed) setRegs(fresh);
      setTodos(prev => prev.map(x => x.key === t.key
        ? { ...x, bibInDex: true, toReg: (!freshFailed && fresh.find(r => r.Id === t.toReg!.Id)) || x.toReg }
        : x));
      if (freshFailed) {
        await showAlert('Übertragen — aber die Teilnehmerliste konnte danach nicht neu geladen werden. Bitte den Dialog später erneut öffnen.', { variant: 'info' });
      }
    } catch {
      await showAlert('Die Startnummer konnte nicht übertragen werden — bitte erneut versuchen.', { variant: 'error' });
    } finally { setBibBusy(''); }
  };

  const open = todos.filter(t => done.indexOf(t.key) < 0);
  const closed = todos.filter(t => done.indexOf(t.key) >= 0);

  const downloadXlsx = async (): Promise<void> => {
    if (xlsxBusy) return;
    setXlsxBusy(true);
    try {
      const rows: string[][] = [['Status', 'Was ist zu tun', 'Startnummer', 'Von', 'E-Mail (von)', 'Auf', 'E-Mail (auf)', 'Sicherheit']];
      for (const t of todos) {
        rows.push([
          done.indexOf(t.key) >= 0 ? 'erledigt' : 'offen',
          `${B2RUN_TODO_LABELS[t.kind]} — ${t.action}`,
          t.bib || '',
          t.fromName || '', t.fromEmail || '',
          t.toName || '', t.toEmail || '',
          t.certain ? 'in DEX aufgezeichnet' : 'erschlossen — bitte prüfen',
        ]);
      }
      const XLSX = await import('xlsx');
      const ws = XLSX.utils.aoa_to_sheet(rows);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = [{ wch: 10 }, { wch: 76 }, { wch: 13 }, { wch: 24 }, { wch: 30 }, { wch: 24 }, { wch: 30 }, { wch: 26 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Offen beim Veranstalter');
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `B2Run_offene_Aufgaben_${new Date().toISOString().slice(0, 10)}.xlsx`;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
    } catch (err) {
      console.warn('[DEX] B2Run-Todo-Export fehlgeschlagen:', err);
      await showAlert('Die Excel-Datei konnte nicht erzeugt werden.', { variant: 'error' });
    } finally { setXlsxBusy(false); }
  };

  // v31.2: Eine Aufgabe ist eine Zeile mit Hover (dex-ui-row) — der Haken
  // links, rechts daneben Art, Nummer, Satz und ggf. der Übertragen-Knopf.
  // Erledigte bleiben sichtbar, aber gedämpft (voll erst beim Überfahren),
  // damit die Person sieht, was sie schon abgehakt hat.
  const renderTodo = (t: B2RunTodo, isDone: boolean): React.ReactElement => (
    <div key={t.key} className={cx('dex-ui-row dex-ui-row--bordered', isDone && 'dex-ui-card--muted')} style={{ alignItems: 'flex-start' }}>
      <input
        type="checkbox"
        checked={isDone}
        disabled={saving}
        onChange={() => { void toggle(t.key); }}
        style={{ marginTop: 3, flexShrink: 0, width: 18, height: 18, cursor: saving ? 'wait' : 'pointer', accentColor: '#86bc25' }}
        aria-label={isDone ? 'Als offen markieren' : 'Als erledigt markieren'}
        title={isDone ? 'Als offen markieren' : 'Als erledigt markieren'}
      />
      <div className="dex-ui-row-main">
        <div className="dex-ui-inline" style={{ marginBottom: 3 }}>
          <span className={cx('dex-ui-pill', KIND_PILL[t.kind])}>{B2RUN_TODO_LABELS[t.kind]}</span>
          {t.bib && (
            <span style={{ fontFamily: "'Courier New',Courier,monospace", fontWeight: 700, fontSize: '0.95rem' }}>{t.bib}</span>
          )}
          {!t.certain && (
            <span className="dex-ui-pill dex-ui-pill--gray" title="Nicht in DEX aufgezeichnet, sondern aus der Teilnehmerliste erschlossen — bitte beim Veranstalter prüfen.">
              erschlossen, nicht aufgezeichnet
            </span>
          )}
        </div>
        <div style={{ fontSize: '0.85rem', lineHeight: 1.5, whiteSpace: 'normal', textDecoration: isDone ? 'line-through' : undefined }}>
          {t.action}
        </div>
        {(t.fromEmail || t.toEmail) && (
          <div className="dex-ui-row-sub">
            {t.fromEmail ? <>von {t.fromEmail}</> : null}
            {t.fromEmail && t.toEmail ? ' · ' : ''}
            {t.toEmail ? <>auf {t.toEmail}</> : null}
          </div>
        )}
        {/* v30.55: Die Ummeldung beim Veranstalter ist das eine — die Nummer
            muss aber auch in DEX bei der richtigen Person stehen, sonst zeigt
            der Check-in für die Person, die wirklich läuft, gar keine Nummer.
            v31.4: Der Knopf hing an `!isDone`. Wer zuerst beim Veranstalter
            ummeldet und abhakt — die naheliegende Reihenfolge —, konnte die
            Nummer danach NIE mehr in DEX nachtragen. Genau daraus entsteht
            die aktive Zeile ohne Nummer, der die Aktion „Startnummern
            zuteilen" später eine zweite gibt. Abgehakt heißt „beim
            Veranstalter erledigt", nicht „in DEX eingetragen". */}
        {t.bib && t.toReg && !t.bibInDex && (
          <div style={{ marginTop: 8 }}>
            <button
              type="button"
              className="btn btn-secondary dex-ui-btn-sm"
              disabled={!!bibBusy}
              style={isDone ? { opacity: 0.75 } : undefined}
              onClick={() => { void moveBibInDex(t); }}
            >
              {bibBusy === t.key ? 'Wird übertragen…' : `Startnummer ${t.bib} in DEX auf ${t.toName} übertragen`}
            </button>
            <div className="dex-ui-help">
              {isDone
                ? <>Beim Veranstalter ist das erledigt — in DEX steht die Nummer aber noch nicht bei {t.toName}. Nur nachtragen, nichts beim Veranstalter ändern.</>
                : <>Danach zeigt der Check-in die Nummer bei {t.toName}; die Aufgabe verschwindet von selbst.</>}
            </div>
          </div>
        )}
        {t.bib && t.bibInDex && (
          <div className="dex-ui-pill dex-ui-pill--green" style={{ marginTop: 6 }}>
            <Check size={12} /> In DEX steht die Nummer bereits bei {t.toName}.
          </div>
        )}
      </div>
    </div>
  );

  return (
    <Modal
      open
      onClose={props.onClose}
      maxWidth={780}
      ariaLabel="Offen beim Veranstalter"
      title="Offen beim B2Run-Veranstalter"
      subtitle="Was du beim Veranstalter noch ummelden, ab- oder nachmelden musst. Bei jedem Öffnen neu aus der Teilnehmerliste berechnet — spätere Abmeldungen tauchen von selbst auf, abgehakte Aufgaben bleiben gespeichert."
      icon={<Hash size={20} />}
      footer={<>
        <div className="dex-ui-modal-foot-left">
          <button
            type="button"
            className="btn btn-secondary"
            disabled={xlsxBusy || todos.length === 0}
            onClick={() => { void downloadXlsx(); }}
          >
            <Download size={16} /> {xlsxBusy ? 'Wird erzeugt…' : 'Als Excel laden'}
          </button>
        </div>
        <button type="button" className="btn btn-primary" onClick={props.onClose}>Schließen</button>
      </>}
    >
      <div className="dex-ui-modal-body">
        {loading ? (
          <p className="dex-ui-muted" style={{ margin: 0 }}>Teilnehmerliste wird gelesen…</p>
        ) : todos.length === 0 ? (
          <div className="dex-ui-empty">
            <div className="dex-ui-empty-icon"><Check size={20} /></div>
            <div className="dex-ui-empty-title">Nichts offen</div>
            Alle Startnummern sind zugeordnet — beim Veranstalter ist nichts zu tun.
          </div>
        ) : (
          <>
            {open.length === 0 ? (
              <div className="dex-ui-callout dex-ui-callout--success">
                <span className="dex-ui-callout-icon"><Check size={16} /></span>
                <span><strong>Alles abgehakt</strong> — {closed.length} Aufgabe{closed.length === 1 ? '' : 'n'} erledigt.</span>
              </div>
            ) : (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  Offen <span className="dex-ui-pill dex-ui-pill--orange">{open.length}</span>
                </div>
                <div className="dex-ui-card" style={{ padding: '4px 6px' }}>
                  {open.map(t => renderTodo(t, false))}
                </div>
              </div>
            )}

            {closed.length > 0 && (
              <div>
                <button
                  type="button"
                  className={cx('dex-ui-disclosure', showDone && 'is-open')}
                  aria-expanded={showDone}
                  onClick={() => setShowDone(v => !v)}
                >
                  <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                  {showDone ? 'Erledigte ausblenden' : 'Erledigte anzeigen'}
                  <span className="dex-ui-disclosure-count">{closed.length}</span>
                </button>
                {showDone && (
                  <div className="dex-ui-disclosure-body">
                    <div className="dex-ui-card dex-ui-card--soft" style={{ padding: '4px 6px' }}>
                      {closed.map(t => renderTodo(t, true))}
                    </div>
                  </div>
                )}
              </div>
            )}
          </>
        )}
      </div>
    </Modal>
  );
}
