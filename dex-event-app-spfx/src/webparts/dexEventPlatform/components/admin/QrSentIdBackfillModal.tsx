/**
 * v31.4: „QR-Nummern nachtragen" — die gedruckten Nummern der bereits
 * verschickten QR-Mails in die Spalte `QrSentId` zurückschreiben.
 *
 * Warum es das gibt: Die Spalte ist neu. Für jedes Event, dessen QR-Mails
 * vorher rausgingen, steht nirgends, welche Nummer in welcher Mail stand —
 * und genau dort tut der Fehler weh, weil dort schon Abmeldungen die
 * `TeilnehmerID` verschoben haben. Die Nummern stehen aber noch in der
 * Mail-Warteschlange `DEX_Emails` (der Body enthält das komplette HTML, die
 * Zeilen werden erst rund einen Monat nach dem Event archiviert).
 *
 * Drei Dinge, die diesen Dialog belastbar machen:
 *  - **Ein Lesefehler ist keine Null.** Bei einem HTTP-Fehler bricht der Lauf
 *    ab und nennt den Status. Und weil `DEX_Emails` zeilenweise gesichert ist
 *    (`ReadSecurity=2`), heißt „0 Mails" auch ohne Fehler nicht „es gab
 *    keine" — die Meldung nennt deshalb IMMER beide Ursachen.
 *  - **Vorschau vor jedem Schreibvorgang.** Wer keinen Fund hat, steht
 *    namentlich da: für diese Personen bleibt der Check-in bei der laufenden
 *    Nummer, und das darf nicht unter den Tisch fallen.
 *  - **Die Klammer ist nie der ganze Pfad.** Gelesen und geschrieben wird über
 *    das Event UND jeden Termin; `EventId` in `DEX_Emails` ist die Id des
 *    Events, zu dem die Mail gehört.
 */
import * as React from 'react';
import Modal from '../Modal';
import { DeloitteEvent } from '../../types';
import { EventService } from '../../services/EventService';
import { QrMailHit } from '../../services/events/qrSentBackfill';
import { useEvents } from '../../context/EventContext';
import { SAMPLE_QR_ID } from './adminConstants';
import { cx } from '../dexUi';
import { AlertCircle, QrCode } from '../Icons';

/** Nur diese Zeilen bekommen eine QR-Mail — Abgemeldete und Wartelistler nicht. */
const BACKFILL_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'No-Show'];

/**
 * v31.4 (Review): `'dup'` ist neu. Fiele dieselbe Nummer auf zwei Zeilen
 * derselben Teilnehmerliste, wäre der ID-Check-in für beide GESPERRT — die
 * Seite wertet zwei Treffer als Datenfehler und verweigert die Eingabe. Vor
 * dem Nachtrag trugen die Zeilen verschiedene laufende Nummern und lösten
 * sauber auf; ein Nachtrag, der eine neue Sperre baut, ist schlimmer als
 * keiner. Zwei Wege dahin: dieselbe Person mit zwei aktiven Zeilen (der
 * Dubletten-Fall aus v30.73) oder zwei Personen, die aus zwei Mails dieselbe
 * Zahl erben (z. B. eine Testmail mit echter Nummer).
 */
type Verdict = 'diff' | 'same' | 'organizer' | 'dup' | 'none';

interface BackfillRow {
  key: string;
  eventTitle: string;
  subsiteUrl: string;
  regId: number;
  name: string;
  email: string;
  /** Die heutige laufende Nummer. */
  current: number | null;
  /** Die Nummer aus der Mail — `null` = kein Fund. */
  qrId: number | null;
  verdict: Verdict;
  /** Nur „Organizer — bitte prüfen"-Zeilen sind abwählbar (s. unten). */
  include: boolean;
}

/** Dreistellig wie in der Mail. padStart gibt es im ES5-Target nicht. */
function pad3(n: number): string { let s = String(n); while (s.length < 3) s = `0${s}`; return s; }

export default function QrSentIdBackfillModal(props: {
  event: DeloitteEvent;
  childEvents: DeloitteEvent[];
  service: EventService;
  isDe: boolean;
  onClose: () => void;
  /** Nach dem Schreiben: die Teilnehmerliste der Seite neu laden. */
  onDone?: () => void;
}): React.ReactElement {
  const { event, childEvents, service, isDe, onClose, onDone } = props;
  const { getAllRegistrations } = useEvents();
  const [phase, setPhase] = React.useState<'scan' | 'preview' | 'writing' | 'done'>('scan');
  const [abortMsg, setAbortMsg] = React.useState('');
  const [rows, setRows] = React.useState<BackfillRow[]>([]);
  const [testSkipped, setTestSkipped] = React.useState(0);
  const [scannedMails, setScannedMails] = React.useState(0);
  const [unparsedMails, setUnparsedMails] = React.useState(0);
  const [progress, setProgress] = React.useState({ done: 0, total: 0 });
  const [resultMsg, setResultMsg] = React.useState('');
  const [resultIsError, setResultIsError] = React.useState(false);
  /**
   * v31.4 (Review): Der Lese-Effect hatte sein `cancelled`-Flag, der
   * Schreiblauf nicht — er lief nach dem Aushängen weiter und setzte State
   * auf eine abgemeldete Komponente. Ein Ref statt einer lokalen Variablen,
   * weil `runWrite` kein Effect ist und beide Pfade dieselbe Antwort
   * brauchen: „gibt es diesen Dialog noch?"
   */
  const aliveRef = React.useRef(true);
  React.useEffect(() => () => { aliveRef.current = false; }, []);

  const targets = React.useMemo(
    () => [event].concat(childEvents).filter(e => !!e && !!e.subsiteUrl),
    [event, childEvents],
  );

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      const collected: BackfillRow[] = [];
      let skipped = 0;
      let scanned = 0;
      let unparsed = 0;
      for (const ev of targets) {
        // 1) Teilnehmerliste — ein Lesefehler beendet den Lauf. Eine leere
        //    Liste wegen fehlender Rechte würde sonst als „niemand angemeldet"
        //    durchgehen und die Vorschau wäre eine Aussage über nichts.
        let readErr = -1;
        const regs = await getAllRegistrations(ev.id, (s: number) => { readErr = s; });
        if (cancelled) return;
        if (readErr >= 0) {
          // Status 0 heißt „kein HTTP-Versuch möglich" (keine Subsite oder
          // Netzfehler) — „HTTP 0" wäre dafür eine irreführende Angabe.
          const why = readErr > 0 ? `HTTP ${readErr}` : (isDe ? 'keine Teilnehmerliste erreichbar' : 'attendee list not reachable');
          setAbortMsg(isDe
            ? `Die Teilnehmerliste von „${ev.title}" konnte nicht gelesen werden (${why}) — abgebrochen, es wurde nichts geschrieben. Bitte Leserechte prüfen und erneut versuchen.`
            : `The attendee list of “${ev.title}” could not be read (${why}) — aborted, nothing was written. Please check read access and try again.`);
          setPhase('done');
          setResultIsError(true);
          return;
        }
        // 2) Die verschickten QR-Mails dieses Events.
        const scan = await service.scanQrMailsForEvent(ev.id);
        if (cancelled) return;
        if (!scan.ok) {
          setAbortMsg(isDe
            ? `Die Mail-Warteschlange DEX_Emails konnte nicht gelesen werden (${scan.status ? `HTTP ${scan.status}` : 'Netzwerkfehler'}) — abgebrochen, es wurde nichts geschrieben.`
            : `The mail queue DEX_Emails could not be read (${scan.status ? `HTTP ${scan.status}` : 'network error'}) — aborted, nothing was written.`);
          setPhase('done');
          setResultIsError(true);
          return;
        }
        scanned += scan.scanned;
        unparsed += scan.unparsed;

        // 3) Organizer-Adressen dieses Events UND des Elternevents — die
        //    Test-Mail geht an die Organizer, und bei einem Termin stehen sie
        //    oft nur auf der Klammer.
        const orgSet: Record<string, boolean> = {};
        const parent = ev.parentEventId === event.id ? event : null;
        [ev, parent].forEach(e => {
          if (!e) return;
          ((e.organizerEmails || []) as string[]).concat((e.coOrganizerEmails || []) as string[])
            .forEach(a => { const k = (a || '').trim().toLowerCase(); if (k) orgSet[k] = true; });
        });

        // 4) Test-Mails aussortieren: `qrTestSendAction` schickt Mails vom Typ
        //    QRCode an die Organizer, mit der Beispiel-ID für alle außer der
        //    angemeldeten Person. Eine Mail an eine Organizer-Adresse mit genau
        //    dieser Nummer ist deshalb sicher ein Test — alles andere nicht.
        const usable: QrMailHit[] = [];
        for (const h of scan.hits) {
          if (!h.redirected && orgSet[h.recipient] && h.qrId === SAMPLE_QR_ID) { skipped++; continue; }
          usable.push(h);
        }
        // 5) Je Person gewinnt die Mail mit der HÖCHSTEN Id — ein zweiter
        //    Versand ersetzt den ersten.
        //    v31.4 (Review): ABER eine Zeile mit Status 'Failed' hat nie eine
        //    Mail zugestellt und damit nie eine Nummer gedruckt. Sie darf eine
        //    tatsächlich verschickte nicht verdrängen; nur wenn es zu dieser
        //    Adresse gar keine zugestellte Zeile gibt, zählt sie ersatzweise.
        const failedFirst = (h: QrMailHit): number => (h.status === 'Failed' ? 1 : 0);
        const byEmail: Record<string, QrMailHit> = {};
        for (const h of usable) {
          const k = h.email;
          if (!k) continue;
          const cur = byEmail[k];
          if (!cur) { byEmail[k] = h; continue; }
          const rank = failedFirst(cur) - failedFirst(h);
          if (rank > 0 || (rank === 0 && h.mailId > cur.mailId)) byEmail[k] = h;
        }
        // 6) Zeilen bilden — erst je Event sammeln, damit Schritt 7 die
        //    Eindeutigkeit auf DIESER Teilnehmerliste prüfen kann.
        const evRows: BackfillRow[] = [];
        for (const r of regs) {
          if (BACKFILL_STATI.indexOf(r.Status || '') < 0) continue;
          const email = (r.ParticipantEmail || '').toLowerCase().trim();
          const hit = email ? byEmail[email] : undefined;
          const cur = (r.TeilnehmerID !== undefined && r.TeilnehmerID !== null && isFinite(Number(r.TeilnehmerID)))
            ? Number(r.TeilnehmerID) : null;
          let verdict: Verdict = 'none';
          if (hit) {
            // Ist die Person selbst Organizer und ging die Mail direkt an sie,
            // lässt sich Test von echt nicht sicher trennen — dann wird die
            // Zeile markiert und der Mensch entscheidet.
            if (orgSet[email] && !hit.redirected) verdict = 'organizer';
            else verdict = (cur !== null && hit.qrId === cur) ? 'same' : 'diff';
          }
          evRows.push({
            key: `${ev.id}:${r.Id}`,
            eventTitle: ev.title,
            subsiteUrl: ev.subsiteUrl as string,
            regId: r.Id,
            name: (r.Vorname && r.Nachname) ? `${r.Vorname} ${r.Nachname}` : (r.ParticipantName || r.ParticipantEmail || '-'),
            email: r.ParticipantEmail || '',
            current: cur,
            qrId: hit ? hit.qrId : null,
            verdict,
            // v31.4 (Review): 'organizer' heißt „unentscheidbar" — und
            // unbekannt SPERRT, statt freizugeben (CLAUDE.md). Vorausgewählt
            // hätte der Knopf per Vorgabe eine Testmail-Nummer geschrieben,
            // und `QrSentId` schlägt am Check-in die laufende Nummer.
            include: !!hit && verdict !== 'organizer',
          });
        }
        // 7) v31.4 (Review): Eindeutigkeit je Teilnehmerliste. Zwei Zeilen mit
        //    derselben `QrSentId` sperren den ID-Check-in für beide — solche
        //    Zeilen werden benannt und NICHT geschrieben.
        const perNumber: Record<string, number> = {};
        evRows.forEach(r => { if (r.qrId !== null) { const k = String(r.qrId); perNumber[k] = (perNumber[k] || 0) + 1; } });
        evRows.forEach(r => {
          if (r.qrId !== null && perNumber[String(r.qrId)] > 1) { r.verdict = 'dup'; r.include = false; }
        });
        evRows.forEach(r => collected.push(r));
      }
      if (cancelled) return;
      // Reihenfolge: erst was abweicht (dort passiert etwas), dann die zu
      // prüfenden, dann die ohne Fund, zuletzt die unveränderten.
      const rank = (v: Verdict): number => (v === 'diff' ? 0 : v === 'dup' ? 1 : v === 'organizer' ? 2 : v === 'none' ? 3 : 4);
      collected.sort((a, b) => (rank(a.verdict) - rank(b.verdict)) || a.name.localeCompare(b.name));
      setRows(collected);
      setTestSkipped(skipped);
      setScannedMails(scanned);
      setUnparsedMails(unparsed);
      setPhase('preview');
    })().catch(() => {
      if (cancelled) return;
      setAbortMsg(isDe
        ? 'Unerwarteter Fehler beim Lesen — abgebrochen, es wurde nichts geschrieben.'
        : 'Unexpected error while reading — aborted, nothing was written.');
      setResultIsError(true);
      setPhase('done');
    });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const found = rows.filter(r => r.qrId !== null);
  const diffCount = rows.filter(r => r.verdict === 'diff').length;
  const sameCount = rows.filter(r => r.verdict === 'same').length;
  const orgCount = rows.filter(r => r.verdict === 'organizer').length;
  const dupRows = rows.filter(r => r.verdict === 'dup');
  const missing = rows.filter(r => r.verdict === 'none');
  const writable = rows.filter(r => r.qrId !== null && r.include);

  const sleep = (ms: number): Promise<void> => new Promise(resolve => window.setTimeout(resolve, ms));

  const runWrite = async (): Promise<void> => {
    setPhase('writing');
    setProgress({ done: 0, total: writable.length });
    let ok = 0;
    const failed: BackfillRow[] = [];
    for (let i = 0; i < writable.length; i++) {
      if (!aliveRef.current) return;
      const row = writable[i];
      // v31.4 (Review): Derselbe geprüfte Weg wie `queueIDReorderChecked` —
      // ein MERGE je Person am Stück ist genau die Last, bei der SharePoint
      // ab der Hälfte drosselt. Ein einzelner Versuch mit `errors++` hätte
      // die Hälfte der Liste still auf die laufende Nummer zurückfallen
      // lassen, während `eventHasQrIds` am Tisch schon true ist.
      const delays = [1500, 4000, 8000];
      let r = await service.setQrSentId(row.subsiteUrl, row.regId, row.qrId as number);
      for (let a = 0; a < delays.length && !r.ok; a++) {
        // 400 (Spalte fehlt), 403 (keine Rechte) heilen nicht durch Warten.
        if (r.status === 400 || r.status === 403 || r.status === 401) break;
        await sleep(delays[a]);
        if (!aliveRef.current) return;
        r = await service.setQrSentId(row.subsiteUrl, row.regId, row.qrId as number);
      }
      if (!r.ok) {
        if (r.status === 400) {
          // Genau eine Abhilfe — und der Lauf bricht ab, statt 87-mal
          // denselben Fehler zu erzeugen.
          if (!aliveRef.current) return;
          setResultIsError(true);
          setResultMsg(isDe
            ? `Abgebrochen bei „${row.eventTitle}": Auf dieser Teilnehmerliste fehlt die Spalte QrSentId. Bitte einmal „Spalten fixen" für dieses Event ausführen und danach erneut nachtragen. Bis dahin geschrieben: ${ok}.`
            : `Aborted at “${row.eventTitle}”: the attendee list is missing the QrSentId column. Please run “Fix columns” for this event once, then backfill again. Written so far: ${ok}.`);
          setPhase('done');
          if (onDone) onDone();
          return;
        }
        failed.push(row);
      } else ok++;
      if (!aliveRef.current) return;
      setProgress({ done: i + 1, total: writable.length });
    }
    if (!aliveRef.current) return;
    // v31.4 (Review): Zahlen allein reichen nicht — „35 fehlgeschlagen" sagt
    // dem Organizer nicht, WEN es trifft, und genau diese Personen fallen am
    // Tisch still auf die laufende Nummer zurück. Der Nachtrag darf beliebig
    // oft laufen (er schreibt nur `QrSentId`), das gehört dazu.
    const nameList = (list: BackfillRow[]): string => {
      const names = list.slice(0, 15).map(r => r.name);
      return names.join(', ') + (list.length > names.length ? ' …' : '');
    };
    setResultIsError(failed.length > 0);
    setResultMsg([
      isDe ? `${ok} Nummer(n) nachgetragen.` : `${ok} number(s) backfilled.`,
      failed.length > 0
        ? (isDe
          ? `${failed.length} fehlgeschlagen (Drosselung oder fehlende Rechte): ${nameList(failed)}. Für sie greift am Check-in weiter die laufende Nummer — du kannst den Nachtrag gefahrlos noch einmal starten, er schreibt nur die Nummer.`
          : `${failed.length} failed (throttling or missing permissions): ${nameList(failed)}. Check-in stays on the running number for them — you can safely run the backfill again, it only writes the number.`)
        : '',
      missing.length > 0
        ? (isDe
          ? `${missing.length} Person(en) ohne Fund: ${nameList(missing)}. Für sie bleibt der Check-in bei der laufenden Nummer.`
          : `${missing.length} person(s) without a match: ${nameList(missing)}. Check-in stays on the running number for them.`)
        : '',
      dupRows.length > 0
        ? (isDe
          ? `${dupRows.length} Zeile(n) wurden ausgelassen, weil dieselbe Nummer auf mehr als eine Zeile fällt: ${nameList(dupRows)}.`
          : `${dupRows.length} row(s) were skipped because the same number falls on more than one row: ${nameList(dupRows)}.`)
        : '',
    ].filter(Boolean).join(' '));
    setPhase('done');
    if (onDone) onDone();
  };

  const verdictPill = (v: Verdict): React.ReactElement => {
    if (v === 'diff') return <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'weicht ab' : 'differs'}</span>;
    if (v === 'same') return <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'gleich' : 'same'}</span>;
    if (v === 'organizer') return <span className="dex-ui-pill dex-ui-pill--blue">{isDe ? 'Organizer — prüfen' : 'organizer — check'}</span>;
    if (v === 'dup') return <span className="dex-ui-pill dex-ui-pill--red">{isDe ? 'doppelte Nummer — nicht schreiben' : 'duplicate number — not written'}</span>;
    return <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'kein Fund' : 'no match'}</span>;
  };

  // v31.4 (Review): Während des Schreibens gibt es KEINEN Schließen-Knopf.
  // `dismissable={phase !== 'writing'}` sperrt Backdrop und Escape genau
  // dafür; ein eigener Knopf daneben hätte diese Absicht umgangen — und der
  // 400-Abbruch weiter oben ist die einzige Stelle, an der die Abhilfe
  // („Spalten fixen") steht. Wer vorher schließt, sieht sie nie.
  const footer = phase === 'preview'
    ? (<>
        <button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
        <button type="button" className="btn btn-primary" disabled={writable.length === 0} onClick={() => { void runWrite(); }}>
          {isDe ? `${writable.length} Nummer(n) schreiben` : `Write ${writable.length} number(s)`}
        </button>
      </>)
    : phase === 'writing'
      ? (<button type="button" className="btn btn-secondary" disabled>{isDe ? 'Bitte warten…' : 'Please wait…'}</button>)
      : (<button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Schließen' : 'Close'}</button>);

  return (
    <Modal
      open
      onClose={onClose}
      maxWidth={860}
      dismissable={phase !== 'writing'}
      ariaLabel={isDe ? 'QR-Nummern nachtragen' : 'Backfill QR numbers'}
      title={isDe ? 'QR-Nummern nachtragen' : 'Backfill QR numbers'}
      subtitle={isDe
        ? 'Holt aus den bereits verschickten QR-Mails zurück, welche Nummer bei welcher Person gedruckt stand. Danach greift am Check-in die Nummer aus der Mail — auch wenn Abmeldungen die laufende Nummer längst verschoben haben.'
        : 'Recovers from the QR emails already sent which number was printed for which person. After that, check-in resolves the number from the email — even when cancellations have long since shifted the running number.'}
      icon={<QrCode size={20} />}
      footer={footer}
    >
      {phase === 'scan' && (
        <p className="dex-ui-muted" style={{ margin: 0 }}>
          {isDe
            ? `Mail-Warteschlange und Teilnehmerlisten werden gelesen (${targets.length} Liste(n))…`
            : `Reading the mail queue and attendee lists (${targets.length} list(s))…`}
        </p>
      )}

      {phase === 'preview' && (
        <>
          <div className="dex-ui-section">
            <div className="dex-ui-section-title">{isDe ? 'Ergebnis der Suche' : 'Search result'}</div>
            <p style={{ margin: '0 0 8px' }}>
              {isDe
                ? <><strong>{found.length}</strong> von <strong>{rows.length}</strong> angemeldeten Personen konnten aus den QR-Mails belegt werden.</>
                : <><strong>{found.length}</strong> of <strong>{rows.length}</strong> registered people could be matched from the QR emails.</>}
            </p>
            <div className="dex-ui-inline">
              <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? `${diffCount} weichen ab` : `${diffCount} differ`}</span>
              <span className="dex-ui-pill dex-ui-pill--green">{isDe ? `${sameCount} gleich` : `${sameCount} same`}</span>
              {orgCount > 0 && <span className="dex-ui-pill dex-ui-pill--blue">{isDe ? `${orgCount} Organizer — prüfen` : `${orgCount} organizer — check`}</span>}
              {dupRows.length > 0 && <span className="dex-ui-pill dex-ui-pill--red">{isDe ? `${dupRows.length} doppelte Nummer` : `${dupRows.length} duplicate number`}</span>}
              {missing.length > 0 && <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? `${missing.length} ohne Fund` : `${missing.length} without a match`}</span>}
              {testSkipped > 0 && <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? `${testSkipped} Test-Mails übersprungen` : `${testSkipped} test emails skipped`}</span>}
            </div>
          </div>

          {/* v31.4 (Review): Der Doppel-Ursachen-Satz hing an „gar kein Fund".
              Im gemischten Fall stand jede „kein Fund"-Zeile als Aussage über
              die Daten da, obwohl es genauso eine Rechte-Frage sein kann: Die
              Auto-QR-Mails der Nachzügler legt die TEILNEHMERIN an, nicht der
              Organizer — wer nicht „Manage Lists" hat, sieht sie nicht. Der
              Satz gehört deshalb an JEDE fundlose Zeile. */}
          {missing.length > 0 && (
            <div className="dex-ui-callout dex-ui-callout--warn">
              <span className="dex-ui-callout-icon" aria-hidden="true"><AlertCircle size={16} /></span>
              <div>
                {found.length === 0
                  ? (isDe
                    ? `In der Mail-Warteschlange wurde keine QR-Mail zu diesem Event gefunden (${scannedMails} Zeile(n) gelesen). Zwei mögliche Ursachen: Es wurden noch keine QR-Mails verschickt — oder du hast keine Leserechte auf DEX_Emails. Die Liste zeigt jeder Person nur die Zeilen, die sie selbst angelegt hat; wenn der Massenversand von einem anderen Organizer lief, sieht ihn nur diese Person oder ein Site-Owner.`
                    : `No QR email for this event was found in the mail queue (${scannedMails} row(s) read). Two possible causes: no QR emails have been sent yet — or you lack read access to DEX_Emails. The list shows each person only the rows they created themselves; if the mass send was run by another organizer, only that person or a site owner can see it.`)
                  : (isDe
                    ? `Bei ${missing.length} Person(en) wurde keine QR-Mail gefunden — das heißt nicht zwingend, dass sie keine bekommen haben. DEX_Emails zeigt jeder Person nur die Zeilen, die sie selbst angelegt hat: Die automatischen QR-Mails der Nachzügler entstehen im Browser der Teilnehmerin, ein anderer Massenversand im Browser der Organizerin, die ihn gestartet hat. Wenn die Zahl unerwartet hoch ist, lass den Nachtrag von einem Site-Owner oder von der Person laufen, die den Versand gemacht hat — statt diesen Personen neue QR-Codes mit neuen Nummern zu schicken.`
                    : `For ${missing.length} person(s) no QR email was found — that does not necessarily mean they never received one. DEX_Emails shows each person only the rows they created themselves: the automatic QR emails of late registrants are queued in the attendee's browser, another mass send in the browser of the organizer who ran it. If the number looks unexpectedly high, have a site owner — or the person who ran the send — run the backfill instead of sending these people new QR codes with new numbers.`)}
              </div>
            </div>
          )}

          {dupRows.length > 0 && (
            <div className="dex-ui-callout dex-ui-callout--warn">
              <span className="dex-ui-callout-icon" aria-hidden="true"><AlertCircle size={16} /></span>
              <div>
                {isDe
                  ? `${dupRows.length} Zeile(n) werden NICHT geschrieben: Dieselbe Nummer fiele dort auf mehr als eine Zeile derselben Teilnehmerliste — am Check-in wäre die Eingabe dieser Zahl danach für alle Betroffenen gesperrt. Ursache ist meist eine doppelte Anmeldezeile derselben Person; die lässt sich in der Teilnehmerliste bereinigen, danach kann der Nachtrag erneut laufen.`
                  : `${dupRows.length} row(s) will NOT be written: the same number would land on more than one row of the same attendee list — typing that number at check-in would then be blocked for everyone involved. The usual cause is a duplicate registration row for the same person; clean that up in the attendee list, then run the backfill again.`}
              </div>
            </div>
          )}

          {unparsedMails > 0 && (
            <p className="dex-ui-muted" style={{ margin: 0, fontSize: '0.76rem' }}>
              {/* v31.4 (Review): `parseQrIdFromBody` liefert auch dann `null`,
                  wenn die Mail gar keinen ID-Block hatte (keine laufende
                  Nummer beim Versand) oder wenn das Markup nicht mehr zum
                  Parser passt. Eine einzige Ursache zu behaupten heißt: Der
                  Organizer hört am Vorabend auf zu suchen. */}
              {isDe
                ? `${unparsedMails} QR-Mail(s) enthielten keine gedruckte Nummer. Mögliche Ursachen: Mails aus der Zeit vor v30.35 (die ID stand noch nicht neben dem Code), Mails an Personen ohne laufende Nummer — oder das Markup passt nicht mehr zum Parser.`
                : `${unparsedMails} QR email(s) contained no printed number. Possible causes: emails predating v30.35 (the ID was not yet shown next to the code), emails to people without a running number — or the markup no longer matches the parser.`}
              {unparsedMails === scannedMails && scannedMails > 0 && (isDe
                ? ' Weil ALLE gelesenen Mails betroffen sind, ist der Parser der wahrscheinlichste Grund — bitte den DEX-Admins melden, bevor du neue Nummern verschickst.'
                : ' Because ALL scanned emails are affected, the parser is the most likely reason — please report this to the DEX admins before sending out new numbers.')}
            </p>
          )}

          {rows.length > 0 && (
            <div className="dex-ui-table-wrap" style={{ maxHeight: 340, overflowY: 'auto' }}>
              <table className="dex-ui-table dex-ui-table--compact">
                <thead>
                  <tr>
                    <th>{isDe ? 'Person' : 'Person'}</th>
                    {targets.length > 1 && <th>{isDe ? 'Termin' : 'Date'}</th>}
                    <th className="is-num">{isDe ? 'aus der Mail' : 'from the email'}</th>
                    <th className="is-num">{isDe ? 'laufend heute' : 'running today'}</th>
                    <th>{isDe ? 'Bewertung' : 'Assessment'}</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map(r => (
                    <tr key={r.key} className={cx(r.verdict === 'none' && 'is-muted')}>
                      <td>
                        <div className="dex-ui-person-name">{r.name}</div>
                        <div className="dex-ui-person-sub">{r.email}</div>
                      </td>
                      {targets.length > 1 && <td>{r.eventTitle}</td>}
                      <td className="is-num" style={{ fontFamily: "'Courier New',Courier,monospace" }}>
                        {r.qrId !== null ? pad3(r.qrId) : '–'}
                      </td>
                      <td className="is-num" style={{ fontFamily: "'Courier New',Courier,monospace" }}>
                        {r.current !== null ? pad3(r.current) : '–'}
                      </td>
                      <td>
                        {verdictPill(r.verdict)}
                        {/* Nur die unsicheren Zeilen sind abwählbar — alles
                            andere wird geschrieben, damit `QrSentId` danach ein
                            vollständiger Datensatz ist und nicht halb gefüllt.
                            v31.4 (Review): Der Haken ist NICHT vorbelegt —
                            hier lässt sich Testmail von echter Mail nicht
                            trennen, und eine Testmail-Nummer schlägt am
                            Check-in die laufende Nummer. */}
                        {r.verdict === 'organizer' && (
                          <label className="dex-ui-inline" style={{ marginTop: 4, fontSize: '0.74rem' }}>
                            <input
                              type="checkbox"
                              checked={r.include}
                              onChange={e => {
                                const on = e.target.checked;
                                setRows(prev => prev.map(x => x.key === r.key ? { ...x, include: on } : x));
                              }}
                            />
                            <span>{isDe ? 'mitschreiben' : 'include'}</span>
                          </label>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {(missing.length > 0 || orgCount > 0) && (
            <p className="dex-ui-muted" style={{ margin: 0, fontSize: '0.78rem' }}>
              {missing.length > 0 && (isDe
                ? `Für die ${missing.length} Person(en) ohne Fund bleibt der Check-in bei der laufenden Nummer — sie stehen oben namentlich in der Tabelle. `
                : `For the ${missing.length} person(s) without a match, check-in stays on the running number — they are listed by name in the table above. `)}
              {orgCount > 0 && (isDe
                ? 'Nicht angehakte Zeilen bleiben ebenfalls bei der laufenden Nummer — es wird nichts geschrieben.'
                : 'Rows that are not ticked also stay on the running number — nothing is written for them.')}
            </p>
          )}
        </>
      )}

      {phase === 'writing' && (
        <p style={{ margin: 0 }}>
          {isDe ? `Wird geschrieben: ${progress.done} von ${progress.total}…` : `Writing: ${progress.done} of ${progress.total}…`}
        </p>
      )}

      {phase === 'done' && (
        <div className={cx('dex-ui-callout', resultIsError ? 'dex-ui-callout--warn' : 'dex-ui-callout--success')}>
          {resultIsError && <span className="dex-ui-callout-icon" aria-hidden="true"><AlertCircle size={16} /></span>}
          <div>{abortMsg || resultMsg}</div>
        </div>
      )}
    </Modal>
  );
}
