/**
 * v30.60: „Benötigte T-Shirts" — die Bestellliste zum Event.
 *
 * Nutzer-Ansage: „Und ich brauche eine Aktion mit benötigte T-Shirts." Gemeint
 * ist die Zahl, die beim Ausstatter bestellt wird — pro Größe, nicht pro
 * Person. Bisher stand die Angabe nur verstreut in den Abfrage-Antworten der
 * Teilnehmerliste; wer bestellen wollte, exportierte nach Excel und zählte von
 * Hand.
 *
 * Zwei Entscheidungen, die die Zahl belastbar machen (Begründung in
 * `utils/checkInExtras.shirtTally`):
 *  - Gezählt wird über DIESELBE Feld-Erkennung wie am Check-in-Tisch.
 *  - Wer keine Größe angegeben hat, taucht namentlich auf, statt aus der
 *    Summe zu verschwinden.
 *
 * Bei einem Klammer-Event mit Terminen wird über ALLE Termine gezählt: Das
 * Trikot bekommt die Person, nicht die Anmeldung. Doppelt gezählt wird
 * niemand — die Zusammenführung läuft über die E-Mail-Adresse (CLAUDE.md:
 * „Die E-Mail-Adresse ist der einzige Schlüssel").
 *
 * v31.3: Diese Zusammenführung behielt je Person die ZUERST gelesene Zeile.
 * Damit gewann bei einem Klammer-Event die Schattenzeile (oft ohne
 * CustomData) und bei mehreren Terminen die zufällig vorne stehende Ebene —
 * eine im Organizer Center korrigierte Größe kam in dieser Liste nie an
 * (Befund 08.09.2026: „L"/„XL" blieben stehen, obwohl längst auf
 * „Herrengröße L"/„XL" geändert). Welche Zeile die Frage beantwortet,
 * entscheidet jetzt `pickShirtAnswerRows` — mit Antwort vor ohne, aktiv vor
 * abgemeldet, sonst die zuletzt geänderte.
 *
 * v30.88: Ist-Bestand je Größe + Gegenvorschläge. Nutzer-Ansage (07.09.2026,
 * B2Run Köln): „neben der Anzahl an Shirts auch ermöglichen, dass man angibt,
 * wie viele Shirts man wirklich hat, und dann wird geguckt, ob es überhaupt
 * passt — und bei jeder Person, wo es nicht mehr passt, ein Gegenvorschlag.
 * Die Funktion brauche ich dann für den Check-in bzw. die Abholung." Der
 * Bestand liegt als Piggyback `_shirtStock` am Hauptevent; die Verteilung
 * rechnet `utils/checkInExtras.shirtAllocate` — dieselbe Funktion, die die
 * Check-in-Seite je Person anzeigt, damit Tisch und Liste dasselbe sagen.
 *
 * v31.4: Die Liste zeigt jetzt auch, was AUSGEGEBEN wurde (Spalte
 * `ShirtIssued`, gesetzt am Check-in-Tisch). Damit sind „Bestand" und
 * „verfügbar" zwei verschiedene Zahlen — beide stehen nebeneinander, sonst
 * bestellt jemand nach einer Zahl, die schon halb im Umlauf ist. Dazu zwei
 * Auskünfte aus dem Live-Fall vom 08.09.2026 („die korrigierte Größe kommt in
 * der Liste nicht an"): Die aufgeklappte Namensliste nennt je Person die
 * Herkunft der gelesenen Zeile und die weiteren Zeilen derselben Person, und
 * von jeder Zeile führt ein Sprung in die Teilnehmerliste, wo „Bearbeiten"
 * sitzt (Nutzer-Wunsch: „Gib mir die Möglichkeit, zu den Teilnehmern zu
 * springen mit diesem Wert").
 *
 * v31.4 (Nachtrag 2): Dieser Sprung läuft je Größen-Zeile über die
 * E-MAIL-ADRESSEN der Zeile (`onJumpToParticipants`), nicht mehr über das
 * Suchfeld. Nutzer-Wunsch 08.09.2026: „Hier würde ich gerne auf ‚L' und ‚XL'
 * klicken können und kriege dann gefiltert die TN mit dieser falschen Größe."
 * Genau das ging über die Textsuche nicht — ein einzelnes „L" trifft jede
 * zweite Adresse, weshalb der Knopf eine Mindestlänge von drei Zeichen hatte
 * und ausgerechnet den beiden Werten fehlte, die korrigiert gehören. Mit dem
 * Adress-Filter fällt die Bedingung weg, und die Zeile „ohne Angabe" bekommt
 * den Knopf ebenfalls.
 *
 * v31.4.3: Zwei Dinge, die derselbe Live-Fall am 08.09.2026 nachgelegt hat —
 * „ich habe die Größe auf ‚Herrengröße L' korrigiert, hier steht weiter ‚L',
 * und der Sprung aus dieser Zeile führt zu genau dieser Person":
 *
 *  1. **Die Ursache lag im Speichern, nicht im Zählen.** Eine Antwort steht an
 *     zwei Stellen derselben Zeile — in `CustomData` und in der SP-Spalte des
 *     Feldes. Der Bearbeiten-Dialog schrieb nur die Spalte, dieses Modul las
 *     nur `CustomData`. Gelesen wird jetzt über `shirtAnswerOf` (Spalte zuerst,
 *     wie in der Teilnehmertabelle), geschrieben werden beide.
 *  2. **Die Feldwahl ist nicht mehr still.** Trifft mehr als ein Abfragefeld
 *     das Trikot-Muster, nennt ein Hinweiskasten alle — mit Herkunft und
 *     Antwortzahl — und der Organizer schaltet um, welches zählt. Die Wahl
 *     fällt EINMAL (`fieldOnly`) und geht in jede Rechnung dieses Dialogs;
 *     vorher wählten `pickShirtAnswerRows`, `shirtTally` und `shirtAllocate`
 *     jeweils selbst, und zwar über verschiedene Zeilenmengen.
 *
 * v31.4 (Nachtrag): Der Ausgabe-Knopf am Tisch ist jünger als der Lauftag —
 * wer vorher eingecheckt wurde, hat sein Trikot trotzdem bekommen. Diese
 * Personen zählen deshalb mit ihrer Wunschgröße als abgeholt (Ansage des
 * Organizers, der am Tisch stand). Es bleibt eine ANNAHME: Sie wird nie in
 * die Spalte geschrieben, und jede Zahl, die sie enthält, nennt sie
 * daneben — Kachel, Spalte, Namensliste und Excel. Ändern lässt sie sich je
 * Person über „Bearbeiten" in der Teilnehmerliste.
 */
import * as React from 'react';
import Modal from '../Modal';
import { useDialog } from '../../context/DialogContext';
import { useEvents } from '../../context/EventContext';
import { useLocaleSafe } from '../../context/LanguageContext';
import { DeloitteEvent } from '../../types';
import { EventService, SPRegistration } from '../../services/EventService';
import {
  shirtTally, ShirtTallyResult, shirtAllocate, ShirtAllocationResult, ShirtStock, parseShirtStock, shirtSizeKey,
  pickShirtAnswerRows, shirtSizeLabel, splitShirtSize, isShirtSizeKey,
  // v31.4: Herkunft der gelesenen Zeile + ausgegebene Trikots.
  shirtFieldOf, parseShirtIssue,
  // v31.4.3: Feldwahl sichtbar machen und EINMAL treffen.
  FieldDef, SHIRT_PATTERN, shirtAnswerOf,
} from '../../utils/checkInExtras';
import { cx } from '../dexUi';
import { Shirt, Download, Plus, Check, ChevronDown, AlertCircle } from '../Icons';

/** v31.4: Woher stammt der Wert, mit dem diese Liste rechnet? */
type ShirtOrigin = {
  /** Titel des Events, dessen Teilnehmerzeile gewonnen hat. */
  source: string;
  /** Ausgegebene Größe, falls die Person ihr Trikot schon hat. */
  issued: string;
  /** Zeilen derselben Person, die NICHT gewonnen haben — mit ihrem Wert. */
  others: Array<{ source: string; value: string }>;
  /**
   * v31.4.3: Abweichende Werte derselben Person in einem ANDEREN Trikot-Feld.
   *
   * Genau dieser Fall war von außen nicht erkennbar: Die Tabelle zeigt das eine
   * Feld, die Bestellliste zählt das andere, beide sagen die Wahrheit über
   * verschiedene Spalten — und niemand sieht, dass es zwei sind.
   */
  otherFields: Array<{ label: string; source: string; value: string }>;
};

/**
 * v31.4.3: Ein Feld, das nach einer Trikot-/Konfektionsgröße aussieht — mit
 * allem, was der Organizer zum Auseinanderhalten braucht.
 */
type ShirtCandidate = {
  id: string;
  label: string;
  spInternalName?: string;
  /** Weitere SP-Spalten desselben Feldes (Klammer + Termine), s. `FieldDef`. */
  spInternalNames: string[];
  /** Titel der Events, auf denen dieses Feld definiert ist. */
  sources: string[];
  /** Zeilen mit irgendeiner Antwort in diesem Feld. */
  answers: number;
  /** Davon Antworten, die wirklich wie eine Größe aussehen. */
  sizeAnswers: number;
};

export default function ShirtSizeModal(props: {
  event: DeloitteEvent;
  onClose: () => void;
  /**
   * v31.4: Sprung in die Teilnehmerliste (Nutzer 08.09.2026: „Gib mir die
   * Möglichkeit, zu den Teilnehmern zu springen mit diesem Wert, und dann kann
   * ich auf Bearbeiten klicken."). Der Aufrufer setzt die Suche und scrollt;
   * ohne die Prop bleibt die Liste wie bisher.
   */
  onJumpToParticipant?: (_query: string) => void;
  /**
   * v31.4: Sprung zu ALLEN Personen einer Größen-Zeile (Nutzer 08.09.2026:
   * „Hier würde ich gerne auf ‚L' und ‚XL' klicken können und kriege dann
   * gefiltert die TN mit dieser falschen Größe."). Bewusst über die
   * E-Mail-Adressen aus `row.people` statt über das Suchfeld: Ein Textsprung
   * mit „L" trifft jede zweite Adresse — deshalb hatte der Knopf bis v31.4
   * eine Mindestlänge von drei Zeichen und ausgerechnet die kurzen Größen,
   * die korrigiert gehören, keinen.
   *
   * Der dritte Parameter zählt die Personen der Zeile OHNE Adresse. Sie
   * lassen sich so nicht filtern; der Aufrufer muss sie in seiner Filterzeile
   * benennen, statt sie stillschweigend wegzulassen.
   */
  onJumpToParticipants?: (_emails: string[], _label: string, _withoutEmail?: number) => void;
}): React.ReactElement {
  // v31.2: Zweisprachig wie jeder andere Dialog — bisher nur Deutsch. Die
  // Props bleiben unverändert (kein `isDe`-Prop), die Sprache kommt aus dem
  // Context; ohne Provider fällt sie auf Deutsch zurück.
  const isDe = useLocaleSafe() === 'de';
  const { getAllRegistrations, events, refreshEvents } = useEvents();
  const { showAlert } = useDialog();
  const [loading, setLoading] = React.useState(true);
  /**
   * v31.4.3: Die ROHEN Zeilen aller Ebenen bleiben liegen — die Auswertung ist
   * jetzt abgeleitet, nicht eingefroren.
   *
   * Grund: Welches Feld gezählt wird, darf der Organizer umschalten (s.
   * `chosenFieldId`). Rechnete der Effect die Zeilen wie bisher einmal fertig,
   * müsste jede Umschaltung neu von SharePoint lesen — und je nach Drosselung
   * mit einem anderen Ergebnis zurückkommen als die Zeile daneben.
   */
  const [rawRegs, setRawRegs] = React.useState<SPRegistration[]>([]);
  /** Aus welchem Event stammt eine Zeile? Über die Objekt-Identität, weil
   *  `pickShirtAnswerRows` filtert und nicht kopiert. */
  const [rowSource, setRowSource] = React.useState<Map<SPRegistration, string>>(() => new Map());
  const [fields, setFields] = React.useState<FieldDef[]>([]);
  /** Feld-Id → Titel der Events, die dieses Feld definieren. */
  const [fieldSources, setFieldSources] = React.useState<Record<string, string[]>>({});
  /**
   * v31.4.3: Vom Organizer gewähltes Größenfeld ('' = die Vorauswahl der App).
   *
   * Vorher entschied `shirtFieldOf` still — bei zwei gleich benannten Feldern
   * sah man weder, dass es zwei gibt, noch welches gewinnt.
   */
  const [chosenFieldId, setChosenFieldId] = React.useState<string>('');
  const [skipped, setSkipped] = React.useState<string[]>([]);
  const [openSize, setOpenSize] = React.useState<string | null>(null);
  const [xlsxBusy, setXlsxBusy] = React.useState(false);
  // v30.88: Bestand — Eingabe als Text je Größe (leer = kein Bestand für die
  // Größe), gespeichert als Zahlen. `stockDirty` hält den Speichern-Knopf an.
  const [stockInput, setStockInput] = React.useState<Record<string, string>>(() => {
    const st = parseShirtStock(props.event.emailTemplateOverrides);
    const out: Record<string, string> = {};
    Object.keys(st).forEach(k => { out[k] = String(st[k]); });
    return out;
  });
  const [savedStock, setSavedStock] = React.useState<ShirtStock>(() => parseShirtStock(props.event.emailTemplateOverrides));
  const [stockSaving, setStockSaving] = React.useState(false);
  const [newSize, setNewSize] = React.useState('');

  React.useEffect(() => {
    let cancelled = false;
    // v31.4.3: Ein anderes Event heißt andere Felder — eine stehengebliebene
    // Feldwahl zeigte sonst auf eine Id, die es hier nicht gibt (dann greift
    // zwar die Vorauswahl, aber der Kasten behauptete eine Wahl, die keine ist).
    setChosenFieldId('');
    (async () => {
      try {
        const children = events.filter(e => e.parentEventId === props.event.id);
        const targets = [props.event, ...children];
        const all: SPRegistration[] = [];
        const failed: string[] = [];
        // v31.4: Woher kam welche Zeile? Die Liste soll ihre Herkunft nennen
        // können — der Live-Fall vom 08.09.2026 („die Korrektur kommt nicht
        // an") ließ sich sonst von außen nicht von „nie gespeichert"
        // unterscheiden. Über die Objekt-Identität, weil `pickShirtAnswerRows`
        // filtert und nicht kopiert.
        const srcOf = new Map<SPRegistration, string>();
        for (const ev of targets) {
          if (!ev.subsiteUrl) continue;
          let ok = true;
          // v30.37-Lehre: Ein leeres Ergebnis ohne geprüften Status ist keine
          // Aussage über die Daten. Ein gesperrter Termin wird deshalb NAMENTLICH
          // gemeldet, statt als „0 Trikots" durchzugehen — sonst bestellt man zu
          // wenig und erfährt den Grund nie.
          const rs = await getAllRegistrations(ev.id, () => { ok = false; });
          if (!ok) { failed.push(ev.title); continue; }
          // v31.3: Hier wird NICHT mehr zusammengeführt — welche Zeile die
          // Größenfrage beantwortet, lässt sich erst sagen, wenn das Feld
          // bekannt ist und alle Ebenen gelesen sind (s. unten).
          for (const r of rs) { all.push(r); srcOf.set(r, ev.title); }
        }
        if (cancelled) return;
        // Die Feld-Definitionen des Hauptevents plus die der Termine: Das
        // Trikot-Feld kann auf beiden Ebenen stehen (CLAUDE.md: Antworten
        // stehen dort, wo angemeldet wurde).
        // v31.4.3: Dazu die Herkunft je Feld-Id. Ohne sie kann der Hinweiskasten
        // unten nicht sagen, WO ein zweites gleichnamiges Feld herkommt — und
        // genau das ist die Frage, die man dann hat.
        const flds: FieldDef[] = [];
        const fsrc: Record<string, string[]> = {};
        targets.forEach(ev => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (((ev as any).eventSpecificFields || []) as FieldDef[]).forEach(f => {
            flds.push(f);
            (fsrc[f.id] = fsrc[f.id] || []).push(ev.title);
          });
        });
        setFields(flds);
        setFieldSources(fsrc);
        setRawRegs(all);
        setRowSource(srcOf);
        setSkipped(failed);
      } catch (err) {
        console.warn('[DEX] Trikot-Auswertung fehlgeschlagen:', err);
        showAlert(isDe ? 'Die Trikotgrößen konnten nicht gelesen werden.' : 'The shirt sizes could not be read.', { variant: 'error' });
      } finally { if (!cancelled) setLoading(false); }
    })().catch(() => { /* im finally behandelt */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.event.id]);

  /**
   * v31.4.3: ALLE Felder, die nach einer Größe aussehen — nicht nur der Sieger.
   *
   * Zusammengefasst wird über die Feld-Id: Dasselbe Feld auf der Klammer und
   * auf drei Terminen ist EINE Frage, keine vier (dieselbe Regel wie
   * `CheckInPage.shirtFieldsFor`). Zwei Felder mit gleicher Bezeichnung, aber
   * verschiedenen Ids sind dagegen wirklich zwei — und genau die kann man von
   * außen nicht auseinanderhalten, weshalb hier Herkunft und Antwortzahl
   * mitgezählt werden.
   */
  const shirtCandidates = React.useMemo((): ShirtCandidate[] => {
    const byId: Record<string, ShirtCandidate> = {};
    const order: string[] = [];
    fields.forEach(f => {
      if (!SHIRT_PATTERN.test(f.label || '')) return;
      const cur = byId[f.id];
      if (cur) {
        // Die SP-Spalte heißt je Liste anders — alle sammeln, sonst liest die
        // Zählung auf der falschen Ebene ins Leere (s. `FieldDef.spInternalNames`).
        const sp = f.spInternalName || '';
        if (sp && sp !== cur.spInternalName && cur.spInternalNames.indexOf(sp) < 0) cur.spInternalNames.push(sp);
        return;
      }
      byId[f.id] = {
        id: f.id, label: f.label, spInternalName: f.spInternalName, spInternalNames: [],
        sources: fieldSources[f.id] || [], answers: 0, sizeAnswers: 0,
      };
      order.push(f.id);
    });
    order.forEach(id => {
      const c = byId[id];
      rawRegs.forEach(r => {
        const v = shirtAnswerOf(c, r);
        if (!v) return;
        c.answers++;
        if (splitShirtSize(v).isSize) c.sizeAnswers++;
      });
    });
    return order.map(id => byId[id]);
  }, [fields, fieldSources, rawRegs]);

  /** Die Vorauswahl der App — dieselbe Regel wie am Check-in-Tisch. */
  const autoField = React.useMemo(() => shirtFieldOf(shirtCandidates, rawRegs), [shirtCandidates, rawRegs]);
  /**
   * v31.4.3: **Die EINE Feldwahl des ganzen Dialogs.**
   *
   * Sie wird bewusst als Liste mit genau EINEM Eintrag weitergereicht:
   * `pickShirtAnswerRows`, `shirtTally` und `shirtAllocate` rufen intern alle
   * wieder `shirtFieldOf` — und zwar mit unterschiedlichen Zeilenmengen (einmal
   * alle Zeilen, einmal die schon je Person zusammengeführten). Bei mehr als
   * einem Treffer entscheiden dort die Antworten, und eine andere Eingabemenge
   * kann einen anderen Sieger ergeben: Dann rechnete die Namensliste mit Feld A
   * und die Bestellsumme mit Feld B, im selben Dialog, ohne dass es irgendwo
   * steht. Mit genau einem Treffer ist `shirtFieldOf` per Vertrag ein
   * Durchreicher (`if (hits.length <= 1 …) return hits[0]`) — die Wahl fällt
   * hier, einmal, sichtbar.
   */
  const field = React.useMemo((): ShirtCandidate | undefined => {
    if (chosenFieldId) {
      const hit = shirtCandidates.filter(c => c.id === chosenFieldId)[0];
      if (hit) return hit;
    }
    return autoField as ShirtCandidate | undefined;
  }, [chosenFieldId, shirtCandidates, autoField]);
  /** Genau diese Liste geht in JEDE Rechnung dieses Dialogs — s. oben. */
  const fieldOnly = React.useMemo((): FieldDef[] => (field ? [field] : []), [field]);

  /**
   * v31.4.3: Zeilenwahl, Ausgabe-Übertrag und Herkunft — abgeleitet statt im
   * Effect gerechnet, damit ein Umschalten des Feldes alles mitzieht.
   */
  const picked = React.useMemo(() => {
    const all = rawRegs;
    // v31.3: Eine Zeile je Person — die, die die Größenfrage wirklich
    // beantwortet (Antwort vor keiner Antwort, aktiv vor abgemeldet, sonst
    // die zuletzt geänderte). Vorher gewann die zuerst gelesene Ebene.
    const p = pickShirtAnswerRows(fieldOnly, all);
    // v31.4: Zwei Nachbesserungen an den gewählten Zeilen:
    //  1. Die AUSGABE steht dort, wo eingecheckt wurde — meist auf der
    //     Termin-Zeile, während die Größenfrage die Klammer-Zeile gewinnen
    //     kann. Ohne diesen Übertrag zählt die Bestellliste ein Trikot als
    //     „noch im Karton", das längst jemand trägt.
    //  2. Die Herkunft je Person, damit die aufgeklappte Namensliste sagen
    //     kann, welche Zeile gemeint ist (und welche es sonst noch gibt).
    const issuedByEmail: Record<string, string> = {};
    // v31.4 (Nachtrag): Der Check-in steht auf derselben Zeile wie die
    // Ausgabe — und damit ebenfalls meist NICHT auf der Zeile, die die
    // Größenfrage gewinnt. Deshalb hier über ALLE Zeilen gesammelt.
    const checkedIn: Record<string, true> = {};
    const rowsByEmail: Record<string, SPRegistration[]> = {};
    all.forEach(r => {
      const em = (r.ParticipantEmail || '').toLowerCase().trim();
      if (!em) return;
      (rowsByEmail[em] = rowsByEmail[em] || []).push(r);
      if (r.ShirtIssued && !issuedByEmail[em]) issuedByEmail[em] = r.ShirtIssued;
      if (r.Status === 'Eingecheckt') checkedIn[em] = true;
    });
    const org: Record<string, ShirtOrigin> = {};
    const rows = p.rows.map(r => {
      const em = (r.ParticipantEmail || '').toLowerCase().trim();
      if (!em) return r;
      const iss = parseShirtIssue(r.ShirtIssued || issuedByEmail[em]);
      const counted = shirtAnswerOf(field, r);
      // v31.4.3: Was steht bei derselben Person in den ANDEREN Trikot-Feldern?
      // Nur Abweichendes wird genannt — ein gleicher Wert ist keine Auskunft,
      // sondern Lärm.
      const otherFields: Array<{ label: string; source: string; value: string }> = [];
      const seen: Record<string, true> = {};
      shirtCandidates.forEach(c => {
        if (!field || c.id === field.id) return;
        (rowsByEmail[em] || []).forEach(row => {
          const v = shirtAnswerOf(c, row);
          if (!v || v.toLowerCase() === (counted || '').toLowerCase()) return;
          const key = `${c.id}|${v.toLowerCase()}`;
          if (seen[key]) return;
          seen[key] = true;
          otherFields.push({ label: c.label, source: rowSource.get(row) || '', value: v });
        });
      });
      org[em] = {
        source: rowSource.get(r) || '',
        issued: iss ? iss.size : '',
        others: (p.othersByEmail[em] || []).map(o => ({ source: rowSource.get(o) || '', value: shirtAnswerOf(field, o) })),
        otherFields,
      };
      return (!r.ShirtIssued && issuedByEmail[em]) ? { ...r, ShirtIssued: issuedByEmail[em] } : r;
    });
    return { rows, origins: org, conflicts: p.conflicts, checkedInEmails: checkedIn };
  }, [rawRegs, fieldOnly, field, shirtCandidates, rowSource]);

  const regs = picked.rows;
  const origins = picked.origins;
  const conflicts = picked.conflicts;
  const checkedInEmails = picked.checkedInEmails;
  const result: ShirtTallyResult = React.useMemo(() => shirtTally(fieldOnly, regs), [fieldOnly, regs]);

  // Bestand aus den Eingaben (nur gültige Zahlen zählen).
  const stockFromInput = React.useMemo((): ShirtStock => {
    const out: ShirtStock = {};
    Object.keys(stockInput).forEach(k => {
      const v = stockInput[k].trim();
      if (v === '') return;
      const n = parseInt(v, 10);
      if (isFinite(n) && n >= 0) out[k] = n;
    });
    return out;
  }, [stockInput]);
  const stockDirty = React.useMemo(() => {
    const a = Object.keys(stockFromInput).sort(); const b = Object.keys(savedStock).sort();
    if (a.join('|') !== b.join('|')) return true;
    return a.some(k => stockFromInput[k] !== savedStock[k]);
  }, [stockFromInput, savedStock]);
  // Die Verteilung rechnet LIVE mit den Eingaben — der Organizer sieht sofort,
  // was eine Zahl mehr oder weniger bedeutet; gespeichert wird bewusst extra.
  // v31.4.3: `fieldOnly` statt aller Felder — dieselbe eine Wahl wie oben,
  // sonst sucht sich die Verteilung ihr Feld selbst (s. `field`).
  const alloc: ShirtAllocationResult | null = React.useMemo(
    () => result.fieldLabel ? shirtAllocate(fieldOnly, regs, stockFromInput, checkedInEmails) : null,
    [result, fieldOnly, regs, stockFromInput, checkedInEmails],
  );

  const saveStock = async (): Promise<void> => {
    if (stockSaving) return;
    setStockSaving(true);
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (!ctx) throw new Error('Kein SPFx-Kontext');
      const svc = new EventService(ctx);
      const ok = await svc.patchEventOverridesValue(Number(props.event.id), '_shirtStock', Object.keys(stockFromInput).length ? stockFromInput : null);
      if (!ok) throw new Error('patch failed');
      setSavedStock({ ...stockFromInput });
      await refreshEvents();
      showAlert(isDe ? 'Bestand gespeichert — die Check-in-Seite zeigt Gegenvorschläge jetzt je Person an.' : 'Stock saved — the check-in page now shows the alternative size per person.', { variant: 'success' });
    } catch (err) {
      console.warn('[DEX] Trikot-Bestand speichern fehlgeschlagen:', err);
      showAlert(isDe ? 'Der Bestand konnte nicht gespeichert werden — bitte erneut versuchen.' : 'The stock could not be saved — please try again.', { variant: 'error' });
    } finally { setStockSaving(false); }
  };

  const downloadXlsx = async (): Promise<void> => {
    if (xlsxBusy || !result) return;
    setXlsxBusy(true);
    try {
      const hasStock = !!(alloc && alloc.hasStock);
      // v31.4: „Ausgegeben" ans ENDE — die bestehende Spaltenfolge bleibt, wer
      // die Datei jedes Jahr gleich liest, findet sich weiter zurecht.
      // v31.4 (Nachtrag): „davon angenommen" als eigene Spalte hinter
      // „Ausgegeben" — in einer weitergereichten Datei muss stehen, welcher
      // Teil der Zahl aus der Spalte kommt und welcher aus dem Check-in-Status.
      const rows: string[][] = [hasStock
        ? ['Größe', 'Benötigt', 'Bestand', 'Fehlt', 'Reserve', 'Personen', 'Ausgegeben', 'davon angenommen']
        : ['Größe', 'Anzahl', 'Personen', 'Ausgegeben', 'davon angenommen']];
      for (const r of result.rows) {
        // v31.3: Eine Abwahl-Antwort ist keine Größe — sie steht mit ihrem
        // Wortlaut in der Datei, bekommt aber keine Bestands-Spalten, damit
        // niemand sie beim Ausstatter bestellt.
        const a = (alloc && r.size && !r.optOut) ? alloc.rows.find(x => x.key === shirtSizeKey(r.size)) : undefined;
        const label = r.size ? (r.optOut ? `${r.size} (kein Shirt nötig)` : r.size) : 'ohne Angabe';
        rows.push(hasStock
          ? [label, String(r.count), a ? String(a.stock) : '—', a ? String(a.missing) : '—', a ? String(a.spare) : '—', r.names.join(', '), a ? String(a.issued) : '—', a ? String(a.issuedAssumed) : '—']
          : [label, String(r.count), r.names.join(', '), a ? String(a.issued) : '—', a ? String(a.issuedAssumed) : '—']);
      }
      // v31.4 (Review): Größen, die NUR ausgegeben wurden, haben keine
      // Tally-Zeile (der Tally kennt nur aktive Wünsche). Ohne diese Ergänzung
      // summiert sich die Spalte „Ausgegeben" nicht auf die Summenzeile
      // „davon bereits ausgegeben" — und wer den Karton nachzählt, sucht die
      // fehlende Zeile.
      if (alloc) {
        const tallyKeys: Record<string, true> = {};
        result.rows.forEach(r => { if (r.size && !r.optOut) tallyKeys[shirtSizeKey(r.size)] = true; });
        alloc.rows.forEach(a2 => {
          if (tallyKeys[a2.key] || a2.issued <= 0) return;
          rows.push(hasStock
            ? [a2.size, '0', String(a2.stock), String(a2.missing), String(a2.spare), '', String(a2.issued), String(a2.issuedAssumed)]
            : [a2.size, '0', '', String(a2.issued), String(a2.issuedAssumed)]);
        });
      }
      rows.push([]);
      rows.push(['Summe', String(result.total), '']);
      // v31.3: Was davon wirklich bestellt wird — Abwahl und fehlende Angaben
      // sind keine Shirts, standen aber bisher stillschweigend in derselben Summe.
      rows.push(['davon kein Shirt nötig', String(result.optOut), '']);
      rows.push(['davon ohne Größenangabe', String(result.missing), '']);
      rows.push(['Shirts zu bestellen', String(result.sizeTotal), '']);
      // v31.4: Was am Tisch wirklich rausgegangen ist — und was danach noch da
      // sein müsste. Die Bestellliste wird weitergereicht; die Abendzahl gehört
      // deshalb in die Datei, nicht nur in den Dialog.
      if (totalIssued > 0) {
        // v31.4 (Review): Dieselbe Beschriftung wie im Dialog. Ist ein Termin
        // nicht lesbar, ist „ausgegeben" eine Untergrenze und „im Karton" eine
        // OBERGRENZE (die fehlenden Ausgaben lassen ihn rechnerisch wachsen) —
        // in einer Datei, die weitergereicht wird, ist eine harte Zahl dafür
        // die schlechteste Wahl.
        rows.push([partial ? 'davon bereits ausgegeben (mind.)' : 'davon bereits ausgegeben', String(totalIssued), '']);
        // v31.4 (Nachtrag): Eine Annahme, die in einer Excel als harte Zahl
        // steht, ist die schlechteste Sorte Zahl — deshalb steht sie hier
        // getrennt und mit ihrem Grund.
        if (totalIssuedAssumed > 0) rows.push([`davon angenommen (eingecheckt, ohne Ausgabe-Eintrag)`, String(totalIssuedAssumed), '']);
        if (hasStock) rows.push([partial ? 'rechnerisch noch im Karton (höchstens)' : 'rechnerisch noch im Karton', String(totalInBox), '']);
      }
      if (skipped.length > 0) {
        // v31.3: Der Hinweis gehört IN die Datei — eine Bestellliste wird
        // weitergereicht, der rote Kasten im Dialog bleibt zurück.
        rows.push([]);
        rows.push(['ACHTUNG: Die Zahlen oben sind UNVOLLSTÄNDIG — für diese Termine konnte die Teilnehmerliste nicht gelesen werden:']);
        skipped.forEach(t => rows.push([`· ${t}`]));
        rows.push(['Die Werte sind damit Untergrenzen. Vor dem Bestellen die Leserechte prüfen und neu laden.']);
      }
      const XLSX = await import('xlsx');
      const ws = XLSX.utils.aoa_to_sheet(rows);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = hasStock
        ? [{ wch: 16 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 8 }, { wch: 90 }, { wch: 12 }, { wch: 18 }]
        : [{ wch: 16 }, { wch: 10 }, { wch: 90 }, { wch: 12 }, { wch: 18 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Trikots');
      if (hasStock && alloc) {
        const prop: string[][] = [['Person', 'E-Mail', 'Wunschgröße', 'Vorschlag']];
        for (const r of regs) {
          const em = (r.ParticipantEmail || '').toLowerCase().trim();
          const a = alloc.byEmail[em];
          if (!a || !a.short) continue;
          prop.push([(r.ParticipantName || em).trim(), r.ParticipantEmail || '', a.wish, a.proposal || 'keine Größe mehr vorrätig']);
        }
        if (skipped.length > 0) {
          prop.push([]);
          prop.push([`ACHTUNG: unvollständig — nicht lesbare Termine: ${skipped.join(', ')}`]);
        }
        const ws2 = XLSX.utils.aoa_to_sheet(prop);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (ws2 as any)['!cols'] = [{ wch: 30 }, { wch: 34 }, { wch: 14 }, { wch: 26 }];
        XLSX.utils.book_append_sheet(wb, ws2, 'Gegenvorschläge');
      }
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Trikots_${new Date().toISOString().slice(0, 10)}.xlsx`;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
    } catch (err) {
      console.warn('[DEX] Trikot-Export fehlgeschlagen:', err);
      showAlert(isDe ? 'Die Excel-Datei konnte nicht erzeugt werden.' : 'The Excel file could not be created.', { variant: 'error' });
    } finally { setXlsxBusy(false); }
  };

  const maxCount = result ? result.rows.reduce((m, r) => Math.max(m, r.count), 0) : 0;
  // Zeilen für den Bestand: alle gewünschten Größen plus alle, für die schon
  // ein Bestand eingetragen ist (auch wenn niemand sie wollte — die Reserve).
  const stockKeys = React.useMemo((): Array<{ key: string; label: string }> => {
    const seen: Record<string, string> = {};
    // v31.3: Nur echte Größen bekommen eine Bestandszeile. Eine Abwahl
    // („T-Shirt bereits aus Vorjahren vorhanden") hatte bis dahin ein
    // Bestandsfeld, als wäre sie eine Bestellposition — 30 von 90 Personen im
    // B2Run-Event standen so in der Bestellung.
    (result ? result.rows : []).forEach(r => { if (r.size && !r.optOut) seen[shirtSizeKey(r.size)] = r.size; });
    Object.keys(stockInput).forEach(k => { if (!seen[k] && isShirtSizeKey(k)) seen[k] = shirtSizeLabel(k); });
    // v31.4 (Review): `shirtAllocate` legt seit v31.4 ausdrücklich auch für
    // Größen eine Zeile an, die NUR ausgegeben wurden (Reserve aus einem
    // anderen Karton, oder die einzige Wünscherin ist inzwischen No-Show).
    // Der Filter `!!seen[k]` unten warf sie wieder weg: Die Kachel zählte die
    // Ausgabe, die Tabelle kannte die Größe nicht — und ohne Zeile gibt es
    // auch kein Bestandsfeld, über das man sie je in die Rechnung holen
    // könnte. Der Vertrag wurde geändert, der Aufrufer war nicht nachgezogen.
    (alloc ? alloc.rows : []).forEach(r => { if (!seen[r.key] && r.issued > 0 && r.size) seen[r.key] = r.size; });
    // v31.2: Die Schlüssel aus `seen` hinten anhängen — eine über „Größe
    // ergänzen" neu angelegte Größe steht mit leerem Wert noch nicht in
    // `stockFromInput`, also auch nicht in `alloc.rows`, und fiel bis dahin
    // aus der Liste, bevor man ihr eine Zahl geben konnte.
    return (alloc ? alloc.rows.map(r => r.key).concat(Object.keys(seen)) : Object.keys(seen))
      .filter((k, i, arr) => arr.indexOf(k) === i && !!seen[k])
      .map(k => ({ key: k, label: seen[k] }));
  }, [result, stockInput, alloc]);
  const proposals = React.useMemo(() => {
    if (!alloc || !alloc.hasStock) return [];
    return regs
      .map(r => {
        const em = (r.ParticipantEmail || '').toLowerCase().trim();
        const a = alloc.byEmail[em];
        return a && a.short ? { name: (r.ParticipantName || em).trim(), tid: r.TeilnehmerID, wish: a.wish, proposal: a.proposal } : null;
      })
      .filter((x): x is { name: string; tid: number | undefined; wish: string; proposal: string | null } => !!x);
  }, [alloc, regs]);

  // v31.2: Bedarf und Bestand stehen in EINER Tabelle je Größe (vorher zwei
  // Blöcke zum selben Thema: Balkenliste oben, Eingabe-Raster unten). Die
  // Zählung je Größe wird dafür über den Schlüssel nachgeschlagen; die Zeile
  // „ohne Angabe" hat keinen Schlüssel und kommt zuletzt.
  const hasStock = !!(alloc && alloc.hasStock);
  const ready = !loading && !!result && !!result.fieldLabel;
  const tallyByKey: Record<string, { count: number; names: string[]; people: Array<{ name: string; email: string }> }> = {};
  (result ? result.rows : []).forEach(r => { if (r.size && !r.optOut) tallyByKey[shirtSizeKey(r.size)] = { count: r.count, names: r.names, people: r.people }; });
  // v31.3: Abwahl-Antworten sind eine eigene Gruppe zwischen Größen und
  // „ohne Angabe" — sichtbar (der Organizer will wissen, wie viele keins
  // brauchen), aber ohne Bestand, ohne Soll/Ist und ohne Bestellposition.
  const optOutRows = result ? result.rows.filter(r => !!r.size && !!r.optOut) : [];
  const optOutTotal = result ? result.optOut : 0;
  const noneRow = result ? result.rows.filter(r => !r.size)[0] : undefined;
  const totalMissing = alloc ? alloc.rows.reduce((s, r) => s + r.missing, 0) : 0;
  // v31.4: Was ist raus, was liegt noch da? „Im Karton" ist Bestand minus
  // Ausgaben — nicht `spare`: Das ist der Rest NACH der geplanten Verteilung
  // und beantwortet eine andere Frage (reicht es für die Wünsche?).
  const totalIssued = alloc ? alloc.rows.reduce((s, r) => s + r.issued, 0) : 0;
  const totalInBox = alloc ? alloc.rows.reduce((s, r) => s + Math.max(0, r.stock - r.issued), 0) : 0;
  // v31.4 (Nachtrag): Davon ist ein Teil ANGENOMMEN — eingecheckte Personen
  // ohne Ausgabe-Eintrag. Die Zahl steht überall neben `totalIssued`, nie
  // allein: „5 ausgegeben" und „5 ausgegeben, davon 3 angenommen" sind zwei
  // verschiedene Auskünfte, und nur die zweite ist wahr.
  const totalIssuedAssumed = alloc ? alloc.rows.reduce((s, r) => s + r.issuedAssumed, 0) : 0;
  // v31.3: Solange ein Termin nicht lesbar ist, ist JEDE dieser Zahlen eine
  // Untergrenze — dann wird sie auch so beschriftet (CLAUDE.md: ein
  // Lesefehler ist keine Null).
  const partial = skipped.length > 0;
  const atLeast = (n: number): string => partial ? (isDe ? `mind. ${n}` : `at least ${n}`) : String(n);
  const pillInTitle: React.CSSProperties = { textTransform: 'none', letterSpacing: 0 };
  const stickyTh: React.CSSProperties = { position: 'sticky', top: 0 };

  // Eine Zeile der Größen-Tabelle; `key` ist der Aufklapp-Schlüssel für die Namen.
  // v31.3: drei Arten — echte Größe, Abwahl („habe schon eins"), ohne Angabe.
  // v31.4: `people` (mit E-Mail) statt nur Namen — daran hängen der Sprung in
  // die Teilnehmerliste und die Herkunfts-Zeile.
  const renderSizeRow = (
    key: string, label: string, count: number, names: string[],
    kind: 'size' | 'optout' | 'none',
    people?: Array<{ name: string; email: string }>,
  ): React.ReactElement => {
    const open = openSize === key;
    const pct = maxCount > 0 ? Math.round((count / maxCount) * 100) : 0;
    // v31.4: Die Zeile der Verteilung wird auch OHNE Bestand gebraucht — die
    // Ausgabe-Spalte hängt nicht daran, ob jemand Kartons gezählt hat.
    const row = (kind === 'size' && alloc) ? alloc.rows.filter(x => x.key === key)[0] : undefined;
    const a = hasStock ? row : undefined;
    const barColor = kind === 'none' ? 'var(--dex-orange, #ed8b00)' : kind === 'optout' ? 'var(--dex-gray-400, #9e9e9e)' : 'var(--dex-green, #86bc25)';
    return (
      <React.Fragment key={key}>
        <tr>
          <td style={{ fontWeight: 700, fontSize: '0.95rem', color: kind === 'none' ? 'var(--dex-orange-dark, #b35a00)' : kind === 'optout' ? 'var(--dex-gray-600)' : undefined, whiteSpace: kind === 'size' ? 'nowrap' : undefined }}>{label}</td>
          <td>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <strong style={{ minWidth: 24, textAlign: 'right', fontSize: '0.95rem' }}>{count}</strong>
              {/* Balken statt nur Zahl: Beim Bestellen zählt vor allem, welche Größen die Masse ausmachen. */}
              {count > 0
                ? <span style={{ flex: 1, minWidth: 50, height: 8, borderRadius: 999, background: 'var(--dex-gray-100)' }}>
                  <span style={{ display: 'block', height: 8, width: `${pct}%`, borderRadius: 999, background: barColor, transition: 'width 0.18s ease' }} />
                </span>
                : <span className="dex-ui-muted">{isDe ? 'niemand gewünscht' : 'nobody asked'}</span>}
            </div>
          </td>
          <td>
            {kind !== 'size' ? <span className="dex-ui-muted">—</span> : (
              <input type="number" min={0} inputMode="numeric" className="dex-ui-input dex-ui-input--sm" placeholder="0"
                aria-label={`${isDe ? 'Bestand' : 'Stock'} ${label}`} value={stockInput[key] ?? ''} style={{ width: 76, textAlign: 'center', fontWeight: 700 }}
                onChange={e => setStockInput(prev => ({ ...prev, [key]: e.target.value.replace(/[^\d]/g, '') }))} />
            )}
          </td>
          <td>
            {/* v31.4: Was wirklich raus ist. „Bestand" ist die Zahl, die der
                Organizer eingetragen hat — im Karton liegt sie MINUS der
                Ausgaben. Beides steht nebeneinander, damit niemand die eine
                für die andere hält. */}
            {kind !== 'size' || !row ? <span className="dex-ui-muted">—</span> : row.issued > 0 ? (
              <span>
                <strong>{row.issued}</strong>
                {/* v31.4 (Nachtrag): Angenommene Ausgaben zählen mit, heißen aber
                    so. Ohne den Zusatz läse der Organizer eine Schätzung als
                    Protokoll — und suchte abends nach Trikots, die nie jemand
                    festgehalten hat. */}
                {row.issuedAssumed > 0 && (
                  <span className="dex-ui-muted"> {isDe ? `(davon ${row.issuedAssumed} angenommen)` : `(${row.issuedAssumed} assumed)`}</span>
                )}
                {row.stock > 0 && (
                  <span className="dex-ui-muted" style={{ whiteSpace: 'nowrap' }}> · {isDe ? `noch ${Math.max(0, row.stock - row.issued)} im Karton` : `${Math.max(0, row.stock - row.issued)} left in the box`}</span>
                )}
              </span>
            ) : <span className="dex-ui-muted">0</span>}
          </td>
          <td>
            {/* v30.88: Soll/Ist-Abgleich je Zeile, sobald ein Bestand eingetragen ist. */}
            {kind === 'none'
              ? <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'nachfragen' : 'ask them'}</span>
              : kind === 'optout'
                ? <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'kein Shirt nötig' : 'no shirt needed'}</span>
                : a
                  ? (a.missing > 0
                    ? <span className="dex-ui-pill dex-ui-pill--red">{isDe ? `fehlt ${a.missing}` : `${a.missing} short`}</span>
                    // v31.4 (Review): Über-Ausgabe ist derselbe Mangel, nur
                    // später bemerkt. Der Annahme-Vorlauf zieht auch dann ein
                    // Stück ab, wenn der Bestand längst leer ist; ohne diesen
                    // Zweig kippte „fehlt 5" am Lauftag auf „reicht" — genau
                    // in dem Moment, in dem die fünf Trikots real fehlen.
                    : a.over > 0
                      ? <span className="dex-ui-pill dex-ui-pill--red">{isDe ? `${a.over} mehr ausgegeben als da` : `${a.over} more handed out than in stock`}</span>
                      // v31.3: Mit einem gesperrten Termin ist „reicht" eine Zusage,
                      // die die Daten nicht decken — dann heißt es „reicht bisher".
                      : partial
                        ? <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'reicht bisher' : 'enough so far'}{a.spare > 0 ? ` (+${a.spare})` : ''}</span>
                        : <span className="dex-ui-pill dex-ui-pill--green"><Check size={12} /> {isDe ? 'reicht' : 'enough'}{a.spare > 0 ? ` (+${a.spare})` : ''}</span>)
                  : <span className="dex-ui-muted">—</span>}
          </td>
          <td style={{ textAlign: 'right', whiteSpace: 'nowrap' }}>
            {/* v31.4: Sprung in die Teilnehmerliste, wo „Bearbeiten" sitzt —
                seit dem Nachtrag über die ADRESSEN dieser Zeile statt über das
                Suchfeld. Damit fällt die alte Mindestlänge von drei Zeichen
                weg: „L" und „XL" waren genau die Werte, die der Organizer
                korrigieren wollte, und ein Textsprung mit „L" hätte jede
                zweite Adresse getroffen. Auch „ohne Angabe" bekommt den Knopf
                — „wer hat keine Größe angegeben?" ist die nächste Frage, nicht
                die letzte. */}
            {/* Ohne Personen-Liste gäbe es nichts zu filtern — dann bliebe ein
                Knopf stehen, der die Teilnehmerliste leer räumt. */}
            {props.onJumpToParticipants && count > 0 && !!(people && people.length > 0) && (
              <button
                type="button"
                className="dex-ui-textbtn"
                onClick={() => {
                  const list = people || [];
                  // Die Beschriftung nennt das FELD mit: „T-Shirt Größe: L"
                  // steht später allein über einer fremden Tabelle, und dort
                  // sagt ein nacktes „L" niemandem mehr, wonach gefiltert ist.
                  // Die Zeile „ohne Angabe" hat keinen Wert, der sich lesen
                  // liesse — dort steht die Frage selbst.
                  const val = kind === 'none' ? (isDe ? 'ohne Angabe' : 'no answer') : label.trim();
                  const flt = (result && result.fieldLabel) ? `${result.fieldLabel}: ${val}` : val;
                  if (props.onJumpToParticipants) {
                    props.onJumpToParticipants(
                      list.map(p => p.email).filter(e => !!e),
                      flt,
                      // Ohne Adresse kein Filter — die Zahl geht mit, damit der
                      // Aufrufer sie nennen kann (CLAUDE.md: keine stille Kürzung).
                      list.filter(p => !p.email).length,
                    );
                  }
                }}
                title={isDe
                  ? `Teilnehmerliste auf diese ${count === 1 ? 'Person' : `${count} Personen`} filtern`
                  : `Filter the attendee list to ${count === 1 ? 'this person' : `these ${count} people`}`}>
                {isDe ? 'Zu den Teilnehmern' : 'Show participants'}
              </button>
            )}
            {/* Ohne den neuen Rückruf bleibt der alte Textsprung — mit seiner
                Mindestlänge, denn über die Suche gilt der Grund von oben. */}
            {!props.onJumpToParticipants && props.onJumpToParticipant && kind !== 'none' && label.trim().length >= 3 && count > 0 && (
              <button type="button" className="dex-ui-textbtn" onClick={() => props.onJumpToParticipant && props.onJumpToParticipant(label.trim())}
                title={isDe ? `Teilnehmerliste nach „${label}“ durchsuchen` : `Search the attendee list for “${label}”`}>
                {isDe ? 'Zu den Teilnehmern' : 'Show participants'}
              </button>
            )}
            {count > 0 && (
              <button type="button" className="dex-ui-iconbtn" aria-expanded={open} title={isDe ? 'Personen anzeigen' : 'Show people'}
                aria-label={kind === 'size'
                  ? (isDe ? `Personen mit Größe ${label} anzeigen` : `Show people with size ${label}`)
                  : (isDe ? `Personen anzeigen: ${label}` : `Show people: ${label}`)} onClick={() => setOpenSize(open ? null : key)}>
                <span style={{ display: 'inline-flex', transition: 'transform 0.2s ease', transform: open ? 'rotate(180deg)' : 'none' }}><ChevronDown size={16} /></span>
              </button>
            )}
          </td>
        </tr>
        {open && (
          <tr><td colSpan={6} style={{ background: 'var(--dex-gray-50, #fafafa)', color: 'var(--dex-gray-600)', fontSize: '0.82rem', lineHeight: 1.6 }}>
            {/* v31.4: Statt einer Namenskette je Person eine Zeile — mit dem
                Sprung in die Teilnehmerliste und der Herkunft des Werts. Die
                Herkunft ist keine Debug-Ausgabe: Sie beantwortet die Frage
                „warum steht hier noch der alte Wert?" (bearbeitet wurde die
                andere Zeile) in zehn Sekunden. */}
            {(people && people.length > 0) ? (
              <div className="dex-ui-stack" style={{ gap: 8 }}>
                {people.map((p, i) => {
                  const o = p.email ? origins[p.email] : undefined;
                  // v31.4 (Nachtrag): Festgehalten oder angenommen? `origins`
                  // kennt nur die Spalte; die Annahme steht in der Verteilung.
                  // Beides muss verschieden aussehen, sonst ist „abgeholt"
                  // eine Behauptung, die niemand nachprüfen kann.
                  const av = (p.email && alloc) ? alloc.byEmail[p.email] : undefined;
                  const assumed = !!(av && av.issued && av.issuedAssumed);
                  return (
                    <div key={p.email || `${i}`}>
                      <span style={{ fontWeight: 600, color: 'var(--dex-gray-800)' }}>{p.name}</span>
                      {props.onJumpToParticipant && p.email && (
                        <button type="button" className="dex-ui-textbtn" style={{ marginLeft: 8 }}
                          onClick={() => props.onJumpToParticipant && props.onJumpToParticipant(p.email)}
                          title={isDe ? 'In der Teilnehmerliste öffnen — dort sitzt „Bearbeiten"' : 'Open in the attendee list — that is where "Edit" sits'}>
                          {isDe ? 'zu dieser Person' : 'go to this person'}
                        </button>
                      )}
                      {(o || assumed) && (
                        <div className="dex-ui-muted" style={{ fontSize: '0.74rem', lineHeight: 1.5 }}>
                          {/* v31.4.3: Zur Zeile gehört das FELD — bei zwei
                              gleich benannten Feldern ist der Event-Titel
                              allein keine Antwort auf „woher kommt der Wert?". */}
                          {result.fieldLabel && <>{isDe ? 'Feld: ' : 'field: '}{result.fieldLabel}</>}
                          {o && o.source && <>{result.fieldLabel ? ' · ' : ''}{isDe ? 'aus: ' : 'from: '}{o.source}</>}
                          {o && o.issued
                            ? <>{(o.source || result.fieldLabel) ? ' · ' : ''}{isDe ? `Trikot ${o.issued} ausgegeben` : `shirt ${o.issued} handed out`}</>
                            : assumed
                              ? <span style={{ fontStyle: 'italic' }}>{((o && o.source) || result.fieldLabel) ? ' · ' : ''}{isDe ? `abgeholt (angenommen — eingecheckt, ${av && av.issued ? av.issued : ''} gewünscht)` : `collected (assumed — checked in, wished ${av && av.issued ? av.issued : ''})`}</span>
                              : null}
                          {(o ? o.others : []).map((x, k) => (
                            <div key={k}>
                              {isDe ? 'weitere Zeile: ' : 'other row: '}{x.source || (isDe ? 'unbekannt' : 'unknown')} · {x.value || (isDe ? 'ohne Angabe' : 'no answer')}
                            </div>
                          ))}
                          {/* v31.4.3: Ein abweichender Wert in einem ANDEREN
                              Trikot-Feld. Genau das war von außen unsichtbar —
                              und genau das ist die Erklärung, wenn Tabelle und
                              Bestellliste sich widersprechen. */}
                          {(o ? o.otherFields : []).map((x, k) => (
                            <div key={`f${k}`} style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>
                              {isDe ? 'anderes Feld: ' : 'other field: '}<strong>{x.label}</strong>
                              {x.source ? <> ({x.source})</> : null} · {x.value}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            ) : names.join(' · ')}
          </td></tr>
        )}
      </React.Fragment>
    );
  };

  // v31.2: Genau ein Primär-Knopf — Speichern schreibt Daten, Excel ist die
  // Nebenaktion links im Fuß. Vorher standen zwei grüne Knöpfe im Inhalt.
  const footer = (
    <>
      {ready && (
        <span className="dex-ui-modal-foot-left">
          <button type="button" className="btn btn-outline dex-ui-btn-sm" disabled={xlsxBusy} onClick={() => { void downloadXlsx(); }}>
            <Download size={14} /> {xlsxBusy ? (isDe ? 'Wird erzeugt…' : 'Generating…') : (isDe ? 'Als Excel laden' : 'Download as Excel')}
          </button>
        </span>
      )}
      <button type="button" className="btn btn-secondary" onClick={props.onClose}>{isDe ? 'Schließen' : 'Close'}</button>
      {ready && (
        <button type="button" className="btn btn-primary" disabled={!stockDirty || stockSaving} onClick={() => { void saveStock(); }}>
          {stockSaving ? (isDe ? 'Speichert…' : 'Saving…') : (stockDirty ? (isDe ? 'Bestand speichern' : 'Save stock') : (isDe ? 'Bestand gespeichert' : 'Stock saved'))}
        </button>
      )}
    </>
  );

  return (
    <Modal open onClose={props.onClose} maxWidth={760} icon={<Shirt size={20} />} footer={footer}
      ariaLabel={isDe ? 'Benötigte T-Shirts' : 'T-shirts needed'} title={isDe ? 'Benötigte T-Shirts' : 'T-shirts needed'}
      subtitle={<>{props.event.title}{result && result.fieldLabel && <> · {isDe ? 'gezählt über das Feld' : 'counted via the field'} <strong>{result.fieldLabel}</strong></>}</>}>
      {loading && <p className="dex-ui-muted" style={{ margin: 0 }}>{isDe ? 'Trikotgrößen werden gelesen…' : 'Reading shirt sizes…'}</p>}

      {!loading && result && !result.fieldLabel && (
        <div className="dex-ui-callout dex-ui-callout--warn">
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>
            {isDe
              ? <>Dieses Event hat kein Abfragefeld, das nach einer Trikot- oder Konfektionsgröße aussieht.
                Lege im Assistenten unter <strong>Felder</strong> ein Feld an, dessen Bezeichnung die Größe
                benennt (z.B. &bdquo;T-Shirt Größe&ldquo; oder &bdquo;Trikotgröße&ldquo;) — danach zählt diese Ansicht
                automatisch mit, und die Größe steht auch am Check-in-Tisch.</>
              : <>This event has no form field that looks like a shirt or clothing size.
                In the wizard, under <strong>Fields</strong>, add a field whose label names the size
                (e.g. &ldquo;T-shirt size&rdquo; or &ldquo;Jersey size&rdquo;) — this view then counts automatically,
                and the size also shows at the check-in desk.</>}
          </span>
        </div>
      )}

      {!loading && result && result.fieldLabel && (
        <>
          {skipped.length > 0 && (
            <div className="dex-ui-callout dex-ui-callout--danger">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span>
                {isDe
                  ? <>Für {skipped.length === 1 ? 'diesen Termin' : 'diese Termine'} konnte die Teilnehmerliste nicht gelesen werden —
                    die Zahlen unten sind deshalb unvollständig: <strong>{skipped.join(', ')}</strong>.</>
                  : <>The attendee list of {skipped.length === 1 ? 'this date' : 'these dates'} could not be read —
                    the numbers below are therefore incomplete: <strong>{skipped.join(', ')}</strong>.</>}
              </span>
            </div>
          )}

          {/* v31.4.3: Mehr als ein Feld, das nach einer Größe aussieht — dann
              muss dastehen, welches zählt, welche es sonst gibt und was das für
              eine Korrektur bedeutet. Vorher entschied `shirtFieldOf` still,
              und eine im Organizer Center geänderte Größe konnte im anderen
              Feld landen als dem, das diese Liste zählt: Die Tabelle zeigte den
              neuen Wert, die Bestellliste den alten, und nichts sagte, warum. */}
          {shirtCandidates.length > 1 && (
            <div className="dex-ui-callout dex-ui-callout--warn">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span>
                {isDe
                  ? <><strong>Dieses Event hat {shirtCandidates.length} Felder, die nach einer Größe aussehen.</strong>{' '}
                    Gezählt wird gerade <strong>{result.fieldLabel}</strong>. Deine Korrektur landet immer in dem Feld, das
                    im Bearbeiten-Dialog der Zeile steht — zählt hier ein anderes Feld, siehst du sie in dieser Liste nicht.
                    Wähle das Feld, das gelten soll:</>
                  : <><strong>This event has {shirtCandidates.length} fields that look like a size.</strong>{' '}
                    Right now <strong>{result.fieldLabel}</strong> is counted. Your correction always lands in the field shown
                    in the edit dialog of that row — if a different field is counted here, you will not see it in this list.
                    Pick the field that should count:</>}
                <span className="dex-ui-inline" style={{ marginTop: 8 }}>
                  {shirtCandidates.map(c => {
                    const active = !!field && field.id === c.id;
                    const auto = !!autoField && autoField.id === c.id;
                    return (
                      <button
                        key={c.id}
                        type="button"
                        className={cx('dex-ui-chip', active && 'is-active')}
                        aria-pressed={active}
                        onClick={() => { setChosenFieldId(c.id); setOpenSize(null); }}
                        title={isDe
                          ? `${c.label} — aus: ${c.sources.join(', ') || 'unbekannt'} · ${c.sizeAnswers} von ${c.answers} Antworten sehen aus wie eine Größe`
                          : `${c.label} — from: ${c.sources.join(', ') || 'unknown'} · ${c.sizeAnswers} of ${c.answers} answers look like a size`}>
                        {c.label}
                        <span style={{ opacity: 0.75, fontWeight: 400 }}>
                          {' · '}{c.sizeAnswers}{isDe ? ' Größen' : ' sizes'}
                          {auto ? (isDe ? ' · Vorauswahl' : ' · default') : ''}
                        </span>
                      </button>
                    );
                  })}
                </span>
                <span style={{ display: 'block', marginTop: 6 }}>
                  {shirtCandidates.map(c => (
                    <span key={c.id} style={{ display: 'block' }} className="dex-ui-muted">
                      <strong>{c.label}</strong>{' — '}
                      {isDe ? 'aus: ' : 'from: '}{c.sources.join(', ') || (isDe ? 'unbekannt' : 'unknown')}
                      {' · '}{isDe
                        ? `${c.answers} ${c.answers === 1 ? 'Antwort' : 'Antworten'}, davon ${c.sizeAnswers} wie eine Größe`
                        : `${c.answers} ${c.answers === 1 ? 'answer' : 'answers'}, ${c.sizeAnswers} of them size-shaped`}
                    </span>
                  ))}
                </span>
              </span>
            </div>
          )}

          {/* v31.3: Widersprüchliche Zeilen benennen, statt still zu entscheiden. */}
          {conflicts.length > 0 && (
            <div className="dex-ui-callout dex-ui-callout--warn">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span>
                {isDe
                  ? <><strong>{conflicts.length === 1 ? 'Eine Person hat' : `${conflicts.length} Personen haben`} auf zwei Anmeldezeilen zwei verschiedene Größen.</strong>{' '}
                    Gerechnet wird mit der zuletzt geänderten Zeile (jeweils zuerst genannt) — bitte kurz nachfragen:{' '}
                    {conflicts.map(c => `${c.name}: ${c.values.join(' / ')}`).join(' · ')}</>
                  : <><strong>{conflicts.length === 1 ? 'One person has' : `${conflicts.length} people have`} two different sizes on two registration rows.</strong>{' '}
                    The most recently changed row is used (named first) — please double-check:{' '}
                    {conflicts.map(c => `${c.name}: ${c.values.join(' / ')}`).join(' · ')}</>}
              </span>
            </div>
          )}

          {/* v31.2: Kennzahlen zuerst — die Antwort auf „reicht es?" steht oben, bevor die Tabelle ins Detail geht.
              v31.3: „Shirts zu bestellen" statt „Größen gewünscht" — die Abwahl-Antworten sind keine Bestellung,
              und bei einem gesperrten Termin steht vor jeder Zahl „mind.". */}
          {/* v31.4: `dex-ui-kpi-row` statt eines starren Dreier-Rasters — sobald
              Trikots ausgegeben sind, kommt eine vierte Kachel dazu. */}
          <div className="dex-ui-kpi-row">
            <div className="dex-ui-kpi"><div className="dex-ui-kpi-value">{atLeast(result.total)}</div>
              <div className="dex-ui-kpi-label">{isDe ? (result.total === 1 ? 'Person angemeldet' : 'Personen angemeldet') : (result.total === 1 ? 'person registered' : 'people registered')}</div></div>
            <div className="dex-ui-kpi"><div className="dex-ui-kpi-value">{atLeast(result.sizeTotal)}</div>
              <div className="dex-ui-kpi-label">{isDe ? 'Shirts zu bestellen' : 'shirts to order'}</div>
              {optOutTotal > 0 && <div className="dex-ui-kpi-sub">{isDe ? `${optOutTotal} brauchen keins` : `${optOutTotal} need none`}</div>}</div>
            {hasStock
              ? <div className={cx('dex-ui-kpi', totalMissing > 0 ? 'dex-ui-kpi--orange' : 'dex-ui-kpi--green')}><div className="dex-ui-kpi-value">{atLeast(totalMissing)}</div>
                <div className="dex-ui-kpi-label">{isDe ? 'ohne Wunschgröße' : 'wished size short'}</div>
                {result.missing > 0 && <div className="dex-ui-kpi-sub">{isDe ? `+ ${result.missing} ohne Angabe` : `+ ${result.missing} without a size`}</div>}</div>
              : <div className={cx('dex-ui-kpi', result.missing > 0 && 'dex-ui-kpi--orange')}><div className="dex-ui-kpi-value">{atLeast(result.missing)}</div>
                <div className="dex-ui-kpi-label">{isDe ? 'ohne Größenangabe' : 'without a size'}</div></div>}
            {/* v31.4: Die Abend-Frage: Wie viele sind raus, was liegt noch da? */}
            {totalIssued > 0 && (
              <div className="dex-ui-kpi dex-ui-kpi--blue"><div className="dex-ui-kpi-value">{atLeast(totalIssued)}</div>
                <div className="dex-ui-kpi-label">{isDe ? 'Trikots ausgegeben' : 'shirts handed out'}</div>
                {/* v31.4 (Nachtrag): Die Kachel zählt die Annahmen mit — und sagt es. */}
                {totalIssuedAssumed > 0 && <div className="dex-ui-kpi-sub">{isDe ? `davon ${totalIssuedAssumed} angenommen` : `${totalIssuedAssumed} of them assumed`}</div>}
                {/* v31.4 (Review): „noch X im Karton" ist bei einem gesperrten
                    Termin keine Untergrenze, sondern eine OBERGRENZE — die dort
                    fehlenden Ausgaben lassen den rechnerischen Inhalt STEIGEN.
                    Deshalb hier „höchstens", nicht „mind." (CLAUDE.md: ein
                    Lesefehler ist keine Null). */}
                {hasStock && <div className="dex-ui-kpi-sub">{partial
                  ? (isDe ? `höchstens ${totalInBox} im Karton` : `at most ${totalInBox} left in the box`)
                  : (isDe ? `noch ${totalInBox} im Karton` : `${totalInBox} left in the box`)}</div>}</div>
            )}
          </div>

          <div className="dex-ui-section">
            <div className="dex-ui-section-title">
              {isDe ? 'Bedarf und Bestand je Größe' : 'Need and stock per size'}
              {stockDirty && <span className="dex-ui-pill dex-ui-pill--orange" style={pillInTitle}>{isDe ? 'Bestand noch nicht gespeichert' : 'stock not saved yet'}</span>}
            </div>
            <p className="dex-ui-section-desc">
              {isDe
                ? <>Trag unter <strong>Bestand</strong> ein, wie viele Shirts du je Größe wirklich hast. Die App prüft, ob es reicht, und macht je Person einen
                  Gegenvorschlag — den sieht das Check-in-Team bei der Abholung direkt an der Person.</>
                : <>Under <strong>Stock</strong>, enter how many shirts you really have per size. The app checks whether that is enough and proposes an
                  alternative per person — the check-in team sees it right at the person when handing out.</>}
            </p>
            <div className="dex-ui-table-wrap">
              <table className="dex-ui-table">
                <thead>
                  <tr>
                    <th>{isDe ? 'Größe' : 'Size'}</th>
                    <th style={{ width: '36%' }}>{isDe ? 'Benötigt' : 'Needed'}</th>
                    <th>{isDe ? 'Bestand' : 'Stock'}</th>
                    <th>{isDe ? 'Ausgegeben' : 'Handed out'}</th>
                    <th>{isDe ? 'Reicht es?' : 'Enough?'}</th>
                    <th><span className="dex-ui-sr-only">{isDe ? 'Personen' : 'People'}</span></th>
                  </tr>
                </thead>
                <tbody>
                  {stockKeys.map(s => { const t = tallyByKey[s.key]; return renderSizeRow(s.key, s.label, t ? t.count : 0, t ? t.names : [], 'size', t ? t.people : []); })}
                  {/* v31.3: Abwahl-Antworten stehen mit ihrem echten Wortlaut in der
                      Liste — der Organizer sieht, wie viele keins brauchen, ohne dass
                      daraus eine Bestellposition wird. */}
                  {optOutRows.map(r => renderSizeRow(`__opt__${shirtSizeKey(r.size)}`, r.size, r.count, r.names, 'optout', r.people))}
                  {noneRow && renderSizeRow('__none__', isDe ? 'ohne Angabe' : 'no answer', noneRow.count, noneRow.names, 'none', noneRow.people)}
                </tbody>
              </table>
            </div>
            {/* v31.4: Sobald ausgegeben wurde, sind „Bestand" und „verfügbar"
                zwei verschiedene Zahlen — das muss dastehen, sonst liest der
                Organizer den eingetragenen Bestand als Kartoninhalt. */}
            {totalIssued > 0 && (
              <div className="dex-ui-callout dex-ui-callout--info" style={{ marginTop: 12 }}>
                <span className="dex-ui-callout-icon"><Shirt size={16} /></span>
                <span>
                  {isDe
                    ? <><strong>{atLeast(totalIssued)} {totalIssued === 1 ? 'Trikot ist' : 'Trikots sind'} ausgegeben</strong>{hasStock ? <> — nach deinem Bestand liegen {partial ? 'höchstens' : 'noch'} <strong>{totalInBox}</strong> im Karton{partial ? <> (für {skipped.length === 1 ? 'einen Termin' : 'einige Termine'} fehlen die Ausgaben — die echte Zahl ist kleiner)</> : null}.</> : <>. Trag oben einen Bestand ein, dann rechnet die App dir aus, was noch da ist.</>}{' '}
                      {/* v31.4 (Review): Der Satz gilt nur für die FESTGEHALTENEN
                          Ausgaben. Der angenommene Teil hängt am Status
                          'Eingecheckt' und fällt bei Abmeldung/No-Show wieder in
                          den Bestand zurück — genau das Gegenteil dessen, was
                          hier stand. (Auf `CheckedInDate` umzustellen geht
                          nicht: `markNoShowParticipant` schreibt dieselbe
                          Spalte, ein No-Show sähe dann aus wie ein Check-in.) */}
                      {totalIssued > totalIssuedAssumed && <>Festgehaltene Ausgaben bleiben abgezogen, auch wenn die Person später abgemeldet oder als No-Show markiert wird — sie hat es ja mitgenommen.{' '}</>}
                      {/* v31.4 (Nachtrag): Die Annahme muss dastehen, sonst hält der Organizer eine Schätzung für ein Protokoll. */}
                      {totalIssuedAssumed > 0 && <> <strong>{totalIssuedAssumed} davon {totalIssuedAssumed === 1 ? 'ist angenommen' : 'sind angenommen'}:</strong>{' '}
                        Wer eingecheckt ist, aber keinen Ausgabe-Eintrag hat, zählt mit seiner Wunschgröße als abgeholt — den Ausgabe-Knopf am
                        Check-in-Tisch gibt es erst seit v31.4. Sobald jemand dort eine Größe festhält, gilt diese. Auf eine ANDERE Größe ändern
                        kannst du das je Person in der Teilnehmerliste über &bdquo;Bearbeiten&ldquo;. Angenommene Ausgaben verschwinden allerdings wieder, sobald die Person
                        abgemeldet oder als No-Show markiert wird — halt sie am Tisch fest, wenn die Zahl den Abend überstehen soll.</>}</>
                    : <><strong>{atLeast(totalIssued)} {totalIssued === 1 ? 'shirt has' : 'shirts have'} been handed out</strong>{hasStock ? <> — going by your stock, {partial ? 'at most' : ''} <strong>{totalInBox}</strong> are still in the box{partial ? <> (handouts are missing for {skipped.length === 1 ? 'one date' : 'some dates'} — the real number is lower)</> : null}.</> : <>. Enter a stock above and the app works out what is left.</>}{' '}
                      {totalIssued > totalIssuedAssumed && <>Recorded handouts stay deducted even if the person is cancelled or marked as a no-show later — they took it with them.{' '}</>}
                      {totalIssuedAssumed > 0 && <> <strong>{totalIssuedAssumed} of {totalIssuedAssumed === 1 ? 'them is assumed' : 'them are assumed'}:</strong>{' '}
                        anyone who is checked in but has no handout entry counts as served with their wished size — the handout button at the
                        check-in desk only exists since v31.4. As soon as someone records a size there, that one counts. You can change it per
                        person in the attendee list via &ldquo;Edit&rdquo;. Assumed handouts do disappear again once the person is cancelled or
                        marked as a no-show — record them at the desk if the number should survive the evening.</>}</>}
                </span>
              </div>
            )}

            <div className="dex-ui-inline" style={{ marginTop: 10 }}>
              <input type="text" className="dex-ui-input dex-ui-input--sm" value={newSize} onChange={e => setNewSize(e.target.value)} style={{ maxWidth: 200 }}
                placeholder={isDe ? 'Weitere Größe, z.B. XXL' : 'Another size, e.g. XXL'} aria-label={isDe ? 'Weitere Größe' : 'Another size'} />
              <button
                type="button"
                className="dex-ui-textbtn"
                // v31.3: Nur echte Größen — sonst legt man eine Bestandszeile für
                // etwas an, das niemand tragen kann.
                disabled={!splitShirtSize(newSize).isSize}
                onClick={() => {
                  if (!splitShirtSize(newSize).isSize) return;
                  const k = shirtSizeKey(newSize);
                  if (!k) return;
                  setStockInput(prev => (prev[k] !== undefined ? prev : { ...prev, [k]: '' }));
                  setNewSize('');
                }}
              >
                <Plus size={14} /> {isDe ? 'Größe ergänzen' : 'Add size'}
              </button>
              <span className="dex-ui-muted">
                {isDe
                  ? 'Erscheint als neue Zeile — für Reserve-Größen, die niemand gewünscht hat. Muss auf ein Größenkürzel enden (S, M, L, XL …), gern mit Vorsatz („Herrengröße XL").'
                  : 'Adds a row — for spare sizes nobody asked for. Has to end in a size (S, M, L, XL …), a prefix is fine (“Men XL”).'}
              </span>
            </div>

            {result.missing > 0 && (
              <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 12 }}>
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>
                  {isDe
                    ? <><strong>{result.missing} {result.missing === 1 ? 'Person hat' : 'Personen haben'} keine Größe angegeben.</strong>{' '}
                      Die Namen stehen in der Zeile &bdquo;ohne Angabe&ldquo; — frag dort nach, bevor du bestellst.
                      Sonst fehlt am Lauftag genau diese Anzahl Trikots.</>
                    : <><strong>{result.missing} {result.missing === 1 ? 'person has' : 'people have'} not given a size.</strong>{' '}
                      Their names are in the &ldquo;no answer&rdquo; row — ask them before you order.
                      Otherwise exactly that many shirts will be missing on race day.</>}
                </span>
              </div>
            )}
          </div>

          {/* v30.88: Gegenvorschläge je Person — erst, wenn ein Bestand eingetragen ist. */}
          {alloc && alloc.hasStock && (
            <div className="dex-ui-section">
              <div className="dex-ui-section-title">
                {isDe ? 'Wer bekommt eine andere Größe?' : 'Who gets a different size?'}
                {proposals.length > 0 && <span className="dex-ui-pill dex-ui-pill--gray" style={pillInTitle}>{proposals.length}</span>}
              </div>
              {proposals.length === 0 && alloc.noneLeft.length === 0 ? (
                /* v31.3: „Reicht für alle" nur, wenn es auch stimmt. Der Kasten hing
                   bis dahin allein an den Gegenvorschlägen — Personen ohne Angabe
                   verbrauchen nichts und tauchen dort nie auf, ein gesperrter Termin
                   ebenso wenig. Beides macht die Aussage falsch, nicht ungenau. */
                (partial || result.missing > 0) ? (
                  <div className="dex-ui-callout dex-ui-callout--warn">
                    <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                    <span>
                      {isDe
                        ? <><strong>Für die abgegebenen Wünsche reicht der Bestand.</strong>{' '}
                          {result.missing > 0 && <>Ob er für ALLE reicht, lässt sich nicht sagen: {result.missing} {result.missing === 1 ? 'Person hat' : 'Personen haben'} keine Größe angegeben —
                            frag über die Zeile &bdquo;ohne Angabe&ldquo; nach und trag die Antwort ein. </>}
                          {partial && <>Außerdem {skipped.length === 1 ? 'ist ein Termin' : 'sind Termine'} nicht lesbar ({skipped.join(', ')}) — die dort Angemeldeten fehlen in dieser Rechnung.</>}</>
                        : <><strong>The stock covers every size that was given.</strong>{' '}
                          {result.missing > 0 && <>Whether it covers EVERYONE cannot be said: {result.missing} {result.missing === 1 ? 'person has' : 'people have'} not given a size —
                            ask them via the &ldquo;no answer&rdquo; row and record the answer. </>}
                          {partial && <>On top of that, {skipped.length === 1 ? 'one date is' : 'some dates are'} unreadable ({skipped.join(', ')}) — those registrations are missing from this calculation.</>}</>}
                    </span>
                  </div>
                ) : (
                  <div className="dex-ui-callout dex-ui-callout--success">
                    <span className="dex-ui-callout-icon"><Check size={16} /></span>
                    <span>
                      {isDe
                        ? <><strong>Der Bestand reicht für alle Wünsche.</strong> Jede Person bekommt ihre Größe.</>
                        : <><strong>The stock covers every wish.</strong> Everyone gets their size.</>}
                    </span>
                  </div>
                )
              ) : (
                <>
                  <p className="dex-ui-section-desc">
                    {isDe
                      ? 'Reihenfolge nach Teilnehmer-ID: Wer zuerst angemeldet war, bekommt seine Wunschgröße. Ausweichgröße = nächste mit Rest, zuerst eine Nummer größer.'
                      : 'Order by attendee ID: whoever registered first gets their wished size. Alternative = next size with stock left, one size up first.'}
                  </p>
                  <div className="dex-ui-table-wrap" style={{ maxHeight: 260, overflowY: 'auto' }}>
                    <table className="dex-ui-table">
                      <thead>
                        <tr>
                          <th style={stickyTh}>ID</th>
                          <th style={stickyTh}>{isDe ? 'Person' : 'Person'}</th>
                          <th style={stickyTh}>{isDe ? 'Wunsch' : 'Wish'}</th>
                          <th style={stickyTh}>{isDe ? 'Vorschlag' : 'Proposal'}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {proposals.map((p, i) => (
                          <tr key={i}>
                            <td style={{ color: 'var(--dex-gray-400)' }}>{p.tid ?? '—'}</td>
                            <td style={{ fontWeight: 600 }}>{p.name}</td>
                            <td><span style={{ textDecoration: 'line-through', color: 'var(--dex-gray-500)' }}>{p.wish}</span></td>
                            <td>
                              {p.proposal
                                ? <span className="dex-ui-pill dex-ui-pill--green">{p.proposal}</span>
                                : <span className="dex-ui-pill dex-ui-pill--red">{isDe ? 'keine Größe mehr vorrätig' : 'no size left in stock'}</span>}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          )}
        </>
      )}
    </Modal>
  );
}
