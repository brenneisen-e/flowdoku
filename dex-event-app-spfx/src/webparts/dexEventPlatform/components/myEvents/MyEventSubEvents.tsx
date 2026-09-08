/* MyEventSubEvents — aus MyEventsPage.tsx ausgelagert (Zeilen 3868-4985 des
 * urspruenglichen Stands, v30.65). Enthaelt die Sub-Event-Ansicht der Karte
 * „Meine Events" samt Kalender-/Listen-Darstellung und das nur hier benutzte
 * PeerCancelCheckboxModal. Der Code ist zeichengleich uebernommen.
 */
import * as React from 'react';
import { DeloitteEvent, EventSpecificField } from '../../types';
import { SPRegistration } from '../../services/EventService';
import { useLanguage } from '../../context/LanguageContext';
import { useCurrentUser } from '../../context/UserContext';
import { useRoles } from '../../context/RoleContext';
import { useDialog } from '../../context/DialogContext';
import { useEvents } from '../../context/EventContext';
import { isoToLocal } from '../../utils/berlinTime';
import { isEventOver, subEventRegDeadline } from '../../utils/eventFormat';
import { selfCancelLocked, selfCancelLockReason } from '../../utils/cancelPolicy';
import { isEventVisibleForUser } from '../EventListPage';
import { InfoTooltip } from '../InfoTooltip';
// v31.8: Gemeinsame UI-Klassen statt handgebauter Inline-Kästen (docs/ui-leitfaden.md).
// `ensureDexUiStyles()` wird hier NICHT gerufen — der Leitfaden (Abschnitt 6)
// legt den Aufruf auf die Seiten-Komponente `MyEventsPage`, damit er genau
// einmal je Seite passiert; jedes geöffnete `Modal` ruft ihn zusätzlich selbst.
import { cx } from '../dexUi';
import { AlertCircle, ChevronDown, FileText, Info, Pencil } from '../Icons';
import Modal from '../Modal';
import StayRangePicker from '../StayRangePickerLazy';
import { FieldAnswerTag, formatDateTimeRange } from './myEventsHelpers';
import { groupSubEventTabs, stripGroupPrefix } from '../../utils/subEventGroups';
import { SubmitOverlay } from '../registration/RegistrationBanners';

// ==================== Sub-Events im "My Events"-Tab ====================
// Seit v6.4: Sub-Events sind eigene DEX_Events-Items (childEventsOf(parentId)).
// Anmeldung/Abmeldung läuft über den normalen registerForEvent/cancelRegistration-
// Pfad — identisch zu einem Top-Level-Event. Eigene Teilnehmerliste pro Sub-Event.
// v10.27: Plus Anzeige + Bearbeitung der Sub-Event-spezifischen Antworten
// (eventSpecificData) — analog zum Edit-Modus auf der Hauptevent-Karte.


export default function MyEventSubEvents(props: {
  parentEvent: DeloitteEvent;
  childEvents: DeloitteEvent[];
  registerForEvent: (eventId: string, customData: Record<string, string>, participantFirstName?: string, participantLastName?: string, participantEmail?: string, preferredStarterType?: string, opts?: { suppressMail?: boolean; suppressOutlook?: boolean; skipReload?: boolean }) => Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  cancelRegistration: (eventId: string, opts?: { suppressNotifications?: boolean }) => Promise<boolean>;
  getMyRegistration: (eventId: string) => Promise<SPRegistration | null>;
  /** v30.67: mit `onHttpError` — ohne Rückruf ist ein Leseverbot von einer leeren Liste nicht zu unterscheiden. */
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
  updateMyRegistration: (eventId: string, customData: Record<string, string>) => Promise<boolean>;
  onMutated: () => Promise<void>;
}): React.ReactElement | null {
  // v15.13: isDe MUSS aus dem UI-Locale kommen, nicht aus event.emailLanguage —
  // letzteres ist die Sprache der Bestätigungsmails (organizer-konfiguriert)
  // und sagt nichts darüber aus, in welcher Sprache der User die App sieht.
  const { locale: __uiLocale } = useLanguage();
  const isDe = __uiLocale === 'de';
  // v22.13: Sub-Events auch in „Meine Events" nach ihrer EIGENEN Sichtbarkeit
  // filtern (gleiche Logik wie Anmeldeseite/Event-Liste, Fix v22.10). Eigene
  // aktive Anmeldungen bleiben IMMER sichtbar (verwalten/abmelden); Organizer
  // des Events und Admins sehen weiterhin alles.
  const { currentUser, groupEmails } = useCurrentUser();
  // v30.4: previewAsUser = „Übersicht als User sehen" — isAdmin kommt aus dem
  // RoleContext bereits abgesenkt, der per-Event-Organizer-Check unten zieht
  // lokal mit; An-/Abmelden ist in der Vorschau gesperrt (handleToggle).
  const { isAdmin, previewAsUser } = useRoles();
  const { showAlert, confirmDialog } = useDialog();
  // v30.67: Belegung aus dem für alle lesbaren Platzzähler — s. refresh().
  const { getLiveCounterStats } = useEvents();
  const [busyId, setBusyId] = React.useState<string | null>(null);
  // v15.21: Voll-Bild-Progress-Modal beim (Peer-)Cancel + Anmeldung. Vorher
  // war nur die einzelne Karte mit „…" markiert — bei Peer-Cancels haben die
  // peers visuell nicht reagiert, weil busyId nur die EINE id hält.
  const [processingMessage, setProcessingMessage] = React.useState<string>('');
  // v30.79: Prozentwert für das SubmitOverlay (0–100), s. handleToggle.
  const [processingProgress, setProcessingProgress] = React.useState<number>(0);
  // v30.19: Während der Verarbeitung warnt der Browser vor dem Schließen des
  // Fensters/Tabs (nativer „Website verlassen?"-Dialog) — ein Abbruch mitten
  // im (Peer-)Cancel oder Anmelden hinterlässt halbe Zustände. Gleiche
  // Mechanik wie das Submit-Overlay der RegistrationPage.
  React.useEffect(() => {
    if (!processingMessage) return undefined;
    const warnBeforeUnload = (e: BeforeUnloadEvent): void => { e.preventDefault(); e.returnValue = ''; };
    window.addEventListener('beforeunload', warnBeforeUnload);
    return () => window.removeEventListener('beforeunload', warnBeforeUnload);
  }, [processingMessage]);
  const [registeredSet, setRegisteredSet] = React.useState<Set<string>>(new Set());
  // v30.67: Der echte Status je Termin — `registeredSet` warf 'Warteliste' mit
  // 'Angemeldet' in einen Topf, und ein Wartelisten-Platz stand grün als
  // „(Angemeldet)" da, mit „Abmelden"-Knopf.
  const [statusById, setStatusById] = React.useState<Record<string, string>>({});
  // v30.67: null = Belegung nicht ermittelbar (weder Platzzähler noch lesbare
  // Liste) — dann wird keine Zahl gezeigt und nichts als „frei" behauptet.
  const [counts, setCounts] = React.useState<Record<string, number | null>>({});
  // v10.27: pro Sub-Event die geparsten Custom-Field-Antworten aus
  // myReg.CustomData. Wird nur für aktive Registrierungen befüllt.
  const [seData, setSeData] = React.useState<Record<string, Record<string, string>>>({});
  // v10.27: aktuell ge-editiertes Sub-Event + Draft der Werte. Beim Save
  // gehen die Werte über updateMyRegistration in die Subsite zurück.
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [editDraft, setEditDraft] = React.useState<Record<string, string>>({});
  const [savingEdit, setSavingEdit] = React.useState(false);
  // v24.30 BUG-FIX: Beim Anmelden für ein Sub-Event aus „Meine Events" wurden
  // die (Pflicht-)Felder nie abgefragt (registerForEvent lief mit leerem
  // customData). Jetzt öffnet sich — wie auf der Anmeldeseite — erst ein
  // Feld-Modal, das die Antworten sammelt. resolve liefert die Antworten oder
  // 'abort'.
  const [regFieldsModal, setRegFieldsModal] = React.useState<{
    ce: DeloitteEvent; resolve: (_v: Record<string, string> | 'abort') => void;
  } | null>(null);
  const [regDraft, setRegDraft] = React.useState<Record<string, string>>({});
  const [regShowErrors, setRegShowErrors] = React.useState(false);
  // v11.34: Cascade-Cancel-Dialog für Sub-Event-Cancel — fragt ob auch
  // die anderen aktiven Sub-Events des gleichen Parents abgemeldet
  // werden sollen (Peer-Cancel).
  // v15.8: Peer-Cancel-Modal mit Checkbox-Liste statt drei-Buttons —
  // User hakt direkt an, welche zusätzlichen Sub-Events er mitabmelden
  // will (Target ist immer fix abgemeldet, weil er ja explizit den
  // Cancel-Button geklickt hat). resolve liefert das Set der peer-IDs
  // zurück oder 'abort' wenn der Modal-User komplett abbricht.
  const [peerCancelDialog, setPeerCancelDialog] = React.useState<{
    targetTitle: string;
    peers: { id: string; title: string; startDate?: string; location?: string }[];
    resolve: (_choice: { peerIds: string[] } | 'abort') => void;
  } | null>(null);
  // v30.9: Hover-State für die Kalender-Zellen (Inline-Styles können kein :hover).
  const [dayHoverKey, setDayHoverKey] = React.useState<string>('');
  // v30.79: Termin-Gruppen („Day 1 - …") wie auf der Anmeldeseite (v30.76) —
  // alle offen, ein Klick auf den Kopf klappt zu. Nutzer-Ansage 07.09.2026:
  // „hier auch noch strukturiert anzeigen".
  const [collapsedGroups, setCollapsedGroups] = React.useState<Set<string>>(() => new Set<string>());

  const refresh = React.useCallback(async (): Promise<void> => {
    try {
      const regs = await Promise.all(props.childEvents.map(async (ce) => {
        const reg = await props.getMyRegistration(ce.id);
        const isActive = !!reg && reg.Status !== 'Abgemeldet';
        // v10.27: Custom-Field-Werte aus reg.CustomData (JSON) parsen.
        let data: Record<string, string> = {};
        if (isActive && reg && reg.CustomData) {
          try { data = JSON.parse(reg.CustomData) as Record<string, string>; } catch { /* fall back to empty */ }
        }
        return { id: ce.id, isActive, status: (reg && reg.Status) || '', data };
      }));
      const s = new Set<string>();
      const statusMap: Record<string, string> = {};
      const dataMap: Record<string, Record<string, string>> = {};
      for (const r of regs) {
        if (r.isActive) {
          s.add(r.id);
          statusMap[r.id] = r.status;
          dataMap[r.id] = r.data;
        }
      }
      setRegisteredSet(s);
      setStatusById(statusMap);
      setSeData(dataMap);

      // v30.67: Belegung wie auf der Anmeldeseite (`occupancyOf`, v30.62):
      // zuerst der Platzzähler DEX_TeilnehmerCounter, den jeder lesen darf.
      // Die Teilnehmerliste ist zeilenweise gesichert (ReadSecurity=2) — ein
      // Teilnehmer zählte dort nur die eigene Zeile, sah „0/20" bei vollem
      // Termin, kein „(voll)", einen aktiven Anmelde-Knopf, und landete still
      // auf der Warteliste. Die Liste bleibt Rückfall, aber nur mit geprüftem
      // Status; trägt keine Quelle, bleibt die Zahl unbekannt (null).
      const occupancyOf = async (ce: DeloitteEvent): Promise<number | null> => {
        try {
          const stats = await getLiveCounterStats(ce.id);
          // `seatsKnown` trennt „null Anmeldungen" von „nie geschrieben".
          if (stats && stats.seatsKnown && typeof stats.active === 'number' && stats.active >= 0) return stats.active;
        } catch (e) { console.warn('[DEX] Platzzähler nicht lesbar, Rückfall auf die Liste:', ce.id, e); }
        let readable = true;
        try {
          const all = await props.getAllRegistrations(ce.id, () => { readable = false; });
          if (!readable) return null;
          return (all || []).filter(r => {
            const st = r.Status || '';
            return st === 'Angemeldet' || st === 'QR versendet' || st === 'Eingecheckt';
          }).length;
        } catch (e) {
          console.warn('[DEX] Belegung nicht ermittelbar:', ce.id, e);
          return null;
        }
      };
      const countPairs = await Promise.all(props.childEvents.map(async (ce) => ({ id: ce.id, count: await occupancyOf(ce) })));
      const map: Record<string, number | null> = {};
      for (const p of countPairs) map[p.id] = p.count;
      setCounts(map);
    } catch (e) {
      // v31.8: `counts` wird erst am ENDE des try gesetzt. Bricht vorher etwas
      // ab, blieb die Map für immer leer — und die Ansicht behauptete
      // dauerhaft „Belegung wird geladen…", obwohl nichts mehr lädt. Ein
      // Fehlschlag ist kein Ladezustand, sondern „unbekannt": jeder noch nicht
      // gelesene Termin bekommt ausdrücklich `null` (CLAUDE.md: ein Lesefehler
      // ist keine Null — und auch kein Wartezustand). Bereits gelesene Zahlen
      // bleiben stehen, ein späterer Fehler entwertet sie nicht.
      console.warn('[DEX] Sub-Event-Stand nicht vollständig ladbar:', e);
      setCounts(prev => {
        const next: Record<string, number | null> = { ...prev };
        for (const ce of props.childEvents) {
          if (!Object.prototype.hasOwnProperty.call(next, ce.id)) next[ce.id] = null;
        }
        return next;
      });
    }
  }, [props.childEvents.map(c => c.id).join(',')]);

  React.useEffect(() => { refresh().catch(() => { /* ignore */ }); }, [refresh]);

  // v15.13: Bezeichnung kommt jetzt direkt aus event.childEventTermPlural /
  // childEventTermSingular (Wizard-Setting). Fallback: „Sub-Events"-Begriff.
  // v29.13: Besteht das Event ausschließlich aus Sub-Events, heißen sie hier
  // wie auf der Anmeldeseite „Events" — es gibt kein Haupt-Event, unter dem
  // sie stehen könnten, und der Teilnehmer hat sich genau für diese Einträge
  // angemeldet. Ein eigener Begriff des Organizers geht weiterhin vor.
  // v31.8: Der Block steht jetzt VOR handleToggle und den Dialogen — die
  // Bezeichnung wird dort ebenfalls gebraucht, statt „Sub-Event" fest
  // zu verdrahten (UI-Leitfaden 6d).
  const subOnly = !!props.parentEvent.subEventsOnlyMode;
  const termPlural = props.parentEvent.childEventTermPlural
    || (subOnly ? (isDe ? 'Events' : 'events') : '');
  const termSingular = props.parentEvent.childEventTermSingular
    || (subOnly ? (isDe ? 'Event' : 'event') : '');
  const headerLabel = termPlural || (isDe ? 'Sub-Events' : 'Sub-events');
  // v29.13: „Eine Session", aber „Ein Event" — der Artikel war fest verdrahtet.
  // v31.8: Das Geschlecht steht jetzt an EINER Stelle in dieser Datei; Artikel
  // und Kasus aller Sätze unten leiten sich daraus ab, statt einzeln
  // verdrahtet zu werden. Ein im Assistenten gepflegtes Geschlecht schlägt die
  // Heuristik (gleiche Reihenfolge wie `childOneDe` seit v29.60); ohne Angabe
  // bleibt die alte Suffix-Liste, damit sich an bestehenden Events nichts
  // still verschiebt.
  const termGenderDe: 'm' | 'f' | 'n' = props.parentEvent.childEventTermGender
    || (/(session|veranstaltung|einheit|runde|reihe|tour|führung|schicht|woche|gruppe|stunde)$/i.test(termSingular) ? 'f' : 'n');
  // v31.8: Ersatz-Bezeichnung und Ersatz-Titel — vorher stand hier fünfmal
  // „Sub-Event" bzw. „Session ohne Titel" im Code. Der Rückfall ist derselbe
  // wie in `headerLabel` (der dokumentierte Default des Wizard-Feldes, s.
  // `childEventTermSingular` in types/index.ts) — sonst hieße dieselbe Sache
  // in der Überschrift „Sub-Events" und drei Zeilen darunter „Termine".
  const termSingularSafe = termSingular || (isDe ? 'Sub-Event' : 'Sub-event');
  const termPluralSafe = headerLabel;
  const untitled = isDe ? `${termSingularSafe} ohne Titel` : `Untitled ${termSingularSafe}`;
  // v31.8: „zu diesem Termin", aber „zu dieser Session" — beim Ersetzen von
  // „Sub-Event" ist der Kasus verlorengegangen (dieselbe Falle, die v29.13
  // schon einmal beim Artikel zugeschnappt hat). Die Form kommt aus derselben
  // Geschlechts-Quelle wie der Artikel, damit beide nicht auseinanderlaufen.
  const termThisDatDe = `${termGenderDe === 'f' ? 'dieser' : 'diesem'} ${termSingularSafe}`;

  const fmt = (iso: string): string => {
    if (!iso) return '';
    try {
      const d = new Date(iso);
      if (isNaN(d.getTime())) return iso;
      return d.toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    } catch { return iso; }
  };

  const handleToggle = async (childEventId: string, currentlyRegistered: boolean): Promise<void> => {
    // v30.4: User-Vorschau ist nur zum Ansehen — An- UND Abmelden gesperrt.
    if (previewAsUser) {
      await showAlert(isDe
        ? 'Vorschau-Modus: Du siehst die Ansicht gerade so, wie reguläre User sie sehen — An- und Abmelden ist hier deaktiviert. Beende die Vorschau über den blauen Balken oben.'
        : 'Preview mode: you are viewing this as regular users see it — registering and cancelling are disabled here. End the preview via the blue bar at the top.');
      return;
    }
    // v29.25: Backstop zur Button-Sperre — Selbst-Abmeldung nach der Frist
    // deaktiviert (Organizer-Option auf dem Parent).
    if (currentlyRegistered && selfCancelLocked(props.childEvents.find(ce => ce.id === childEventId), props.parentEvent)) return;
    // v30.20: Explizite Bestätigung VOR jeder Einzel-Abmeldung — im Kalender
    // meldete ein Klick auf einen grünen Tag bisher SOFORT ab (ohne Peers gab
    // es gar keine Nachfrage). Nutzer-Befund: „nicht intuitiv". Der Dialog
    // nennt den Termin und was passiert (Bestätigungs-Mail, Outlook-Rückzug).
    if (currentlyRegistered) {
      const cancelTarget = props.childEvents.find(ce => ce.id === childEventId);
      const cancelTitle = (cancelTarget && cancelTarget.title) || (isDe ? 'diesen Termin' : 'this date');
      const okCancel = await confirmDialog(
        isDe
          ? `„${cancelTitle}" wirklich abmelden?\n\nDu bekommst eine Abmeldebestätigung per Mail und der Outlook-Termin wird zurückgezogen.`
          : `Really cancel „${cancelTitle}"?\n\nYou will receive a cancellation confirmation by email and the Outlook invite will be withdrawn.`,
        { danger: true, confirmLabel: isDe ? 'Abmelden' : 'Cancel registration' }
      );
      if (!okCancel) return;
    }
    // v11.34: Beim Cancel eines Sub-Events fragen, ob die anderen aktiven
    // Sub-Events des gleichen Parents auch abgemeldet werden sollen
    // (Peer-Cancel-Cascade) — gleicher Pattern wie der Parent-Cancel.
    let peerIdsToCancel: string[] = [];
    if (currentlyRegistered) {
      const peers = props.childEvents
        .filter(ce => ce.id !== childEventId && registeredSet.has(ce.id))
        .map(ce => ({
          id: ce.id,
          title: ce.title || untitled,
          startDate: ce.startDate || '',
          endDate: ce.endDate || '',
          location: ce.location || '',
        }));
      if (peers.length > 0) {
        const target = props.childEvents.find(ce => ce.id === childEventId);
        // v31.8: Derselbe Ersatz-Titel wie bei den Peers eine Zeile höher.
        // „diesen ${termSingularSafe}" war doppelt falsch: Der Wert steht im
        // Dialog als Name der angehakten Zeile, also im NOMINATIV, und der
        // Akkusativ „diesen" passt ohnehin nur zu maskulinen Begriffen.
        const targetTitle = (target && target.title) || untitled;
        const choice = await new Promise<{ peerIds: string[] } | 'abort'>(resolve => {
          setPeerCancelDialog({ targetTitle, peers, resolve });
        });
        setPeerCancelDialog(null);
        if (choice === 'abort') return;
        peerIdsToCancel = choice.peerIds;
      }
    }
    // v24.30 BUG-FIX: Vor der Sub-Event-Anmeldung die (Pflicht-)Felder abfragen
    // — sonst gingen sie verloren (registerForEvent lief mit leerem customData).
    let registerCustomData: Record<string, string> = {};
    if (!currentlyRegistered) {
      const target = props.childEvents.find(ce => ce.id === childEventId);
      const fields = (target?.eventSpecificFields || []).filter(f => f && f.label);
      if (target && fields.length > 0) {
        const result = await new Promise<Record<string, string> | 'abort'>(resolve => {
          setRegDraft({});
          setRegShowErrors(false);
          setRegFieldsModal({ ce: target, resolve });
        });
        setRegFieldsModal(null);
        if (result === 'abort') return;
        registerCustomData = result;
      }
    }
    setBusyId(childEventId);
    // v15.21: Voll-Bild-Progress mit Beschreibung — bei Peer-Cancels
    // zählen wir die Fortschritte fortlaufend mit, damit der User sieht
    // dass auch die peers verarbeitet werden.
    // v30.79: echter Prozent-Balken (SubmitOverlay wie beim Anmelden) —
    // ein Schritt mehr für das Nachladen am Ende.
    const totalSteps = (currentlyRegistered ? 1 + peerIdsToCancel.length : 1);
    const progressTotal = totalSteps + 1;
    const stepPct = (done: number): number => Math.min(99, Math.round((done / progressTotal) * 100));
    const initialMsg = currentlyRegistered
      ? (isDe
          ? `Abmeldung wird verarbeitet… (1/${totalSteps})`
          : `Cancellation in progress… (1/${totalSteps})`)
      : (isDe ? 'Anmeldung wird verarbeitet…' : 'Registration in progress…');
    setProcessingProgress(stepPct(0));
    setProcessingMessage(initialMsg);
    try {
      if (currentlyRegistered) {
        await props.cancelRegistration(childEventId);
        let done = 1;
        setProcessingProgress(stepPct(done));
        // v30.79: Sechs Abmeldungen hintereinander sind rund vierzig
        // SharePoint-Schreibvorgänge am Stück (Status, Register, Mail-Queue,
        // Outlook-Queue, Reorder-Queue, Platz-Sync je Termin) — SharePoint
        // drosselte (Nutzer 07.09.2026: „hat zu throttle geführt"), und ein
        // gedrosselter Termin blieb still angemeldet. Jetzt: kurze Pause
        // zwischen den Terminen, und ein gescheiterter Termin wird nach
        // einer längeren Pause ein zweites Mal versucht; was dann noch
        // scheitert, wird gemeldet statt geschluckt.
        const failedPeers: string[] = [];
        for (const peerId of peerIdsToCancel) {
          done++;
          setProcessingMessage(isDe
            ? `Abmeldung wird verarbeitet… (${done}/${totalSteps})`
            : `Cancellation in progress… (${done}/${totalSteps})`);
          await new Promise<void>(resolve => setTimeout(resolve, 600));
          let ok = false;
          try { ok = await props.cancelRegistration(peerId); }
          catch (err) { console.warn('[DEX] peer-cancel failed:', peerId, err); }
          if (!ok) {
            setProcessingMessage(isDe
              ? `SharePoint ist ausgelastet — zweiter Versuch… (${done}/${totalSteps})`
              : `SharePoint is busy — retrying… (${done}/${totalSteps})`);
            await new Promise<void>(resolve => setTimeout(resolve, 2500));
            try { ok = await props.cancelRegistration(peerId); }
            catch (err) { console.warn('[DEX] peer-cancel retry failed:', peerId, err); }
          }
          if (!ok) failedPeers.push(peerId);
          setProcessingProgress(stepPct(done));
        }
        if (failedPeers.length > 0) {
          const names = failedPeers.map(id => (props.childEvents.find(c => c.id === id) || { title: '' }).title || id).join(', ');
          await showAlert(isDe
            ? `${failedPeers.length} Termin${failedPeers.length === 1 ? '' : 'e'} konnte${failedPeers.length === 1 ? '' : 'n'} nicht abgemeldet werden (SharePoint ausgelastet): ${names}. Die Anmeldung dort bleibt bestehen — bitte in ein paar Minuten erneut abmelden.`
            : `${failedPeers.length} date${failedPeers.length === 1 ? '' : 's'} could not be cancelled (SharePoint busy): ${names}. Those registrations remain — please cancel again in a few minutes.`, { variant: 'error' });
        }
        // v15.26: Im subEventsOnlyMode war das Hauptevent als Schatten-
        // Registrierung angelegt. Wenn die letzte aktive Sub-Event-
        // Registrierung jetzt entfernt wurde, soll auch die Schatten-
        // Parent-Reg wegfallen — sonst behauptet „Meine Events"
        // weiterhin „Registered" obwohl der User kein einziges Sub-
        // Event mehr hat.
        if (props.parentEvent.subEventsOnlyMode) {
          const remainingActive = props.childEvents
            .map(ce => ce.id)
            .filter(id => id !== childEventId && peerIdsToCancel.indexOf(id) < 0)
            .filter(id => registeredSet.has(id));
          if (remainingActive.length === 0) {
            setProcessingMessage(isDe ? 'Hauptevent-Eintrag wird entfernt…' : 'Removing main-event entry…');
            // v17.22: Schatten-Parent-Cancel → Notifications unterdrücken
            // (die Sub-Event-Abmeldung hat ihre eigene Mail schon geschickt).
            try { await props.cancelRegistration(props.parentEvent.id, { suppressNotifications: true }); }
            catch (err) { console.warn('[DEX] shadow-parent cancel failed:', err); }
          }
        }
      } else {
        const regRes = await props.registerForEvent(childEventId, registerCustomData);
        // v30.67: Das Ergebnis MELDEN. Vorher kam nichts zurück — weder bei
        // einem Fehlschlag (volle Gruppe ohne Warteliste, Doppel-Check) noch
        // beim Wartelisten-Platz; die Person hielt sich für angemeldet, plante
        // die Anreise und las in der Wartelisten-Mail etwas anderes als hier.
        if (!regRes.ok) {
          const why = regRes.reason === 'full'
            ? (isDe ? 'Der Termin ist voll und hat keine Warteliste.' : 'This date is full and has no waitlist.')
            : regRes.reason === 'already-registered'
            ? (isDe ? 'Du bist für diesen Termin bereits angemeldet.' : 'You are already registered for this date.')
            : regRes.reason === 'dup-check-failed'
            ? (isDe ? 'Die Anmeldung wurde nicht angelegt — SharePoint ist gerade ausgelastet. Bitte versuche es in ein paar Minuten erneut; es wurde nichts gespeichert.' : 'The registration was not created — SharePoint is busy right now. Please try again in a few minutes; nothing was saved.')
            : (isDe ? 'Die Anmeldung wurde nicht gespeichert — bitte erneut versuchen oder die Organizer ansprechen.' : 'The registration was not saved — please try again or contact the organizers.');
          await showAlert(why, { variant: 'error' });
        } else if (regRes.status === 'Warteliste') {
          await showAlert(isDe
            ? 'Der Termin ist bereits voll — du stehst jetzt auf der Warteliste und wirst benachrichtigt, sobald ein Platz frei wird.'
            : 'This date is already full — you are now on the waitlist and will be notified as soon as a seat becomes free.');
        }
        // v30.14 → v30.68: Die Klammer-Zeile stellt registerForEvent selbst
        // sicher — VOR dem Termin, mit zweitem Versuch danach und Nachzug-
        // Merker (utils/shadowHeal). Der eigene Nachzug hier ist entfallen.
      }
      setProcessingMessage(isDe ? 'Aktualisiere…' : 'Refreshing…');
      setProcessingProgress(stepPct(totalSteps));
      await refresh();
      await props.onMutated();
      setProcessingProgress(100);
    } finally {
      setBusyId(null);
      setProcessingMessage('');
      setProcessingProgress(0);
    }
  };

  // v22.13: nur Sub-Events zeigen, die der User laut Sub-Event-Sichtbarkeit
  // sehen darf — ODER in denen er bereits aktiv angemeldet ist.
  const myEmailLc = (currentUser.email || '').toLowerCase();
  // v30.4: In der User-Vorschau zählt auch die per-Event-Organizer-Eigenschaft
  // nicht — sonst blieben Sichtbarkeit und Freischalt-Bypässe die des Organizers.
  const isParentOrganizer = !previewAsUser && (props.parentEvent.organizerEmails || []).some(e => (e || '').toLowerCase() === myEmailLc);
  const visibleChildren = (isAdmin || isParentOrganizer)
    ? props.childEvents
    : props.childEvents.filter(ce =>
        registeredSet.has(ce.id)
        || isEventVisibleForUser(ce, currentUser.email, currentUser.location, groupEmails, currentUser.jobTitle));

  if (visibleChildren.length === 0) return null;

  // v29.25: Bei aktiver Abmelde-Sperre nicht „jederzeit abmelden" versprechen.
  const anyCancelLocked = visibleChildren.some(ce => registeredSet.has(ce.id) && selfCancelLocked(ce, props.parentEvent));
  // v31.8: Der Hinweis sagt jetzt ZUERST, was du tun kannst, und erst danach
  // die Mechanik (Mail, Outlook) — vorher stand die Technik vorn und dazu das
  // Versprechen „jederzeit", obwohl je Termin eine Frist gilt
  // (UI-Leitfaden 6d, „Erst was du tun kannst, dann die Mechanik").
  // v31.8 (Nacharbeit): Die beiden Richtungen hängen an ZWEI verschiedenen
  // Fristen — anmelden an `subEventRegDeadline`, abmelden an
  // `lastDeregisterDate` bzw. an der Organizer-Sperre (utils/cancelPolicy).
  // Der erste Entwurf hängte beides an die Anmeldefrist und sperrte damit im
  // Kopf des Lesers etwas, das der Code erlaubt: Ist die Anmeldefrist durch,
  // kann man sich normalerweise weiterhin abmelden. Der Satz nennt deshalb je
  // Richtung nur die Frist, die dort wirklich gilt; wenn eine greift, steht
  // der Grund am jeweiligen Termin (`disabledReason`) bzw. im Warnkasten.
  // v31.8: Der Satz nennt den PLURAL — der braucht im Deutschen kein
  // Geschlecht. Zwei Anlaeufe davor waren falsch: „Ein Tag kannst du anmelden"
  // (Akkusativ fehlt) und „fuer jedes Tag" (Geschlecht geraten). Ohne
  // gepflegtes `childEventTermGender` raet die Heuristik nie „m", also ist
  // JEDE Formulierung falsch, die ein Geschlecht braucht — v29.60 hat die
  // Heuristik bewusst so gelassen, damit sich an Bestandsevents nichts still
  // verschiebt. „fuer einzelne Tage / Sessions / Workshops" stimmt immer.
  const hintDe = (subject: string): string =>
    `Du kannst dich hier für einzelne ${subject} an- und abmelden. Fürs Anmelden zählt die Anmeldefrist des jeweiligen Termins. Fürs Abmelden zählt die Abmeldefrist — und die sperrt nur, wenn die Organizer es so eingestellt haben; dann steht der Grund direkt beim Termin. Zu jeder An- und Abmeldung bekommst du eine Bestätigungs-Mail; der Outlook-Termin wird dabei angelegt bzw. zurückgezogen.`;
  const hintEn = (subject: string): string =>
    `You can register for individual ${subject} here and cancel them again. Registering follows the registration deadline of that date; cancelling follows the cancellation deadline — and that one only blocks if the organizers set it up that way, in which case the reason is shown right at the date. Every registration and cancellation sends you a confirmation email and adds or withdraws the Outlook entry.`;
  const hintLabel = isDe ? hintDe(termPluralSafe) : hintEn(termPluralSafe);
  // v31.8: Die Abmelde-Sperre hängt nicht mehr als vierter Nebensatz am
  // Erklärtext, sondern steht als Warnkasten — im Fließtext ging sie auf dem
  // Handy unter, und sie ist die einzige Einschränkung, die man VORHER wissen muss.
  const lockNotice = anyCancelLocked
    ? (props.parentEvent.noSelfCancel
      ? (isDe
          ? 'Die Organizer haben die Selbst-Abmeldung deaktiviert. Zum Abmelden wende dich bitte an die Organizer.'
          : 'The organizers have disabled self-cancellation. Please contact the organizers to cancel.')
      : (isDe
          ? 'Bei Terminen mit abgelaufener Abmeldefrist ist die Selbst-Abmeldung deaktiviert. Wende dich dafür bitte an die Organizer.'
          : 'For dates whose cancellation deadline has passed, self-cancellation is disabled. Please contact the organizers.'))
    : '';
  // v31.8: „noch nicht geladen" ist nicht dasselbe wie „nicht ermittelbar" —
  // `counts` ist bis zum ersten refresh() leer. Ein fehlender Eintrag darf
  // deshalb nicht als „Belegung unbekannt" behauptet werden; erst der
  // geschriebene `null`-Wert ist die Aussage „nicht lesbar" (CLAUDE.md:
  // ein Lesefehler ist keine Null — und auch keine Zahl).
  const countLoaded = (id: string): boolean => Object.prototype.hasOwnProperty.call(counts, id);
  const unknownOccupancyCount = visibleChildren.filter(ce =>
    typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0
    && countLoaded(ce.id) && counts[ce.id] === null).length;
  return (
    <div style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
      <div className="dex-ui-section-title">{headerLabel}</div>
      <div className="dex-ui-section-desc">{hintLabel}</div>
      {!!lockNotice && (
        <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 10 }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>{lockNotice}</span>
        </div>
      )}
      {unknownOccupancyCount > 0 && (
        <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 10 }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>
            {isDe
              ? `Bei ${unknownOccupancyCount} ${unknownOccupancyCount === 1 ? 'Termin' : 'Terminen'} konnten wir die Belegung gerade nicht lesen. Wie viele Plätze dort noch frei sind, wissen wir nicht — eine Anmeldung kann deshalb auch auf der Warteliste landen.`
              : `We could not read the occupancy of ${unknownOccupancyCount} ${unknownOccupancyCount === 1 ? 'date' : 'dates'} just now. We do not know how many seats are left there — a registration may therefore end up on the waitlist.`}
          </span>
        </div>
      )}
      {/* v30.9: Termin-Kalender auch in „Meine Events" — dieselbe Monats-
          Darstellung wie auf der Anmeldeseite (v28.91). Bei einer Office-Tage-
          Reihe mit vielen Terminen ist die Zeilen-Liste kaum zu erfassen;
          im Raster sieht man die eigenen (grünen) Tage auf einen Blick.
          Anders als auf der Anmeldeseite schaltet ein Klick hier DIREKT um:
          er läuft über handleToggle und damit über denselben Feld-Modal-,
          Peer-Cancel- und Sperr-Pfad wie die Buttons der Listen-Ansicht. */}
      {!!props.parentEvent.subEventCalendar && (() => {
        // v30.67: Berliner Kalendertag statt Browser-Tag — ein Termin „Di 00:00"
        // steht als 22:00Z des Vortags in der Liste, und auf einem UTC-Browser
        // rutschte die Kachel auf den Montag (gleiche Rechnung wie auf der
        // Anmeldeseite, `isoToLocal` ist die Umrechnung des Wizards).
        const dayOf = (iso?: string): string => {
          if (!iso) return '';
          return isoToLocal(iso).slice(0, 10);
        };
        const entries = visibleChildren
          .map(ce => ({ ce, key: dayOf(ce.startDate) }))
          .filter(e => !!e.key);
        if (entries.length === 0) return null;
        const byDay: Record<string, DeloitteEvent> = {};
        entries.forEach(e => { byDay[e.key] = e.ce; });
        const monthKeys: string[] = [];
        entries.forEach(e => {
          const mk = e.key.slice(0, 7);
          if (monthKeys.indexOf(mk) < 0) monthKeys.push(mk);
        });
        monthKeys.sort();
        const weekdays = isDe
          ? ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So']
          : ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
        const regCount = entries.filter(e => registeredSet.has(e.ce.id)).length;
        // v10.27-Anschluss: Antworten bearbeiten geht im Raster nicht direkt —
        // die angemeldeten Termine mit Abfragefeldern stehen deshalb als
        // kompakte Zeilen unter dem Kalender (öffnet dasselbe Edit-Modal).
        const editableRegs = entries
          .map(e => e.ce)
          .filter(ce => registeredSet.has(ce.id) && (ce.eventSpecificFields || []).some(f => !!f.label));
        return (
          <div>
            {monthKeys.map(mk => {
              const [my, mm] = mk.split('-').map(n => parseInt(n, 10));
              const first = new Date(my, mm - 1, 1);
              const daysInMonth = new Date(my, mm, 0).getDate();
              // Montag als erster Wochentag (getDay: So=0).
              const lead = (first.getDay() + 6) % 7;
              const cells: Array<string | null> = [];
              for (let i = 0; i < lead; i++) cells.push(null);
              for (let d = 1; d <= daysInMonth; d++) {
                cells.push(`${my}-${String(mm).padStart(2, '0')}-${String(d).padStart(2, '0')}`);
              }
              return (
                <div key={mk} style={{ marginBottom: 14 }}>
                  <div style={{ fontWeight: 700, fontSize: '0.9rem', marginBottom: 6, color: 'var(--dex-gray-700, #444)' }}>
                    {first.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { month: 'long', year: 'numeric' })}
                  </div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4 }}>
                    {weekdays.map(w => (
                      <div key={w} style={{ fontSize: '0.68rem', fontWeight: 700, color: 'var(--dex-gray-400)', textAlign: 'center', padding: '2px 0' }}>{w}</div>
                    ))}
                    {cells.map((key, i) => {
                      if (!key) return <div key={`e${i}`} />;
                      const ce = byDay[key];
                      const dayNum = parseInt(key.slice(8), 10);
                      if (!ce) {
                        return (
                          <div key={key} style={{
                            textAlign: 'center', padding: '8px 0', borderRadius: 8,
                            fontSize: '0.8rem', color: 'var(--dex-gray-300, #ccc)',
                          }}>{dayNum}</div>
                        );
                      }
                      const isReg = registeredSet.has(ce.id);
                      // v30.67: Wartelisten-Platz sichtbar machen (orange) — s. statusById.
                      const isWait = isReg && statusById[ce.id] === 'Warteliste';
                      const isBusy = busyId === ce.id;
                      // v30.67: null = Belegung unbekannt → weder „voll" noch „N frei".
                      // v31.8: … und „noch nicht gelesen" ist wieder etwas
                      // anderes als „nicht lesbar". `counts` ist bis zum Ende
                      // des ersten refresh() leer; ohne diese Unterscheidung
                      // behauptet der Kalender im ersten Moment „Belegung
                      // unbekannt". Derselbe Schutz wie in der Listen-Ansicht.
                      const countReady = countLoaded(ce.id);
                      const count: number | null = counts[ce.id] ?? null;
                      const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
                      const isFull = hasCap && count !== null && count >= (ce.maxParticipants || 0);
                      const free: number | null = hasCap ? (count !== null ? Math.max(0, (ce.maxParticipants || 0) - count) : null) : -1;
                      // Frist-/Freischalt-Rechnung wie in der Listen-Ansicht (v30.8/v30.2).
                      const effRegDeadline = subEventRegDeadline(props.parentEvent, ce);
                      const perDayRules = !!(props.parentEvent.subDeadlineRule && props.parentEvent.subDeadlineRule.reg) || !!props.parentEvent.subEventOpenRule;
                      const deadlinePassed = !!(effRegDeadline && new Date(effRegDeadline) < new Date())
                        || (!perDayRules && !!(props.parentEvent.klammerDeadline && new Date(props.parentEvent.klammerDeadline) < new Date()));
                      const deadlineLocked = deadlinePassed && !isAdmin && !isParentOrganizer;
                      const openFrom = ((): Date | null => {
                        const rule = props.parentEvent.subEventOpenRule;
                        if (!rule) return null;
                        if (rule.mode === 'fixed') {
                          const dd = new Date(rule.date || '');
                          return isFinite(dd.getTime()) ? dd : null;
                        }
                        if (!((rule.days || 0) > 0)) return null;
                        // v30.67: Berliner Kalendertag (s. dayOf) — sonst öffnet der
                        // Tag auf einem UTC-Browser einen Tag zu früh.
                        const baseLocal = isoToLocal(ce.startDate || '');
                        if (!baseLocal) return null;
                        const [by, bm, bd] = baseLocal.slice(0, 10).split('-').map(n => parseInt(n, 10));
                        const dd = new Date(by, bm - 1, bd);
                        if (rule.mode === 'week') dd.setDate(dd.getDate() - ((dd.getDay() + 6) % 7));
                        dd.setDate(dd.getDate() - (rule.days || 0));
                        dd.setHours(0, 0, 0, 0);
                        return dd;
                      })();
                      const notYetOpen = !isReg && !!openFrom && new Date() < openFrom;
                      const openLocked = notYetOpen && !isAdmin && !isParentOrganizer;
                      const openFromLabel = openFrom
                        ? openFrom.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit' })
                        : '';
                      const lockReason = isReg ? selfCancelLockReason(ce, props.parentEvent) : null;
                      const cancelLocked = !!lockReason;
                      const past = isEventOver(ce);
                      const disabled = isBusy || past
                        || (isReg ? cancelLocked : ((isFull) || deadlineLocked || openLocked));
                      const dlLabel = (() => {
                        const dl = new Date(effRegDeadline || '');
                        return isFinite(dl.getTime())
                          ? dl.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit' })
                          : '';
                      })();
                      const title = [
                        ce.title || '',
                        isReg
                          ? (isWait
                            ? (isDe ? 'Warteliste — Klick meldet ab' : 'Waitlist — click to cancel')
                            : (isDe ? 'Angemeldet — Klick meldet ab' : 'Registered — click to cancel'))
                          : (isDe ? 'Klick meldet an' : 'Click to register'),
                        hasCap
                          ? (free !== null
                            ? (isDe ? `${free} von ${ce.maxParticipants} Plätzen frei` : `${free} of ${ce.maxParticipants} seats free`)
                            : countReady
                            ? (isDe ? `${ce.maxParticipants} Plätze · Belegung nicht ermittelbar` : `${ce.maxParticipants} seats · occupancy unknown`)
                            : (isDe ? `${ce.maxParticipants} Plätze · Belegung wird noch geladen` : `${ce.maxParticipants} seats · occupancy still loading`))
                          : '',
                        cancelLocked
                          ? (lockReason === 'always'
                            ? (isDe ? 'Selbst-Abmeldung von den Organizern deaktiviert' : 'Self-cancellation disabled by the organizers')
                            : (isDe ? 'Abmeldefrist abgelaufen — Selbst-Abmeldung gesperrt' : 'Cancellation deadline passed — self-cancellation locked'))
                          : '',
                        !isReg && deadlinePassed
                          ? (deadlineLocked
                            ? (isDe ? 'Anmeldefrist abgelaufen' : 'Registration deadline passed')
                            : (isDe ? 'Anmeldefrist abgelaufen — als Organizer/Admin trotzdem wählbar' : 'Deadline passed — still selectable as organizer/admin'))
                          : '',
                        notYetOpen
                          ? (openLocked
                            ? (isDe ? `Anmeldung ab ${openFrom!.toLocaleDateString('de-DE')} möglich` : `Registration opens on ${openFrom!.toLocaleDateString('en-GB')}`)
                            : (isDe ? `Anmeldung öffnet regulär am ${openFrom!.toLocaleDateString('de-DE')} — als Organizer/Admin trotzdem wählbar` : `Opens on ${openFrom!.toLocaleDateString('en-GB')} — still selectable as organizer/admin`))
                          : '',
                        past ? (isDe ? 'Termin liegt in der Vergangenheit' : 'Date is in the past') : '',
                      ].filter(Boolean).join(' · ');
                      // v31.8: Der Zustand des Tages steht IMMER im Tag, nie
                      // nur im Hover oder im `title` — auf dem Handy gibt es
                      // beides nicht (UI-Leitfaden 6b). Statt „✓", „fix",
                      // „war bis" und „—" stehen hier ganze Angaben; die Folge
                      // des Tippens erklärt der Kasten unter dem Kalender.
                      // v31.8 (Nacharbeit): „vorbei" kam bisher ausschließlich
                      // im `title` vor — ein vergangener Tag sah aus wie ein
                      // gesperrter. Er steht jetzt als eigene Angabe im Tag und
                      // ergänzt bei einer bestehenden Anmeldung den Status,
                      // statt ihn zu ersetzen (wer teilgenommen hat, will das
                      // weiterhin sehen).
                      const pastSuffix = past ? (isDe ? ' · vorbei' : ' · past') : '';
                      const dayStateLabel = ((): string => {
                        if (isBusy) return isDe ? 'läuft…' : 'working…';
                        if (isReg) {
                          const base = cancelLocked
                            ? (isDe ? 'Selbst-Abmeldung gesperrt' : 'self-cancellation locked')
                            : (isWait ? (isDe ? 'Warteliste' : 'waitlist') : (isDe ? 'angemeldet' : 'registered'));
                          return base + pastSuffix;
                        }
                        if (past) return isDe ? 'vorbei' : 'past';
                        if (notYetOpen) return isDe ? `Anmeldung ab ${openFromLabel}` : `opens ${openFromLabel}`;
                        if (deadlinePassed) {
                          return dlLabel
                            ? (isDe ? `Anmeldefrist war am ${dlLabel}` : `registration closed ${dlLabel}`)
                            : (isDe ? 'Anmeldung geschlossen' : 'registration closed');
                        }
                        if (isFull) return isDe ? 'ausgebucht' : 'fully booked';
                        if (hasCap && free !== null) return isDe ? `${free} frei` : `${free} free`;
                        // v31.8: „unbekannt" ist ein eigener Zustand (6c) — vorher
                        // stand hier „—" und der Grund nur im `title`. Und solange
                        // noch gelesen wird, ist es nicht „unbekannt", sondern
                        // schlicht noch nichts.
                        if (hasCap) {
                          return countReady
                            ? (isDe ? 'Belegung unbekannt' : 'seats unknown')
                            : (isDe ? 'wird geladen…' : 'loading…');
                        }
                        return isDe ? 'anmelden' : 'register';
                      })();
                      return (
                        <button
                          key={key}
                          type="button"
                          onClick={() => { if (!disabled) void handleToggle(ce.id, isReg); }}
                          onMouseEnter={() => { if (!disabled) setDayHoverKey(key); }}
                          onMouseLeave={() => setDayHoverKey(h => (h === key ? '' : h))}
                          onFocus={() => { if (!disabled) setDayHoverKey(key); }}
                          onBlur={() => setDayHoverKey(h => (h === key ? '' : h))}
                          disabled={disabled}
                          title={title}
                          aria-pressed={isReg}
                          style={{
                            display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
                            gap: 2, padding: '6px 3px 5px', borderRadius: 8, minHeight: 56,
                            border: `1px solid ${isReg
                              ? (isWait ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)')
                              : (dayHoverKey === key ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)')}`,
                            background: isReg
                              ? (isWait
                                ? (dayHoverKey === key ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-orange, #ed8b00)')
                                : (dayHoverKey === key ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-green, #86bc25)'))
                              : (dayHoverKey === key ? 'rgba(134,188,37,0.10)' : '#fff'),
                            color: isReg ? '#fff' : (disabled ? 'var(--dex-gray-400)' : 'var(--dex-gray-800, #333)'),
                            cursor: disabled ? 'not-allowed' : 'pointer',
                            // v29.72-Regel: der Zustand ist für ALLE sichtbar, der
                            // Organizer/Admin-Bypass öffnet nur das Klicken.
                            opacity: disabled && !isReg ? 0.55 : (notYetOpen ? 0.65 : 1),
                            fontWeight: 700, fontSize: '0.82rem',
                            transition: 'background 120ms ease, border-color 120ms ease, transform 120ms ease',
                            transform: (dayHoverKey === key && !disabled) ? 'translateY(-1px)' : 'none',
                          }}
                        >
                          <span>{dayNum}</span>
                          {/* Zustandszeile des Tages — Text s. `dayStateLabel`. */}
                          <span style={{
                            fontSize: '0.58rem', fontWeight: 600, opacity: 0.9, lineHeight: 1.2,
                            display: 'block', width: '100%', textAlign: 'center',
                            wordBreak: 'break-word', hyphens: 'auto',
                          }}>
                            {dayStateLabel}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </div>
              );
            })}
            {/* v31.8: Was ein Tipp bewirkt, stand bisher nur im Hover eines
                Tags und im `title` — auf dem Handy also nirgends. Jetzt sagt
                der Kasten es in ganzen Sätzen (UI-Leitfaden 6b). */}
            <div className="dex-ui-callout dex-ui-callout--neutral">
              <span className="dex-ui-callout-icon"><Info size={16} /></span>
              <span>
                {regCount === 0
                  ? (isDe ? 'Du hast noch keinen Termin gebucht. ' : 'You have not booked a date yet. ')
                  : (isDe
                    ? `Du bist für ${regCount} ${regCount === 1 ? 'Termin' : 'Termine'} angemeldet — grün markiert. `
                    : `You are registered for ${regCount} ${regCount === 1 ? 'date' : 'dates'} — marked green. `)}
                {isDe
                  ? 'Ein Tipp auf einen grünen Tag meldet dich von diesem Termin ab; vorher wirst du gefragt. Ein Tipp auf einen weißen Tag meldet dich an. Blasse Tage kannst du gerade nicht buchen — warum, steht im Tag.'
                  : 'Tapping a green day cancels that date; you are asked first. Tapping a white day registers you. Pale days cannot be booked right now — the reason is written in the day.'}
              </span>
            </div>
            {editableRegs.length > 0 && (
              <div className="dex-ui-stack" style={{ marginTop: 10 }}>
                {editableRegs.map(ce => (
                  <div key={ce.id} className="dex-ui-card dex-ui-card--accent" style={{ padding: '10px 12px' }}>
                    <div style={{ fontWeight: 600, fontSize: '0.85rem', color: 'var(--dex-gray-800, #333)' }}>{ce.title || fmt(ce.startDate || '')}</div>
                    <div className="dex-ui-inline" style={{ marginTop: 6 }}>
                      {(ce.eventSpecificFields || [])
                        .filter(f => f.label && ((seData[ce.id] || {})[f.id] || '').trim())
                        .map(f => (
                          <FieldAnswerTag key={f.id} label={f.label} value={(seData[ce.id] || {})[f.id]} type={f.type} small />
                        ))}
                    </div>
                    {/* v31.8: Der Knopf steht links beim Inhalt statt rechts
                        außen in einer sonst leeren Zeilenhälfte (2a′). */}
                    <button
                      className="btn btn-secondary dex-ui-btn-sm"
                      style={{ marginTop: 10 }}
                      disabled={busyId === ce.id}
                      onClick={() => { setEditingId(ce.id); setEditDraft({ ...(seData[ce.id] || {}) }); }}
                    >
                      {isDe ? 'Antworten bearbeiten' : 'Edit answers'}
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })()}

      {!props.parentEvent.subEventCalendar && (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
        {(() => {
        const renderRow = (ce: DeloitteEvent, shownTitle: string): React.ReactElement => {
          const isReg = registeredSet.has(ce.id);
          const isBusy = busyId === ce.id;
          // v28.20: Auch die explizite Klammer-Frist des Hauptevents sperrt
          // das NACHTRÄGLICHE Anmelden (Abmelden bleibt möglich).
          // v30.8: effektive Tages-Frist (materialisierte Spalte, sonst
          // Fallback aus der rollierenden Regel). Die Klammer-Frist sperrt
          // nur noch, wenn KEINE Je-Termin-Logik konfiguriert ist — gleiche
          // Regel wie isRegistrationFullyClosed.
          const effRegDeadline = subEventRegDeadline(props.parentEvent, ce);
          const perDayRules = !!(props.parentEvent.subDeadlineRule && props.parentEvent.subDeadlineRule.reg) || !!props.parentEvent.subEventOpenRule;
          const deadlinePassed = !!(effRegDeadline && new Date(effRegDeadline) < new Date())
            || (!perDayRules && !!(props.parentEvent.klammerDeadline && new Date(props.parentEvent.klammerDeadline) < new Date()));
          // v30.2: „Anmeldung ab" (Freischalt-Regel der Klammer) galt bisher
          // nur auf der Anmeldeseite — ueber „Meine Events" liess sich JEDER
          // Termin sofort buchen. Gleiche Rechnung wie subOpenFrom in
          // RegistrationPage; Anzeige fuer alle Rollen, Klick-Sperre nur fuer
          // Teilnehmer (v29.72-Regel: Bypass oeffnet Interaktion, nie den
          // Anblick). Bestehende Anmeldungen bleiben verwaltbar.
          const openFrom = ((): Date | null => {
            const rule = props.parentEvent.subEventOpenRule;
            if (!rule) return null;
            if (rule.mode === 'fixed') {
              const dd = new Date(rule.date || '');
              return isFinite(dd.getTime()) ? dd : null;
            }
            if (!((rule.days || 0) > 0)) return null;
            // v30.67: Berliner Kalendertag (s. dayOf im Kalender) — sonst öffnet
            // der Tag auf einem UTC-Browser einen Tag zu früh.
            const baseLocal = isoToLocal(ce.startDate || '');
            if (!baseLocal) return null;
            const [by, bm, bd] = baseLocal.slice(0, 10).split('-').map(n => parseInt(n, 10));
            const dd = new Date(by, bm - 1, bd);
            if (rule.mode === 'week') dd.setDate(dd.getDate() - ((dd.getDay() + 6) % 7));
            dd.setDate(dd.getDate() - (rule.days || 0));
            dd.setHours(0, 0, 0, 0);
            return dd;
          })();
          const notYetOpen = !isReg && !!openFrom && new Date() < openFrom;
          const openLocked = notYetOpen && !isAdmin && !isParentOrganizer;
          // v30.67: null = Belegung unbekannt → weder „voll" noch eine Zahl.
          // v31.8: „noch nicht gelesen" ist davon zu trennen — s. `countLoaded`.
          const countReady = countLoaded(ce.id);
          const count: number | null = counts[ce.id] ?? null;
          const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
          const isFull = hasCap && count !== null && count >= (ce.maxParticipants || 0);
          // v30.67: Wartelisten-Platz sichtbar machen (orange) — s. statusById.
          const isWait = isReg && statusById[ce.id] === 'Warteliste';
          // v22.22: Vergangene Sessions sind weder an- noch abmeldbar
          // (Abmelde-Sperre für vergangene Events; EventContext blockt zusätzlich).
          // v29.25: Abmelde-Sperre (Organizer-Option, komplett oder nach der
          // Frist) — die Flags liegen auf dem Parent, maßgeblich ist die
          // eigene Sub-Event-Frist.
          const lockReason = isReg ? selfCancelLockReason(ce, props.parentEvent) : null;
          const cancelLocked = !!lockReason;
          const disabled = isBusy || (deadlinePassed && !isReg) || (isFull && !isReg) || isEventOver(ce) || cancelLocked || openLocked;
          // v11.31: Custom-Field-Antworten gehören INS Sub-Event-Karten-
          // Layout, nicht ausserhalb. Maintainer-Wunsch: Tags zwischen
          // der Datums-/Adress-Zeile und den Action-Buttons (rechts) als
          // kleiner grüner Pastell-Stripe IN der Karte.
          const filledFieldTags = (() => {
            if (!isReg) return null;
            const data = seData[ce.id] || {};
            const filled = (ce.eventSpecificFields || [])
              .filter(f => f.label && (data[f.id] || '').trim())
              .map(f => ({ label: f.label, value: data[f.id], type: f.type }));
            if (filled.length === 0) return null;
            return (
              <div className="dex-ui-inline" style={{ marginTop: 8 }}>
                {filled.map(({ label, value, type }) => (
                  <FieldAnswerTag key={label} label={label} value={value} type={type} small />
                ))}
              </div>
            );
          })();
          // v31.8: Datum der Anmeldefrist auch in der Listen-Ansicht — für den
          // Grund unter dem Knopf. Gerechnet wird nichts Neues: `effRegDeadline`
          // steht schon oben, hier wird es nur lesbar gemacht.
          const dlLabel = ((): string => {
            const dl = new Date(effRegDeadline || '');
            return isFinite(dl.getTime())
              ? dl.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit' })
              : '';
          })();
          // v31.8: Der Knopf behält seine Beschriftung („Abmelden"), auch wenn
          // er gesperrt ist — vorher HIESS er „Abmeldung gesperrt", also war die
          // Fehlermeldung die Beschriftung. Der Grund steht jetzt als Satz
          // darunter und ist damit auch auf dem Handy da (UI-Leitfaden 6b).
          // Reihenfolge wie in `disabled` — sonst nennt der Satz einen anderen
          // Grund als den, der tatsächlich sperrt.
          const disabledReason = ((): string => {
            if (isBusy) return '';
            if (deadlinePassed && !isReg) {
              return isDe
                ? (dlLabel
                  ? `Die Anmeldefrist war am ${dlLabel} — anmelden geht hier nicht mehr.`
                  : 'Die Anmeldefrist ist abgelaufen — anmelden geht hier nicht mehr.')
                : (dlLabel
                  ? `The registration deadline was on ${dlLabel} — you can no longer register here.`
                  : 'The registration deadline has passed — you can no longer register here.');
            }
            if (isFull && !isReg) {
              return isDe
                ? 'Dieser Termin ist ausgebucht.'
                : 'This date is fully booked.';
            }
            if (isEventOver(ce)) {
              return isDe
                ? 'Dieser Termin ist vorbei — an- und abmelden geht nicht mehr.'
                : 'This date is over — registering and cancelling are no longer possible.';
            }
            if (cancelLocked) {
              return lockReason === 'always'
                ? (isDe
                  ? 'Selbst-Abmeldung gesperrt: Die Organizer haben sie für dieses Event abgeschaltet. Zum Abmelden wende dich bitte an die Organizer.'
                  : 'Self-cancellation locked: the organizers have disabled it for this event. Please contact the organizers to cancel.')
                : (isDe
                  ? 'Selbst-Abmeldung gesperrt: Die Abmeldefrist ist abgelaufen. Zum Abmelden wende dich bitte an die Organizer.'
                  : 'Self-cancellation locked: the cancellation deadline has passed. Please contact the organizers to cancel.');
            }
            if (openLocked) {
              return isDe
                ? `Die Anmeldung für diesen Termin öffnet am ${openFrom!.toLocaleDateString('de-DE')}.`
                : `Registration for this date opens on ${openFrom!.toLocaleDateString('en-GB')}.`;
            }
            return '';
          })();
          return (
            // v31.8: Karte statt handgebautem Kasten; die grüne bzw. orange
            // Kante links trägt den Zustand (dex-ui-card--accent), damit der
            // Rahmen nicht die ganze Zeile einfärbt (Grundsatz 1.1).
            <div
              key={ce.id}
              className={cx('dex-ui-card', isReg ? 'dex-ui-card--accent' : 'dex-ui-card--soft')}
              style={{
                padding: '12px 14px',
                borderLeftColor: isReg && isWait ? 'var(--dex-orange, #ed8b00)' : undefined,
                // v30.2: noch nicht freigeschaltet → gedimmt (alle Rollen).
                opacity: notYetOpen ? 0.7 : 1,
              }}
            >
              <div style={{ fontWeight: 700, fontSize: '0.9rem', color: 'var(--dex-gray-800, #333)', lineHeight: 1.35 }}>{shownTitle}</div>
              <div className="dex-ui-row-sub" style={{ whiteSpace: 'normal' }}>
                {ce.startDate && <>{fmt(ce.startDate)}{ce.endDate ? ` – ${fmt(ce.endDate)}` : ''}</>}
                {ce.location && <>&nbsp;·&nbsp;{ce.location}</>}
              </div>
              {/* v31.8: Zustand und Belegung als Pillen statt als Klammer-
                  Zusätze in der Metazeile — auf dem Handy war die Zeile sonst
                  ein einziger Umbruch. Reine Anzeige, deshalb `dex-ui-pill`. */}
              <div className="dex-ui-inline" style={{ marginTop: 8 }}>
                {isReg && (
                  <span className={cx('dex-ui-pill', 'dex-ui-pill--wrap', isWait ? 'dex-ui-pill--orange' : 'dex-ui-pill--green')}>
                    {isWait
                      ? (isDe ? 'Du stehst auf der Warteliste' : 'You are on the waitlist')
                      : (isDe ? 'Du bist angemeldet' : 'You are registered')}
                  </span>
                )}
                {hasCap && !countReady && (
                  <span className="dex-ui-pill dex-ui-pill--wrap dex-ui-pill--gray">
                    {isDe ? `${ce.maxParticipants} Plätze · Belegung wird geladen…` : `${ce.maxParticipants} seats · loading occupancy…`}
                  </span>
                )}
                {/* v31.8: „unbekannt" ist ein eigener Zustand und wird benannt —
                    vorher stand „—" da und der Grund nur im `title` (6c).
                    Niemals als 0 oder als freier Platz rendern. */}
                {hasCap && countReady && count === null && (
                  <span className="dex-ui-pill dex-ui-pill--wrap dex-ui-pill--orange">
                    {isDe ? `${ce.maxParticipants} Plätze · Belegung unbekannt` : `${ce.maxParticipants} seats · occupancy unknown`}
                  </span>
                )}
                {/* v31.8: Zahl und „voll" standen als zwei Zeichen nebeneinander
                    („12/20" plus „(voll)") — jetzt eine Pille, damit nicht zwei
                    rote Marken dasselbe sagen. Beide Aussagen bleiben. */}
                {hasCap && count !== null && (
                  <span className={cx('dex-ui-pill', 'dex-ui-pill--wrap', isFull && !isReg ? 'dex-ui-pill--red' : 'dex-ui-pill--gray')}>
                    {isFull && !isReg
                      ? (isDe ? `Ausgebucht — ${count} von ${ce.maxParticipants} Plätzen belegt` : `Fully booked — ${count} of ${ce.maxParticipants} seats taken`)
                      : (isDe ? `${count} von ${ce.maxParticipants} Plätzen belegt` : `${count} of ${ce.maxParticipants} seats taken`)}
                  </span>
                )}
                {notYetOpen && (
                  <span className="dex-ui-pill dex-ui-pill--wrap dex-ui-pill--orange">
                    {openLocked
                      ? (isDe ? `Anmeldung ab ${openFrom!.toLocaleDateString('de-DE')}` : `Registration opens ${openFrom!.toLocaleDateString('en-GB')}`)
                      : (isDe ? `Anmeldung regulär ab ${openFrom!.toLocaleDateString('de-DE')} — als Organizer/Admin wählbar` : `Opens ${openFrom!.toLocaleDateString('en-GB')} — selectable as organizer/admin`)}
                  </span>
                )}
              </div>
              {filledFieldTags}
              {/* v31.8: Die Knöpfe stehen links beim Inhalt statt per
                  space-between am rechten Rand (2a′) — auf dem Handy war die
                  Knopfspalte sonst eine eigene, halb leere Zeile. */}
              <div className="dex-ui-inline" style={{ marginTop: 10 }}>
                <button
                  className={cx('btn', isReg ? 'btn-secondary' : 'btn-primary', 'dex-ui-btn-sm')}
                  disabled={disabled}
                  title={cancelLocked
                    ? (lockReason === 'always'
                      ? (isDe
                        ? 'Die Organizer haben die Selbst-Abmeldung für dieses Event deaktiviert. Bitte wende dich zum Abmelden an die Organizer.'
                        : 'The organizers have disabled self-cancellation for this event. Please contact the organizers to cancel.')
                      : (isDe
                        ? 'Die Abmeldefrist ist abgelaufen — eine Selbst-Abmeldung ist bei diesem Event danach nicht mehr möglich. Bitte wende dich an die Organizer.'
                        : 'The cancellation deadline has passed — self-cancellation is no longer possible for this event. Please contact the organizers.'))
                    : undefined}
                  onClick={() => handleToggle(ce.id, isReg)}
                >
                  {isBusy
                    ? (isReg ? (isDe ? 'Wird abgemeldet…' : 'Cancelling…') : (isDe ? 'Wird angemeldet…' : 'Registering…'))
                    : (isReg ? (isDe ? 'Abmelden' : 'Cancel') : (isDe ? 'Anmelden' : 'Register'))}
                </button>
                {isReg && (ce.eventSpecificFields || []).filter(f => f.label).length > 0 && (
                  <button
                    className="btn btn-secondary dex-ui-btn-sm"
                    disabled={isBusy}
                    onClick={() => {
                      setEditingId(ce.id);
                      setEditDraft({ ...(seData[ce.id] || {}) });
                    }}
                  >
                    {isDe ? 'Antworten bearbeiten' : 'Edit answers'}
                  </button>
                )}
              </div>
              {!!disabledReason && (
                <div className="dex-ui-help" style={{ marginTop: 6 }}>{disabledReason}</div>
              )}
            </div>
          );
        };
        const grouping = groupSubEventTabs(visibleChildren.map(ce => ce.title || ''));
        if (!grouping.grouped) return visibleChildren.map(ce => renderRow(ce, ce.title || untitled));
        // v30.79: gruppiert — Kopf je Gruppe mit Anzahl und eigenen Anmeldungen,
        // darunter die Zeilen ohne das wiederholte Präfix.
        return grouping.groups.map(g => {
          const members = g.idxs.map(i => visibleChildren[i]).filter(Boolean);
          const mine = members.filter(ce => registeredSet.has(ce.id)).length;
          const collapsed = collapsedGroups.has(g.label);
          const label = g.label === 'Weitere' ? (isDe ? 'Weitere' : 'Other') : g.label;
          return (
            <div key={g.label} style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              {/* v31.8: Aufklapper wie überall sonst — `dex-ui-disclosure` mit
                  ChevronDown statt ▸/▾ als Text. Der Pfeil im Text war kein
                  Bedienhinweis, und der Kopf hatte weder Hover noch sichtbaren
                  Tastaturfokus (UI-Leitfaden 6c). */}
              <button
                type="button"
                aria-expanded={!collapsed}
                className={cx('dex-ui-disclosure', !collapsed && 'is-open')}
                onClick={() => setCollapsedGroups(prev => { const next = new Set(prev); if (next.has(g.label)) next.delete(g.label); else next.add(g.label); return next; })}
              >
                <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                <span style={{ fontWeight: 800, fontSize: '0.9rem', color: 'var(--dex-gray-800, #333)' }}>{label}</span>
                <span className="dex-ui-muted">{members.length} {termPluralSafe}</span>
                {mine > 0 && (
                  <span className="dex-ui-pill dex-ui-pill--green" style={{ marginLeft: 'auto' }}>
                    {mine} {isDe ? 'angemeldet' : 'registered'}
                  </span>
                )}
              </button>
              {!collapsed && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6, paddingLeft: 14 }}>
                  {members.map(ce => renderRow(ce, stripGroupPrefix(ce.title || untitled, g.label)))}
                </div>
              )}
            </div>
          );
        });
        })()}
      </div>
      )}

      {/* v10.27: Sub-Event-Edit-Modal — analog zum Hauptevent-Inline-Edit
          (siehe weiter oben), aber als Modal damit es übersichtlich
          bleibt. Beim Speichern geht der Draft per
          updateMyRegistration(subEventId, ...) in die Subsite zurück. */}
      {editingId && (() => {
        const ce = props.childEvents.find(c => c.id === editingId);
        if (!ce) return null;
        // v24.68 BUG-FIX: Sichtbarkeitsbedingung (showIf) auch beim NACHTRÄGLICHEN
        // Bearbeiten respektieren — vorher wurden ALLE bedingten Felder gezeigt
        // (z.B. „Auswahl für Rice Bowl" UND „Auswahl für Com Curry"), egal welches
        // Gericht gewählt war. Quelle der Bedingungs-Antwort ist der live
        // editDraft, damit das Umschalten des steuernden Felds die abhängigen
        // Felder sofort ein-/ausblendet (gleiche Logik wie das Anmelde-Modal).
        const isFieldVisible = (f: EventSpecificField): boolean => {
          if (!f.showIf || !f.showIf.fieldId) return true;
          const raw = (editDraft[f.showIf.fieldId] || '').trim();
          if (!raw) return false;
          const answers = raw.indexOf(' | ') >= 0 ? raw.split(' | ').map(s => s.trim()).filter(Boolean) : [raw];
          return answers.some(a => f.showIf!.values.indexOf(a) >= 0);
        };
        const allFields = (ce.eventSpecificFields || []).filter(f => f && f.label);
        const fields = allFields.filter(isFieldVisible);
        const closeModal = (): void => { setEditingId(null); setEditDraft({}); };
        const saveEdit = async (): Promise<void> => {
          setSavingEdit(true);
          try {
            // v24.68 BUG-FIX: Aktuell ausgeblendete bedingte Felder beim Speichern
            // LEEREN — sonst bleibt eine alte Auswahl (z.B. „Auswahl für Com Curry")
            // erhalten und „ergänzt" nur die neue, obwohl das Gericht gewechselt
            // wurde.
            const payload: Record<string, string> = { ...editDraft };
            for (const f of allFields) {
              if (f.showIf && f.showIf.fieldId && !isFieldVisible(f)) payload[f.id] = '';
            }
            await props.updateMyRegistration(editingId, payload);
            await refresh();
            await props.onMutated();
            closeModal();
          } finally {
            setSavingEdit(false);
          }
        };
        return (
          // v31.8: Kopf und Fußzeile über die Modal-Props statt eigenem <h3>
          // und eigener Knopfzeile (UI-Leitfaden, Abschnitt 4 „Modale").
          <Modal
            open={true}
            onClose={closeModal}
            dismissable={!savingEdit}
            maxWidth={520}
            padding={24}
            ariaLabel={ce.title || untitled}
            title={ce.title || untitled}
            subtitle={isDe
              ? `Hier änderst du deine Antworten zu ${termThisDatDe}. Gespeichert wird erst, wenn du unten auf Speichern tippst.`
              : `Change your answers for this ${termSingularSafe} here. Nothing is saved until you tap Save below.`}
            icon={<Pencil size={20} />}
            footer={<>
              <button className="btn btn-secondary" onClick={closeModal} disabled={savingEdit}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button className="btn btn-primary" onClick={saveEdit} disabled={savingEdit}>
                {savingEdit ? (isDe ? 'Wird gespeichert…' : 'Saving…') : (isDe ? 'Speichern' : 'Save')}
              </button>
            </>}
          >
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                {fields.map(f => {
                  const val = editDraft[f.id] || '';
                  const setVal = (v: string): void => setEditDraft(prev => ({ ...prev, [f.id]: v }));
                  return (
                    <div key={f.id}>
                      <label className="form-label" style={{ display: 'block', fontSize: '0.85rem', marginBottom: 4 }}>
                        {f.label}
                        {f.required && <span style={{ color: 'var(--dex-red, #c00)', marginLeft: 4 }}>*</span>}
                        {/* v11.16: konsistenter InfoTooltip — gleicher
                            Look wie auf der Register-Page und im
                            Sub-Event-Modal. */}
                        {f.helpText && <InfoTooltip text={f.helpText} />}
                      </label>
                      {f.type === 'select' ? (
                        f.multi ? (
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                            {(f.options || []).map(opt => {
                              const current = val.split(' | ').map(s => s.trim()).filter(Boolean);
                              const checked = current.indexOf(opt) >= 0;
                              return (
                                <label key={opt} style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                                  <input
                                    type="checkbox"
                                    checked={checked}
                                    onChange={e => {
                                      const next = e.target.checked
                                        ? [...current, opt]
                                        : current.filter(x => x !== opt);
                                      setVal(next.join(' | '));
                                    }}
                                  />
                                  {opt}
                                </label>
                              );
                            })}
                          </div>
                        ) : (
                          <select className="form-select" value={val} onChange={e => setVal(e.target.value)}>
                            <option value="">{isDe ? '— bitte wählen —' : '— please select —'}</option>
                            {(f.options || []).map(opt => <option key={opt} value={opt}>{opt}</option>)}
                          </select>
                        )
                      ) : f.type === 'checkbox' ? (
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                          <input
                            type="checkbox"
                            checked={val === 'true'}
                            onChange={e => setVal(e.target.checked ? 'true' : 'false')}
                          />
                          {f.label}
                        </label>
                      ) : f.type === 'number' ? (
                        <input className="form-input" type="number" value={val} onChange={e => setVal(e.target.value)} />
                      ) : (
                        <input className="form-input" type="text" value={val} onChange={e => setVal(e.target.value)} />
                      )}
                    </div>
                  );
                })}
              </div>
          </Modal>
        );
      })()}

      {/* v15.8: Peer-Cancel-Modal mit Checkbox-Liste. User hakt direkt
          an, welche zusätzlichen Sub-Events er mitabmelden will.
          Default: target ist immer abgemeldet (lock-Checkbox); peers
          unchecked. Klick auf „Abmelden" verarbeitet die Auswahl. */}
      {peerCancelDialog && <PeerCancelCheckboxModal dlg={peerCancelDialog} isDe={isDe} isSectionedEvent={!!props.parentEvent.requireSubEventSelection} termPlural={termPluralSafe} />}
      {/* v24.30 BUG-FIX: Feld-Modal beim Anmelden für ein Sub-Event aus „Meine
          Events" — fragt die (Pflicht-)Felder ab (vorher gingen sie verloren).
          Gleiche Felder-Logik (inkl. Sichtbarkeitsbedingung) wie das Sub-Event-
          Modal auf der Anmeldeseite. */}
      {regFieldsModal && (() => {
        const ce = regFieldsModal.ce;
        const fields = (ce.eventSpecificFields || [])
          .filter(f => f && f.label)
          .filter(f => {
            if (!f.showIf || !f.showIf.fieldId) return true;
            const raw = (regDraft[f.showIf.fieldId] || '').trim();
            if (!raw) return false;
            const answers = raw.indexOf(' | ') >= 0 ? raw.split(' | ').map(s => s.trim()).filter(Boolean) : [raw];
            return answers.some(a => f.showIf!.values.indexOf(a) >= 0);
          });
        const setVal = (id: string, v: string): void => setRegDraft(prev => ({ ...prev, [id]: v }));
        const isMissing = (f: EventSpecificField): boolean =>
          !!f.required && (f.type === 'checkbox' ? regDraft[f.id] !== 'true' : !((regDraft[f.id] || '').trim()));
        const onSubmit = (): void => {
          if (fields.some(isMissing)) { setRegShowErrors(true); return; }
          regFieldsModal.resolve({ ...regDraft });
        };
        const onCancel = (): void => regFieldsModal.resolve('abort');
        const errStyle = (f: EventSpecificField): React.CSSProperties =>
          regShowErrors && isMissing(f) ? { borderColor: 'var(--dex-red, #c00)', boxShadow: '0 0 0 1px var(--dex-red, #c00) inset' } : {};
        return (
          // v31.8: Kopf und Fußzeile über die Modal-Props (UI-Leitfaden 4).
          <Modal
            open={true}
            onClose={onCancel}
            maxWidth={520}
            padding={24}
            ariaLabel={ce.title || untitled}
            title={ce.title || untitled}
            subtitle={isDe
              ? `Beantworte die Fragen zu ${termThisDatDe} — danach melden wir dich an.`
              : `Answer the questions about this ${termSingularSafe} — we register you right after.`}
            icon={<FileText size={20} />}
            footer={<>
              <button className="btn btn-secondary" onClick={onCancel}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
              <button className="btn btn-primary" onClick={onSubmit}>{isDe ? 'Anmelden' : 'Register'}</button>
            </>}
          >
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              {fields.map(f => {
                const val = regDraft[f.id] || '';
                return (
                  <div key={f.id}>
                    <label className="form-label" style={{ display: 'block', fontSize: '0.85rem', marginBottom: 4 }}>
                      {f.required && <span style={{ color: 'var(--dex-red, #c00)', marginRight: 4 }}>*</span>}
                      {f.label}
                      {f.helpText && <InfoTooltip text={f.helpText} />}
                    </label>
                    {f.type === 'select' ? (
                      f.multi ? (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          {(f.options || []).map(opt => {
                            const current = val.split(' | ').map(s => s.trim()).filter(Boolean);
                            const checked = current.indexOf(opt) >= 0;
                            return (
                              <label key={opt} style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                                <input type="checkbox" checked={checked} onChange={e => setVal(f.id, (e.target.checked ? [...current, opt] : current.filter(x => x !== opt)).join(' | '))} />
                                {opt}
                              </label>
                            );
                          })}
                        </div>
                      ) : (
                        <select className="form-select" value={val} onChange={e => setVal(f.id, e.target.value)} style={errStyle(f)}>
                          <option value="">{isDe ? '— bitte wählen —' : '— please select —'}</option>
                          {(f.options || []).map(opt => <option key={opt} value={opt}>{opt}</option>)}
                        </select>
                      )
                    ) : f.type === 'checkbox' ? (
                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                        <input type="checkbox" checked={val === 'true'} onChange={e => setVal(f.id, e.target.checked ? 'true' : 'false')} />
                        {f.label}
                      </label>
                    ) : f.type === 'number' ? (
                      <input className="form-input" type="number" value={val} onChange={e => setVal(f.id, e.target.value)} style={errStyle(f)} />
                    ) : f.type === 'daterange' ? (
                      <StayRangePicker
                        value={val}
                        onChange={(next: string) => setVal(f.id, next)}
                        isDe={isDe}
                        rangeStart={f.rangeStart}
                        rangeEnd={f.rangeEnd}
                        maxNights={f.maxNights}
                        required={f.required}
                        compact
                      />
                    ) : f.type === 'date' ? (
                      <input className="form-input" type={f.withTime ? 'datetime-local' : 'date'} value={val} onChange={e => setVal(f.id, e.target.value)} style={errStyle(f)} />
                    ) : (
                      <input className="form-input" type="text" value={val} onChange={e => setVal(f.id, e.target.value)} style={errStyle(f)} />
                    )}
                  </div>
                );
              })}
            </div>
            {regShowErrors && fields.some(isMissing) && (
              <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>{isDe ? 'Bitte fülle die rot markierten Pflichtfelder aus.' : 'Please fill in the required fields marked in red.'}</span>
              </div>
            )}
          </Modal>
        );
      })()}
      {/* v15.21: Globaler Progress-Overlay während (Peer-)Cancel +
          Registrierung — blockiert die ganze Seite, damit der User nicht
          versehentlich nochmal klickt und sieht, dass die Aktion läuft. */}
      {/* v30.79: dasselbe Overlay wie beim Anmelden (SubmitOverlay) — mit
          echtem Prozent-Balken statt der laufenden Animation. Der
          beforeunload-Guard (Hook oben) bleibt. */}
      {processingMessage && (
        <SubmitOverlay
          displayProgress={processingProgress}
          locale={isDe ? 'de' : 'en'}
          submitProgressLabel={processingMessage}
          title={isDe ? 'Wird verarbeitet …' : 'Processing …'}
        />
      )}
    </div>
  );
}

// v15.8: Peer-Cancel-Modal mit Checkbox-Liste pro Sub-Event.
// Target ist locked und immer ausgewählt (User hat ja explizit Cancel
// geklickt); Peers sind initial unchecked. Eine „Bestätigen"-Aktion
// schickt die peer-IDs der angehakten Sub-Events an den resolver.
function PeerCancelCheckboxModal(props: {
  dlg: {
    targetTitle: string;
    peers: { id: string; title: string; startDate?: string; endDate?: string; location?: string }[];
    resolve: (_choice: { peerIds: string[] } | 'abort') => void;
  };
  isDe: boolean;
  isSectionedEvent: boolean;
  /** v31.8: Bezeichnung des Organizers (childEventTermPlural) — statt des
   *  fest verdrahteten „Sub-Events". Der Sections-Zweig bleibt, weil er an
   *  `requireSubEventSelection` hängt und einen eigenen Begriff meint. */
  termPlural: string;
}): React.ReactElement {
  const { dlg, isDe, isSectionedEvent, termPlural } = props;
  const [selectedPeerIds, setSelectedPeerIds] = React.useState<Set<string>>(new Set());
  const peerTermPl = isSectionedEvent
    ? (isDe ? 'Sections' : 'sections')
    : termPlural;
  const togglePeer = (id: string): void => {
    setSelectedPeerIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  // v30.79: von–bis statt nur Start (gleicher Tag: „12.10.2026, 13:00 – 14:30").
  const formatLine = (startDate?: string, endDate?: string, location?: string): string =>
    [formatDateTimeRange(startDate, endDate, isDe), location].filter(Boolean).join(' · ');
  // v30.79: dieselbe Präfix-Gruppierung wie in der Termin-Liste — bei vielen
  // Terminen sonst eine lange Liste gleich beginnender Namen. Gruppen sind
  // hier nur Zwischenüberschriften, nichts klappt zu (Auswahl muss sichtbar sein).
  const peerGrouping = groupSubEventTabs(dlg.peers.map(p => p.title));
  const renderPeer = (p: { id: string; title: string; startDate?: string; endDate?: string; location?: string }, shownTitle: string): React.ReactElement => {
    const checked = selectedPeerIds.has(p.id);
    const meta = formatLine(p.startDate, p.endDate, p.location);
    return (
      <label key={p.id} style={{
        display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
        borderRadius: 8,
        border: `1px solid ${checked ? 'var(--dex-red, #d62828)' : 'var(--dex-gray-200)'}`,
        background: checked ? 'rgba(214,40,40,0.04)' : '#fff',
        cursor: 'pointer',
        transition: 'border-color 0.15s, background 0.15s',
      }}>
        <input
          type="checkbox"
          checked={checked}
          onChange={() => togglePeer(p.id)}
          style={{ marginTop: 3, accentColor: 'var(--dex-red, #d62828)', cursor: 'pointer' }}
        />
        <div style={{ flex: 1 }}>
          <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)' }}>{shownTitle}</div>
          {meta && (
            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>{meta}</div>
          )}
        </div>
      </label>
    );
  };
  const totalCount = 1 + selectedPeerIds.size;
  return (
    // v31.8: Kopf und Fußzeile über die Modal-Props (UI-Leitfaden 4). Der
    // Bestätigen-Knopf ist `btn-danger` statt eines lokal rot eingefärbten
    // `btn-primary` — `.btn-danger` ist in dieser App bewusst grau, die
    // Warnung trägt der Text (6b), und ein selbst gebauter roter Knopf ist
    // genau das, was der Leitfaden verbietet.
    <Modal
      open={true}
      onClose={() => dlg.resolve('abort')}
      maxWidth={560}
      ariaLabel={isDe ? 'Abmeldung — Auswahl' : 'Cancellation — selection'}
      title={isDe ? `Welche ${peerTermPl} möchtest du abmelden?` : `Which ${peerTermPl} do you want to cancel?`}
      subtitle={isDe
        ? 'Wähle die Einträge an, die du abmelden willst. Nicht angewählte Anmeldungen bleiben erhalten.'
        : 'Tick the entries you want to cancel. Unticked registrations stay active.'}
      icon={<AlertCircle size={20} />}
      footer={<>
        <button
          type="button"
          className="btn btn-secondary"
          onClick={() => dlg.resolve('abort')}
        >
          {isDe ? 'Abbrechen' : 'Cancel'}
        </button>
        <button
          type="button"
          className="btn btn-danger"
          onClick={() => dlg.resolve({ peerIds: Array.from(selectedPeerIds) })}
        >
          {isDe ? `Abmelden (${totalCount})` : `Cancel ${totalCount}`}
        </button>
      </>}
    >
      <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 320, overflowY: 'auto' }}>
        <label style={{
          display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
          borderRadius: 8,
          border: '1px solid var(--dex-red, #d62828)',
          background: 'rgba(214,40,40,0.06)',
        }}>
          <input type="checkbox" checked={true} disabled style={{ marginTop: 3, accentColor: 'var(--dex-red, #d62828)' }} />
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)' }}>{dlg.targetTitle}</div>
            <div style={{ fontSize: '0.75rem', color: 'var(--dex-red-dark, #8b1414)', marginTop: 2 }}>
              {isDe ? 'Dieses hast du gerade zur Abmeldung gewählt' : 'You just selected this one to cancel'}
            </div>
          </div>
        </label>
        {!peerGrouping.grouped
          ? dlg.peers.map(p => renderPeer(p, p.title))
          : peerGrouping.groups.map(g => {
            const members = g.idxs.map(i => dlg.peers[i]).filter(Boolean);
            const picked = members.filter(p => selectedPeerIds.has(p.id)).length;
            const label = g.label === 'Weitere' ? (isDe ? 'Weitere' : 'Other') : g.label;
            const allOn = picked === members.length;
            return (
              <div key={g.label} style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '6px 4px 2px', borderBottom: '1px solid var(--dex-gray-200)' }}>
                  <span style={{ fontWeight: 800, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>{label}</span>
                  <span style={{ fontSize: '0.74rem', color: 'var(--dex-gray-600)' }}>{picked}/{members.length} {isDe ? 'gewählt' : 'picked'}</span>
                  {/* Alle Termine der Gruppe auf einmal an- oder abwählen —
                      bei „Day 2 komplett absagen" sonst fünf Einzelklicks.
                      v31.8: steht links beim Inhalt statt per marginLeft:auto
                      am rechten Rand (2a′). */}
                  <button
                    type="button"
                    className="dex-ui-textbtn"
                    onClick={() => setSelectedPeerIds(prev => {
                      const next = new Set(prev);
                      members.forEach(p => { if (allOn) next.delete(p.id); else next.add(p.id); });
                      return next;
                    })}
                  >
                    {allOn ? (isDe ? 'Keine' : 'None') : (isDe ? 'Alle' : 'All')}
                  </button>
                </div>
                {members.map(p => renderPeer(p, stripGroupPrefix(p.title, g.label)))}
              </div>
            );
          })}
      </div>
    </Modal>
  );
}
