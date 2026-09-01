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
import { isEventOver, subEventRegDeadline } from '../../utils/eventFormat';
import { selfCancelLocked, selfCancelLockReason } from '../../utils/cancelPolicy';
import { addPendingShadowParent } from '../../utils/shadowHeal';
import { isEventVisibleForUser } from '../EventListPage';
import { InfoTooltip } from '../InfoTooltip';
import Modal from '../Modal';
import StayRangePicker from '../StayRangePickerLazy';
import { FieldAnswerTag } from './myEventsHelpers';

// ==================== Sub-Events im "My Events"-Tab ====================
// Seit v6.4: Sub-Events sind eigene DEX_Events-Items (childEventsOf(parentId)).
// Anmeldung/Abmeldung läuft über den normalen registerForEvent/cancelRegistration-
// Pfad — identisch zu einem Top-Level-Event. Eigene Teilnehmerliste pro Sub-Event.
// v10.27: Plus Anzeige + Bearbeitung der Sub-Event-spezifischen Antworten
// (eventSpecificData) — analog zum Edit-Modus auf der Hauptevent-Karte.
export default function MyEventSubEvents(props: {
  parentEvent: DeloitteEvent;
  childEvents: DeloitteEvent[];
  registerForEvent: (eventId: string, customData: Record<string, string>, participantFirstName?: string, participantLastName?: string, participantEmail?: string, preferredStarterType?: string, opts?: { suppressMail?: boolean; suppressOutlook?: boolean; skipReload?: boolean }) => Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' }>;
  cancelRegistration: (eventId: string, opts?: { suppressNotifications?: boolean }) => Promise<boolean>;
  getMyRegistration: (eventId: string) => Promise<SPRegistration | null>;
  getAllRegistrations: (eventId: string) => Promise<SPRegistration[]>;
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
  const [busyId, setBusyId] = React.useState<string | null>(null);
  // v15.21: Voll-Bild-Progress-Modal beim (Peer-)Cancel + Anmeldung. Vorher
  // war nur die einzelne Karte mit „…" markiert — bei Peer-Cancels haben die
  // peers visuell nicht reagiert, weil busyId nur die EINE id hält.
  const [processingMessage, setProcessingMessage] = React.useState<string>('');
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
  const [counts, setCounts] = React.useState<Record<string, number>>({});
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
        return { id: ce.id, isActive, data };
      }));
      const s = new Set<string>();
      const dataMap: Record<string, Record<string, string>> = {};
      for (const r of regs) {
        if (r.isActive) {
          s.add(r.id);
          dataMap[r.id] = r.data;
        }
      }
      setRegisteredSet(s);
      setSeData(dataMap);

      const countPairs = await Promise.all(props.childEvents.map(async (ce) => {
        const all = await props.getAllRegistrations(ce.id);
        const active = (all || []).filter(r => {
          const st = r.Status || '';
          return st === 'Angemeldet' || st === 'QR versendet' || st === 'Eingecheckt';
        }).length;
        return { id: ce.id, count: active };
      }));
      const map: Record<string, number> = {};
      for (const p of countPairs) map[p.id] = p.count;
      setCounts(map);
    } catch { /* ignore */ }
  }, [props.childEvents.map(c => c.id).join(',')]);

  React.useEffect(() => { refresh().catch(() => { /* ignore */ }); }, [refresh]);

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
          title: ce.title || (isDe ? 'Sub-Event' : 'Sub-event'),
          startDate: ce.startDate || '',
          location: ce.location || '',
        }));
      if (peers.length > 0) {
        const target = props.childEvents.find(ce => ce.id === childEventId);
        const targetTitle = (target && target.title) || (isDe ? 'dieses Sub-Event' : 'this sub-event');
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
    const totalSteps = (currentlyRegistered ? 1 + peerIdsToCancel.length : 1);
    const initialMsg = currentlyRegistered
      ? (isDe
          ? `Abmeldung wird verarbeitet… (1/${totalSteps})`
          : `Cancellation in progress… (1/${totalSteps})`)
      : (isDe ? 'Anmeldung wird verarbeitet…' : 'Registration in progress…');
    setProcessingMessage(initialMsg);
    try {
      if (currentlyRegistered) {
        await props.cancelRegistration(childEventId);
        let done = 1;
        for (const peerId of peerIdsToCancel) {
          done++;
          setProcessingMessage(isDe
            ? `Abmeldung wird verarbeitet… (${done}/${totalSteps})`
            : `Cancellation in progress… (${done}/${totalSteps})`);
          try { await props.cancelRegistration(peerId); }
          catch (err) { console.warn('[DEX] peer-cancel failed:', peerId, err); }
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
        // v30.14: Im Klammer-Modus (subEventsOnlyMode) braucht die Person auch
        // die Schatten-Klammer-Zeile — dieser Pfad (An­melden über „Meine
        // Events", seit v30.9 auch per Kalender-Klick) legte sie nie an und
        // produzierte „Fehlende Klammer-Anmeldung" im Organizer Center.
        // registerForEvent ist für die Klammer idempotent (aktive Schatten-
        // Zeile → kein zweiter Insert); still, weil reine Datenvollständigkeit.
        if (regRes.ok && props.parentEvent.subEventsOnlyMode) {
          try {
            await props.registerForEvent(props.parentEvent.id, {}, undefined, undefined, undefined, undefined,
              { suppressMail: true, suppressOutlook: true, skipReload: true });
          } catch (err) {
            console.warn('[DEX] shadow-parent ensure failed:', err);
            // v30.16: Nachzug-Merker — der EventContext holt die Klammer-Zeile
            // beim nächsten App-Start still nach (utils/shadowHeal).
            addPendingShadowParent({
              eventId: props.parentEvent.id, customData: {},
              firstName: '', lastName: '', email: currentUser.email || '', ts: Date.now(),
            });
          }
        }
      }
      setProcessingMessage(isDe ? 'Aktualisiere…' : 'Refreshing…');
      await refresh();
      await props.onMutated();
    } finally {
      setBusyId(null);
      setProcessingMessage('');
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

  // v15.13: Bezeichnung kommt jetzt direkt aus event.childEventTermPlural /
  // childEventTermSingular (Wizard-Setting). Fallback: „Sub-Events"-Begriff.
  // v29.13: Besteht das Event ausschließlich aus Sub-Events, heißen sie hier
  // wie auf der Anmeldeseite „Events" — es gibt kein Haupt-Event, unter dem
  // sie stehen könnten, und der Teilnehmer hat sich genau für diese Einträge
  // angemeldet. Ein eigener Begriff des Organizers geht weiterhin vor.
  const subOnly = !!props.parentEvent.subEventsOnlyMode;
  const termPlural = props.parentEvent.childEventTermPlural
    || (subOnly ? (isDe ? 'Events' : 'events') : '');
  const termSingular = props.parentEvent.childEventTermSingular
    || (subOnly ? (isDe ? 'Event' : 'event') : '');
  const headerLabel = termPlural || (isDe ? 'Sub-Events' : 'Sub-events');
  // v29.13: „Eine Session", aber „Ein Event" — der Artikel war fest verdrahtet.
  const termArticleDe = /(session|veranstaltung|einheit|runde|reihe|tour|führung|schicht|woche|gruppe|stunde)$/i.test(termSingular) ? 'Eine' : 'Ein';
  // v29.25: Bei aktiver Abmelde-Sperre nicht „jederzeit abmelden" versprechen.
  const anyCancelLocked = visibleChildren.some(ce => registeredSet.has(ce.id) && selfCancelLocked(ce, props.parentEvent));
  const lockSuffix = anyCancelLocked
    ? (props.parentEvent.noSelfCancel
      ? (isDe
          ? ' Hinweis: Die Organizer haben die Selbst-Abmeldung deaktiviert — zum Abmelden wende dich bitte an die Organizer.'
          : ' Note: the organizers have disabled self-cancellation — please contact the organizers to cancel.')
      : (isDe
          ? ' Hinweis: Bei Terminen mit abgelaufener Abmeldefrist ist die Selbst-Abmeldung deaktiviert — bitte wende dich dafür an die Organizer.'
          : ' Note: for dates whose cancellation deadline has passed, self-cancellation is disabled — please contact the organizers.'))
    : '';
  const hintLabel = (termSingular
    ? (isDe
        ? `${termArticleDe} ${termSingular} kannst du jederzeit nachträglich an- oder abmelden. Bei jeder Aktion bekommst du eine Bestätigungs-Mail und der Termin wird in Outlook angelegt bzw. zurückgezogen.`
        : `You can register or cancel a ${termSingular} at any time. Every action triggers a confirmation mail and creates or removes the Outlook calendar entry.`)
    : (isDe
        ? 'Sub-Events kannst du jederzeit nachträglich an- oder abmelden. Bei jeder Aktion bekommst du eine Bestätigungs-Mail und der Termin wird in Outlook angelegt bzw. zurückgezogen.'
        : 'You can register or cancel sub-events at any time. Every action triggers a confirmation mail and creates or removes the Outlook calendar entry.')) + lockSuffix;
  return (
    <div style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4 }}>
        {headerLabel}
      </div>
      <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginBottom: 10, lineHeight: 1.45 }}>
        {hintLabel}
      </div>
      {/* v30.9: Termin-Kalender auch in „Meine Events" — dieselbe Monats-
          Darstellung wie auf der Anmeldeseite (v28.91). Bei einer Office-Tage-
          Reihe mit vielen Terminen ist die Zeilen-Liste kaum zu erfassen;
          im Raster sieht man die eigenen (grünen) Tage auf einen Blick.
          Anders als auf der Anmeldeseite schaltet ein Klick hier DIREKT um:
          er läuft über handleToggle und damit über denselben Feld-Modal-,
          Peer-Cancel- und Sperr-Pfad wie die Buttons der Listen-Ansicht. */}
      {!!props.parentEvent.subEventCalendar && (() => {
        const dayOf = (iso?: string): string => {
          if (!iso) return '';
          const d = new Date(iso);
          if (isNaN(d.getTime())) return '';
          return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
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
                      const isBusy = busyId === ce.id;
                      const count = counts[ce.id] || 0;
                      const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
                      const isFull = hasCap && count >= (ce.maxParticipants || 0);
                      const free = hasCap ? Math.max(0, (ce.maxParticipants || 0) - count) : -1;
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
                        const base = new Date(ce.startDate || '');
                        if (!isFinite(base.getTime())) return null;
                        const dd = new Date(base.getFullYear(), base.getMonth(), base.getDate());
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
                          ? (isDe ? 'Angemeldet — Klick meldet ab' : 'Registered — click to cancel')
                          : (isDe ? 'Klick meldet an' : 'Click to register'),
                        hasCap
                          ? (isDe ? `${free} von ${ce.maxParticipants} Plätzen frei` : `${free} of ${ce.maxParticipants} seats free`)
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
                            gap: 1, padding: '6px 0 5px', borderRadius: 8, minHeight: 46,
                            border: `1px solid ${isReg
                              ? 'var(--dex-green, #86bc25)'
                              : (dayHoverKey === key ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)')}`,
                            background: isReg
                              ? (dayHoverKey === key ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-green, #86bc25)')
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
                          <span style={{ fontSize: '0.6rem', fontWeight: 600, opacity: 0.85 }}>
                            {isBusy
                              ? '…'
                              : isReg
                              // v30.20: Beim Überfahren eines grünen Tags steht
                              // „abmelden" statt ✓ — die Klick-Geste war sonst
                              // nicht zu erraten (Nutzer-Befund: „nicht intuitiv").
                              ? (cancelLocked ? (isDe ? 'fix' : 'fixed') : (dayHoverKey === key ? (isDe ? 'abmelden' : 'cancel') : '✓'))
                              : notYetOpen
                              ? (isDe ? `ab ${openFromLabel}` : `from ${openFromLabel}`)
                              : deadlinePassed
                              ? (dlLabel
                                ? (isDe ? `war bis ${dlLabel}` : `until ${dlLabel}`)
                                : (isDe ? 'zu' : 'closed'))
                              : isFull
                              ? (isDe ? 'voll' : 'full')
                              : hasCap
                              ? (isDe ? `${free} frei` : `${free} free`)
                              : '—'}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </div>
              );
            })}
            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
              {regCount === 0
                ? (isDe ? 'Noch kein Termin angemeldet.' : 'No date registered yet.')
                : (isDe
                  ? `${regCount} ${regCount === 1 ? 'Termin' : 'Termine'} angemeldet (grün).`
                  : `${regCount} ${regCount === 1 ? 'date' : 'dates'} registered (green).`)}
            </div>
            {editableRegs.length > 0 && (
              <div style={{ marginTop: 10, display: 'flex', flexDirection: 'column', gap: 6 }}>
                {editableRegs.map(ce => (
                  <div key={ce.id} style={{
                    display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12,
                    padding: '6px 12px', borderRadius: 8,
                    background: 'rgba(134,188,37,0.08)', border: '1px solid var(--dex-green, #86bc25)',
                  }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: '0.82rem' }}>{ce.title || fmt(ce.startDate || '')}</div>
                      <div style={{ marginTop: 4, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                        {(ce.eventSpecificFields || [])
                          .filter(f => f.label && ((seData[ce.id] || {})[f.id] || '').trim())
                          .map(f => (
                            <FieldAnswerTag key={f.id} label={f.label} value={(seData[ce.id] || {})[f.id]} type={f.type} small />
                          ))}
                      </div>
                    </div>
                    <button
                      className="btn btn-secondary"
                      style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                      disabled={busyId === ce.id}
                      onClick={() => { setEditingId(ce.id); setEditDraft({ ...(seData[ce.id] || {}) }); }}
                    >
                      {isDe ? 'Bearbeiten' : 'Edit'}
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
        {visibleChildren.map(ce => {
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
            const base = new Date(ce.startDate || '');
            if (!isFinite(base.getTime())) return null;
            const dd = new Date(base.getFullYear(), base.getMonth(), base.getDate());
            if (rule.mode === 'week') dd.setDate(dd.getDate() - ((dd.getDay() + 6) % 7));
            dd.setDate(dd.getDate() - (rule.days || 0));
            dd.setHours(0, 0, 0, 0);
            return dd;
          })();
          const notYetOpen = !isReg && !!openFrom && new Date() < openFrom;
          const openLocked = notYetOpen && !isAdmin && !isParentOrganizer;
          const count = counts[ce.id] || 0;
          const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
          const isFull = hasCap && count >= (ce.maxParticipants || 0);
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
              <div style={{ marginTop: 6, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                {filled.map(({ label, value, type }) => (
                  <FieldAnswerTag key={label} label={label} value={value} type={type} small />
                ))}
              </div>
            );
          })();
          return (
            <div key={ce.id} style={{
              display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12,
              padding: '8px 12px', borderRadius: 8,
              background: isReg ? 'rgba(134,188,37,0.08)' : 'var(--dex-gray-50, #fafafa)',
              border: `1px solid ${isReg ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
              // v30.2: noch nicht freigeschaltet → gedimmt (alle Rollen).
              opacity: notYetOpen ? 0.7 : 1,
            }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 600, fontSize: '0.88rem' }}>{ce.title || (isDe ? 'Session ohne Titel' : 'Untitled session')}</div>
                <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                  {ce.startDate && <>{fmt(ce.startDate)}{ce.endDate ? ` – ${fmt(ce.endDate)}` : ''}</>}
                  {ce.location && <>&nbsp;·&nbsp;{ce.location}</>}
                  {hasCap && (
                    <>&nbsp;·&nbsp;<span style={{ color: isFull ? 'var(--dex-red, #c00)' : 'inherit' }}>{count}/{ce.maxParticipants}</span></>
                  )}
                  {isReg && (
                    <span style={{ marginLeft: 8, color: 'var(--dex-green)', fontWeight: 600 }}>
                      ({isDe ? 'Angemeldet' : 'Registered'})
                    </span>
                  )}
                  {isFull && !isReg && (
                    <span style={{ marginLeft: 8, color: 'var(--dex-red, #c00)', fontWeight: 600 }}>
                      ({isDe ? 'voll' : 'full'})
                    </span>
                  )}
                  {notYetOpen && (
                    <span style={{ marginLeft: 8, color: 'var(--dex-orange, #ed8b00)', fontWeight: 600 }}>
                      ({openLocked
                        ? (isDe ? `Anmeldung ab ${openFrom!.toLocaleDateString('de-DE')}` : `Registration opens ${openFrom!.toLocaleDateString('en-GB')}`)
                        : (isDe ? `Anmeldung regulär ab ${openFrom!.toLocaleDateString('de-DE')} — als Organizer/Admin wählbar` : `Opens ${openFrom!.toLocaleDateString('en-GB')} — selectable as organizer/admin`)})
                    </span>
                  )}
                </div>
                {filledFieldTags}
              </div>
              <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                {isReg && (ce.eventSpecificFields || []).filter(f => f.label).length > 0 && (
                  <button
                    className="btn btn-secondary"
                    style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                    disabled={isBusy}
                    onClick={() => {
                      setEditingId(ce.id);
                      setEditDraft({ ...(seData[ce.id] || {}) });
                    }}
                  >
                    {isDe ? 'Bearbeiten' : 'Edit'}
                  </button>
                )}
                <button
                  className={`btn ${isReg ? 'btn-secondary' : 'btn-primary'}`}
                  style={{ fontSize: '0.75rem', padding: '4px 12px', minWidth: 92 }}
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
                  {isBusy ? '...' : (isReg ? (cancelLocked ? (isDe ? 'Abmeldung gesperrt' : 'Locked') : (isDe ? 'Abmelden' : 'Cancel')) : (isDe ? 'Anmelden' : 'Register'))}
                </button>
              </div>
            </div>
          );
        })}
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
          <Modal
            open={true}
            onClose={closeModal}
            dismissable={!savingEdit}
            maxWidth={520}
            padding={24}
            ariaLabel={ce.title || (isDe ? 'Sub-Event' : 'Sub-event')}
          >
              <h3 style={{ margin: '0 0 6px', fontSize: '1.1rem' }}>
                {ce.title || (isDe ? 'Sub-Event' : 'Sub-event')}
              </h3>
              <p style={{ margin: '0 0 18px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {isDe
                  ? 'Hier kannst du deine Antworten zu diesem Sub-Event aktualisieren.'
                  : 'Update your answers for this sub-event here.'}
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12, marginBottom: 18 }}>
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
              <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button className="btn btn-secondary" onClick={closeModal} disabled={savingEdit}>
                  {isDe ? 'Abbrechen' : 'Cancel'}
                </button>
                <button className="btn btn-primary" onClick={saveEdit} disabled={savingEdit}>
                  {savingEdit ? '…' : (isDe ? 'Speichern' : 'Save')}
                </button>
              </div>
          </Modal>
        );
      })()}

      {/* v15.8: Peer-Cancel-Modal mit Checkbox-Liste. User hakt direkt
          an, welche zusätzlichen Sub-Events er mitabmelden will.
          Default: target ist immer abgemeldet (lock-Checkbox); peers
          unchecked. Klick auf „Abmelden" verarbeitet die Auswahl. */}
      {peerCancelDialog && <PeerCancelCheckboxModal dlg={peerCancelDialog} isDe={isDe} isSectionedEvent={!!props.parentEvent.requireSubEventSelection} />}
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
          <Modal open={true} onClose={onCancel} maxWidth={520} padding={24} ariaLabel={ce.title || (isDe ? 'Sub-Event' : 'Sub-event')}>
            <h3 style={{ margin: '0 0 6px', fontSize: '1.1rem' }}>{ce.title || (isDe ? 'Sub-Event' : 'Sub-event')}</h3>
            <p style={{ margin: '0 0 18px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              {isDe ? 'Bitte beantworte die Fragen für dieses Sub-Event, dann wirst du angemeldet:' : 'Please answer the questions for this sub-event, then you’ll be registered:'}
            </p>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12, marginBottom: 18 }}>
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
              <div style={{ color: 'var(--dex-red, #c00)', fontSize: '0.8rem', marginBottom: 10 }}>
                {isDe ? 'Bitte fülle die Pflichtfelder aus.' : 'Please fill in the required fields.'}
              </div>
            )}
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              <button className="btn btn-secondary" onClick={onCancel}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
              <button className="btn btn-primary" onClick={onSubmit}>{isDe ? 'Anmelden' : 'Register'}</button>
            </div>
          </Modal>
        );
      })()}
      {/* v15.21: Globaler Progress-Overlay während (Peer-)Cancel +
          Registrierung — blockiert die ganze Seite, damit der User nicht
          versehentlich nochmal klickt und sieht, dass die Aktion läuft. */}
      {processingMessage && (
        <div
          role="dialog"
          aria-modal="true"
          aria-label={isDe ? 'Wird verarbeitet' : 'Processing'}
          style={{
            position: 'fixed', inset: 0, zIndex: 2000,
            background: 'rgba(0,0,0,0.55)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            padding: 16,
          }}
        >
          <div style={{
            background: '#fff', borderRadius: 12, padding: '28px 32px',
            maxWidth: 420, width: '100%',
            boxShadow: '0 16px 48px rgba(0,0,0,0.35)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16,
            textAlign: 'center',
          }}>
            <div aria-hidden="true" style={{
              width: '100%', height: 6, borderRadius: 3,
              background: '#e5e5e5',
              overflow: 'hidden', position: 'relative',
            }}>
              <div style={{
                position: 'absolute', top: 0, bottom: 0,
                width: '40%',
                background: 'var(--dex-green, #86bc25)',
                borderRadius: 3,
                animation: 'dexProgressSlide 1.2s ease-in-out infinite',
              }} />
            </div>
            <div style={{ fontWeight: 700, fontSize: '1rem', color: 'var(--dex-gray-800)' }}>
              {processingMessage}
            </div>
            <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>
              {isDe ? 'Bitte einen Moment Geduld…' : 'Please wait a moment…'}
            </div>
            {/* v30.19: pulsierender Warnhinweis — wie im Anmelde-Overlay der
                RegistrationPage; zusätzlich beforeunload-Guard (Hook oben). */}
            <div style={{
              display: 'flex', alignItems: 'center', gap: 8,
              padding: '8px 16px', borderRadius: 999,
              background: 'rgba(237,139,0,0.12)', border: '1px solid var(--dex-orange, #ed8b00)',
              color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 700, fontSize: '0.85rem',
              animation: 'dexWaitPulse 1.5s ease-in-out infinite',
            }}>
              {isDe
                ? 'Bitte warten — Fenster nicht schließen'
                : 'Please wait — do not close this window'}
            </div>
          </div>
          <style>{`@keyframes dexProgressSlide { 0% { left: -40%; } 100% { left: 100%; } }
@keyframes dexWaitPulse { 0%, 100% { opacity: 1; transform: scale(1); } 50% { opacity: 0.55; transform: scale(1.04); } }`}</style>
        </div>
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
    peers: { id: string; title: string; startDate?: string; location?: string }[];
    resolve: (_choice: { peerIds: string[] } | 'abort') => void;
  };
  isDe: boolean;
  isSectionedEvent: boolean;
}): React.ReactElement {
  const { dlg, isDe, isSectionedEvent } = props;
  const [selectedPeerIds, setSelectedPeerIds] = React.useState<Set<string>>(new Set());
  const peerTermPl = isSectionedEvent
    ? (isDe ? 'Sections' : 'sections')
    : (isDe ? 'Sub-Events' : 'sub-events');
  const togglePeer = (id: string): void => {
    setSelectedPeerIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  const formatLine = (startDate?: string, location?: string): string => {
    const dateStr = startDate
      ? new Date(startDate).toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })
      : '';
    return [dateStr, location].filter(Boolean).join(' · ');
  };
  const totalCount = 1 + selectedPeerIds.size;
  return (
    <Modal
      open={true}
      onClose={() => dlg.resolve('abort')}
      maxWidth={560}
      ariaLabel={isDe ? 'Abmeldung — Auswahl' : 'Cancellation — selection'}
    >
      <h3 style={{ margin: 0, fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
        {isDe ? `Welche ${peerTermPl} möchtest du abmelden?` : `Which ${peerTermPl} do you want to cancel?`}
      </h3>
      <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
        {isDe
          ? 'Wähle die Einträge an, die du abmelden willst. Nicht angewählte Anmeldungen bleiben erhalten.'
          : 'Tick the entries you want to cancel. Unticked registrations stay active.'}
      </p>
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
        {dlg.peers.map(p => {
          const checked = selectedPeerIds.has(p.id);
          const meta = formatLine(p.startDate, p.location);
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
                <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)' }}>{p.title}</div>
                {meta && (
                  <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>{meta}</div>
                )}
              </div>
            </label>
          );
        })}
      </div>
      <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 4 }}>
        <button
          type="button"
          className="btn btn-secondary"
          onClick={() => dlg.resolve('abort')}
        >
          {isDe ? 'Abbrechen' : 'Cancel'}
        </button>
        <button
          type="button"
          className="btn btn-primary"
          onClick={() => dlg.resolve({ peerIds: Array.from(selectedPeerIds) })}
          style={{ background: 'var(--dex-red, #d62828)', borderColor: 'var(--dex-red, #d62828)' }}
        >
          {isDe ? `Abmelden (${totalCount})` : `Cancel ${totalCount}`}
        </button>
      </div>
    </Modal>
  );
}
