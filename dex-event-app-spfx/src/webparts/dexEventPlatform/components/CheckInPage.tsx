/**
 * Check-In Seite - QR-Code Scanner + Foto-Upload + Manuelle Eingabe
 *
 * Versucht zuerst die Kamera direkt zu nutzen (funktioniert in vielen SPFx-Umgebungen).
 * Fallback 1: Foto-Upload (User macht Foto vom QR-Code, App liest es aus).
 * Fallback 2: Manuelle Code-Eingabe.
 */

import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useNavigation } from '../context/NavigationContext';
// v20.1: Self-Check-in direkt aus der Check-in-Seite (Live-QR + PDF-Druck).
import { generateSelfCheckInToken } from '../utils/selfCheckIn';
import { downloadSelfCheckInPdf } from '../utils/selfCheckInPdf';
// v20.4: modernes Alert-Modal statt window.alert.
import { useDialog } from '../context/DialogContext';
import { useRoles } from '../context/RoleContext';
import { useCurrentUser } from '../context/UserContext';
import { EventService, SPRegistration } from '../services/EventService';
import {
  checkInExtras, parseCustomData, CheckInExtra, shirtAllocate, parseShirtStock, ShirtAllocationResult,
  // v31.4: Trikot-Ausgabe am Tisch — was rausgegeben wurde, steht in der Zeile.
  parseShirtIssue, ShirtIssue, shirtFieldOf, splitShirtSize,
} from '../utils/checkInExtras';
// v31.4: „0412" und „412" sind auf dem Zettel dieselbe Nummer — der
// Überkleben-Hinweis muss beide finden.
import { bibKey } from '../utils/b2runBibPool';
import { parseAgendaCheckIns, parseAgendaMarks, parseAgendaNoShows, formatMarkTime, suggestCurrentAgendaItem } from '../utils/agendaCheckIns';
import Modal from './Modal';
// v31.4: Klassensatz aus docs/ui-leitfaden.md (Kästen, Pillen, Werkzeugleiste).
import { ensureDexUiStyles } from './dexUi';
import { agendaGroups, groupLabel, groupDateLabel } from '../utils/agendaGroups';
import { useLanguage } from '../context/LanguageContext';
import { useIsMobile } from '../utils/useIsMobile';
import OrganizerList from './OrganizerList';
import { AlertCircle, ChevronDown, ChevronUp } from './Icons';
// v20.0 (Audit): qr-scanner nur noch als Typ statisch importieren — die
// eigentliche Bibliothek wird erst beim Kamera-Start dynamisch nachgeladen.
import type QrScanner from 'qr-scanner';
import { shortSubEventTitle } from '../utils/subEventTitle';


/**
 * v31.4: Zwei Nummern, die dieselbe Zahl sein KÖNNEN — und es nach der ersten
 * Abmeldung nicht mehr sind.
 *
 * `TeilnehmerID` ist der laufende Rang; `reorderParticipantIDs` und der Flow
 * `DEX_IDReorder_TeilnehmerIDs` vergeben ihn bei jeder Abmeldung neu.
 * `QrSentId` ist die Zahl, die in der versendeten QR-Mail GEDRUCKT steht —
 * die, die der Teilnehmer am Einlass vorliest, wenn der Kamera-Scan scheitert
 * (auf Android in der SharePoint-App der Normalfall, s. CLAUDE.md
 * „Kamera-Scan"). Sie ändert sich nie. Der QR-Code selbst enthält
 * `DEX|<EventNr>|<E-Mail>` und ist von alldem nicht betroffen; wer hier etwas
 * ändert, hat mit dem Scan-Weg nichts zu tun.
 */
function numOrNull(v: unknown): number | null {
  if (v === undefined || v === null || v === '') return null;
  const n = Number(v);
  return (isFinite(n) && n > 0) ? n : null;
}
/** Die Nummer aus der versendeten QR-Mail (fest) — `null`, wenn nicht hinterlegt. */
function qrSentIdOf(r: SPRegistration): number | null { return numOrNull(r.QrSentId); }
/** Der laufende Rang von heute (wandert bei jeder Abmeldung). */
function runningIdOf(r: SPRegistration): number | null { return numOrNull(r.TeilnehmerID); }
/** Dreistellig wie in der Mail („012"). padStart gibt es im ES5-Target nicht. */
function pad3(n: number): string { let s = String(n); while (s.length < 3) s = `0${s}`; return s; }
/** Anzeigename einer Zeile — überall gleich, damit Karte und Liste dasselbe sagen. */
function regDisplayName(r: SPRegistration): string {
  return (r.Vorname && r.Nachname) ? `${r.Vorname} ${r.Nachname}` : (r.ParticipantName || r.ParticipantEmail || '-');
}

/** v31.4: Was der Ausgabetisch über EINE Person wissen muss. `issued` ist die
 *  Tatsache (schlägt alles andere), `preset` die Vorbelegung der Größenwahl:
 *  der Gegenvorschlag, sonst die Wunschgröße — aber nur, wenn die Antwort
 *  überhaupt eine Größe ist („T-Shirt bereits vorhanden" wäre keine). */
type ShirtDeskInfo = {
  wish: string;
  proposal: string | null;
  preset: string;
  issued: ShirtIssue | null;
  sizes: string[];
};

/**
 * v31.4: Größenwahl für die Trikot-Ausgabe.
 *
 * Eigene kleine Komponente, weil dieselbe Wahl an ZWEI Stellen steht
 * (Bestätigungskarte nach dem Scan, Dialog aus der Trefferliste) — zwei
 * Kopien liefen sonst irgendwann auseinander, und der „andere Größe"-Modus
 * bräuchte je Stelle einen eigenen State.
 *
 * Die Liste kommt aus der Verteilung des Events (Wünsche + Bestand + bereits
 * Ausgegebenes); „Andere Größe" bleibt trotzdem möglich, weil der Helfer am
 * Tisch manchmal einfach in einen anderen Karton greift.
 */
function ShirtSizePicker(props: {
  sizes: string[];
  value: string;
  onChange: (_size: string) => void;
  isDe: boolean;
}): React.ReactElement {
  const { sizes, value, onChange, isDe } = props;
  const OTHER = '__other__';
  const known = (v: string): boolean => sizes.some(s => s.toLowerCase() === (v || '').toLowerCase());
  const [free, setFree] = React.useState<boolean>(sizes.length === 0 || (!!value && !known(value)));
  const label = isDe ? 'Ausgegebene Größe' : 'Handed-out size';
  if (free) {
    return (
      <>
        <input
          type="text"
          className="dex-ui-input dex-ui-input--sm"
          value={value}
          aria-label={label}
          placeholder={isDe ? 'z.B. Herrengröße XL' : 'e.g. Men XL'}
          onChange={e => onChange(e.target.value)}
          style={{ maxWidth: 200 }}
        />
        {sizes.length > 0 && (
          <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={() => { setFree(false); onChange(sizes[0]); }}>
            {isDe ? 'aus der Liste wählen' : 'pick from the list'}
          </button>
        )}
      </>
    );
  }
  return (
    <select
      className="dex-ui-select dex-ui-select--sm"
      value={known(value) ? sizes.filter(s => s.toLowerCase() === value.toLowerCase())[0] : ''}
      aria-label={label}
      onChange={e => { if (e.target.value === OTHER) { setFree(true); onChange(''); } else onChange(e.target.value); }}
      style={{ maxWidth: 220 }}
    >
      {!known(value) && <option value="">{isDe ? '— Größe wählen —' : '— pick a size —'}</option>}
      {sizes.map(s => <option key={s} value={s}>{s}</option>)}
      <option value={OTHER}>{isDe ? 'Andere Größe…' : 'Other size…'}</option>
    </select>
  );
}

export default function CheckInPage(): React.ReactElement {
  const { events, getAllRegistrations, updateEvent } = useEvents();
  // v20.4: App-Modal statt nativem Browser-Alert.
  const { showAlert, confirmDialog } = useDialog();
  const { selectedEventId, navigate } = useNavigation();
  const { isAdmin, isOrganizer, siteUrl } = useRoles();
  const { currentUser } = useCurrentUser();
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  const isMobile = useIsMobile();
  // v31.4: Die Seite nutzt `dex-ui-*` (Kästen, Pillen, Werkzeugleiste) auch
  // außerhalb der Modals — das Stylesheet darf also nicht davon abhängen,
  // dass gerade eins offen war. Der Aufruf ist idempotent (wie AdminPage).
  ensureDexUiStyles();
  // v6.22 / v13.11: aktueller User-E-Mail über UserContext — der respektiert
  // die Demo-Impersonation (sonst greift hier immer die echte SPFx-Identität
  // des Admins, und Demo-Modus „Check-In-Team" käme nie an die Check-In-
  // Seite ran, weil die qrScannerEmails-Injektion auf die Demo-Mail zeigt).
  const currentEmailLc: string = (currentUser.email || '').toLowerCase();
  // Events, die der aktuelle User einchecken darf: Admin = alle, sonst nur die
  // Events in denen er Organizer oder QR-Code-Scanner ist (per E-Mail-Match).
  const accessibleEvents = React.useMemo(() => {
    return (events || []).filter(e => {
      if (isAdmin) return true;
      const orgMatch = (e.organizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
      const qrMatch = (e.qrScannerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
      return orgMatch || qrMatch;
    });
  }, [events, isAdmin, currentEmailLc]);
  // v15.3: Picker zeigt per Default nur aktive Events (gleicher Filter wie
  // EventListPage). Toggle „Nur aktive" oben rechts, der das aufweicht.
  const [onlyActiveCheckIn, setOnlyActiveCheckIn] = React.useState<boolean>(true);
  const visibleCheckInEvents = React.useMemo(() => {
    if (!onlyActiveCheckIn) return accessibleEvents;
    const now = Date.now();
    return accessibleEvents.filter(e => {
      if (e.status !== 'Active') return false;
      if (e.isFictive) return false;
      const activeFromTs = e.activeFrom ? new Date(e.activeFrom).getTime() : 0;
      if (activeFromTs > 0 && activeFromTs > now) return false;
      // Vergangene Events ebenfalls ausblenden (EndDate < heute - 1 Tag)
      const endTs = e.endDate ? new Date(e.endDate).getTime() : 0;
      if (endTs > 0 && endTs < now - 24 * 60 * 60 * 1000) return false;
      return true;
    });
  }, [accessibleEvents, onlyActiveCheckIn]);
  // v23.45: Sub-Events werden im Picker unter ihrem Hauptevent gruppiert und
  // standardmäßig eingeklappt — statt flacher Liste. Ein Hauptevent mit
  // Sub-Events bekommt einen Aufklapp-Pfeil; Klick auf die Karte selbst checkt
  // weiterhin direkt in das (Haupt-)Event ein.
  const groupedCheckInEvents = React.useMemo(() => {
    const visibleIds = new Set(visibleCheckInEvents.map(e => e.id));
    const childrenByParent: Record<string, typeof visibleCheckInEvents> = {};
    visibleCheckInEvents.forEach(e => {
      if (e.parentEventId && visibleIds.has(e.parentEventId)) {
        (childrenByParent[e.parentEventId] = childrenByParent[e.parentEventId] || []).push(e);
      }
    });
    // Top-Level = Events ohne sichtbaren Parent (echte Hauptevents ODER
    // verwaiste Sub-Events, deren Parent ausgeblendet/gefiltert ist).
    const topLevel = visibleCheckInEvents.filter(e => !(e.parentEventId && visibleIds.has(e.parentEventId)));
    return topLevel.map(parent => ({ parent, children: childrenByParent[parent.id] || [] }));
  }, [visibleCheckInEvents]);
  const [expandedCheckInParents, setExpandedCheckInParents] = React.useState<Record<string, boolean>>({});
  const scannerRef = React.useRef<QrScanner | null>(null);
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const [resultMessage, setResultMessage] = React.useState('');
  const [resultType, setResultType] = React.useState<'success' | 'error' | 'info' | ''>('');
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [isScanning, setIsScanning] = React.useState(false);
  const [cameraError, setCameraError] = React.useState('');
  // v30.29: Merkt sich, ob der Kamera-Fehler aus der iframe-Einbettung kam —
  // nur dann hilft der „Seite im Browser öffnen"-Knopf, sonst wäre er ein
  // Angebot, das nichts löst.
  const [cameraErrorInIframe, setCameraErrorInIframe] = React.useState(false);
  // v30.30: Foto-Weg — native Kamera-App per File-Input statt getUserMedia.
  const photoInputRef = React.useRef<HTMLInputElement>(null);
  // v30.31: Nur für die `accept`-Erweiterung unten — bewusst eng gefasste
  // UA-Prüfung, weil der Workaround eine Android-Eigenheit adressiert.
  const isAndroid = typeof navigator !== 'undefined' && /Android/i.test(navigator.userAgent || '');
  const [photoBusy, setPhotoBusy] = React.useState(false);
  const [photoHint, setPhotoHint] = React.useState('');
  // v30.35: Direkteingabe der Teilnehmer-ID (s. Block unter dem Scan-Knopf).
  const [idInput, setIdInput] = React.useState('');
  const [idError, setIdError] = React.useState('');
  // v31.4 (Review): Die Herkunftsangabe („gefunden über die laufende Nummer")
  // war ein eigener grauer Satz unter dem Eingabefeld — gesetzt NACH dem
  // Schreibvorgang, nie zurückgesetzt, und damit stand er noch da, als längst
  // die nächste Person vor dem Tisch war. Sie steht jetzt im Warnkasten der
  // Bestätigungskarte (`QrIdConflict.viaRunning`): vor dem Einchecken, und sie
  // verschwindet mit der Karte.
  const [checkedInCount, setCheckedInCount] = React.useState(0);
  const confirmCardRef = React.useRef<HTMLDivElement>(null);
  type PendingCheckInInfo = {
    name: string; email: string; event: { id?: string; subsiteUrl: string; title: string };
    regId: number; status: string; department?: string; jobTitle?: string; location?: string; photoUrl?: string;
    /** v30.53: Startnummer, Trikotgröße, Gruppe — was am Tisch gebraucht wird. */
    extras?: CheckInExtra[];
    /** v30.91: Programmpunkt, an dem eingecheckt wird (statt Event-Status). */
    agendaItemId?: string;
    agendaLabel?: string;
    /** v31.4: Alles, was die Trikot-Ausgabe an dieser Person braucht. */
    shirt?: ShirtDeskInfo | null;
    /** v31.4: „QR-Nr. 012 · laufend 009" — nur wenn beide Nummern auseinanderlaufen. */
    qrNote?: string;
    /** v31.4: Die getippte Nummer war zweideutig — s. QrIdConflict. */
    qrConflict?: QrIdConflict;
  };
  /**
   * v31.4: Warum eine getippte Zahl NICHT direkt eingecheckt wird.
   *
   * Die v31.1-Begründung für den Direkt-Check-in war „die Person ist hier
   * schon eindeutig gewählt". Genau das gilt für eine Zahl in mehreren Fällen
   * nicht — und dann entscheidet der Helfer vor dem Schreibvorgang, nicht die
   * Zahl danach. Die Gründe können zusammen auftreten, deshalb Flags statt
   * eines einzelnen Falls:
   *
   * `qrName`        — die Zahl stand in der QR-Mail dieser Person. Sie ist
   *                   vorbelegt, weil die Mail die Zusage ist, die die Person
   *                   in der Hand hält.
   * `altName`/`alt` — eine ANDERE aktive Zeile trägt heute dieselbe laufende
   *                   Nummer; für sie gibt es den Umschaltknopf.
   * `cancelledName` — eine abgemeldete Zeile trägt die Zahl. Einchecken lässt
   *                   sie sich nicht (sie löst die Nummer seit v31.4 auch
   *                   nicht mehr auf), aber sie erklärt, warum die Zahl auf
   *                   jemand anderen zeigt, als die Person erwartet.
   * `viaRunning`    — gefunden über die laufende Nummer, obwohl dieses Event
   *                   QR-Nummern hinterlegt hat. Der Treffer ist damit eine
   *                   Ebene schwächer als eine gedruckte Mail-Nummer.
   */
  type QrIdConflict = {
    typed: number;
    qrName?: string;
    altName?: string;
    alt?: SPRegistration;
    cancelledName?: string;
    viaRunning?: boolean;
  };
  const [pendingCheckIn, setPendingCheckIn] = React.useState<PendingCheckInInfo | null>(null);
  // v31.4: Trikot-Ausgabe — Größenwahl der Bestätigungskarte, Größenwahl des
  // Zeilen-Dialogs und ein gemeinsames Busy-Flag. Zwei getrennte Werte, weil
  // beide Stellen gleichzeitig offen sein können (Karte oben, Liste darunter).
  const [shirtAsk, setShirtAsk] = React.useState<{ reg: import('../services/EventService').SPRegistration; name: string; eventId: string } | null>(null);
  const [cardShirtSize, setCardShirtSize] = React.useState('');
  const [askShirtSize, setAskShirtSize] = React.useState('');
  const [shirtBusy, setShirtBusy] = React.useState(false);
  // v31.1: „Letzte Check-ins" dieser Sitzung — mit Rückgängig. Nur was HIER
  // eingecheckt wurde (Scan, ID, Liste); der vorherige Status wird gemerkt,
  // damit der Revert nichts erfindet.
  // v31.2: `kind` — auch ein No-Show landet hier und ist rücknehmbar (Nutzer
  // 07.09.2026: „No-Show soll auch rückgängig machbar sein").
  // v31.4: `shirt` — auch die Trikot-Ausgabe steht hier und ist rücknehmbar
  // (falsche Größe getippt, Shirt wieder eingesammelt).
  type RecentCheckIn = {
    key: string; at: string; name: string; regId: number; eventId: string; subsiteUrl: string;
    prevStatus: string; agendaItemId?: string; agendaLabel?: string; kind?: 'checkin' | 'noshow' | 'shirt';
    /** v31.4: Ausgegebene Größe — die Zeile nennt sie, sonst weiß niemand, was er zurücknimmt. */
    shirtSize?: string;
  };
  const [recentCheckIns, setRecentCheckIns] = React.useState<RecentCheckIn[]>([]);
  const [undoBusyKey, setUndoBusyKey] = React.useState<string>('');
  // v31.2: Rückfrage im Programmpunkt-Modus — No-Show nur am gewählten Punkt
  // oder für das ganze Event? Beides sind andere Daten (Marke vs. Status).
  const [noShowAsk, setNoShowAsk] = React.useState<{ reg: import('../services/EventService').SPRegistration; name: string } | null>(null);
  // v31.1: Live-Scanner und Self-Check-in sind Kacheln zum Aufklappen —
  // Standard zu (Nutzer 07.09.2026). Läuft der Scanner, ist die Kachel offen.
  const [scannerOpen, setScannerOpen] = React.useState(false);
  const [selfOpen, setSelfOpen] = React.useState(false);


  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const context = (window as any).__dexSpfxContext;
  const eventService = React.useMemo(() => context ? new EventService(context) : null, []);

  const selectedEvent = selectedEventId ? events.find(e => e.id === selectedEventId) : null;

  /**
   * v30.53: Startnummer + Trikotgröße + Gruppe zu einer Zeile.
   *
   * Am Check-in-Tisch eines Laufs wird nicht nur abgehakt, sondern auch das
   * Trikot ausgegeben und die Startnummer genannt. Ohne diese Angaben muss
   * daneben eine Excel offen sein — genau dort entsteht die Schlange. Die
   * Regel, WELCHE Felder das sind, steht in utils/checkInExtras.
   */
  // v30.88: Trikot-Verteilung je Event — dieselbe Rechnung wie in der Aktion
  // „Benötigte T-Shirts" (utils/checkInExtras.shirtAllocate), gecacht je
  // Registrierungs-Array. Der Cache der Teilnehmerlisten wird weiter unten
  // deklariert; extrasFor liest ihn über eine Ref, damit die Closure nie einen
  // veralteten Stand sieht.
  const searchRegsCacheRef = React.useRef<Record<string, import('../services/EventService').SPRegistration[]>>({});
  const shirtAllocCacheRef = React.useRef<Map<object, ShirtAllocationResult>>(new Map());
  /**
   * v31.3: Bestand und Größenfeld fallen auf das Elternevent zurück.
   *
   * `ShirtSizeModal` speichert `_shirtStock` am gewählten Event — im Organizer
   * Center ist das bei einem Klammer-Event die Klammer. Am Lauftag steht das
   * Team aber auf dem TERMIN: Der Dialog meldete „Bestand gespeichert — die
   * Check-in-Seite zeigt Gegenvorschläge jetzt je Person an", und am Tisch kam
   * nie einer an. Dasselbe gilt für die Feld-Definition; das Größenfeld steht
   * oft nur im Formular der Klammer.
   *
   * Bewusst nur ein Rückgriff: Ein eigener Bestand am Termin gewinnt weiter
   * (jemand hat dort Kartons stehen), und die Antworten selbst werden NICHT
   * vom Elternevent nachgeladen — die Teilnehmerliste der Klammer ist eine
   * andere Liste (s. Bericht).
   */
  const shirtParentOf = React.useCallback((eventId: string) => {
    const ev = events.find(e => e.id === eventId);
    if (!ev || !ev.parentEventId) return undefined;
    return events.find(e => e.id === ev.parentEventId);
  }, [events]);
  const shirtFieldsFor = React.useCallback((eventId: string): Array<{ id: string; label: string }> => {
    const ev = events.find(e => e.id === eventId);
    const out: Array<{ id: string; label: string }> = ((ev && ev.eventSpecificFields) || []).slice();
    const parent = shirtParentOf(eventId);
    ((parent && parent.eventSpecificFields) || []).forEach(f => {
      if (!out.some(x => x.id === f.id)) out.push(f);
    });
    return out;
  }, [events, shirtParentOf]);
  const shirtAllocFor = React.useCallback((eventId: string): ShirtAllocationResult | null => {
    const ev = events.find(e => e.id === eventId);
    const rs = searchRegsCacheRef.current[eventId];
    if (!ev || !rs) return null;
    let stock = parseShirtStock(ev.emailTemplateOverrides);
    if (Object.keys(stock).length === 0) {
      const parent = shirtParentOf(eventId);
      if (parent) stock = parseShirtStock(parent.emailTemplateOverrides);
    }
    // v31.4: Ohne Bestand wurde hier früher abgebrochen — es gab ja nichts zu
    // verteilen. Seit die Ausgabe festgehalten wird, steht in der Verteilung
    // aber auch, WER sein Shirt schon hat und welche Größen es an diesem Event
    // überhaupt gibt. Beides braucht der Tisch auch ohne gepflegten Bestand;
    // `hasStock` bleibt false, es gibt also weiterhin keine Gegenvorschläge.
    const hit = shirtAllocCacheRef.current.get(rs);
    if (hit) return hit;
    const res = shirtAllocate(shirtFieldsFor(eventId), rs, stock);
    shirtAllocCacheRef.current.set(rs, res);
    return res;
  }, [events, shirtParentOf, shirtFieldsFor]);
  /**
   * v31.4: Welche Startnummer trägt noch den Namen einer ANDEREN Person?
   *
   * Eine in DEX frei gewordene Nummer ist immer eine gebrauchte: Sie wurde
   * beim Veranstalter für jemand anderen gedruckt, und die Ummeldung dort
   * ändert den Aufdruck nicht. Am Ausgabetisch muss deshalb jemand den Namen
   * überkleben — sonst läuft die Person mit einem fremden Namen auf der Brust
   * und wird bei der Zeitnahme der falschen Person zugeordnet.
   *
   * Quelle ist der bereits geladene `_b2runTodo`-Blob des Events (kein
   * zusätzlicher Lesevorgang am Tisch, wo das Netz am schlechtesten ist).
   * **Abgehakte Aufgaben zählen ausdrücklich MIT**: Abgehakt heißt „beim
   * Veranstalter umgemeldet", nicht „Aufdruck geändert" — der Zettel bleibt
   * falsch, bis jemand klebt. Deshalb wird `_b2runTodoDone` hier gar nicht
   * gelesen.
   *
   * Der Cache ist kein Luxus: `extrasFor` läuft je Zeile der Trefferliste, und
   * `emailTemplateOverrides` trägt bei Events mit eingebettetem Mail-Logo
   * mehrere hundert Kilobyte. Neu geparst wird nur, wenn sich der Blob ändert.
   */
  const bibRelabelCacheRef = React.useRef<Record<string, { src: string; map: Record<string, string> }>>({});
  const bibRelabelNames = React.useCallback((eventId: string): Record<string, string> => {
    const ev = events.find(e => e.id === eventId);
    const parent = shirtParentOf(eventId);
    const src = `${(ev && ev.emailTemplateOverrides) || ''} ${(parent && parent.emailTemplateOverrides) || ''}`;
    const hit = bibRelabelCacheRef.current[eventId];
    if (hit && hit.src === src) return hit.map;
    const map: Record<string, string> = {};
    const collect = (raw: string | undefined): void => {
      if (!raw) return;
      try {
        const o = JSON.parse(raw);
        const list = (o && Array.isArray(o._b2runTodo)) ? o._b2runTodo : [];
        for (const t of list) {
          if (!t || typeof t !== 'object') continue;
          if (t.kind !== 'transfer' && t.kind !== 'assign') continue;
          const bib = String(t.bib || '').trim();
          const from = String(t.fromName || '').trim();
          if (!bib || !from) continue;
          const key = bibKey(bib);
          if (key && !map[key]) map[key] = from;
        }
      } catch { /* kein oder kaputtes Piggyback — dann gibt es keinen Hinweis */ }
    };
    // Eigenes Event zuerst: Bei einem Klammer-Event kann die Aufgabenliste auf
    // der Klammer liegen, der Tisch steht aber auf dem Termin.
    collect(ev && ev.emailTemplateOverrides);
    collect(parent && parent.emailTemplateOverrides);
    bibRelabelCacheRef.current[eventId] = { src, map };
    return map;
  }, [events, shirtParentOf]);
  const extrasFor = React.useCallback((
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    reg: any,
    eventId: string,
  ): CheckInExtra[] => {
    // v31.4: Der Überkleben-Hinweis hängt an der Nummer AUF DER ZEILE — trägt
    // die Person keine, gibt es am Tisch auch nichts zu kleben.
    const bibRaw = String((reg && reg.Startnummer) || '').trim();
    const otherName = bibRaw ? (bibRelabelNames(eventId)[bibKey(bibRaw)] || '') : '';
    const out = checkInExtras(
      shirtFieldsFor(eventId),
      parseCustomData(reg?.CustomData),
      reg,
      { bib: isDe ? 'Startnummer' : 'Bib number', group: isDe ? 'Gruppe' : 'Group' },
      otherName
        ? {
          label: isDe ? 'Überkleben' : 'Relabel',
          text: isDe
            ? `Die Nummer läuft beim Veranstalter noch auf ${otherName} — bitte den Namen auf der Startnummer überkleben.`
            : `With the organiser this number is still registered to ${otherName} — please cover the name on the bib.`,
        }
        : null,
    );
    // v30.88: Gegenvorschlag, wenn die Wunschgröße laut Bestand nicht reicht —
    // die Antwort auf „passt es überhaupt?" gehört an den Tisch, nicht in eine Excel.
    const alloc = shirtAllocFor(eventId);
    const em = String((reg && reg.ParticipantEmail) || '').toLowerCase().trim();
    const a = alloc && em ? alloc.byEmail[em] : undefined;
    if (a && a.short) {
      // v31.3: Ein Satz statt eines Chips. Nutzer-Ansage (07.09.2026): „hier ist
      // eine Person, die ein T-Shirt hat, was es zu wenig gibt — also biete XX an
      // (das, was wir zu viel haben)." Der Helfer soll nicht rechnen müssen; die
      // Rest-Stückzahl kommt aus derselben Verteilung (`spare`), damit „übrig"
      // belegt ist und nicht geraten.
      const spare = (a.proposalKey && alloc)
        ? (alloc.rows.filter(r => r.key === a.proposalKey)[0] || { spare: 0 }).spare
        : 0;
      out.push({
        label: isDe ? 'Trikot' : 'Shirt',
        value: a.proposal
          ? (isDe
            ? `Wunschgröße ${a.wish} ist vergeben — bitte ${a.proposal} anbieten.${spare > 0 ? ` Davon sind nach der Planung noch ${spare} übrig.` : ' Dieses Stück ist für diese Person eingeplant.'}`
            : `Wished size ${a.wish} is gone — please offer ${a.proposal} instead.${spare > 0 ? ` ${spare} of those are still spare after planning.` : ' That one is reserved for this person.'}`)
          : (isDe
            ? `Wunschgröße ${a.wish} ist vergeben, und es ist auch keine Ausweichgröße mehr da — bitte am Ausgabetisch klären.`
            : `Wished size ${a.wish} is gone and there is no alternative left — please sort this out at the handout desk.`),
        tone: a.proposal ? 'warn' : 'danger',
      });
    }
    return out;
  }, [isDe, shirtAllocFor, shirtFieldsFor, bibRelabelNames]);

  /**
   * v31.4: Trikot-Angaben zu EINER Person am Ausgabetisch.
   *
   * Gibt `null` zurück, wenn das Event gar kein Größenfeld hat — dann gibt es
   * am Tisch auch nichts auszugeben, und weder Karte noch Zeile zeigen etwas.
   *
   * Die ausgegebene Größe wird notfalls aus der zwischengespeicherten
   * Teilnehmerliste nachgeschlagen: Der QR-Weg (`getRegistrationByEmail`)
   * liest mit einer festen Spaltenliste, die `ShirtIssued` bewusst NICHT
   * nennt — eine unbekannte Spalte im `$select` beantwortet SharePoint mit
   * einem Fehler auf die ganze Abfrage, und auf einer Bestandsliste ohne die
   * Spalte wäre damit der Scan tot. Der Cache liest über `$select=*` und hat
   * den Wert, sobald die Liste geladen ist.
   *
   * v31.4 (Review): Ohne geladene Liste gibt es deshalb GAR KEINE Ausgabe-UI.
   * Vorher zeigte die Karte nach einem Scan auf ein Event, dessen Liste nicht
   * im Cache liegt (Klammer gescannt, Termin unten gewählt), „Welches Trikot
   * gibst du aus?" mit leerem Feld — obwohl das Shirt morgens schon ausgegeben
   * und in `ShirtIssued` festgehalten war. Ein zweites Trikot aus dem Karton
   * ist teurer als ein fehlender Knopf; unbekannt sperrt, statt freizugeben.
   */
  const shirtDeskInfoFor = React.useCallback((
    reg: { Id?: number; ParticipantEmail?: string; ShirtIssued?: string } | null | undefined,
    eventId: string,
  ): ShirtDeskInfo | null => {
    if (!reg || !eventId) return null;
    const rs = searchRegsCacheRef.current[eventId];
    if (!rs) return null;
    if (!shirtFieldOf(shirtFieldsFor(eventId), rs)) return null;
    const alloc = shirtAllocFor(eventId);
    const em = (reg.ParticipantEmail || '').toLowerCase().trim();
    const a = alloc && em ? alloc.byEmail[em] : undefined;
    const cached = (rs || []).filter(x => (reg.Id !== undefined && x.Id === reg.Id) || (!!em && (x.ParticipantEmail || '').toLowerCase().trim() === em))[0];
    const issued = parseShirtIssue(reg.ShirtIssued || (cached ? cached.ShirtIssued : ''));
    const wish = a ? a.wish : '';
    return {
      wish,
      proposal: (a && a.proposal) || null,
      preset: (a && a.proposal) || (splitShirtSize(wish).isSize ? wish : ''),
      issued,
      sizes: alloc ? alloc.rows.map(r => r.size) : [],
    };
  }, [shirtAllocFor, shirtFieldsFor]);

  /**
   * v31.4: „QR-Nr. 012 · laufend 009" — leer, solange beide Nummern gleich sind.
   *
   * Der Helfer vergleicht mit dem Handy in der Hand; eine Zahl ohne Herkunft
   * hilft ihm nicht. Stimmen die beiden überein (der Normalfall vor der ersten
   * Abmeldung), wird bewusst NICHTS gezeigt — sonst steht an jeder Zeile eine
   * Angabe, die nichts unterscheidet, und die eine Zeile, die es tut, geht
   * darin unter.
   */
  const qrNoteOf = React.useCallback((reg: SPRegistration | null | undefined): string => {
    if (!reg) return '';
    const q = qrSentIdOf(reg);
    const t2 = runningIdOf(reg);
    if (q === null || t2 === null || q === t2) return '';
    return isDe ? `QR-Nr. ${pad3(q)} · laufend ${pad3(t2)}` : `QR no. ${pad3(q)} · current ${pad3(t2)}`;
  }, [isDe]);

  /**
   * v31.4 (Review): Dieselbe Angabe für eine Zeile, die NICHT aus der
   * Teilnehmerliste kommt.
   *
   * Der Scan-Weg liest über `getRegistrationByEmail`, und dessen fester
   * `$select` nennt weder `TeilnehmerID` noch `QrSentId` — auf der
   * Bestätigungskarte nach einem Scan lieferte `qrNoteOf` deshalb IMMER den
   * leeren String, obwohl der Kommentar dort das Gegenteil versprach. Die
   * Spalten in den `$select` aufzunehmen wäre der falsche Ort: Eine unbekannte
   * Spalte beantwortet SharePoint mit einem Fehler auf die GANZE Abfrage, und
   * damit wäre der Scan auf jeder Bestandsliste tot. Also derselbe Umweg wie
   * bei `ShirtIssued` — nachschlagen in der zwischengespeicherten Liste. Ist
   * sie nicht geladen, bleibt der Satz leer; geraten wird nichts.
   */
  const qrNoteFromCache = React.useCallback((
    reg: { Id?: number; ParticipantEmail?: string } | null | undefined,
    eventId: string,
  ): string => {
    if (!reg || !eventId) return '';
    const rs = searchRegsCacheRef.current[eventId];
    if (!rs) return '';
    const em = (reg.ParticipantEmail || '').toLowerCase().trim();
    const row = rs.filter(x => (reg.Id !== undefined && x.Id === reg.Id) || (!!em && (x.ParticipantEmail || '').toLowerCase().trim() === em))[0];
    return row ? qrNoteOf(row) : '';
  }, [qrNoteOf]);

  /**
   * v31.4 (Review): In der Trefferliste steht bei einer Zahlen-Suche an JEDER
   * Zeile, welche Nummer sie trägt — nicht nur an denen, deren Nummern
   * auseinanderlaufen.
   *
   * Seit die Suche auch über `QrSentId` filtert, können auf eine getippte Zahl
   * zwei Zeilen erscheinen: die eine, weil die Zahl in ihrer QR-Mail stand,
   * die andere, weil sie heute diese laufende Nummer trägt. Trug eine davon
   * gar keine Angabe (nie eine QR-Mail bekommen), sah der Helfer zwei Namen
   * ohne jeden Anhaltspunkt und klickte auf den, bei dem nichts von der
   * getippten Zahl abwich — die falsche Person. Ohne Zahlen-Suche bleibt es
   * bei der alten Regel: nur die Abweichung, sonst stünde an jeder Zeile eine
   * Angabe, die nichts unterscheidet.
   */
  const qrRowNote = React.useCallback((reg: SPRegistration, numericSearch: boolean): string => {
    if (!numericSearch) return qrNoteOf(reg);
    const q = qrSentIdOf(reg);
    const t2 = runningIdOf(reg);
    const parts: string[] = [];
    if (q !== null) parts.push(isDe ? `QR-Nr. ${pad3(q)}` : `QR no. ${pad3(q)}`);
    if (t2 !== null) parts.push(isDe ? `laufend ${pad3(t2)}` : `current ${pad3(t2)}`);
    if (parts.length === 0) return isDe ? 'keine Nummer hinterlegt' : 'no number on file';
    return parts.join(' · ');
  }, [isDe, qrNoteOf]);

  // v7.12: Name-Suche für manuelles Einchecken — wenn der QR-Scanner in der
  // SP-App nicht funktioniert (Camera-API gesperrt) oder der Teilnehmer den
  // QR-Code nicht zur Hand hat, kann der Helfer nach Namen / E-Mail suchen
  // und per Tap "Einchecken" auslösen. Die Registrierungen werden pro Event
  // lazy nachgeladen und in `searchRegsCache` zwischengespeichert.
  const [nameSearchQuery, setNameSearchQuery] = React.useState('');
  const [nameSearchEventId, setNameSearchEventId] = React.useState<string>(selectedEventId || '');
  const [searchRegsCache, setSearchRegsCache] = React.useState<Record<string, import('../services/EventService').SPRegistration[]>>({});
  searchRegsCacheRef.current = searchRegsCache; // v30.88 (s. shirtAllocFor)

  /**
   * v31.4: Was dieses Gerät gerade selbst geschrieben hat.
   *
   * Am Lauftag lädt die Liste alle 45 Sekunden nach (s. unten). SharePoint
   * antwortet auf ein GET direkt nach einem MERGE aber gelegentlich noch mit
   * dem alten Wert — ohne Schutz macht der Hintergrund-Lauf aus einer gerade
   * eingecheckten Person wieder eine offene. Das ist am Tisch der schlimmste
   * Fall: Der Helfer sieht „Angemeldet", checkt ein zweites Mal ein, und der
   * Kollege am anderen Tablet sucht die Person weiter.
   *
   * Deshalb wird jeder lokale Schreibvorgang gemerkt und nach jedem Laden für
   * 120 Sekunden wieder über die frische Zeile gelegt. Der Schlüssel enthält
   * die Event-Id: Item-Ids sind nur JE LISTE eindeutig, und ein Klammer-Event
   * hat so viele Listen wie Termine.
   *
   * `agenda` wird bewusst je Punkt gemerkt, nicht als fertiger JSON-String:
   * Setzt ein zweites Gerät in derselben Zeile einen ANDEREN Punkt, darf unser
   * Merker ihn nicht wieder wegräumen (dieselbe Lesen-Ändern-Schreiben-Falle
   * wie in `parseAgendaMarks`, v31.2).
   */
  type RegPatch = {
    status?: string;
    shirtIssued?: string;
    agenda?: Record<string, { at: string; by: string; noShow?: boolean } | null>;
  };
  const recentWritesRef = React.useRef<Record<string, { at: number; patch: RegPatch }>>({});
  const RECENT_WRITE_MS = 120000;
  const applyRegPatch = (row: SPRegistration, patch: RegPatch): SPRegistration => {
    let out = row;
    if (patch.status !== undefined) out = { ...out, Status: patch.status };
    if (patch.shirtIssued !== undefined) out = { ...out, ShirtIssued: patch.shirtIssued };
    if (patch.agenda) {
      const marks = parseAgendaMarks(out.AgendaCheckIns);
      const a = patch.agenda;
      Object.keys(a).forEach(pid => {
        const mark = a[pid];
        if (mark) marks[pid] = mark; else delete marks[pid];
      });
      out = { ...out, AgendaCheckIns: Object.keys(marks).length ? JSON.stringify(marks) : '' };
    }
    return out;
  };
  /** v31.4: Der EINE Weg, eine Zeile lokal zu ändern — Anzeige und Merker
   *  bleiben so zwangsläufig beieinander. Wer nur den State patcht, verliert
   *  seine Änderung beim nächsten Hintergrund-Lauf. */
  const patchCachedReg = (eventId: string, regId: number, patch: RegPatch): void => {
    if (!eventId || !regId) return;
    const key = `${eventId}|${regId}`;
    const known = recentWritesRef.current[key];
    const merged: RegPatch = known ? { ...known.patch, ...patch } : { ...patch };
    if (known && known.patch.agenda) merged.agenda = { ...known.patch.agenda, ...patch.agenda };
    recentWritesRef.current[key] = { at: Date.now(), patch: merged };
    setSearchRegsCache(prev => {
      const list = prev[eventId];
      if (!list) return prev;
      return { ...prev, [eventId]: list.map(x => x.Id === regId ? applyRegPatch(x, patch) : x) };
    });
  };
  /** v31.4: Frisch Geschriebenes über eine frisch geladene Liste legen. */
  const withRecentWrites = (eventId: string, rows: SPRegistration[]): SPRegistration[] => {
    const store = recentWritesRef.current;
    const now = Date.now();
    Object.keys(store).forEach(k => { if (now - store[k].at > RECENT_WRITE_MS) delete store[k]; });
    let touched = false;
    const out = rows.map(r => {
      const hit = store[`${eventId}|${r.Id}`];
      if (!hit) return r;
      touched = true;
      return applyRegPatch(r, hit.patch);
    });
    return touched ? out : rows;
  };

  // v30.91: Programmpunkte (Konzept docs/konzept-programmpunkte.md, Stufe 2).
  // Bei einem Event mit `agendaCheckIn` wird nicht das Event, sondern EIN
  // Programmpunkt eingecheckt: Das Team wählt oben den Punkt, jeder Scan,
  // jede ID und jeder Klick setzt nur diesen Punkt in `AgendaCheckIns`. Der
  // Event-Status bleibt unberührt (Nutzer-Entscheidung 07.09.2026).
  const agendaEv = React.useMemo(() => events.find(e => e.id === nameSearchEventId) || null, [events, nameSearchEventId]);
  const agendaItems = React.useMemo(() => (agendaEv && agendaEv.agendaCheckIn
    ? (agendaEv.agenda || []).slice().sort((a, b) => ((a.date || '') + (a.time || '')).localeCompare((b.date || '') + (b.time || '')))
    : []), [agendaEv]);
  const agendaMode = agendaItems.length > 0;
  const agendaTermSingular = (agendaEv && agendaEv.agendaTermSingular) || (isDe ? 'Programmpunkt' : 'Agenda item');
  const [agendaPointId, setAgendaPointId] = React.useState<string>('');
  React.useEffect(() => {
    if (!agendaMode) { setAgendaPointId(''); return; }
    let stored = '';
    try { stored = window.localStorage.getItem(`dex_checkin_point_${nameSearchEventId}`) || ''; } catch { /* */ }
    if (agendaItems.some(a => a.id === stored)) { setAgendaPointId(stored); return; }
    const s = suggestCurrentAgendaItem(agendaItems);
    setAgendaPointId(s ? s.id : '');
  }, [agendaMode, nameSearchEventId, agendaItems]);
  const choosePoint = (id: string): void => {
    setAgendaPointId(id);
    try { window.localStorage.setItem(`dex_checkin_point_${nameSearchEventId}`, id); } catch { /* */ }
  };
  const agendaPoint = agendaItems.find(a => a.id === agendaPointId) || null;
  const presentAt = (reg: { AgendaCheckIns?: string } | null | undefined, pointId: string): boolean =>
    !!(reg && pointId && parseAgendaCheckIns(reg.AgendaCheckIns)[pointId]);
  // Zähler je Punkt (aktive Anmeldungen) — für die Chips der Punkt-Wahl.
  const agendaCounts = React.useMemo((): Record<string, number> => {
    const out: Record<string, number> = {};
    if (!agendaMode) return out;
    const regs = searchRegsCache[nameSearchEventId] || [];
    for (const r of regs) {
      if (r.Status === 'Abgemeldet') continue;
      const marks = parseAgendaCheckIns(r.AgendaCheckIns);
      Object.keys(marks).forEach(k => { out[k] = (out[k] || 0) + 1; });
    }
    return out;
  }, [agendaMode, searchRegsCache, nameSearchEventId]);
  const [isLoadingSearchRegs, setIsLoadingSearchRegs] = React.useState(false);
  const [searchLoadError, setSearchLoadError] = React.useState('');
  // v31.4: Stand der Liste je Event (Zeitpunkt des letzten ERFOLGREICHEN
  // Ladens) plus die beiden Zustände des Nachladens. Bewusst getrennt von
  // `isLoadingSearchRegs`/`searchLoadError`: Die beiden gehören zur
  // Erstladung, bei der es noch keine Liste gibt — sie blenden die Liste aus.
  // Beim Nachladen darf genau das nicht passieren, die alte Liste ist das
  // Beste, was der Tisch hat (CLAUDE.md: ein Lesefehler ist keine Null).
  const [regsLoadedAt, setRegsLoadedAt] = React.useState<Record<string, number>>({});
  const [refreshBusy, setRefreshBusy] = React.useState(false);
  const [refreshError, setRefreshError] = React.useState('');
  const loadInFlightRef = React.useRef(false);
  // v31.4: Damit „Stand: 08:42" altern KANN. Scheitert der Hintergrund-Lauf,
  // ändert sich sonst gar nichts am State — die Uhrzeit bliebe grau und
  // harmlos, obwohl sie längst nicht mehr stimmt.
  const [nowTick, setNowTick] = React.useState<number>(Date.now());
  // v20.1: Busy-Flag für die Self-Check-in-Aktionen (Live-QR / PDF).
  const [selfCheckInBusy, setSelfCheckInBusy] = React.useState(false);
  React.useEffect(() => {
    if (selectedEventId && !nameSearchEventId) setNameSearchEventId(selectedEventId);
  }, [selectedEventId]); // eslint-disable-line react-hooks/exhaustive-deps

  /**
   * v31.4: Nachladen ist der Normalfall, nicht die Ausnahme.
   *
   * Befund 08.09.2026: „Der Aktualisieren-Button klappt nicht so gut. Wenn
   * jemand anderes jemanden eingecheckt hat, sieht man das erst, wenn man
   * wieder zurückgeht und wieder öffnet." Der Grund stand in der ersten Zeile
   * dieser Funktion: Bei gefülltem Cache kehrte sie sofort zurück — der Knopf
   * tat also NICHTS, und nur das Unmounten der Seite leerte den State.
   *
   * `force` überspringt die Cache-Prüfung, `silent` unterdrückt Spinner und
   * Fehlermeldung (für den Lauf alle 45 Sekunden). Scheitert ein Lauf, bleibt
   * die alte Liste stehen und der Stand-Zeitstempel wird NICHT erneuert — die
   * gealterte Uhrzeit ist die einzige ehrliche Aussage, wenn im Hintergrund
   * etwas klemmt.
   */
  const loadRegsForSearch = React.useCallback(async (
    eventId: string,
    opts?: { force?: boolean; silent?: boolean },
  ): Promise<void> => {
    if (!eventId) return;
    const force = !!(opts && opts.force);
    const silent = !!(opts && opts.silent);
    const hadList = !!searchRegsCacheRef.current[eventId];
    if (!force && hadList) return; // bereits geladen
    // Der Hintergrund-Lauf drängelt nicht: Läuft schon eine Abfrage, ist sein
    // Ergebnis in Sekunden ohnehin da. Ein Klick des Helfers läuft dagegen
    // immer — sonst ist der Knopf wieder der aus dem Befund.
    if (silent && loadInFlightRef.current) return;
    loadInFlightRef.current = true;
    if (!silent) {
      if (hadList) { setRefreshBusy(true); setRefreshError(''); }
      else { setIsLoadingSearchRegs(true); setSearchLoadError(''); }
    }
    try {
      // v30.67: `getAllRegistrations` wirft bei HTTP-Fehlern NICHT, sondern
      // liefert die bis dahin gelesenen Zeilen — bei 403/429 also `[]`. Das
      // landete hier als gültige, leere Liste im Cache: „Keine Teilnehmer",
      // „Keine Anmeldung mit der Teilnehmer-ID …", und weil der Cache-
      // Schlüssel gesetzt war, kein zweiter Versuch mehr bis zum Seiten-
      // Reload. Nur der Rückruf unterscheidet „leer" von „verboten"; ohne
      // geprüften Status kommt nichts in den Cache (CLAUDE.md, dritter Pfad).
      let readable = true;
      let httpStatus = 0;
      const regs = await getAllRegistrations(eventId, (status) => { readable = false; httpStatus = status; });
      if (!readable) {
        // v30.67 (Review): zweisprachig wie der Nachbarpfad checkInByParticipantId.
        const msg = httpStatus === 403
          ? (isDe
            ? 'Keine Leseberechtigung auf der Teilnehmerliste dieses Termins — bitte Organizer/Admin um Freigabe bitten.'
            : 'No read permission on this date\'s attendee list — please ask an organizer/admin for access.')
          : (isDe
            ? `Teilnehmerliste konnte nicht gelesen werden (${httpStatus ? 'HTTP ' + httpStatus : 'keine Teilnehmerliste gefunden'}) — bitte erneut versuchen.`
            : `The attendee list could not be read (${httpStatus ? 'HTTP ' + httpStatus : 'no attendee list found'}) — please try again.`);
        if (!silent) { if (hadList) setRefreshError(msg); else setSearchLoadError(msg); }
      } else {
        // v31.4: Frisch Geschriebenes gewinnt gegen eine Antwort, die es noch
        // nicht kennt — und der Trikot-Cache ist über die Array-Identität
        // geschlüsselt, die alten Einträge sind ab hier tote Last. Die
        // Verschmelzung läuft VOR `setSearchRegsCache`: Sie räumt abgelaufene
        // Merker weg, und ein Updater darf keine Nebenwirkung haben.
        const merged = withRecentWrites(eventId, regs);
        setSearchRegsCache(prev => ({ ...prev, [eventId]: merged }));
        setRegsLoadedAt(prev => ({ ...prev, [eventId]: Date.now() }));
        shirtAllocCacheRef.current.clear();
        if (!silent) setRefreshError('');
      }
    } catch {
      const msg = isDe ? 'Teilnehmerliste konnte nicht geladen werden.' : 'The attendee list could not be loaded.';
      if (!silent) { if (hadList) setRefreshError(msg); else setSearchLoadError(msg); }
    }
    if (!silent) { setIsLoadingSearchRegs(false); setRefreshBusy(false); }
    loadInFlightRef.current = false;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [getAllRegistrations, isDe]);
  // v31.4: Der Timer unten darf nicht an der Identität dieser Funktion hängen
  // — `getAllRegistrations` kommt aus dem EventContext und wird bei jedem
  // Render dort neu gebaut. Ein Effekt mit dieser Abhängigkeit würde sein
  // Intervall häufiger neu aufsetzen, als es feuert: Der Hintergrund-Lauf
  // käme dann nie zustande.
  const loadRegsRef = React.useRef(loadRegsForSearch);
  loadRegsRef.current = loadRegsForSearch;

  // v7.14: Sobald nameSearchEventId gesetzt ist, Teilnehmerliste vorab laden,
  // damit die Live-Liste ohne Vorab-Tippen sichtbar ist.
  React.useEffect(() => {
    if (nameSearchEventId && !searchRegsCache[nameSearchEventId]) {
      loadRegsForSearch(nameSearchEventId);
    } else if (searchLoadError) {
      // v30.67: Der Ladefehler gehört zu dem Event, bei dem er entstand —
      // beim Wechsel auf ein bereits geladenes Event darf er nicht stehen
      // bleiben (loadRegsForSearch kehrt bei Cache-Treffer vorher zurück).
      setSearchLoadError('');
    }
  }, [nameSearchEventId, searchRegsCache, loadRegsForSearch]);

  /**
   * v31.4: Alle 45 Sekunden nachladen, solange der Tisch hinschaut.
   *
   * Zwei Tablets an einem Eingang sind der Normalfall; wer auf Tablet 2 steht,
   * muss sehen, dass Tablet 1 gerade eingecheckt hat. `visibilityState` ist
   * die Bedingung, nicht der Komfort: Ein Tablet in der Tasche würde sonst
   * stundenlang gegen dieselbe Liste laufen, und Drosselung trifft am
   * Lauftag alle.
   */
  React.useEffect(() => {
    if (!nameSearchEventId) return undefined;
    const tick = (): void => {
      setNowTick(Date.now());
      if (typeof document !== 'undefined' && document.visibilityState !== 'visible') return;
      void loadRegsRef.current(nameSearchEventId, { force: true, silent: true });
    };
    const iv = window.setInterval(tick, 45000);
    // Zurück aus dem Sperrbildschirm heißt: Der Helfer will den Stand von
    // JETZT, nicht den von vor 44 Sekunden.
    const onVisible = (): void => { if (document.visibilityState === 'visible') tick(); };
    document.addEventListener('visibilitychange', onVisible);
    return () => {
      window.clearInterval(iv);
      document.removeEventListener('visibilitychange', onVisible);
    };
  }, [nameSearchEventId]);

  /**
   * v31.4: „Letzte Check-ins" überleben das Verlassen der Seite.
   *
   * Befund 08.09.2026: „Wenn man bei Check-in ist, dann zurück, dann sieht man
   * nicht mehr seine letzten Check-ins." Die Liste war reiner Komponenten-
   * State; jede Navigation unmountete die Seite — und mit der Liste war auch
   * das Rückgängigmachen weg, obwohl jeder Eintrag alles trägt, was der Revert
   * braucht (`regId`, `subsiteUrl`, `prevStatus`, `agendaItemId`, `kind`).
   *
   * Gespeichert wird je Event und nur für eine Schicht: Ein Lauftag ist keine
   * Woche, und ein Eintrag von gestern lädt nur dazu ein, den falschen
   * Check-in zurückzunehmen. Jeder Zugriff in try/catch — im privaten Fenster
   * wirft schon das Lesen.
   */
  const RECENT_MAX = 30;
  const RECENT_MAX_AGE_MS = 12 * 60 * 60 * 1000;
  const recentKey = (eventId: string): string => `dex_recent_checkins_v1_${eventId}`;
  const readRecent = (eventId: string): RecentCheckIn[] => {
    try {
      const raw = window.localStorage.getItem(recentKey(eventId));
      if (!raw) return [];
      const arr = JSON.parse(raw) as RecentCheckIn[];
      if (!Array.isArray(arr)) return [];
      const now = Date.now();
      return arr
        .filter(e => !!e && typeof e.key === 'string' && typeof e.at === 'string' && !!e.regId
          && (now - Date.parse(e.at)) < RECENT_MAX_AGE_MS)
        .slice(0, RECENT_MAX);
    } catch { return []; }
  };
  // Spiegel des States, damit `setRecent` ohne Updater-Funktion auskommt —
  // sonst müsste der localStorage-Schreibvorgang IN den Updater, und der darf
  // in React keine Nebenwirkung haben.
  const recentRef = React.useRef<RecentCheckIn[]>([]);
  recentRef.current = recentCheckIns;
  const setRecent = (next: RecentCheckIn[]): void => {
    const list = next.slice(0, RECENT_MAX);
    recentRef.current = list;
    setRecentCheckIns(list);
    if (!nameSearchEventId) return;
    try { window.localStorage.setItem(recentKey(nameSearchEventId), JSON.stringify(list)); } catch { /* voll oder gesperrt */ }
  };
  React.useEffect(() => {
    setRecentCheckIns(nameSearchEventId ? readRecent(nameSearchEventId) : []);
    // v31.4: Der Nachlade-Fehler gehört zu dem Event, bei dem er entstand —
    // dieselbe Regel wie für `searchLoadError` seit v30.67.
    setRefreshError('');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nameSearchEventId]);

  // v7.16: Quick-Filter "Nur offene" — wenn aktiv, werden Eingecheckte,
  // Wartelistler und Abgemeldete aus der Liste ausgeblendet, sodass der
  // Helfer am Eingang nur die noch offenen Anmeldungen sieht.
  const [onlyOpen, setOnlyOpen] = React.useState(false);

  // KPI-Zähler: angemeldet (= Status Angemeldet/QR versendet/Eingecheckt)
  // und eingecheckt. Wird im KPI-Bereich der Check-In-Page angezeigt.
  const checkInKpis = React.useMemo(() => {
    if (!nameSearchEventId) return { registered: 0, checkedIn: 0, noShow: 0 };
    const regs = searchRegsCache[nameSearchEventId] || [];
    let registered = 0;
    let checkedIn = 0;
    let noShow = 0;
    for (const r of regs) {
      if (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt') registered++;
      // v30.91: Im Programmpunkt-Modus zählt die Kachel die Anwesenden am
      // gewählten Punkt — der Event-Status sagt dort nichts.
      if (agendaMode ? presentAt(r, agendaPointId) : r.Status === 'Eingecheckt') checkedIn++;
      // v31.2: Im Programmpunkt-Modus zählt die Kachel die No-Shows AM PUNKT
      // (Marke), sonst den Event-Status.
      if (agendaMode && agendaPointId ? !!parseAgendaNoShows(r.AgendaCheckIns)[agendaPointId] : r.Status === 'No-Show') noShow++;
    }
    return { registered, checkedIn, noShow };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nameSearchEventId, searchRegsCache, agendaMode, agendaPointId]);

  // v7.14: Live-Filter über die ganze Liste — leerer Query zeigt alle
  // Teilnehmer. Sortiert nach Status (Aktive zuerst), dann Nachname.
  const searchHits = React.useMemo(() => {
    if (!nameSearchEventId) return [];
    const regs = searchRegsCache[nameSearchEventId] || [];
    const q = nameSearchQuery.trim().toLowerCase();
    // v30.33: Die Teilnehmer-ID ist jetzt suchbar — und zwar EXAKT, nicht als
    // Teiltreffer. Sie steht schon heute unter jedem QR-Code in der Mail; sie
    // eintippen zu können macht den Check-in unabhängig von der Kamera, die auf
    // verwalteten Geräten nicht überall erreichbar ist.
    //
    // Exakt deshalb, weil ein Teiltreffer bei „17" auch 117 und 170 liefern
    // würde — am Einlass die falsche Person einzuchecken ist schlimmer als
    // einmal mehr zu tippen. Die Namens-/Mail-Suche bleibt zusätzlich als
    // Teiltreffer bestehen, damit „17" auch eine Mail mit 17 darin findet.
    //
    // Eindeutig ist die Zahl, weil der Check-in immer AUF EIN EVENT bezogen
    // ist (Event-Picker davor) und TeilnehmerID je Event fortlaufend vergeben
    // wird. Ein Event-Präfix braucht es hier deshalb nicht.
    // v30.83: numerisch vergleichen — „005" (so steht die Nummer mit
    // führenden Nullen in der QR-Mail) ist dieselbe ID wie 5. Der String-
    // Vergleich lieferte „Kein Treffer" (Befund 07.09.2026).
    // v31.4: Die Nummer aus der QR-Mail (`QrSentId`) zählt genauso wie die
    // laufende. Sonst findet die Liste unten die Person NICHT, die oben über
    // die Mail-Nummer gesucht wird — und der Helfer glaubt, sie sei gar nicht
    // angemeldet. Beide Treffer sind gewollt: Die Zahl kann zu zwei Zeilen
    // gehören, und dann sollen beide dastehen.
    const numericQ = /^\d+$/.test(q) ? parseInt(q, 10) : NaN;
    const matchesQuery = q.length === 0
      ? regs
      : regs.filter(r => {
          if (isFinite(numericQ) && (runningIdOf(r) === numericQ || qrSentIdOf(r) === numericQ)) return true;
          const full = `${r.Vorname || ''} ${r.Nachname || ''} ${r.ParticipantName || ''} ${r.ParticipantEmail || ''}`.toLowerCase();
          return full.indexOf(q) >= 0;
        });
    // v7.16: optionaler Quick-Filter "nur offene Anmeldungen"
    const filtered = onlyOpen
      ? matchesQuery.filter(r => agendaMode
        ? (r.Status !== 'Abgemeldet' && r.Status !== 'Warteliste' && !presentAt(r, agendaPointId))
        : (r.Status === 'Angemeldet' || r.Status === 'QR versendet'))
      : matchesQuery;
    // Sortierung: Angemeldet/QR versendet zuerst, dann Eingecheckt, dann
    // Warteliste, dann Abgemeldet. Innerhalb der Gruppe alphabetisch nach
    // Nachname.
    const statusRank = (s: string): number => {
      if (s === 'Angemeldet' || s === 'QR versendet') return 0;
      if (s === 'Eingecheckt') return 1;
      if (s === 'Warteliste') return 2;
      return 3; // Abgemeldet & Sonstige
    };
    return filtered.slice().sort((a, b) => {
      const sa = statusRank(a.Status);
      const sb = statusRank(b.Status);
      if (sa !== sb) return sa - sb;
      const na = (a.Nachname || a.ParticipantName || '').toLowerCase();
      const nb = (b.Nachname || b.ParticipantName || '').toLowerCase();
      return na.localeCompare(nb);
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nameSearchQuery, nameSearchEventId, searchRegsCache, onlyOpen, agendaMode, agendaPointId]);
  // v31.4 (Review): Wird nach einer ZAHL gesucht, zeigt jede Trefferzeile,
  // welche Nummer sie trägt — sonst stehen zwei Namen da und nur einer nennt
  // eine Zahl (s. qrRowNote).
  const numericSearch = /^\d+$/.test(nameSearchQuery.trim());

  /**
   * v30.35: Check-in über die eingetippte Teilnehmer-ID.
   *
   * Nutzt bewusst denselben Weg wie ein Klick in der Trefferliste
   * (`startManualCheckInFromSearch`) — also dieselbe Bestätigungskarte mit
   * Foto und Status. Am Einlass will man vor dem Einchecken sehen, WEN man
   * eincheckt; eine ID ohne Gesicht wäre schneller und riskanter.
   *
   * Die Liste ist zu diesem Zeitpunkt geladen (`loadRegsForSearch` läuft beim
   * Auswählen des Events). Ist sie es nicht, sagt die Meldung genau das,
   * statt „ID nicht gefunden" zu behaupten.
   *
   * v31.4: Die getippte Zahl wird ZUERST über die versendete QR-Mail
   * aufgelöst (`QrSentId`), erst danach über die laufende `TeilnehmerID`.
   *
   * Grund (Befund 08.09.2026, laufendes Event): Jede Abmeldung nummeriert die
   * ganze Liste neu, die gedruckte Mail nicht. Nach der ersten Abmeldung
   * zeigte die Mail von Person A auf die Zahl, die inzwischen Person B trägt —
   * der Tisch tippte sie und checkte B ein. Die Mail ist die Zusage, die die
   * Person in der Hand hält; sie gewinnt. Der Scan-Weg ist davon nie betroffen
   * (im Code steht die E-Mail-Adresse, keine Nummer).
   *
   * v31.4 (Review): Die Rangfolge ist damit
   *   1. aktive Zeile mit dieser `QrSentId` — direkt einchecken,
   *   2. …plus eine zweite aktive Zeile mit dieser laufenden Nummer → Karte,
   *   3. nur eine aktive laufende Nummer → Karte, sobald das Event QR-Nummern
   *      führt oder eine abgemeldete Zeile dieselbe Zahl trägt,
   *   4. sonst wie bisher direkt einchecken.
   * Abgemeldete Zeilen lösen die Zahl NIE auf (s. Kommentar unten).
   */
  const checkInByParticipantId = async (): Promise<void> => {
    const raw = idInput.trim();
    if (!raw) return;
    setIdError('');
    if (!nameSearchEventId) {
      setIdError(isDe ? 'Bitte zuerst oben das Event auswählen.' : 'Please pick the event above first.');
      return;
    }
    const regs = searchRegsCache[nameSearchEventId];
    if (!regs) {
      // v30.67: Ohne Cache-Eintrag gibt es zwei Zustände — „lädt noch" und
      // „konnte nicht gelesen werden" (403/429, s. loadRegsForSearch). Beide
      // erlauben KEINE Aussage über die ID; der zweite darf aber nicht als
      // Warten verkauft werden, das nie endet.
      setIdError(searchLoadError
        ? (isDe
          ? `Die Teilnehmerliste konnte nicht gelesen werden — ob die ID ${raw} angemeldet ist, lässt sich so nicht sagen. ${searchLoadError}`
          : `The attendee list could not be read — whether ID ${raw} is registered cannot be told. Please retry or ask an organizer/admin for access.`)
        : (isDe ? 'Teilnehmerliste wird noch geladen — bitte kurz warten und erneut versuchen.' : 'Attendee list is still loading — please wait a moment and try again.'));
      void loadRegsForSearch(nameSearchEventId);
      return;
    }
    // v30.83: numerisch — „005" aus der QR-Mail ist ID 5.
    const rawNum = parseInt(raw, 10);
    /**
     * v31.4 (Review): Eine ABGEMELDETE Zeile löst die Zahl nie auf.
     *
     * `cancelRegistration` setzt `TeilnehmerID` auf null, lässt `QrSentId`
     * aber stehen (richtig so — die Mail existiert weiter). Eine abgemeldete
     * Zeile kann deshalb nie in `tidHits` landen, wohl aber in `qrHits`, und
     * `qrHits.length === 1` gewann vor jeder Statusprüfung. Ergebnis am Tisch:
     * `startManualCheckInFromSearch` brach mit „A — Anmeldung storniert" ab,
     * BEVOR die Konfliktkarte gebaut wurde — die aktive Person B, die heute
     * genau diese laufende Nummer trägt, war nicht mehr erreichbar, und der
     * Alternativ-Knopf existierte ausgerechnet in dem Fall nicht, für den er
     * gedacht war. Vor v31.4 fand derselbe Tastendruck B über `tidHits`.
     *
     * Also: Nur einbuchbare Zeilen lösen auf. Die abgemeldete verschwindet
     * dabei nicht, sie wird BENANNT — sie erklärt, warum die vorgelesene Zahl
     * heute auf jemand anderen zeigt.
     */
    const isCancelled = (r: SPRegistration): boolean => r.Status === 'Abgemeldet';
    const qrAll = regs.filter(r => qrSentIdOf(r) === rawNum);
    const tidAll = regs.filter(r => runningIdOf(r) === rawNum);
    const qrHits = qrAll.filter(r => !isCancelled(r));
    const tidHits = tidAll.filter(r => !isCancelled(r));
    const cancelledHit = qrAll.filter(isCancelled)[0] || tidAll.filter(isCancelled)[0];
    const cancelledName = cancelledHit ? regDisplayName(cancelledHit) : undefined;
    // Hat dieses Event überhaupt hinterlegt, welche Nummern verschickt wurden?
    // Ohne diese Unterscheidung klingt jede Meldung gleich — dabei ist „nichts
    // hinterlegt" ein ganz anderer Zustand als „diese Nummer stand in keiner
    // Mail", und nur der erste hat eine Abhilfe. Abgemeldete Zeilen zählen
    // hier nicht mit: Für die Leute, die heute vor dem Tisch stehen, ist dann
    // eben nichts hinterlegt.
    const eventHasQrIds = regs.some(r => !isCancelled(r) && qrSentIdOf(r) !== null);

    if (qrHits.length > 1) {
      setIdError(isDe
        ? `Zwei aktive Anmeldungen tragen die Nummer ${raw} aus einer versendeten QR-Mail — das ist ein Datenfehler, keine Verwechslung am Tisch. Bitte unten über den Namen einchecken und das den DEX-Admins melden.`
        : `Two active registrations carry number ${raw} from a sent QR email — that is a data error, not a mix-up at the desk. Please check in by name below and report this to the DEX admins.`);
      return;
    }
    if (qrHits.length === 1) {
      const person = qrHits[0];
      // Trägt eine ANDERE aktive Zeile heute dieselbe laufende Nummer, ist die
      // Eingabe zweideutig — dann zeigt die Karte beide Namen (s. QrIdConflict).
      const other = tidHits.filter(r => r.Id !== person.Id)[0];
      setIdInput('');
      // v31.4 (Review): Der Live-Filter bleibt stehen, solange die Karte eine
      // Entscheidung verlangt — sonst stehen beide Kandidaten zur Auswahl,
      // aber keiner mehr in der Liste darunter.
      if (!other) setNameSearchQuery('');
      startManualCheckInFromSearch(person, other
        ? { typed: rawNum, qrName: regDisplayName(person), altName: regDisplayName(other), alt: other }
        : undefined);
      return;
    }
    if (tidHits.length > 1) {
      // Sollte nicht vorkommen (ID ist je Event fortlaufend), wäre aber ein
      // Datenfehler, den man am Einlass nicht stillschweigend raten darf.
      setIdError(isDe
        ? `Mehrere aktive Anmeldungen mit der ID ${raw} gefunden — bitte unten über den Namen einchecken und das den DEX-Admins melden.`
        : `Several active registrations share ID ${raw} — please check in by name below and report this to the DEX admins.`);
      return;
    }
    if (tidHits.length === 0) {
      // v31.4: Ohne EINE hinterlegte QR-Nummer im ganzen Event ist die
      // wahrscheinlichste Ursache nicht die Eingabe, sondern der fehlende
      // Datensatz — und dafür gibt es genau eine Abhilfe.
      const backfillHint = eventHasQrIds ? '' : (isDe
        ? ' Für dieses Event ist noch nicht hinterlegt, welche Nummern in den QR-Mails standen — ein Organizer kann das im Organizer Center über „QR-Nummern nachtragen" nachziehen.'
        : ' For this event it is not yet on file which numbers were printed in the QR emails — an organizer can add them in the organizer center via “Backfill QR numbers”.');
      // v31.4 (Review): Trägt eine abgemeldete Zeile die Zahl, ist das die
      // Erklärung — und sie gehört in die Meldung, sonst sucht der Tisch einen
      // Tippfehler, den es nicht gibt.
      const cancelHint = cancelledName ? (isDe
        ? ` Die Nummer ${raw} gehört zur stornierten Anmeldung von ${cancelledName} — abgemeldete Zeilen lassen sich nicht einchecken.`
        : ` Number ${raw} belongs to the cancelled registration of ${cancelledName} — cancelled rows cannot be checked in.`) : '';
      setIdError((isDe
        ? `Keine aktive Anmeldung mit der Teilnehmer-ID ${raw} bei diesem Event. Bitte die Nummer aus der QR-Mail prüfen — oder unten nach dem Namen suchen.`
        : `No active registration with attendee ID ${raw} for this event. Please check the number in the QR email — or search by name below.`) + cancelHint + backfillHint);
      return;
    }
    // Genau eine aktive laufende Nummer, und keine versendete QR-Mail dazu.
    const viaRunning = tidHits[0];
    setIdInput('');
    if (eventHasQrIds || cancelledName) {
      /**
       * v31.4 (Review): Hier wird NICHT mehr direkt eingecheckt.
       *
       * Das Event hat nachweislich hinterlegt, welche Nummern gedruckt wurden
       * — und die getippte Zahl steht in keiner dieser Mails. Die
       * wahrscheinlichste Ursache ist dann kein Tippfehler, sondern eine
       * Person, deren `QrSentId` nicht erfasst werden konnte (Mail nicht
       * parsbar, Versand vor v30.35, Status außerhalb der Nachtrag-Stati).
       * Ihre gedruckte Nummer wird auf die HEUTIGE laufende Nummer einer
       * anderen Person aufgelöst — genau der Fehler, den v31.4 abschafft.
       * Der Hinweis stand vorher unter dem Eingabefeld, und zwar NACH dem
       * Schreibvorgang; er kam damit zu spät, um noch etwas zu ändern.
       */
      startManualCheckInFromSearch(viaRunning, { typed: rawNum, viaRunning: eventHasQrIds, cancelledName });
      return;
    }
    setNameSearchQuery(''); // v30.87: Live-Filter der Liste zurücksetzen
    startManualCheckInFromSearch(viaRunning);
  };

  /**
   * @param forEventId v31.4 (Review): Das Event, zu dem `reg` gehört — nicht
   * das, das gerade unten im Auswahlfeld steht.
   *
   * `reg.Id` ist eine Listen-Item-Id, und die ist nur JE TEILNEHMERLISTE
   * eindeutig. Der Alternativ-Knopf der Konfliktkarte reicht eine Zeile
   * durch, die beim Öffnen der Karte gelesen wurde; wechselt der Helfer
   * zwischendurch das Event, zeigte dieselbe Id auf der anderen Subsite auf
   * eine völlig andere Person — und `checkInParticipant` schrieb genau
   * dorthin. Die Karte wird beim Event-Wechsel inzwischen geleert; dieses
   * Argument ist der Riegel für den Rest.
   */
  const startManualCheckInFromSearch = (reg: SPRegistration, qrConflict?: QrIdConflict, forEventId?: string): void => {
    const ev = events.find(e => e.id === (forEventId || nameSearchEventId));
    if (!ev || !ev.subsiteUrl) return;
    if (reg.Status === 'Abgemeldet') {
      setResultMessage(`${reg.ParticipantName || reg.ParticipantEmail} — ${t('checkin.cancelled')}`);
      setResultType('error');
      return;
    }
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || reg.ParticipantEmail);
    // v31.4 (Review): Der Programmpunkt wird UNTEN gewählt und gehört damit zu
    // `nameSearchEventId`. Gilt der Aufruf einem anderen Event, passt weder
    // der gewählte Punkt noch seine Abwesenheit — dann lieber nichts
    // schreiben und es sagen.
    const evAgendaMode = agendaMode && ev.id === nameSearchEventId;
    if (!evAgendaMode && ev.agendaCheckIn && (ev.agenda || []).length > 0) {
      setResultMessage(isDe
        ? `„${ev.title}" hat Programmpunkte — bitte unten dieses Event und den Punkt wählen, an dem eingecheckt wird.`
        : `“${ev.title}” has agenda items — please pick this event and the item below first.`);
      setResultType('error');
      return;
    }
    // v30.91: Programmpunkt-Modus — ohne gewählten Punkt kein Check-in;
    // schon erfasst → Hinweis mit Uhrzeit, kein zweiter Schreibvorgang.
    if (evAgendaMode) {
      if (!agendaPoint) {
        setResultMessage(isDe ? `Bitte oben den ${agendaTermSingular} wählen, an dem eingecheckt wird.` : `Please pick the ${agendaTermSingular.toLowerCase()} above first.`);
        setResultType('error');
        return;
      }
      const marks = parseAgendaCheckIns(reg.AgendaCheckIns);
      if (marks[agendaPoint.id]) {
        setResultMessage(`${name} — ${isDe ? 'bereits erfasst' : 'already recorded'} (${agendaPoint.title}, ${formatMarkTime(marks[agendaPoint.id].at)})`);
        setResultType('info');
        return;
      }
    }
    let photoUrl = '';
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (ctx) {
        const siteBase = ctx.pageContext.web.absoluteUrl;
        photoUrl = `${siteBase}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(reg.ParticipantEmail || '')}`;
      }
    } catch { /* */ }
    // v31.1: Liste und Teilnehmer-ID checken DIREKT ein — die Person ist hier
    // schon eindeutig gewählt, eine zweite Bestätigung war nur ein Klick mehr.
    // v31.4: EINE Ausnahme — die getippte Zahl war zweideutig (`qrConflict`).
    // Dann greift die v31.1-Begründung nicht: Es sind zwei Personen im Spiel,
    // und wer eincheckt, soll vorher beide Namen gesehen haben.
    const info: PendingCheckInInfo = {
      name,
      email: reg.ParticipantEmail || '',
      event: { id: ev.id, subsiteUrl: ev.subsiteUrl, title: ev.title },
      regId: reg.Id,
      status: reg.Status,
      agendaItemId: evAgendaMode && agendaPoint ? agendaPoint.id : undefined,
      agendaLabel: evAgendaMode && agendaPoint ? agendaPoint.title : undefined,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      department: (reg as any).Department || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      jobTitle: (reg as any).JobTitle || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      location: (reg as any).Location || '',
      photoUrl,
      extras: extrasFor(reg, ev.id),
      qrNote: qrNoteOf(reg),
      qrConflict,
    };
    setResultMessage('');
    setResultType('');
    if (qrConflict) {
      const shirt = shirtDeskInfoFor(reg, ev.id);
      setCardShirtSize(shirt ? shirt.preset : '');
      setPendingCheckIn({ ...info, shirt });
      setTimeout(() => { confirmCardRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' }); }, 100);
      return;
    }
    // v31.4 (Review): Der Live-Filter wird erst hier zurückgesetzt — bei einer
    // offenen Karte bleiben die Kandidaten in der Liste darunter stehen.
    setNameSearchQuery('');
    setIsProcessing(true);
    void performCheckIn(info).finally(() => setIsProcessing(false));
  };

  // v23.28: Teilnehmer als „No-Show" markieren (nicht erschienen). Direkt aus
  // der Suchliste; nach Bestätigung wird der lokale Cache aktualisiert.
  // v31.2: Der No-Show landet in „Letzte Check-ins" und ist dort rücknehmbar.
  const rememberNoShow = (reg: import('../services/EventService').SPRegistration, name: string, ev: { id: string; subsiteUrl?: string }, agendaItemId?: string, agendaLabel?: string): void => {
    const entry: RecentCheckIn = {
      key: `${reg.Id}:${agendaItemId || 'status'}:noshow:${Date.now()}`,
      at: new Date().toISOString(), name, regId: reg.Id, eventId: ev.id || '', subsiteUrl: ev.subsiteUrl || '',
      prevStatus: reg.Status, agendaItemId, agendaLabel, kind: 'noshow',
    };
    setRecent([entry, ...recentRef.current]);
  };
  // v31.2: No-Show NUR am gewählten Programmpunkt — Marke mit noShow, der
  // Event-Status bleibt (die Person kann beim nächsten Punkt wieder da sein).
  const applyPointNoShow = async (reg: import('../services/EventService').SPRegistration, name: string): Promise<void> => {
    const ev = events.find(e => e.id === nameSearchEventId);
    if (!ev || !ev.subsiteUrl || !eventService || !agendaPoint) return;
    const pointId = agendaPoint.id;
    const r = await eventService.markAgendaNoShow(ev.subsiteUrl, reg.Id, pointId);
    if (!r.ok) {
      setResultMessage(isDe
        ? `${name} — No-Show konnte nicht gespeichert werden${r.status ? ` (HTTP ${r.status})` : ''}. ${r.status === 400 ? 'Fehlt die Spalte AgendaCheckIns? Organizer: „Spalten fixen" ausführen.' : 'Bitte erneut versuchen.'}`
        : `${name} — no-show could not be saved${r.status ? ` (HTTP ${r.status})` : ''}.`);
      setResultType('error');
      return;
    }
    const at = r.at || new Date().toISOString();
    patchCachedReg(nameSearchEventId, reg.Id, { agenda: { [pointId]: { at, by: '', noShow: true } } });
    rememberNoShow(reg, name, ev, pointId, agendaPoint.title);
    setResultMessage(isDe ? `${name} — No-Show bei ${agendaPoint.title}.` : `${name} — no-show at ${agendaPoint.title}.`);
    setResultType('info');
  };
  const applyEventNoShow = async (reg: import('../services/EventService').SPRegistration, name: string): Promise<void> => {
    const ev = events.find(e => e.id === nameSearchEventId);
    if (!ev || !ev.subsiteUrl || !eventService) return;
    const success = await eventService.markNoShowParticipant(ev.subsiteUrl, reg.Id);
    if (success) {
      patchCachedReg(nameSearchEventId, reg.Id, { status: 'No-Show' });
      rememberNoShow(reg, name, ev);
      setResultMessage(isDe ? `${name} — als No-Show markiert.` : `${name} — marked as no-show.`);
      setResultType('info');
    } else {
      // v23.29: No-Show gibt es nur für Events, die ab v23.28 NEU angelegt
      // wurden — Bestands-Events kennen die Choice nicht (HTTP 400).
      setResultMessage(isDe
        ? `${name} — No-Show ist für dieses Event nicht verfügbar (nur für neu angelegte Events).`
        : `${name} — no-show is not available for this event (only for newly created events).`);
      setResultType('error');
    }
  };
  const markNoShowFromSearch = async (reg: import('../services/EventService').SPRegistration): Promise<void> => {
    const ev = events.find(e => e.id === nameSearchEventId);
    if (!ev || !ev.subsiteUrl || !eventService) return;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || reg.ParticipantEmail || '-');
    // v31.2: Mit gewähltem Programmpunkt erst fragen — Punkt oder Event?
    if (agendaMode && agendaPoint) { setNoShowAsk({ reg, name }); return; }
    const ok = await confirmDialog(
      isDe
        ? `„${name}" als nicht erschienen (No-Show) markieren?`
        : `Mark „${name}" as a no-show?`,
      { confirmLabel: isDe ? 'Als No-Show markieren' : 'Mark as no-show' }
    );
    if (!ok) return;
    await applyEventNoShow(reg, name);
  };

  /**
   * v31.4: Trikot-Ausgabe festhalten.
   *
   * Ein 400 heißt hier IMMER dasselbe: Auf dieser Teilnehmerliste fehlt die
   * Spalte `ShirtIssued` (Bestands-Event, nie „Spalten fixen" gelaufen). Das
   * ist keine „hat nicht geklappt"-Meldung, sondern eine mit genau einer
   * Abhilfe — und die gehört in den Satz, sonst probiert das Team es am
   * Lauftag zehnmal.
   */
  const shirtFailMsg = (name: string, status: number, undo: boolean): string => {
    const fix = status === 400
      ? (isDe
        ? 'Auf dieser Teilnehmerliste fehlt die Spalte ShirtIssued — ein Organizer führt im Organizer Center einmal „Spalten fixen" aus, danach klappt es.'
        : 'This attendee list is missing the ShirtIssued column — an organizer runs "Fix columns" in the organizer center once, then it works.')
      : (isDe ? 'Bitte erneut versuchen.' : 'Please try again.');
    const what = isDe
      ? (undo ? 'die Rücknahme der Trikot-Ausgabe' : 'die Trikot-Ausgabe')
      : (undo ? 'undoing the shirt handout' : 'the shirt handout');
    return isDe
      ? `${name} — ${what} konnte nicht gespeichert werden${status ? ` (HTTP ${status})` : ''}. ${fix}`
      : `${name} — ${what} could not be saved${status ? ` (HTTP ${status})` : ''}. ${fix}`;
  };
  /** v31.4: Zeile schreiben, Cache-Zeile ersetzen, Merker setzen. Der
   *  Cache-Patch ist Pflicht: Die Verteilung rechnet aus genau diesen Zeilen,
   *  und ohne den Patch verplant sie das Stück ein zweites Mal (dasselbe
   *  Muster wie `applyPointNoShow`). */
  const applyShirtIssue = async (
    target: { regId: number; name: string; eventId: string; subsiteUrl: string; prevStatus: string },
    size: string,
  ): Promise<ShirtIssue | null> => {
    const clean = (size || '').trim();
    if (!eventService || !clean || !target.subsiteUrl) return null;
    setShirtBusy(true);
    try {
      const r = await eventService.setShirtIssued(target.subsiteUrl, target.regId, clean);
      if (!r.ok) {
        setResultMessage(shirtFailMsg(target.name, r.status, false));
        setResultType('error');
        return null;
      }
      const issue: ShirtIssue = { size: clean, at: r.at || new Date().toISOString(), by: currentEmailLc };
      patchCachedReg(target.eventId, target.regId, { shirtIssued: JSON.stringify(issue) });
      const entry: RecentCheckIn = {
        key: `${target.regId}:shirt:${Date.now()}`,
        at: issue.at, name: target.name, regId: target.regId,
        eventId: target.eventId, subsiteUrl: target.subsiteUrl,
        prevStatus: target.prevStatus, kind: 'shirt', shirtSize: clean,
      };
      setRecent([entry, ...recentRef.current]);
      setResultMessage(isDe
        ? `${target.name} — Trikot ${clean} ausgegeben.`
        : `${target.name} — shirt ${clean} handed out.`);
      setResultType('success');
      return issue;
    } finally { setShirtBusy(false); }
  };
  /** v31.4: Ausgabe zurücknehmen — schreibt, patcht den Cache und räumt den
   *  Merker in „Letzte Check-ins" mit weg. */
  const clearShirtIssue = async (
    target: { regId: number; name: string; eventId: string; subsiteUrl: string },
  ): Promise<boolean> => {
    if (!eventService || !target.subsiteUrl) return false;
    setShirtBusy(true);
    try {
      const r = await eventService.clearShirtIssued(target.subsiteUrl, target.regId);
      if (!r.ok) {
        setResultMessage(shirtFailMsg(target.name, r.status, true));
        setResultType('error');
        return false;
      }
      patchCachedReg(target.eventId, target.regId, { shirtIssued: '' });
      setRecent(recentRef.current.filter(x => !(x.kind === 'shirt' && x.regId === target.regId && x.subsiteUrl === target.subsiteUrl)));
      setResultMessage(isDe
        ? `${target.name} — Trikot-Ausgabe zurückgenommen, die Größe zählt wieder zum Bestand.`
        : `${target.name} — shirt handout reverted, the size counts towards the stock again.`);
      setResultType('info');
      return true;
    } finally { setShirtBusy(false); }
  };
  const openShirtAsk = (reg: import('../services/EventService').SPRegistration, name: string, info: ShirtDeskInfo): void => {
    setAskShirtSize(info.preset);
    setShirtAsk({ reg, name, eventId: nameSearchEventId });
  };


  // URL für den Browser-Link generieren (für zukünftige Nutzung)
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const getBrowserUrl = (): string => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    if (ctx) {
      return ctx.pageContext.site.absoluteUrl + ctx.pageContext.site.serverRelativeUrl.replace(ctx.pageContext.site.serverRelativeUrl, '') + '/SitePages/' +
        (window.location.pathname.split('/').pop() || 'DEX.aspx');
    }
    return window.location.href;
  };

  // Scanner in neuem Fenster öffnen (aspx-Seite in SiteAssets, für zukünftige Nutzung)
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const openExternalScanner = (): void => {
    const checkinUrl = `${siteUrl}/SiteAssets/checkin.aspx`;
    const params = new URLSearchParams();
    params.set('siteUrl', siteUrl);
    if (selectedEvent) {
      params.set('subsiteUrl', selectedEvent.subsiteUrl || '');
      params.set('eventTitle', selectedEvent.title);
      params.set('eventNumber', (selectedEvent.eventNumber || 0).toString());
    }
    window.open(`${checkinUrl}?${params.toString()}`, '_blank');
  };

  const eventByNumber = React.useMemo(() => {
    const map: Record<number, { id: string; title: string; subsiteUrl: string; eventNumber: number }> = {};
    for (const e of events) {
      if (e.eventNumber) map[e.eventNumber] = { id: e.id, title: e.title, subsiteUrl: e.subsiteUrl || '', eventNumber: e.eventNumber };
    }
    return map;
  }, [events]);

  const lastScannedRef = React.useRef<string>('');
  const processingRef = React.useRef<boolean>(false);

  // qr-scanner starten
  const startCamera = async (): Promise<void> => {
    setCameraError('');
    setCameraErrorInIframe(false); // v30.29
    if (!videoRef.current) return;

    // 1. Browser-API-Check: manche Embedded-WebViews (SharePoint-Mobile-App)
    //    stellen mediaDevices gar nicht bereit. Dort gleich mit klarer Meldung
    //    abbrechen statt auf qr-scanner-Fehler zu warten.
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      setCameraError(
        'Dein Browser stellt keinen Kamera-Zugriff bereit. '
        + 'Bitte öffne diese Seite direkt in Edge, angemeldet mit deinem Arbeitskonto (nicht in der SharePoint-App / Teams) — oder nimm den Foto-Weg unten.'
      );
      return;
    }

    // 2. Explizit Berechtigung anfragen (triggert Permission-Prompt). Damit
    //    bekommen wir sauber unterscheidbare Fehler statt einer generischen
    //    qr-scanner-Exception. Das Test-Stream wird danach sofort geschlossen
    //    und qr-scanner startet seinen eigenen Stream.
    try {
      const testStream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: 'environment' } },
        audio: false,
      });
      // Test-Stream sofort stoppen, damit qr-scanner seinen eigenen aufbauen kann
      testStream.getTracks().forEach(track => track.stop());
    } catch (err) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const e = err as any;
      const name = e?.name || '';
      let msg: string;
      if (name === 'NotAllowedError' || name === 'PermissionDeniedError') {
        // v30.29: `NotAllowedError` hat ZWEI sehr verschiedene Ursachen, und
        // die alte Meldung nannte immer nur die eine.
        //
        //  (a) Die Person hat den Dialog wirklich abgelehnt → Browser-
        //      Einstellung zurücksetzen hilft.
        //  (b) Die Seite steckt in einem iframe OHNE `allow="camera"`
        //      (Teams-Registerkarte, eingebettete Web-Part-Ansicht, manche
        //      SharePoint-Rahmen). Dann wird gar nicht erst GEFRAGT — Chrome
        //      wirft denselben Fehlernamen. Wer daraufhin das Schloss-Icon
        //      sucht, findet dort keinen Kamera-Eintrag und hält das Gerät
        //      für kaputt. Der Rahmen gibt die Kamera nicht frei; helfen
        //      kann nur, die Seite direkt im Browser zu öffnen.
        //
        // Unterschieden wird über die Permissions-API: `prompt` heißt „nie
        // gefragt" und damit (b). Fehlt die API (Safari < 16), entscheidet
        // die iframe-Erkennung allein.
        let inIframe = false;
        try { inIframe = window.self !== window.top; } catch { inIframe = true; /* Cross-Origin-Zugriff wirft → sicher im iframe */ }
        let neverAsked = false;
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const perms = (navigator as any).permissions;
          if (perms?.query) {
            const st = await perms.query({ name: 'camera' });
            neverAsked = st?.state === 'prompt';
          }
        } catch { /* Permissions-API kennt 'camera' nicht — dann bleibt es bei der iframe-Erkennung */ }

        setCameraErrorInIframe(inIframe);
        if (inIframe) {
          msg = 'Diese Seite läuft in einem eingebetteten Rahmen (Teams-Registerkarte oder SharePoint-App). '
            + 'Der Rahmen gibt die Kamera nicht frei — du wurdest deshalb gar nicht erst gefragt. '
            + 'Öffne die Seite direkt in Edge (Arbeitskonto) und starte den Scan dort erneut — oder nimm den Foto-Weg, der ohne Kamera-Freigabe auskommt. '
            + 'In Teams: die drei Punkte oben rechts an der Registerkarte → „Im Browser öffnen".';
        } else if (neverAsked) {
          msg = 'Der Browser hat die Kamera blockiert, ohne zu fragen. '
            + 'Das passiert bei unsicheren Verbindungen oder wenn die Seite eingebettet ist — '
            + 'öffne sie direkt in Edge (Arbeitskonto) — oder nimm den Foto-Weg, der ohne Kamera-Freigabe auskommt.';
        } else {
          msg = 'Kamera-Berechtigung wurde abgelehnt. Bitte in den Browser-Einstellungen '
            + 'für diese Seite die Kamera erlauben und dann erneut versuchen. '
            + '(iOS Safari: aA-Icon links in der Adresszeile → Website-Einstellungen → Kamera: Erlauben. '
            + 'Android Chrome: Schloss-Icon in der Adresszeile → Berechtigungen → Kamera: Zulassen.)';
        }
      } else if (name === 'NotFoundError' || name === 'DevicesNotFoundError') {
        msg = 'Keine Kamera gefunden. Stelle sicher, dass dein Gerät eine Kamera hat und kein anderes Programm sie blockiert.';
      } else if (name === 'NotReadableError' || name === 'TrackStartError') {
        msg = 'Die Kamera ist bereits in Benutzung (z.B. Teams-Anruf oder eine andere App). Bitte schließe andere Apps und versuche es erneut.';
      } else if (name === 'OverconstrainedError' || name === 'ConstraintNotSatisfiedError') {
        // Kein Environment-Facing-Camera verfügbar -> Fallback auf beliebige Kamera
        try {
          const fallback = await navigator.mediaDevices.getUserMedia({ video: true, audio: false });
          fallback.getTracks().forEach(track => track.stop());
        } catch {
          setCameraError('Keine passende Kamera gefunden.');
          return;
        }
        // Fallback OK -> weiter mit qr-scanner-Start (ohne early return)
        msg = '';
      } else if (name === 'SecurityError') {
        msg = 'Kamera-Zugriff vom Browser blockiert (vermutlich unsichere Verbindung oder eingebetteter iframe). Öffne die Seite direkt in Edge (Arbeitskonto) — oder nimm den Foto-Weg.';
      } else {
        msg = `Kamera konnte nicht gestartet werden: ${e?.message || String(err) || 'Unbekannter Fehler'}`;
      }
      if (msg) {
        setCameraError(msg);
        return;
      }
    }

    // 3. qr-scanner starten (nutzt jetzt die bereits erteilte Permission)
    try {
      // v20.0: Bibliothek lazy laden (eigener Chunk, nur bei Kamera-Start).
      const QrScannerCls = (await import('qr-scanner')).default;
      const scanner = new QrScannerCls(
        videoRef.current,
        async (result) => {
          const code = result.data;
          if (!code || code === lastScannedRef.current || processingRef.current) return;
          lastScannedRef.current = code;
          processingRef.current = true;
          // Vibration bei Erkennung
          try { navigator.vibrate(200); } catch { /* */ }
          await processCode(code);
          processingRef.current = false;
          setTimeout(() => { lastScannedRef.current = ''; }, 3000);
        },
        {
          preferredCamera: 'environment',
          highlightScanRegion: true,
          highlightCodeOutline: true,
          maxScansPerSecond: 5,
        }
      );
      scannerRef.current = scanner;
      await scanner.start();
      setIsScanning(true);
    } catch (err) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const error = err as any;
      const msg = typeof error === 'string' ? error : error?.message || 'Unbekannter Fehler';
      setCameraError(`Scanner konnte nicht gestartet werden: ${msg}`);
    }
  };

  const stopCamera = (): void => {
    if (scannerRef.current) {
      scannerRef.current.stop();
      scannerRef.current.destroy();
      scannerRef.current = null;
    }
    setIsScanning(false);
  };

  // Code verarbeiten und einchecken
  /**
   * v30.30: Foto-Weg. Der Live-Scanner braucht `getUserMedia` und hängt damit
   * an einer Kette, die wir nicht kontrollieren (Browserwahl → CA-Policy →
   * WebView → Kamera-Permission). Ein File-Input mit `capture="environment"`
   * ruft stattdessen die native Kamera-App auf: Die App fragt DEX nie nach
   * einer Kamera-Berechtigung, sie bekommt ein fertiges Bild.
   *
   * Ab hier ist alles identisch zum Live-Scan — dieselbe `processCode`, dieselbe
   * Ergebniskarte, dieselbe Doppel-Scan-Erkennung. Der Foto-Weg ist bewusst nur
   * eine zweite EINGABE, kein zweiter Ablauf.
   *
   * Die Bibliothek kommt aus demselben Lazy-Chunk wie beim Live-Scanner (v20.0),
   * das Bundle wächst dadurch nicht.
   */
  const handlePhotoPicked = async (file: File | null): Promise<void> => {
    setPhotoHint('');
    if (!file) {
      // Kein File trotz Klick: Auf verwalteten Geräten kann die Richtlinie den
      // Rückweg aus der Kamera-App unterbinden („Receive data from other apps").
      // Das sieht aus wie ein Abbruch durch den Nutzer und ist keiner — deshalb
      // beide Möglichkeiten nennen statt zu raten.
      setPhotoHint(isDe
        ? 'Es kam kein Foto zurück. Entweder hast du abgebrochen — oder eine Geräte-Richtlinie erlaubt die Übergabe aus der Kamera-App nicht. Dann bleibt die Suche in der Teilnehmerliste unten.'
        : 'No photo was returned. Either you cancelled — or a device policy blocks handing files over from the camera app. In that case use the participant search below.');
      return;
    }
    setPhotoBusy(true);
    try {
      const QrScannerCls = (await import('qr-scanner')).default;
      const res = await QrScannerCls.scanImage(file, { returnDetailedScanResult: true });
      await processCode(res.data);
    } catch {
      // scanImage wirft, wenn im Bild kein Code steckt — das ist der Normalfall
      // eines unscharfen oder zu weit entfernten Fotos, kein technischer Fehler.
      setPhotoHint(isDe
        ? 'Kein QR-Code im Foto erkannt. Geh näher ran, halte das Handy ruhig und achte darauf, dass der Code nicht gespiegelt oder überstrahlt ist.'
        : 'No QR code found in the photo. Move closer, hold steady, and make sure the code is not mirrored or washed out by glare.');
    } finally {
      setPhotoBusy(false);
    }
  };

  const processCode = async (code: string): Promise<void> => {
    if (isProcessing) return;
    setIsProcessing(true);
    setResultMessage('');
    setResultType('');

    const parts = code.split('|');
    if (parts.length !== 3 || parts[0] !== 'DEX') {
      setResultMessage(`Ungültiger QR-Code: "${code}"`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    const eventNumber = parseInt(parts[1], 10);
    const email = parts[2];

    // EventNumber oder SP-ID Lookup
    let event = eventByNumber[eventNumber];

    // Fallback: Wenn EventNumber 0 oder nicht gefunden, versuche SP-ID
    if (!event) {
      const byId = events.find(e => e.id === parts[1]);
      if (byId) {
        event = { id: byId.id, title: byId.title, subsiteUrl: byId.subsiteUrl || '', eventNumber: byId.eventNumber };
      }
    }

    if (!event || !event.subsiteUrl || !eventService) {
      setResultMessage(`Event #${eventNumber} nicht gefunden. (Code: ${code})`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    const reg = await eventService.getRegistrationByEmail(event.subsiteUrl, email);
    if (!reg) {
      setResultMessage(`${email} — nicht registriert.`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;

    if (reg.Status === 'Abgemeldet') {
      setResultMessage(`${name} — ${t('checkin.cancelled')}`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }
    // v30.91: Programmpunkt-Modus des GESCANNTEN Events. Der Punkt wird
    // unten beim Event gewählt; ein Scan zu einem anderen Event oder ohne
    // Punkt wird nicht geraten, sondern abgelehnt.
    const scanEv = events.find(e => e.id === event.id);
    const scanAgenda = !!(scanEv && scanEv.agendaCheckIn && (scanEv.agenda || []).length > 0);
    if (scanAgenda) {
      if (event.id !== nameSearchEventId || !agendaPoint) {
        setResultMessage(isDe
          ? `Dieses Event hat Programmpunkte — bitte unten das Event und den ${agendaTermSingular} wählen, an dem eingecheckt wird.`
          : `This event has agenda items — please pick the event and the ${agendaTermSingular.toLowerCase()} below first.`);
        setResultType('error');
        setIsProcessing(false);
        return;
      }
      const marks = parseAgendaCheckIns(reg.AgendaCheckIns);
      if (marks[agendaPoint.id]) {
        setResultMessage(`${name} — ${isDe ? 'bereits erfasst' : 'already recorded'} (${agendaPoint.title}, ${formatMarkTime(marks[agendaPoint.id].at)})`);
        setResultType('info');
        setIsProcessing(false);
        return;
      }
    } else if (reg.Status === 'Eingecheckt') {
      setResultMessage(`${name} — ${t('checkin.alreadycheckedin')}`);
      setResultType('info');
      setIsProcessing(false);
      return;
    }

    // Statt sofort einzuchecken, Info-Karte anzeigen und auf Bestätigung warten
    // Profilbild laden
    let photoUrl = '';
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (ctx) {
        const siteBase = ctx.pageContext.web.absoluteUrl;
        photoUrl = `${siteBase}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`;
      }
    } catch { /* */ }

    // v31.4: Trikot-Angaben der Person für die Bestätigungskarte — die
    // Größenwahl steht dort vorbelegt bereit (Gegenvorschlag vor Wunsch).
    const shirt = shirtDeskInfoFor(reg, event.id);
    setCardShirtSize(shirt ? shirt.preset : '');
    setPendingCheckIn({
      name,
      email,
      event: { id: event.id, subsiteUrl: event.subsiteUrl, title: event.title },
      regId: reg.Id,
      status: reg.Status,
      agendaItemId: scanAgenda && agendaPoint ? agendaPoint.id : undefined,
      agendaLabel: scanAgenda && agendaPoint ? agendaPoint.title : undefined,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      department: (reg as any).Department || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      jobTitle: (reg as any).JobTitle || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      location: (reg as any).Location || '',
      photoUrl,
      // v30.53: auch auf dem QR-Weg — der Tisch braucht dieselben Angaben,
      // egal ob gescannt oder gesucht wurde.
      extras: extrasFor(reg, event.id),
      shirt,
      // v31.4: Auch nach einem Scan — der Helfer schaut nachher auf dieselbe
      // Mail und soll die beiden Zahlen dort wiederfinden.
      // v31.4 (Review): über den Cache, nicht über `reg` — die gescannte Zeile
      // trägt weder `QrSentId` noch `TeilnehmerID` (fester `$select`).
      qrNote: qrNoteFromCache(reg, event.id),
    });
    setResultMessage('');
    setResultType('');
    setIsProcessing(false);
    // Zum Bestätigungs-Dialog scrollen
    setTimeout(() => {
      confirmCardRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 100);
  };

  // Check-in ausführen. v31.1: aus `confirmCheckIn` herausgelöst — Liste und
  // Teilnehmer-ID checken direkt ein (Nutzer 07.09.2026: „wenn ich auf
  // Einchecken klicke, soll er direkt eingecheckt sein"); nur der Scan zeigt
  // vorher die Bestätigungskarte, weil dort die Person erst identifiziert wird.
  const performCheckIn = async (pendingCheckIn: PendingCheckInInfo): Promise<void> => {
    if (!eventService) return;
    const remember = (): void => {
      setRecent([{
        key: `${pendingCheckIn.regId}:${pendingCheckIn.agendaItemId || 'status'}:${Date.now()}`,
        at: new Date().toISOString(), name: pendingCheckIn.name, regId: pendingCheckIn.regId,
        eventId: pendingCheckIn.event.id || '', subsiteUrl: pendingCheckIn.event.subsiteUrl,
        prevStatus: pendingCheckIn.status, agendaItemId: pendingCheckIn.agendaItemId, agendaLabel: pendingCheckIn.agendaLabel,
      }, ...recentRef.current]);
    };
    try {
      // v30.67: `checkInParticipant` wirft nie — es liefert `response.ok` bzw.
      // false. Der catch unten fängt also nur Unerwartetes; ob der MERGE
      // (403 ohne Schreibrecht, 429 in der Einlasswelle, 400 ohne Spalte)
      // durchging, steht NUR im Rückgabewert. Vorher stieg der Zähler und die
      // grüne Meldung kam auch dann, wenn in der Liste weiter „Angemeldet"
      // stand — und die Person tauchte später in der No-Show-Auswertung auf.
      // v30.91: Programmpunkt — nur der Punkt wird gesetzt, der Status bleibt.
      if (pendingCheckIn.agendaItemId) {
        const pointId = pendingCheckIn.agendaItemId;
        const label = pendingCheckIn.agendaLabel || '';
        const r = await eventService.checkInAgendaItem(pendingCheckIn.event.subsiteUrl, pendingCheckIn.regId, pointId);
        if (!r.ok) {
          setResultMessage(isDe
            ? `${pendingCheckIn.name} — Anwesenheit konnte nicht gespeichert werden${r.status ? ` (HTTP ${r.status})` : ''}. ${r.status === 400 ? 'Fehlt die Spalte AgendaCheckIns? Organizer: „Spalten fixen" ausführen.' : 'Bitte erneut versuchen.'}`
            : `${pendingCheckIn.name} — attendance could not be saved${r.status ? ` (HTTP ${r.status})` : ''}.`);
          setResultType('error');
          setPendingCheckIn(null);
          processingRef.current = false;
          return;
        }
        const at = r.already || new Date().toISOString();
        patchCachedReg(pendingCheckIn.event.id || '', pendingCheckIn.regId, { agenda: { [pointId]: { at, by: '' } } });
        if (!r.already) { setCheckedInCount(prev => prev + 1); remember(); }
        setResultMessage(r.already
          ? `${pendingCheckIn.name} — ${isDe ? 'bereits erfasst' : 'already recorded'} (${label}, ${formatMarkTime(r.already)})`
          : `${pendingCheckIn.name} — ${isDe ? 'anwesend bei' : 'present at'} ${label}`);
        setResultType(r.already ? 'info' : 'success');
        setPendingCheckIn(null);
        processingRef.current = false;
        return;
      }
      const ok = await eventService.checkInParticipant(pendingCheckIn.event.subsiteUrl, pendingCheckIn.regId);
      if (!ok) {
        setResultMessage(isDe
          ? `${pendingCheckIn.name} — Check-in fehlgeschlagen (bitte erneut versuchen oder Organizer informieren).`
          : `${pendingCheckIn.name} — check-in failed (please retry or inform an organizer).`);
        setResultType('error');
        setPendingCheckIn(null);
        processingRef.current = false;
        return;
      }
      setCheckedInCount(prev => prev + 1);
      remember();
      // v30.67: Den neuen Status auch in der Trefferliste nachführen — sie
      // liest nur aus dem Cache, und der wurde bisher nur beim No-Show
      // gepatcht. Ohne den Patch blieb die Person „Angemeldet", die KPI-
      // Kachel stand, und unter „Nur offene" stand sie weiter bei den Offenen;
      // ein zweiter Helfer checkte sie erneut ein.
      patchCachedReg(pendingCheckIn.event.id || '', pendingCheckIn.regId, { status: 'Eingecheckt' });
      setResultMessage(`${pendingCheckIn.name} — ${t('checkin.success')}`);
      setResultType('success');
    } catch {
      setResultMessage(`${pendingCheckIn.name} — Check-in fehlgeschlagen.`);
      setResultType('error');
    }
    setPendingCheckIn(null);
    processingRef.current = false;
  };
  const confirmCheckIn = (): Promise<void> => pendingCheckIn ? performCheckIn(pendingCheckIn) : Promise.resolve();

  /** v31.1: Check-in aus „Letzte Check-ins" zurücknehmen. Programmpunkt →
   *  Anwesenheit entfernen; Event-Status → zurück auf den gemerkten Stand.
   *  Das Ergebnis wird geprüft, der Cache nachgeführt, der Zähler korrigiert. */
  const undoCheckIn = async (e: RecentCheckIn): Promise<void> => {
    if (!eventService || undoBusyKey) return;
    setUndoBusyKey(e.key);
    try {
      let ok = false;
      const isNoShow = e.kind === 'noshow';
      // v31.4: Eine Trikot-Ausgabe wird über die Spalte zurückgenommen, nicht
      // über den Status — `clearShirtIssue` meldet das Ergebnis selbst (der
      // 400-Fall braucht seinen eigenen Satz) und räumt den Merker weg.
      if (e.kind === 'shirt') {
        await clearShirtIssue({ regId: e.regId, name: e.name, eventId: e.eventId, subsiteUrl: e.subsiteUrl });
        return;
      }
      if (e.agendaItemId) {
        // v31.2: entfernt die Marke am Punkt — Anwesenheit ODER No-Show.
        ok = (await eventService.removeAgendaCheckIn(e.subsiteUrl, e.regId, e.agendaItemId)).ok;
        // v31.4: `null` heißt „diese Marke ist weg" — auch gegenüber einer
        // frisch geladenen Liste, die sie noch führt.
        if (ok) patchCachedReg(e.eventId, e.regId, { agenda: { [e.agendaItemId]: null } });
      } else {
        ok = await eventService.revertCheckIn(e.subsiteUrl, e.regId, e.prevStatus);
        if (ok) {
          const back = (e.prevStatus === 'QR versendet' || (isNoShow && e.prevStatus === 'Eingecheckt')) ? e.prevStatus : 'Angemeldet';
          patchCachedReg(e.eventId, e.regId, { status: back });
        }
      }
      if (!ok) {
        setResultMessage(isDe ? `${e.name} — Rückgängig fehlgeschlagen, bitte erneut versuchen.` : `${e.name} — undo failed, please retry.`);
        setResultType('error');
        return;
      }
      setRecent(recentRef.current.filter(x => x.key !== e.key));
      if (!isNoShow) setCheckedInCount(prev => Math.max(0, prev - 1));
      setResultMessage(isDe
        ? `${e.name} — ${isNoShow ? 'No-Show' : 'Check-in'} zurückgenommen${e.agendaLabel ? ` (${e.agendaLabel})` : ''}.`
        : `${e.name} — ${isNoShow ? 'no-show' : 'check-in'} reverted${e.agendaLabel ? ` (${e.agendaLabel})` : ''}.`);
      setResultType('info');
    } finally { setUndoBusyKey(''); }
  };

  const cancelCheckIn = (): void => {
    setPendingCheckIn(null);
    setCardShirtSize(''); // v31.4: nächste Person, nächste Größe
    lastScannedRef.current = '';
    processingRef.current = false;
  };


  React.useEffect(() => {
    return () => {
      if (scannerRef.current) {
        scannerRef.current.stop();
        scannerRef.current.destroy();
      }
    };
  }, []);

  // v6.22: Zugriffskontrolle — wer kein zugängliches Event hat UND nicht Admin ist,
  // kommt gar nicht erst in die Scanner-Maske. Deckt User ohne Rolle ab UND auch
  // Organizer/Admin ohne eigene Events.
  if (!isAdmin && !isOrganizer && accessibleEvents.length === 0) {
    return (
      <div className="page-container text-center">
        <p style={{ color: 'var(--dex-gray-400)', padding: 48 }}>
          {t('checkin.noaccess') || 'Du hast keinen Zugriff auf den Check-In. Wende dich an einen Admin, wenn du als Organizer oder QR-Scanner eingetragen werden solltest.'}
        </p>
        <button className="btn btn-secondary" onClick={() => navigate('landing')}>{t('reg.backtoevents') || 'Zurück'}</button>
      </div>
    );
  }

  // v23.45: gemeinsamer Karten-Renderer für Haupt- und Sub-Events im Picker.
  // `isChild` rendert eine kompaktere, eingerückte Variante.
  const renderCheckInEventCard = (ev: typeof events[number], isChild: boolean): React.ReactElement => {
    const thumb = isChild ? 60 : 84;
    return (
      <button
        key={ev.id}
        className="card"
        onClick={() => navigate('check-in', ev.id)}
        style={{
          textAlign: 'left', padding: isChild ? 12 : 14, width: '100%',
          background: isChild ? 'var(--dex-gray-50, #f7f8f9)' : '#fff',
          border: '1px solid var(--dex-gray-200)',
          borderRadius: 12, cursor: 'pointer',
          display: 'flex', alignItems: 'center', gap: 14,
          transition: 'border-color 0.15s, box-shadow 0.15s, transform 0.1s',
        }}
        onMouseEnter={e => { e.currentTarget.style.borderColor = 'var(--dex-green, #86bc25)'; e.currentTarget.style.boxShadow = '0 4px 12px rgba(134,188,37,0.12)'; }}
        onMouseLeave={e => { e.currentTarget.style.borderColor = 'var(--dex-gray-200)'; e.currentTarget.style.boxShadow = 'none'; }}
      >
        {/* v15.3: Event-Bild als Thumbnail links — Fallback grünes Gradient mit Initialen. */}
        <div style={{
          flex: '0 0 auto', width: thumb, height: thumb, borderRadius: 10,
          background: ev.imageUrl
            ? `url(${ev.imageUrl}) center/cover no-repeat`
            : 'linear-gradient(135deg, var(--dex-green, #86bc25), var(--dex-blue, #0076a8))',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          color: '#fff', fontWeight: 700, fontSize: isChild ? '1.1rem' : '1.4rem',
          overflow: 'hidden', flexShrink: 0,
        }}>
          {!ev.imageUrl && ev.title ? ev.title.charAt(0).toUpperCase() : ''}
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          {isChild && (
            <span style={{
              display: 'inline-block', fontSize: '0.66rem', fontWeight: 700, letterSpacing: 0.3,
              color: 'var(--dex-green-dark, #4a7c1f)', background: 'rgba(134,188,37,0.14)',
              borderRadius: 6, padding: '1px 7px', marginBottom: 4, textTransform: 'uppercase',
            }}>{isDe ? 'Sub-Event' : 'Sub-event'}</span>
          )}
          <div style={{ fontWeight: 700, fontSize: isChild ? '0.92rem' : '1rem', color: 'var(--dex-gray-800)', marginBottom: 2 }}>{shortSubEventTitle(ev.title)}</div>
          <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
            {ev.startDate ? new Date(ev.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) : ''}
            {ev.location ? ` · ${ev.location}` : ''}
          </div>
          {/* v15.3: Organizer-Chips mit Hover-Profilfoto (gleiches OrganizerList wie auf der Registrierungsseite). */}
          {/* v29.48: „Organizer ausblenden" wurde hier nicht ausgewertet — die
              Selbst-Check-in-Seite ist teilnehmersichtbar, also gilt die
              Einstellung auch hier (gleiche Prüfung wie Anmeldeseite/Kachel). */}
          {!isChild && (ev.organizers && ev.organizers.length > 0) && !(ev.hideOrganizer && !ev.hideOrganizerIndividualOnly) && (
            <div style={{ marginTop: 6 }}>
              <OrganizerList
                names={ev.organizers}
                emails={ev.organizerEmails || []}
                hiddenEmails={(ev.hideOrganizer && ev.hideOrganizerIndividualOnly) ? (ev.hiddenOrganizerEmails || []) : []}
                size="sm"
                compact
              />
            </div>
          )}
        </div>
      </button>
    );
  };

  // Kein Event ausgewählt → Event-Picker (nur relevante Events: nur die, die
  // der User einchecken darf; bei exakt einem Event wird automatisch weiter
  // navigiert, weil die LandingPage das schon macht. Wir fangen hier aber
  // trotzdem den Fall ab, wenn jemand direkt via Header-Button ohne Eventauswahl
  // herkommt.)
  if (!selectedEvent) {
    if (accessibleEvents.length === 1) {
      // Auto-select: direkt weiterleiten statt Liste mit einem Eintrag zeigen.
      navigate('check-in', accessibleEvents[0].id);
      return (
        <div className="page-container text-center">
          <p style={{ color: 'var(--dex-gray-400)', padding: 48 }}>…</p>
        </div>
      );
    }
    return (
      <div className="page-container" role="main" style={{ maxWidth: 1100, marginLeft: 'auto', marginRight: 'auto' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 12, marginBottom: 8 }}>
          <h2 style={{ margin: 0 }}>{t('checkin.title') || 'Check-In'}</h2>
          {/* v15.3: Toggle „Nur aktive" — identisches Pattern wie EventListPage. */}
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
            <input
              type="checkbox"
              checked={onlyActiveCheckIn}
              onChange={e => setOnlyActiveCheckIn(e.target.checked)}
              style={{ width: 16, height: 16, accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
            />
            <span>{isDe ? 'Nur aktive Events' : 'Active events only'}</span>
          </label>
        </div>
        <p style={{ color: 'var(--dex-gray-600)', marginBottom: 16, fontSize: '0.9rem' }}>
          {t('checkin.pickevent') || 'Wähle das Event, für das du eincheckst:'}
        </p>
        {visibleCheckInEvents.length === 0 ? (
          <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
            {onlyActiveCheckIn
              ? (isDe ? 'Keine aktiven Events. Toggle oben deaktivieren, um auch ältere zu sehen.' : 'No active events. Untoggle „Active events only" to see older ones.')
              : (t('checkin.noevents') || 'Keine Events verfügbar.')}
          </p>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            {groupedCheckInEvents.map(({ parent, children }) => {
              const expanded = !!expandedCheckInParents[parent.id];
              const hasChildren = children.length > 0;
              return (
                <div key={parent.id} style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ position: 'relative' }}>
                    {renderCheckInEventCard(parent, false)}
                    {hasChildren && (
                      <button
                        type="button"
                        onClick={e => { e.stopPropagation(); setExpandedCheckInParents(p => ({ ...p, [parent.id]: !p[parent.id] })); }}
                        title={isDe ? (expanded ? 'Sub-Events einklappen' : 'Sub-Events anzeigen') : (expanded ? 'Collapse sub-events' : 'Show sub-events')}
                        style={{
                          ...(isMobile
                            ? { marginTop: 8, alignSelf: 'flex-start' }
                            : { position: 'absolute', top: '50%', right: 14, transform: 'translateY(-50%)' }),
                          display: 'inline-flex', alignItems: 'center', gap: 6,
                          background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 999,
                          padding: '6px 12px', cursor: 'pointer', color: 'var(--dex-gray-700)',
                          fontSize: '0.78rem', fontWeight: 600,
                        }}
                      >
                        <span>{children.length} {isDe ? 'Sub-Events' : 'sub-events'}</span>
                        {expanded ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
                      </button>
                    )}
                  </div>
                  {expanded && hasChildren && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginLeft: 18, paddingLeft: 18, borderLeft: '2px solid var(--dex-gray-200)' }}>
                      {children.map(ch => renderCheckInEventCard(ch, true))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  // v20.1: Self-Check-in — grundsätzlich immer verfügbar. Hat das Event noch
  // keinen aktiven Token (Wizard-Toggle nie gesetzt), wird Self-Check-in beim
  // ersten Klick automatisch aktiviert: Token erzeugen + am Event speichern.
  // Das Persistieren braucht Schreibrechte auf DEX_Events (Organizer/Admin) —
  // reine Check-in-Helfer bekommen in dem Fall einen klaren Hinweis. Existiert
  // der Token bereits, funktionieren beide Aktionen ohne Schreibzugriff.
  const ensureSelfCheckInReady = async (): Promise<string | null> => {
    if (selectedEvent.selfCheckInEnabled && selectedEvent.selfCheckInToken) {
      return selectedEvent.selfCheckInToken;
    }
    const token = selectedEvent.selfCheckInToken || generateSelfCheckInToken();
    let ok = false;
    try {
      ok = await updateEvent(selectedEvent.id, { 'SelfCheckInEnabled': true, 'SelfCheckInToken': token });
    } catch { ok = false; }
    if (!ok) {
      showAlert(isDe
        ? 'Der Self-Check-in-QR für dieses Event konnte noch nicht eingerichtet werden (fehlende Berechtigung zum Speichern am Event). Bitte einmalig von einem Organizer oder Admin öffnen lassen — danach kann auch das Check-in-Team QR-Anzeige und PDF nutzen.'
        : 'The self check-in QR for this event could not be set up yet (missing permission to save on the event). Please have an organizer or admin open it once — afterwards the check-in team can use the QR display and PDF as well.', { variant: 'error' });
      return null;
    }
    return token;
  };
  const openSelfCheckInDisplay = async (): Promise<void> => {
    if (selfCheckInBusy) return;
    setSelfCheckInBusy(true);
    try {
      const token = await ensureSelfCheckInReady();
      if (token) navigate('self-checkin-display', selectedEvent.id);
    } finally { setSelfCheckInBusy(false); }
  };
  const downloadSelfCheckInQrPdf = async (): Promise<void> => {
    if (selfCheckInBusy) return;
    setSelfCheckInBusy(true);
    try {
      const token = await ensureSelfCheckInReady();
      if (!token) return;
      const dateLabel = selectedEvent.startDate
        ? new Date(selectedEvent.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })
        : '';
      await downloadSelfCheckInPdf({
        eventTitle: selectedEvent.title || 'Event',
        eventDateLabel: dateLabel,
        locationLabel: selectedEvent.location || '',
        token,
      });
    } finally { setSelfCheckInBusy(false); }
  };

  // v31.1: Aufbau wie das Anmeldeformular (Nutzer 07.09.2026: „oben den Kreis
  // des Event-Logos und dann kommen die sinnvollen Check-in-Sektionen") —
  // nummerierte Abschnitte, das Event mit rundem Bild zuerst. Die Reihenfolge
  // kommt über CSS `order`, damit die bestehenden Karten (Scanner, Self-
  // Check-in, Liste) nicht im JSX verschoben werden müssen.
  const heroEv = events.find(e => e.id === nameSearchEventId) || selectedEvent || null;
  const heroImg = ((): string => {
    if (!heroEv) return '';
    if (heroEv.imageUrl) return heroEv.imageUrl;
    try {
      const o = JSON.parse(heroEv.emailTemplateOverrides || '{}');
      if (o && typeof o._eventLogo === 'string' && o._eventLogo) return o._eventLogo;
    } catch { /* */ }
    return heroEv.mailImageBase64 || '';
  })();
  const heroDate = ((): string => {
    if (!heroEv || !heroEv.startDate) return '';
    const f = (iso: string): string => new Date(iso).toLocaleDateString('de-DE', { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric' });
    const a = f(heroEv.startDate);
    const b = heroEv.endDate ? f(heroEv.endDate) : '';
    return b && b !== a ? `${a} – ${b}` : a;
  })();
  const sectionLabel = (n: number, label: string): React.ReactElement => (
    <div style={{ display: 'flex', alignItems: 'center', gap: 12, margin: '6px 0 10px' }}>
      <span style={{ width: 34, height: 34, borderRadius: '50%', background: 'var(--dex-green, #86bc25)', color: '#fff', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: '0.95rem', boxShadow: '0 0 0 4px rgba(134,188,37,0.18)', flexShrink: 0 }}>{n}</span>
      <span style={{ fontSize: '0.78rem', fontWeight: 700, letterSpacing: '0.08em', textTransform: 'uppercase', color: 'var(--dex-green-dark, #4a7c1f)' }}>{label}</span>
    </div>
  );
  // v31.2: Trenner OHNE Nummer für die Werkzeuge unter dem Ablauf. Nutzer
  // 07.09.2026: „das sollte nicht Schritt 4 und 5 sein, sondern einfach
  // optisch getrennt sein — sonst denkt man, das wäre chronologisch."
  const sideLabel = (label: string): React.ReactElement => (
    <div style={{ display: 'flex', alignItems: 'center', gap: 12, margin: '22px 0 10px' }}>
      <span style={{ fontSize: '0.74rem', fontWeight: 700, letterSpacing: '0.08em', textTransform: 'uppercase', color: 'var(--dex-gray-500)' }}>{label}</span>
      <span style={{ flex: 1, height: 1, background: 'var(--dex-gray-200)' }} />
    </div>
  );
  const cardToggle = (title: string, open: boolean, onToggle: () => void, hint: string): React.ReactElement => (
    <button type="button" onClick={onToggle} aria-expanded={open} style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 12, background: 'none', border: 'none', padding: 0, cursor: 'pointer', textAlign: 'left', color: 'inherit', font: 'inherit' }}>
      <h3 style={{ margin: 0 }}>{title}</h3>
      {!open && <span style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{hint}</span>}
      {open && <span style={{ flex: 1 }} />}
      <span style={{ color: 'var(--dex-gray-400)', display: 'inline-flex' }}>{open ? <ChevronUp size={18} /> : <ChevronDown size={18} />}</span>
    </button>
  );

  // v20.3: Breite begrenzen — vorher lief die Check-in-Ansicht auf großen
  // Monitoren über die volle Viewport-Breite (Event-Picker-Branch hatte
  // bereits maxWidth 1100, die Haupt-Ansicht nicht). Schmaler = lesbarer.
  return (
    <div className="page-container" role="main" style={{ maxWidth: 900, marginLeft: 'auto', marginRight: 'auto', width: '100%' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
        <h2 style={{ margin: 0 }}>
          {t('checkin.title')} {selectedEvent ? `— ${selectedEvent.title}` : ''}
        </h2>
        {checkedInCount > 0 && (
          <span style={{ fontSize: '0.9rem', fontWeight: 600, color: 'var(--dex-green)' }}>
            {checkedInCount} {t('checkin.sessioncount')}
          </span>
        )}
      </div>

      {/* Ergebnis-Anzeige */}
      {resultMessage && (
        <div style={{
          padding: '16px 20px', borderRadius: 12, marginBottom: 16, fontWeight: 600, fontSize: '1rem',
          background: resultType === 'success' ? '#e8f5e9' : resultType === 'error' ? '#ffebee' : '#e3f2fd',
          color: resultType === 'success' ? '#2e7d32' : resultType === 'error' ? '#c62828' : '#1565c0',
          border: resultType === 'success' ? '2px solid #86bc25' : resultType === 'error' ? '2px solid #ef5350' : '2px solid #42a5f5',
        }}>
          {resultMessage}
        </div>
      )}

      {/* v30.35: Der orange Warnkasten „Kamera nicht verfuegbar in der
          SharePoint App" ist raus. Er war ein Alarm fuer einen Zustand,
          der in der SharePoint-App der NORMALFALL ist — und seit v30.33
          gibt es dort einen Weg, der einfach funktioniert (Teilnehmer-ID).
          Eine Warnung, die bei jedem Start erscheint und auf etwas
          hinweist, das man nicht aendern kann, wird ueberlesen und macht
          die Seite nur enger. Die Geraete-Erwartung steht jetzt ruhig im
          Hinweistext unter dem Scan-Knopf. */}

      {/* Bestätigungs-Dialog nach Scan */}
      {pendingCheckIn && (
        <div ref={confirmCardRef} className="card" role="dialog" aria-modal="true" aria-label={isDe ? 'Check-in bestätigen' : 'Confirm check-in'} style={{
          padding: 24, marginBottom: 16, border: '2px solid var(--dex-green)',
          borderRadius: 16, background: '#fff',
        }}>
          <div style={{ display: 'flex', gap: 16, alignItems: 'center', marginBottom: 16 }}>
            {pendingCheckIn.photoUrl ? (
              <img
                src={pendingCheckIn.photoUrl}
                alt={pendingCheckIn.name}
                style={{ width: 72, height: 72, borderRadius: '50%', objectFit: 'cover', flexShrink: 0 }}
                onError={(e) => { (e.target as HTMLImageElement).style.display = 'none'; }}
              />
            ) : (
              <div style={{
                width: 72, height: 72, borderRadius: '50%', flexShrink: 0,
                background: 'linear-gradient(135deg, #86bc25, #0076a8)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                color: '#fff', fontWeight: 700, fontSize: '1.4rem',
              }}>
                {pendingCheckIn.name.split(' ').map(n => n[0]).join('').substring(0, 2)}
              </div>
            )}
            <div style={{ flex: 1 }}>
              <h3 style={{ margin: '0 0 4px', fontSize: '1.2rem' }}>{pendingCheckIn.name}</h3>
              <p style={{ margin: '0 0 2px', color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>{pendingCheckIn.email}</p>
              {/* v31.4: Beide Nummern, sobald sie auseinanderlaufen — links
                  die, die in der Mail steht, rechts die von heute. */}
              {pendingCheckIn.qrNote && (
                <p style={{ margin: '0 0 2px', color: 'var(--dex-gray-500)', fontSize: '0.78rem', fontFamily: "'Courier New',Courier,monospace" }}>
                  {pendingCheckIn.qrNote}
                </p>
              )}
              {pendingCheckIn.jobTitle && (
                <p style={{ margin: '0 0 2px', fontSize: '0.85rem' }}>{pendingCheckIn.jobTitle}</p>
              )}
              {pendingCheckIn.department && (
                <p style={{ margin: '0 0 2px', fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>{pendingCheckIn.department}</p>
              )}
              {pendingCheckIn.location && (
                <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>{pendingCheckIn.location}</p>
              )}
              {/* v30.53: Startnummer + Trikotgröße gut sichtbar — sie werden
                  hier vorgelesen bzw. ausgegeben, nicht nur nachgeschlagen.
                  v31.3: Angaben zum Nachschlagen bleiben Chips; eine
                  Handlungsanweisung (Ausweichgröße) ist ein Satz in Warnfarbe —
                  zwischen sechs grauen Chips liest sie sonst niemand. */}
              {(pendingCheckIn.extras || []).filter(x => !x.tone).length > 0 && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginTop: 8 }}>
                  {(pendingCheckIn.extras || []).filter(x => !x.tone).map((x, i) => (
                    <span key={i} style={{
                      display: 'inline-flex', alignItems: 'baseline', gap: 6,
                      padding: x.strong ? '5px 12px' : '4px 10px',
                      borderRadius: 8,
                      background: x.strong ? 'rgba(134,188,37,0.14)' : 'var(--dex-gray-100, #f5f5f5)',
                      border: x.strong ? '1px solid var(--dex-green, #86bc25)' : '1px solid var(--dex-gray-200)',
                    }}>
                      <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-600)' }}>{x.label}</span>
                      <strong style={{
                        fontSize: x.strong ? '1.15rem' : '0.9rem',
                        fontFamily: x.strong ? "'Courier New',Courier,monospace" : 'inherit',
                        letterSpacing: x.strong ? '0.04em' : undefined,
                        color: x.strong ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-800)',
                      }}>{x.value}</strong>
                    </span>
                  ))}
                </div>
              )}
              {(pendingCheckIn.extras || []).filter(x => !!x.tone).map((x, i) => (
                <div key={`tone-${i}`} style={{
                  marginTop: 8, padding: '8px 12px', borderRadius: 8, fontSize: '0.86rem', lineHeight: 1.45,
                  background: x.tone === 'danger' ? 'var(--dex-red-light, #fce8e6)' : '#fff7e6',
                  border: `1px solid ${x.tone === 'danger' ? '#f2b1ab' : '#f5c77a'}`,
                  color: x.tone === 'danger' ? '#9b2018' : '#7a4a00',
                }}>
                  <strong>{x.label}</strong> — {x.value}
                </div>
              ))}
              {/* v31.4: Die Ausgabe direkt unter dem Ausweich-Satz — dort steht,
                  WAS angeboten werden soll, hier wird festgehalten, was die
                  Person tatsächlich mitgenommen hat. Eigener Knopf statt eine
                  Ausgabe am Check-in mitzuschreiben: Ausgabe ohne Check-in gibt
                  es (Abholung am Vortag) und Check-in ohne Ausgabe erst recht
                  (wer schon eins hat). */}
              {pendingCheckIn.shirt && (
                <div style={{ marginTop: 10, padding: '10px 12px', borderRadius: 10, background: 'var(--dex-gray-50, #fafafa)', border: '1px solid var(--dex-gray-200)' }}>
                  {pendingCheckIn.shirt.issued ? (
                    <div className="dex-ui-inline">
                      <span style={{ fontSize: '0.86rem' }}>
                        {isDe ? 'Trikot ausgegeben: ' : 'Shirt handed out: '}
                        <strong>{pendingCheckIn.shirt.issued.size}</strong>
                        {pendingCheckIn.shirt.issued.at ? ` · ${formatMarkTime(pendingCheckIn.shirt.issued.at)}` : ''}
                      </span>
                      <button
                        type="button"
                        className="dex-ui-textbtn dex-ui-textbtn--danger"
                        disabled={shirtBusy}
                        onClick={() => {
                          const p = pendingCheckIn;
                          void clearShirtIssue({ regId: p.regId, name: p.name, eventId: p.event.id || '', subsiteUrl: p.event.subsiteUrl })
                            .then(ok => { if (ok) setPendingCheckIn(prev => (prev && prev.regId === p.regId && prev.shirt) ? { ...prev, shirt: { ...prev.shirt, issued: null } } : prev); });
                        }}
                      >
                        {isDe ? 'Ausgabe zurücknehmen' : 'Undo handout'}
                      </button>
                    </div>
                  ) : (
                    <>
                      <div style={{ fontSize: '0.78rem', fontWeight: 700, color: 'var(--dex-gray-600)', marginBottom: 6 }}>
                        {isDe ? 'Welches Trikot gibst du aus?' : 'Which shirt are you handing out?'}
                      </div>
                      <div className="dex-ui-inline">
                        {/* key: Bei der nächsten Person soll die Wahl wieder bei
                            der Liste anfangen, nicht im Freitext der vorigen. */}
                        <ShirtSizePicker key={pendingCheckIn.regId} sizes={pendingCheckIn.shirt.sizes} value={cardShirtSize} onChange={setCardShirtSize} isDe={isDe} />
                        <button
                          type="button"
                          className="btn btn-secondary dex-ui-btn-sm"
                          disabled={shirtBusy || !cardShirtSize.trim()}
                          onClick={() => {
                            const p = pendingCheckIn;
                            void applyShirtIssue(
                              { regId: p.regId, name: p.name, eventId: p.event.id || '', subsiteUrl: p.event.subsiteUrl, prevStatus: p.status },
                              cardShirtSize,
                            ).then(iss => { if (iss) setPendingCheckIn(prev => (prev && prev.regId === p.regId && prev.shirt) ? { ...prev, shirt: { ...prev.shirt, issued: iss } } : prev); });
                          }}
                        >
                          {shirtBusy ? (isDe ? 'Wird gespeichert…' : 'Saving…') : (isDe ? 'Ausgabe festhalten' : 'Record handout')}
                        </button>
                      </div>
                      <div className="dex-ui-muted" style={{ fontSize: '0.74rem', marginTop: 6 }}>
                        {isDe
                          ? 'Wird dauerhaft vom Bestand abgezogen — auch wenn die Person später abgemeldet oder als No-Show markiert wird.'
                          : 'Permanently deducted from the stock — even if the person is cancelled or marked as a no-show later.'}
                      </div>
                    </>
                  )}
                </div>
              )}
            </div>
          </div>
          {/* v31.4: Warum diese Zahl eine Entscheidung verlangt — VOR dem
              Schreibvorgang. Warnton wie der v31.3-Trikot-Satz, weil hier
              genau wie dort eine ENTSCHEIDUNG ansteht und nicht nur eine
              Angabe nachzuschlagen ist.
              v31.4 (Review): Die Karte deckt jetzt drei Gründe ab (zwei
              Personen, keine QR-Mail zur Zahl, abgemeldete Zeile mit dieser
              Zahl) — sie können zusammen auftreten, deshalb Satz für Satz. */}
          {pendingCheckIn.qrConflict && (() => {
            const c = pendingCheckIn.qrConflict;
            const nr = pad3(c.typed);
            const cardEventId = pendingCheckIn.event.id;
            return (
              <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 14 }}>
                <span className="dex-ui-callout-icon" aria-hidden="true"><AlertCircle size={16} /></span>
                <div>
                  {c.qrName && c.altName && (isDe ? (<>
                    Nummer <strong>{nr}</strong> stand in der QR-Mail von{' '}
                    <strong>{c.qrName}</strong>. Die laufende Nummer {nr} trägt inzwischen{' '}
                    <strong>{c.altName}</strong> (nach Abmeldungen neu vergeben).
                    Eingecheckt wird <strong>{c.qrName}</strong> — so steht es in ihrer Mail.
                  </>) : (<>
                    Number <strong>{nr}</strong> was printed in the QR email of{' '}
                    <strong>{c.qrName}</strong>. The current running number {nr} now belongs to{' '}
                    <strong>{c.altName}</strong> (reassigned after cancellations).
                    We are checking in <strong>{c.qrName}</strong> — that is what their email says.
                  </>))}
                  {c.viaRunning && !c.qrName && (isDe ? (<>
                    Zu der Nummer <strong>{nr}</strong> gibt es keine versendete QR-Mail — gefunden wurde{' '}
                    <strong>{pendingCheckIn.name}</strong> über die laufende Nummer, und die vergibt jede
                    Abmeldung neu. Bei diesem Event ist für andere Personen sehr wohl hinterlegt, was in
                    ihrer Mail stand; dieser Treffer ist also der schwächere. Bitte kurz den Namen mit der
                    Person abgleichen, bevor du eincheckst.
                  </>) : (<>
                    No sent QR email carries number <strong>{nr}</strong> — <strong>{pendingCheckIn.name}</strong>{' '}
                    was found via the running number, which is reassigned on every cancellation. For other
                    people at this event the printed number is on file, so this match is the weaker one.
                    Please check the name with the person before you check them in.
                  </>))}
                  {c.cancelledName && (isDe ? (<>
                    {(c.qrName && c.altName) || c.viaRunning ? ' ' : ''}
                    Die Nummer <strong>{nr}</strong> gehört außerdem zur stornierten Anmeldung von{' '}
                    <strong>{c.cancelledName}</strong> — abgemeldete Zeilen behalten ihre gedruckte Nummer,
                    einchecken lassen sie sich nicht.
                  </>) : (<>
                    {(c.qrName && c.altName) || c.viaRunning ? ' ' : ''}
                    Number <strong>{nr}</strong> also belongs to the cancelled registration of{' '}
                    <strong>{c.cancelledName}</strong> — cancelled rows keep their printed number but
                    cannot be checked in.
                  </>))}
                  {c.alt && c.altName && (
                    <div style={{ marginTop: 8 }}>
                      <button
                        type="button"
                        className="btn btn-secondary dex-ui-btn-sm"
                        disabled={isProcessing}
                        onClick={() => {
                          // v31.4 (Review): Die Karte bleibt stehen, bis der
                          // Alternativ-Pfad wirklich eingecheckt hat —
                          // `startManualCheckInFromSearch` kann abbrechen
                          // (kein Programmpunkt gewählt, bereits erfasst), und
                          // dann wären sonst beide Namen weg und die getippte
                          // Zahl auch. `performCheckIn` schließt sie selbst.
                          // Das Event kommt aus der Karte, nicht aus dem
                          // Auswahlfeld unten: `alt.Id` gilt nur auf DIESER
                          // Teilnehmerliste.
                          if (c.alt) startManualCheckInFromSearch(c.alt, undefined, cardEventId);
                        }}
                      >
                        {isDe
                          ? `Stattdessen ${c.altName} einchecken`
                          : `Check in ${c.altName} instead`}
                      </button>
                    </div>
                  )}
                </div>
              </div>
            );
          })()}
          <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', margin: '0 0 16px' }}>
            {isDe ? 'Event: ' : 'Event: '}<strong>{pendingCheckIn.event.title}</strong>
            {pendingCheckIn.agendaLabel && (
              <><br />{agendaTermSingular}: <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>{pendingCheckIn.agendaLabel}</strong></>
            )}
          </p>
          <div style={{ display: 'flex', gap: 12 }}>
            <button
              className="btn btn-primary"
              onClick={confirmCheckIn}
              // v31.4 (Review): Solange der Alternativ-Knopf schreibt, bleibt
              // die Karte stehen (sie darf erst weg, wenn feststeht, dass es
              // geklappt hat) — dann darf dieser Knopf aber nicht die ANDERE
              // Person danebenschreiben.
              disabled={isProcessing}
              style={{ flex: 1, fontSize: '1rem', padding: '12px 0', background: 'var(--dex-green)' }}
            >
              {pendingCheckIn.agendaItemId ? (isDe ? 'Anwesenheit erfassen' : 'Record attendance') : (isDe ? 'Einchecken' : 'Check in')}
            </button>
            <button
              className="btn btn-secondary"
              onClick={cancelCheckIn}
              style={{ padding: '12px 20px' }}
            >
              {isDe ? 'Abbrechen' : 'Cancel'}
            </button>
          </div>
        </div>
      )}

      <div style={{ display: 'flex', flexDirection: 'column' }}>
      {/* v31.1 — Abschnitt 1: das Event, wie auf der Anmeldeseite mit rundem Bild. */}
      <div style={{ order: 1 }}>
        {sectionLabel(1, isDe ? 'Event' : 'Event')}
        <div className="card" style={{ padding: heroImg ? '86px 24px 20px' : 24, marginBottom: 16, marginTop: heroImg ? 72 : 0, position: 'relative', overflow: 'visible' }}>
          {heroImg && (
            <div style={{ position: 'absolute', top: -72, left: '50%', transform: 'translateX(-50%)', width: 144, height: 144, borderRadius: '50%', background: '#fff', boxShadow: '0 8px 26px rgba(0,0,0,0.14)', padding: 6 }}>
              <img src={heroImg} alt={heroEv ? heroEv.title : ''} style={{ width: '100%', height: '100%', borderRadius: '50%', objectFit: 'cover', display: 'block' }} onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }} />
            </div>
          )}
          {heroEv ? (
            <div style={{ textAlign: 'center', marginBottom: (accessibleEvents.length > 1 || agendaMode) ? 14 : 0 }}>
              <div style={{ fontSize: '1.15rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>{heroEv.title}</div>
              <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                {heroDate}{heroEv.location ? ` · ${heroEv.location}` : ''}
              </div>
            </div>
          ) : (
            <p style={{ margin: 0, textAlign: 'center', color: 'var(--dex-gray-500)', fontSize: '0.9rem' }}>{isDe ? 'Bitte ein Event wählen.' : 'Please pick an event.'}</p>
          )}
          {accessibleEvents.length > 1 && (
            <select
              className="form-input"
              value={nameSearchEventId}
              onChange={e => {
                setNameSearchEventId(e.target.value);
                setNameSearchQuery('');
                // v31.4 (Review): Eine offene Bestätigungskarte gehört zu dem
                // Event, unter dem sie entstanden ist. Blieb sie beim Wechsel
                // stehen, zeigte ihr Alternativ-Knopf auf eine Item-Id, die es
                // auf der neuen Teilnehmerliste zwar gibt — nur mit einer
                // anderen Person dahinter.
                setPendingCheckIn(null);
                setCardShirtSize('');
                setIdInput('');
                setIdError('');
              }}
              style={{ marginBottom: agendaMode ? 12 : 0, padding: '8px 12px', fontSize: '0.9rem', width: '100%' }}
            >
              <option value="">— Event auswählen —</option>
              {accessibleEvents.map(ev => (
                <option key={ev.id} value={ev.id}>{ev.title}</option>
              ))}
            </select>
          )}
          {/* v30.91: Programmpunkt wählen — hier wird eingecheckt. Vorschlag ist
              der Punkt, der gerade läuft (suggestCurrentAgendaItem); die Wahl
              bleibt je Event im Gerät gespeichert. */}
          {agendaMode && (
            <div style={{ padding: '10px 12px', borderRadius: 10, background: 'rgba(134,188,37,0.08)', border: '1px solid rgba(134,188,37,0.45)' }}>
              <div style={{ fontSize: '0.8rem', fontWeight: 700, marginBottom: 6 }}>
                {isDe ? `${agendaTermSingular} wählen — hier wird eingecheckt` : `Pick the ${agendaTermSingular.toLowerCase()} — attendance is recorded there`}
              </div>
              {/* v30.94: Chips je Cluster (utils/agendaGroups). */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {agendaGroups(agendaItems).map((grp, gi) => (
                  <div key={grp.key}>
                    {(agendaGroups(agendaItems).length > 1 || grp.cluster) && (
                      <div style={{ fontSize: '0.74rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', marginBottom: 4 }}>
                        {groupLabel(grp, gi, isDe)} <span style={{ fontWeight: 500, color: 'var(--dex-gray-500)' }}>· {groupDateLabel(grp, isDe, false)}</span>
                      </div>
                    )}
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                      {grp.items.map(it => {
                        const active = it.id === agendaPointId;
                        const cnt = agendaCounts[it.id] || 0;
                        return (
                          <button key={it.id} type="button" onClick={() => choosePoint(it.id)} style={{
                            display: 'inline-flex', alignItems: 'center', gap: 6, padding: '6px 10px', borderRadius: 999, cursor: 'pointer',
                            border: `1px solid ${active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                            background: active ? 'var(--dex-green, #86bc25)' : '#fff', color: active ? '#fff' : 'var(--dex-gray-800)',
                            fontSize: '0.8rem', fontWeight: active ? 700 : 500, textAlign: 'left',
                          }}>
                            <span style={{ fontVariantNumeric: 'tabular-nums', opacity: 0.85 }}>{it.time}</span>
                            <span>{it.title || (isDe ? '(ohne Titel)' : '(untitled)')}</span>
                            <span style={{ fontSize: '0.72rem', padding: '1px 7px', borderRadius: 999, background: active ? 'rgba(255,255,255,0.25)' : 'var(--dex-gray-100)', color: active ? '#fff' : 'var(--dex-gray-600)' }}>{cnt}</span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>
              <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-600)', marginTop: 6 }}>
                {isDe
                  ? 'Jeder Scan, jede ID und jeder Klick setzt nur diesen Punkt. Der Event-Status der Person bleibt unverändert.'
                  : 'Every scan, ID and click records only this item. The person\'s event status stays unchanged.'}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* v31.1 — Live-Scanner, eingeklappt bis gebraucht. v31.2: ohne Nummer —
          Scanner und Self-Check-in sind Werkzeuge neben dem Ablauf, keine
          Schritte danach. */}
      <div style={{ order: 4 }}>
      {sideLabel(isDe ? 'Weitere Wege zum Einchecken' : 'Other ways to check in')}
      {/* Live-Scanner — Kamerabild + Steuerung */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        {!isScanning ? (
          <>
            {cardToggle(isDe ? 'Live-Scanner' : 'Live scanner', scannerOpen, () => setScannerOpen(o => !o), isDe ? 'QR-Code mit der Kamera scannen oder Foto vom Code machen' : 'Scan the QR code with the camera or take a photo of it')}
            {scannerOpen && (
            <div style={{ textAlign: 'center', marginTop: 14 }}>
              {/* v30.35: EIN Knopf statt zwei. Zwei gleich große Scan-Knöpfe
                  nebeneinander sind zwei Bedienwege für dieselbe Absicht — man
                  muss am Einlass erst entscheiden, statt zu scannen. Der
                  Live-Scanner ist der Standard; der Foto-Weg steht als schmaler
                  Text-Link darunter, für den Fall, dass der Scanner nicht
                  aufgeht. Wenn beides scheitert, trägt die Teilnehmer-ID
                  (v30.33) — die steht im Hinweis darunter. */}
              <button className="btn btn-primary" onClick={startCamera} style={{ fontSize: '1.1rem', padding: '14px 36px' }}>
                {t('checkin.scan')}
              </button>
              <div style={{ marginTop: 10 }}>
                <button
                  type="button"
                  onClick={() => photoInputRef.current?.click()}
                  disabled={photoBusy}
                  style={{
                    background: 'none', border: 'none', padding: 4, cursor: photoBusy ? 'default' : 'pointer',
                    color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.88rem', textDecoration: 'underline',
                    font: 'inherit', fontFamily: 'inherit',
                  }}
                >
                  {photoBusy
                    ? (isDe ? 'Foto wird gelesen…' : 'Reading photo…')
                    : (isDe ? 'Stattdessen Foto vom QR-Code machen' : 'Take a photo of the QR code instead')}
                </button>
              </div>

              {/* v31.1: Das Teilnehmer-ID-Feld steht jetzt in Abschnitt 2
                  „Einchecken" — dort, wo auch die Liste ist. */}
              {/* v30.30: `capture="environment"` öffnet die NATIVE Kamera-App des
                  Geräts und liefert eine Datei zurück — ohne getUserMedia, ohne
                  Kamera-Berechtigung für die Seite. Damit ist es der einzige
                  Scan-Weg, der weder an der WebView-Sperre der SharePoint-App
                  noch an der iframe-Freigabe einer Teams-Registerkarte hängt.
                  Bewusst KEIN `multiple` — eine Person pro Foto.

                  v30.31: `accept` wird auf Android erweitert. Seit Android 14
                  leiten Chrome und Edge ein reines `accept="image/*"` in den
                  System-Foto-Picker um — und der hat gar keinen Kamera-Eintrag,
                  auch mit `capture` nicht. Ein zusätzlicher, nicht-standardisierter
                  Wert bricht diese Umleitung und bringt die Kamera zurück (der
                  dokumentierte Workaround, s. Release Notes). Auf iOS bleibt es
                  beim sauberen `image/*` — dort greift `capture` normal, und ein
                  unbekannter MIME-Wert wäre nur ein Risiko ohne Nutzen. */}
              <input
                ref={photoInputRef}
                type="file"
                accept={isAndroid ? 'image/*,android/allowCamera' : 'image/*'}
                capture="environment"
                style={{ display: 'none' }}
                onChange={e => { void handlePhotoPicked(e.target.files?.[0] || null); e.target.value = ''; }}
              />
              {photoHint && (
                <p style={{ color: 'var(--dex-orange)', fontSize: '0.85rem', margin: '10px 0 0' }}>{photoHint}</p>
              )}
              <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', margin: '10px 0 0' }}>
                {/* v30.33: Erwartung ehrlich setzen statt jeden erst scheitern
                    lassen. Auf iPhones läuft der Scanner in aller Regel; auf
                    Android hängt er an Dingen, die wir nicht kontrollieren
                    (WebView der SharePoint-App, Foto-Picker ab Android 14).
                    Deshalb dort gleich den Weg nennen, der immer funktioniert. */}
                {isDe
                  ? <>Auf <strong>iPhones</strong> funktioniert der Live-Scanner in der Regel zuverlässig.{isAndroid ? <> Auf <strong>Android</strong> ist er oft blockiert — dann tippe die <strong>Teilnehmer-ID</strong> unten ins Suchfeld: Sie steht in jeder QR-Mail unter dem Code und funktioniert immer.</> : <> Klappt er nicht, tippe die <strong>Teilnehmer-ID</strong> unten ins Suchfeld — sie steht in jeder QR-Mail unter dem Code.</>} Der Foto-Weg benutzt die normale Kamera-App und kommt ohne Kamera-Freigabe für die Seite aus.{isAndroid ? <> Zeigt Android eine Auswahl statt der Kamera, tippe dort auf <strong>Kamera</strong>.</> : null}</>
                  : <>On <strong>iPhones</strong> the live scanner usually works reliably.{isAndroid ? <> On <strong>Android</strong> it is often blocked — then type the <strong>attendee ID</strong> into the search field below: it is printed under the code in every QR email and always works.</> : <> If it does not start, type the <strong>attendee ID</strong> into the search field below — it is printed under the code in every QR email.</>} The photo route uses the regular camera app and needs no camera permission for the page.{isAndroid ? <> If Android shows a chooser instead of the camera, pick <strong>Camera</strong> there.</> : null}</>}
              </p>
              {cameraError && (
                <div style={{ marginTop: 12 }}>
                  <p style={{ color: 'var(--dex-orange)', fontSize: '0.85rem', margin: 0 }}>{cameraError}</p>
                  {/* v30.29: Steckt die Seite im iframe, ist „direkt im Browser
                      öffnen" die einzige Abhilfe — dann auch einen Knopf dafür
                      anbieten statt nur den Weg zu beschreiben. window.open aus
                      dem iframe erzeugt einen Top-Level-Tab, dort greift die
                      Kamera-Freigabe der Seite selbst. */}
                  {cameraErrorInIframe && (
                    <button
                      className="btn btn-secondary"
                      style={{ marginTop: 10, fontSize: '0.85rem' }}
                      onClick={() => window.open(window.location.href, '_blank', 'noopener')}
                    >
                      Seite im Browser öffnen
                    </button>
                  )}
                  <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)', marginTop: 8, fontWeight: 600 }}>
                    Alternativ: Du kannst Teilnehmer auch in der Liste oben suchen
                    und per Klick auf &quot;Einchecken&quot; manuell einchecken.
                  </p>
                </div>
              )}
            </div>
            )}
          </>
        ) : (
          <>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
              <h3 style={{ margin: 0, color: 'var(--dex-green)' }}>{t('checkin.scanning')}</h3>
              <button className="btn btn-secondary" onClick={stopCamera} style={{ fontSize: '0.85rem' }}>
                Scanner stoppen
              </button>
            </div>
          </>
        )}
        {/* Wichtig: Video nicht mit fixer Höhe + objectFit:cover croppen, sonst landet
            die vom qr-scanner eingeblendete Scan-Region-Box verschoben, weil die
            Library Overlay-Koordinaten aus dem nativen Video-Aspect berechnet.
            Video soll seine natürliche Aspect Ratio behalten (width:100%, height:auto).
            Vor dem Scan-Start verstecken wir den Container via max-height:0 - so
            bleibt der Video-Ref stabil und die Library kann nach getUserMedia die
            Dimensionen korrekt bestimmen. */}
        <div style={{
          position: 'relative', width: '100%', maxWidth: 500, margin: '0 auto',
          overflow: 'hidden', borderRadius: 12,
          maxHeight: isScanning ? '80vh' : 0,
          border: isScanning ? '3px solid var(--dex-green)' : 'none',
          background: '#000',
          transition: 'max-height 0.3s ease',
        }}>
          <video
            ref={videoRef}
            style={{
              width: '100%', height: 'auto', display: 'block',
              background: '#000',
              borderRadius: 9,
            }}
            playsInline
            muted
          />
        </div>
      </div>

      </div>

      {/* v31.1 — Self-Check-in, eingeklappt bis gebraucht (v31.2: ohne
          Nummer, direkt unter dem Live-Scanner). */}
      <div style={{ order: 5 }}>
      {/* v20.1: Self-Check-in — prominent direkt unter dem Live-Scanner.
          Teilnehmer scannen den Event-QR mit der NATIVEN Handy-Kamera (kein
          Kamera-Zugriff in der App nötig) und checken sich selbst ein. */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        {cardToggle('Self-Check-in', selfOpen, () => setSelfOpen(o => !o), isDe ? 'Live-QR für einen Bildschirm am Eingang oder PDF zum Aushängen' : 'Live QR for a screen at the entrance or a printable PDF')}
        {selfOpen && (<>
        <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', margin: '12px 0 14px' }}>
          {isDe
            ? 'Teilnehmer scannen den Event-QR mit der normalen Handy-Kamera und checken sich selbst ein — ohne Scanner-Team und ohne Kamera-Freigabe in der App. Live-Anzeige für einen Bildschirm am Eingang (Code rotiert, foto-sicher) oder PDF zum Ausdrucken und Aushängen.'
            : 'Attendees scan the event QR with their regular phone camera and check themselves in — no scanner team and no in-app camera access needed. Live display for a screen at the entrance (rotating code, photo-safe) or a printable PDF to post.'}
        </p>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'center' }}>
          <button
            className="btn btn-primary"
            disabled={selfCheckInBusy}
            onClick={openSelfCheckInDisplay}
            style={{ fontSize: '0.95rem', padding: '12px 24px' }}
          >
            {isDe ? 'Live-QR anzeigen' : 'Show live QR'}
          </button>
          <button
            className="btn btn-secondary"
            disabled={selfCheckInBusy}
            onClick={downloadSelfCheckInQrPdf}
            style={{ fontSize: '0.95rem', padding: '12px 24px' }}
          >
            {isDe ? 'QR-PDF herunterladen (drucken)' : 'Download QR PDF (print)'}
          </button>
        </div>
        {selfCheckInBusy && (
          <p style={{ textAlign: 'center', fontSize: '0.8rem', color: 'var(--dex-gray-400)', margin: '10px 0 0' }}>
            {isDe ? 'Wird vorbereitet…' : 'Preparing…'}
          </p>
        )}
        </>)}
      </div>
      </div>

      {/* v31.1 — Abschnitt 2: Einchecken (ID, Suche, Liste). Event-Auswahl und
          Programmpunkt stehen jetzt in Abschnitt 1. */}
      <div style={{ order: 2 }}>
      {sectionLabel(2, isDe ? 'Einchecken' : 'Check in')}
      {/* v7.14: Live-Teilnehmerliste mit Foto / Position / Standort + Filter.
          Liste wird sofort beim Auswählen des Events geladen, kein "ab 2
          Zeichen tippen" mehr — der Helfer sieht direkt alle Leute, kann den
          gesuchten Eintrag scrollen oder das Suchfeld zum Filtern nutzen. */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        {/* v30.35/v31.1: Die Teilnehmer-ID steht in der Mail groß unter dem
            QR-Code — also gehört sie hier genauso groß hin. „Einchecken" checkt
            direkt ein (kein zweiter Dialog). Die Nummer filtert die Liste
            darunter live mit (v30.87). */}
        <form
          onSubmit={e => { e.preventDefault(); void checkInByParticipantId(); }}
          style={{ display: 'flex', gap: 8, justifyContent: 'center', alignItems: 'center', flexWrap: 'wrap', marginBottom: 6 }}
        >
          <label htmlFor="dex-checkin-id" style={{ fontSize: '0.78rem', letterSpacing: '0.06em', textTransform: 'uppercase', color: 'var(--dex-gray-500)', flex: '1 1 100%', textAlign: 'center' }}>
            {isDe ? 'Teilnehmer-ID aus der QR-Mail' : 'Attendee ID from the QR email'}
          </label>
          <input
            id="dex-checkin-id"
            value={idInput}
            onChange={e => { const v = e.target.value.replace(/\D/g, ''); setIdInput(v); setIdError(''); setNameSearchQuery(v); }}
            inputMode="numeric"
            pattern="[0-9]*"
            placeholder="17"
            disabled={!nameSearchEventId}
            aria-label={isDe ? 'Teilnehmer-ID' : 'Attendee ID'}
            style={{
              width: 130, padding: '12px 14px', textAlign: 'center',
              fontFamily: "'Courier New', Courier, monospace", fontSize: '1.5rem', fontWeight: 700,
              border: '2px solid var(--dex-gray-300)', borderRadius: 10,
            }}
          />
          <button type="submit" className="btn btn-primary" disabled={!idInput.trim() || isProcessing} style={{ fontSize: '1rem', padding: '12px 24px' }}>
            {agendaMode ? (isDe ? 'Anwesend erfassen' : 'Record') : (isDe ? 'Einchecken' : 'Check in')}
          </button>
        </form>
        {idError && (
          <p style={{ color: 'var(--dex-orange)', fontSize: '0.85rem', margin: '4px 0 10px', textAlign: 'center' }}>{idError}</p>
        )}
        <div style={{ borderTop: '1px solid var(--dex-gray-200)', margin: '12px 0 14px' }} />
        {/* v7.16: KPI-Bereich — angemeldet vs. eingecheckt, plus Quick-Filter
            "Nur offene anzeigen" der die Liste auf Angemeldet/QR versendet
            reduziert. Sichtbar nur wenn Event gewählt + Liste geladen. */}
        {nameSearchEventId && (searchRegsCache[nameSearchEventId] || []).length > 0 && (
          <div style={{
            display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(96px, 1fr))', gap: 8,
            marginBottom: 10,
          }}>
            <div style={{
              padding: '10px 12px', borderRadius: 10,
              background: 'rgba(21,101,192,0.08)', textAlign: 'center',
            }}>
              <div style={{ fontSize: '1.4rem', fontWeight: 700, color: '#1565c0', lineHeight: 1 }}>
                {checkInKpis.registered}
              </div>
              <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                Angemeldet
              </div>
            </div>
            <div style={{
              padding: '10px 12px', borderRadius: 10,
              background: 'rgba(134,188,37,0.12)', textAlign: 'center',
            }}>
              <div style={{ fontSize: '1.4rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', lineHeight: 1 }}>
                {checkInKpis.checkedIn}
                <span style={{ fontSize: '0.85rem', color: 'var(--dex-gray-400)', fontWeight: 500 }}>
                  /{checkInKpis.registered}
                </span>
              </div>
              <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                {agendaMode ? (agendaPoint ? `${isDe ? 'Anwesend' : 'Present'} · ${agendaPoint.title}` : (isDe ? 'Anwesend' : 'Present')) : 'Eingecheckt'}
              </div>
            </div>
            {/* v23.28: No-Show-Zähler. */}
            <div style={{
              padding: '10px 12px', borderRadius: 10,
              background: 'rgba(96,96,96,0.10)', textAlign: 'center',
            }}>
              <div style={{ fontSize: '1.4rem', fontWeight: 700, color: 'var(--dex-gray-700, #444)', lineHeight: 1 }}>
                {checkInKpis.noShow}
              </div>
              <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                No-Show
              </div>
            </div>
          </div>
        )}
        <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
          <input
            className="form-input"
            value={nameSearchQuery}
            onChange={e => setNameSearchQuery(e.target.value)}
            placeholder="Teilnehmer-ID, Vorname, Nachname oder E-Mail…"
            inputMode="text"
            disabled={!nameSearchEventId}
            style={{ flex: '1 1 200px', padding: '10px 14px', fontSize: '0.95rem' }}
          />
          <button
            type="button"
            onClick={() => setOnlyOpen(v => !v)}
            disabled={!nameSearchEventId}
            title={onlyOpen ? 'Alle Teilnehmer zeigen' : 'Nur offene Anmeldungen zeigen (Eingecheckte ausblenden)'}
            style={{
              padding: '8px 14px', borderRadius: 10,
              border: `1px solid ${onlyOpen ? 'var(--dex-green)' : 'var(--dex-gray-300)'}`,
              background: onlyOpen ? 'rgba(134,188,37,0.10)' : '#fff',
              color: onlyOpen ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-700)',
              cursor: nameSearchEventId ? 'pointer' : 'not-allowed',
              fontSize: '0.82rem', fontWeight: 600, fontFamily: 'inherit',
              whiteSpace: 'nowrap', flexShrink: 0,
              opacity: nameSearchEventId ? 1 : 0.5,
            }}
          >
            {onlyOpen ? '✓ Nur offene' : 'Nur offene'}
          </button>
        </div>
        {!nameSearchEventId && accessibleEvents.length > 1 && (
          <p style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
            Bitte zuerst ein Event wählen.
          </p>
        )}
        {isLoadingSearchRegs && (
          <p style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
            Teilnehmerliste wird geladen…
          </p>
        )}
        {searchLoadError && (
          <p style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-red)' }}>
            {searchLoadError}
            {/* v30.67: Ohne Cache-Eintrag ist ein zweiter Versuch möglich — der
                Knopf erspart den Seiten-Reload, der vorher der einzige Weg war. */}
            {nameSearchEventId && !isLoadingSearchRegs && (
              <button
                type="button"
                className="btn btn-secondary"
                style={{ marginLeft: 8, fontSize: '0.74rem', padding: '2px 8px' }}
                onClick={() => { void loadRegsForSearch(nameSearchEventId, { force: true }); }}
              >
                {isDe ? 'Erneut laden' : 'Reload'}
              </button>
            )}
          </p>
        )}
        {/* v30.67: Bei einem Ladefehler gibt es keine Liste — also auch kein
            „Keine Teilnehmer für dieses Event", das wäre eine Aussage über
            Daten, die nie gelesen wurden. */}
        {nameSearchEventId && !isLoadingSearchRegs && !searchLoadError && (
          <div style={{ marginTop: 12 }}>
            {/* v31.4: Der Stand der Liste — und ein Knopf, der wirklich lädt.
                Am Lauftag stehen zwei Tablets am selben Eingang; ohne diese
                Zeile weiß niemand, ob er den Check-in des Kollegen schon sieht.
                Älter als drei Minuten wird gedämpft rot: Dann klemmt der Lauf
                im Hintergrund, und das ist die einzige ehrliche Aussage, die
                die Seite dazu machen kann. */}
            {!!regsLoadedAt[nameSearchEventId] && (() => {
              const at = regsLoadedAt[nameSearchEventId];
              const stale = (nowTick - at) > 180000;
              return (
                <div className="dex-ui-toolbar" style={{ marginBottom: 8 }}>
                  <span style={{ fontSize: '0.74rem', color: stale ? 'var(--dex-red, #c00)' : 'var(--dex-gray-500)' }}>
                    {isDe ? 'Stand: ' : 'As of: '}{formatMarkTime(new Date(at).toISOString())}
                    <span className="dex-ui-muted" style={{ fontSize: '0.72rem' }}>
                      {isDe ? ' · lädt automatisch alle 45 Sekunden nach' : ' · reloads automatically every 45 seconds'}
                    </span>
                  </span>
                  <span className="dex-ui-toolbar-spacer" />
                  <button
                    type="button"
                    className="dex-ui-textbtn"
                    disabled={refreshBusy}
                    onClick={() => { void loadRegsForSearch(nameSearchEventId, { force: true }); }}
                  >
                    {refreshBusy ? (isDe ? 'Wird geladen…' : 'Loading…') : (isDe ? 'Aktualisieren' : 'Refresh')}
                  </button>
                </div>
              );
            })()}
            {/* v31.4: Ein gescheitertes Nachladen blendet die Liste NICHT aus —
                die alte Liste ist das Beste, was der Tisch hat. Gesagt wird es
                trotzdem, sonst hält er einen alten Stand für den aktuellen. */}
            {refreshError && (
              <p style={{ fontSize: '0.74rem', color: 'var(--dex-red, #c00)', margin: '0 0 8px' }}>
                {isDe ? 'Nachladen fehlgeschlagen: ' : 'Reload failed: '}{refreshError}
              </p>
            )}
            {searchHits.length === 0 ? (
              <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)', fontStyle: 'italic', margin: 0 }}>
                {nameSearchQuery.trim().length > 0
                  ? 'Kein Treffer — bitte anders schreiben oder per QR-Code einchecken.'
                  : 'Keine Teilnehmer für dieses Event.'}
              </p>
            ) : (
              <>
                <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-400)', marginBottom: 6 }}>
                  {searchHits.length} {searchHits.length === 1 ? 'Teilnehmer' : 'Teilnehmer'}
                  {nameSearchQuery.trim().length > 0 ? ' (gefiltert)' : ''}
                </div>
                <div style={{
                  display: 'flex', flexDirection: 'column', gap: 8,
                  // Maximalhöhe + Scroll, damit lange Events nicht die ganze Seite sprengen.
                  maxHeight: 480, overflowY: 'auto',
                  paddingRight: 4,
                }}>
                  {searchHits.map(reg => {
                    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || reg.ParticipantEmail || '-');
                    const status = reg.Status;
                    // v30.91: Im Programmpunkt-Modus heißt „drin": am gewählten Punkt erfasst.
                    const alreadyIn = agendaMode ? presentAt(reg, agendaPointId) : status === 'Eingecheckt';
                    const cancelled = status === 'Abgemeldet';
                    const waitlist = status === 'Warteliste';
                    // v31.2: Im Programmpunkt-Modus zählt der No-Show AM PUNKT
                    // (Marke) — der Event-Status bleibt sichtbar, wenn er es ist.
                    const noShowAtPoint = agendaMode && !!agendaPointId && !!parseAgendaNoShows(reg.AgendaCheckIns)[agendaPointId];
                    const noShow = status === 'No-Show' || noShowAtPoint;
                    const statusBg = alreadyIn ? 'rgba(134,188,37,0.15)'
                      : cancelled ? 'rgba(204,0,0,0.10)'
                      : noShow ? 'rgba(96,96,96,0.14)'
                      : waitlist ? 'rgba(237,139,0,0.12)'
                      : 'rgba(21,101,192,0.10)';
                    const statusFg = alreadyIn ? 'var(--dex-green-dark, #4a7c1f)'
                      : cancelled ? 'var(--dex-red, #c00)'
                      : noShow ? 'var(--dex-gray-700, #444)'
                      : waitlist ? 'var(--dex-orange, #ed8b00)'
                      : '#1565c0';
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const jobTitle = (reg as any).JobTitle || '';
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const location = (reg as any).Location || '';
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const ctx = (window as any).__dexSpfxContext;
                    const photoSrc = ctx && reg.ParticipantEmail
                      ? `${ctx.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(reg.ParticipantEmail)}`
                      : '';
                    return (
                      <div
                        key={reg.Id}
                        style={{
                          display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap',
                          padding: '10px 12px', border: '1px solid var(--dex-gray-200)',
                          borderRadius: 10, background: '#fff',
                        }}
                      >
                        {/* Foto-Avatar mit Initials-Fallback. Initialen liegen
                            im Hintergrund, das <img> deckt sie ab — schlägt
                            der Image-Load fehl, kommen sie zum Vorschein. */}
                        <div style={{
                          position: 'relative',
                          width: 44, height: 44, borderRadius: '50%',
                          flexShrink: 0, overflow: 'hidden',
                          background: 'var(--dex-gray-100)',
                          display: 'flex', alignItems: 'center', justifyContent: 'center',
                          fontWeight: 600, fontSize: '0.85rem', color: 'var(--dex-gray-500)',
                        }}>
                          <span style={{ pointerEvents: 'none' }}>
                            {(`${(reg.Vorname || reg.ParticipantName || '').charAt(0)}${(reg.Nachname || '').charAt(0)}`.toUpperCase()) || '?'}
                          </span>
                          {photoSrc && (
                            <img
                              src={photoSrc}
                              alt={name}
                              style={{
                                position: 'absolute', inset: 0,
                                width: '100%', height: '100%', objectFit: 'cover',
                              }}
                              onError={e => {
                                (e.currentTarget as HTMLImageElement).style.display = 'none';
                              }}
                            />
                          )}
                        </div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 600, fontSize: '0.9rem', color: 'var(--dex-gray-800)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                            {name}
                          </div>
                          {(jobTitle || location) && (
                            <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-600)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                              {[jobTitle, location].filter(Boolean).join(' · ')}
                            </div>
                          )}
                          <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                            {reg.ParticipantEmail}
                          </div>
                          {/* v31.4: Läuft die Nummer aus der QR-Mail von der
                              heutigen laufenden Nummer auseinander, stehen
                              beide da — sonst sucht der Helfer die Zahl vom
                              Handy in einer Liste, die eine andere zeigt.
                              Sind sie gleich, steht hier bewusst nichts —
                              außer bei einer Zahlen-Suche, s. qrRowNote. */}
                          {qrRowNote(reg, numericSearch) && (
                            <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', fontFamily: "'Courier New',Courier,monospace", whiteSpace: 'nowrap' }}>
                              {qrRowNote(reg, numericSearch)}
                            </div>
                          )}
                          {/* v30.53: Startnummer + Trikotgröße schon in der
                              Trefferliste — beim B2Run wird beides am selben
                              Tisch gebraucht wie der Check-in selbst. */}
                          {(() => {
                            const ex = extrasFor(reg, nameSearchEventId);
                            if (ex.length === 0) return null;
                            // v31.3: Der Ausweich-Satz steht auch hier in Warnfarbe
                            // unter den Chips — der Helfer sieht schon in der Liste,
                            // bei wem er nachdenken muss.
                            const chips = ex.filter(x => !x.tone);
                            const notes = ex.filter(x => !!x.tone);
                            return (
                              <>
                                {chips.length > 0 && (
                                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 4 }}>
                                    {chips.map((x, i) => (
                                      <span key={i} style={{
                                        display: 'inline-flex', alignItems: 'baseline', gap: 5,
                                        fontSize: '0.7rem', padding: '2px 8px', borderRadius: 6,
                                        background: x.strong ? 'rgba(134,188,37,0.14)' : 'var(--dex-gray-100, #f5f5f5)',
                                        color: 'var(--dex-gray-700)',
                                      }}>
                                        <span style={{ color: 'var(--dex-gray-500)' }}>{x.label}</span>
                                        <strong style={{
                                          fontFamily: x.strong ? "'Courier New',Courier,monospace" : 'inherit',
                                          color: x.strong ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-800)',
                                        }}>{x.value}</strong>
                                      </span>
                                    ))}
                                  </div>
                                )}
                                {notes.map((x, i) => (
                                  <div key={`tone-${i}`} style={{
                                    marginTop: 4, padding: '4px 8px', borderRadius: 6, fontSize: '0.72rem', lineHeight: 1.4,
                                    whiteSpace: 'normal',
                                    background: x.tone === 'danger' ? 'var(--dex-red-light, #fce8e6)' : '#fff7e6',
                                    border: `1px solid ${x.tone === 'danger' ? '#f2b1ab' : '#f5c77a'}`,
                                    color: x.tone === 'danger' ? '#9b2018' : '#7a4a00',
                                  }}>
                                    <strong>{x.label}</strong> — {x.value}
                                  </div>
                                ))}
                              </>
                            );
                          })()}
                        </div>
                        <span style={{
                          fontSize: '0.7rem', padding: '3px 8px', borderRadius: 999,
                          background: statusBg, color: statusFg, fontWeight: 600, whiteSpace: 'nowrap',
                        }}>{noShowAtPoint && status !== 'No-Show' ? (isDe ? 'No-Show hier' : 'No-show here') : status}</span>
                        {/* v23.28: Check-in UND No-Show nebeneinander. */}
                        <div style={{ display: 'flex', gap: 6, flexShrink: 0, flex: isMobile ? '1 1 100%' : undefined }}>
                          <button
                            type="button"
                            className="btn btn-primary"
                            style={{ fontSize: '0.78rem', padding: '6px 12px', whiteSpace: 'nowrap' }}
                            disabled={alreadyIn || cancelled || isProcessing}
                            onClick={() => startManualCheckInFromSearch(reg)}
                          >
                            {alreadyIn ? (agendaMode ? (isDe ? '✓ Anwesend' : '✓ Present') : '✓ Eingecheckt') : cancelled ? 'Abgemeldet' : (agendaMode ? (isDe ? 'Anwesend erfassen' : 'Record') : 'Einchecken')}
                          </button>
                          <button
                            type="button"
                            className="btn btn-secondary"
                            style={{ fontSize: '0.78rem', padding: '6px 12px', whiteSpace: 'nowrap', color: 'var(--dex-gray-700, #444)' }}
                            disabled={cancelled || noShow || alreadyIn || isProcessing}
                            onClick={() => { void markNoShowFromSearch(reg); }}
                            title={isDe
                              ? (agendaMode ? 'Nicht erschienen — für diesen Punkt oder das ganze Event' : 'Teilnehmer als nicht erschienen markieren')
                              : (agendaMode ? 'No-show — for this point or the whole event' : 'Mark attendee as a no-show')}
                          >
                            {noShow ? (isDe ? 'No-Show' : 'No-show') : 'No-Show'}
                          </button>
                          {/* v31.4: Die Trikot-Ausgabe steht in derselben Zeile wie
                              Einchecken und No-Show — am Tisch passiert beides in
                              einem Handgriff. Wer sein Trikot hat, sieht statt des
                              Knopfs die Tatsache (Größe, Uhrzeit) und kann sie
                              zurücknehmen. Bewusst NICHT gesperrt bei Abgemeldet/
                              No-Show: Ein ausgegebenes Trikot ist aus dem Karton,
                              egal was der Status sagt. */}
                          {(() => {
                            const si = shirtDeskInfoFor(reg, nameSearchEventId);
                            if (!si) return null;
                            const ev = events.find(e => e.id === nameSearchEventId);
                            const sub = (ev && ev.subsiteUrl) || '';
                            if (si.issued) {
                              return (
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                                  <span className="dex-ui-pill dex-ui-pill--green" title={isDe ? 'Trikot ausgegeben' : 'Shirt handed out'}>
                                    {isDe ? 'Trikot' : 'Shirt'} {si.issued.size}{si.issued.at ? ` · ${formatMarkTime(si.issued.at)}` : ''}
                                  </span>
                                  <button
                                    type="button"
                                    className="dex-ui-textbtn dex-ui-textbtn--danger"
                                    disabled={shirtBusy}
                                    onClick={() => { void clearShirtIssue({ regId: reg.Id, name, eventId: nameSearchEventId, subsiteUrl: sub }); }}
                                  >
                                    {isDe ? 'Rücknehmen' : 'Undo'}
                                  </button>
                                </span>
                              );
                            }
                            return (
                              <button
                                type="button"
                                className="btn btn-secondary"
                                style={{ fontSize: '0.78rem', padding: '6px 12px', whiteSpace: 'nowrap' }}
                                disabled={shirtBusy}
                                onClick={() => openShirtAsk(reg, name, si)}
                                title={isDe ? 'Festhalten, welche Größe diese Person bekommen hat' : 'Record which size this person received'}
                              >
                                {isDe ? 'Trikot ausgeben' : 'Hand out shirt'}
                              </button>
                            );
                          })()}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </div>
        )}
      </div>
      </div>

      {/* v31.1 — Abschnitt 3: Letzte Check-ins, mit Rückgängig (Nutzer
          07.09.2026). Nur was auf diesem Gerät eingecheckt wurde — der Revert
          setzt den gemerkten vorherigen Status bzw. entfernt die Punkt-
          Anwesenheit.
          v31.4: Die Liste überlebt jetzt das Verlassen der Seite (localStorage
          je Event, 12 Stunden) — deshalb heißt sie nicht mehr „diese Sitzung",
          das wäre nach dem Zurückgehen falsch. */}
      <div style={{ order: 3 }}>
        {sectionLabel(3, isDe ? 'Letzte Check-ins' : 'Recent check-ins')}
        <div className="card" style={{ padding: recentCheckIns.length ? '14px 20px' : '14px 20px', marginBottom: 16 }}>
          {recentCheckIns.length === 0 ? (
            <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
              {isDe
                ? 'Noch kein Check-in in den letzten 12 Stunden. Jeder Check-in erscheint hier und lässt sich zurücknehmen — auch nachdem du die Seite zwischendurch verlassen hast.'
                : 'No check-in in the past 12 hours. Every check-in appears here and can be reverted — also after you have left the page in between.'}
            </p>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4, maxHeight: 260, overflowY: 'auto' }}>
              {recentCheckIns.map(e => (
                <div key={e.key} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '6px 8px', borderRadius: 8, background: 'var(--dex-gray-50, #fafafa)' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontVariantNumeric: 'tabular-nums', width: 44, flexShrink: 0 }}>{formatMarkTime(e.at)}</span>
                  <span style={{ fontWeight: 600, fontSize: '0.88rem', flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {e.name}
                    {e.agendaLabel && <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}> · {e.agendaLabel}</span>}
                  </span>
                  {/* v31.2: No-Shows stehen mit in der Liste — als graue Pille erkennbar. */}
                  {e.kind === 'noshow' && (
                    <span className="dex-ui-pill dex-ui-pill--gray">No-Show</span>
                  )}
                  {/* v31.4: Ausgaben ebenso — mit der Größe, sonst weiß niemand,
                      was „Rückgängig" hier zurücknimmt. */}
                  {e.kind === 'shirt' && (
                    <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'Trikot' : 'Shirt'} {e.shirtSize}</span>
                  )}
                  <button type="button" className="btn btn-secondary" disabled={!!undoBusyKey} onClick={() => { void undoCheckIn(e); }} style={{ fontSize: '0.76rem', padding: '4px 10px', whiteSpace: 'nowrap' }}
                    title={isDe
                      ? (e.kind === 'shirt'
                        ? `Trikot-Ausgabe (${e.shirtSize}) zurücknehmen — die Größe zählt wieder zum Bestand`
                        : e.agendaItemId
                          ? (e.kind === 'noshow' ? 'No-Show an diesem Punkt entfernen' : 'Anwesenheit an diesem Punkt entfernen')
                          : `Status zurück auf „${(e.prevStatus === 'QR versendet' || (e.kind === 'noshow' && e.prevStatus === 'Eingecheckt')) ? e.prevStatus : 'Angemeldet'}“`)
                      : (e.kind === 'shirt' ? 'Revert this shirt handout' : e.kind === 'noshow' ? 'Revert this no-show' : 'Revert this check-in')}>
                    {undoBusyKey === e.key ? '…' : (isDe ? 'Rückgängig' : 'Undo')}
                  </button>
                </div>
              ))}
            </div>
          )}
          {recentCheckIns.length > 0 && (
            <p className="dex-ui-muted" style={{ fontSize: '0.72rem', margin: '8px 0 0' }}>
              {isDe
                ? 'Die letzten 12 Stunden auf diesem Gerät, für dieses Event.'
                : 'The past 12 hours on this device, for this event.'}
            </p>
          )}
        </div>
      </div>
      </div>

      {/* v7.14: Die alte "Manuell QR-Code als String tippen"-Card ist raus.
          Die Live-Teilnehmerliste oben deckt das Manuelle Einchecken
          benutzerfreundlicher ab. */}

      {/* v31.2: Rückfrage im Programmpunkt-Modus — Nutzer 07.09.2026: „man soll
          gefragt werden, ob das für das ganze Event No-Show ist oder nur für
          den Programmpunkt". Zwei Kacheln, jede sagt, was sie tut. */}
      <Modal
        open={!!noShowAsk}
        onClose={() => setNoShowAsk(null)}
        maxWidth={520}
        title={isDe ? `${noShowAsk ? noShowAsk.name : ''} — nicht erschienen` : `${noShowAsk ? noShowAsk.name : ''} — no-show`}
        subtitle={isDe ? 'Wofür gilt der No-Show?' : 'What does the no-show apply to?'}
        footer={<button type="button" className="btn btn-secondary" onClick={() => setNoShowAsk(null)}>{isDe ? 'Abbrechen' : 'Cancel'}</button>}
      >
        <div className="dex-ui-stack">
          <button
            type="button"
            className="dex-ui-choice"
            onClick={() => { const a = noShowAsk; setNoShowAsk(null); if (a) void applyPointNoShow(a.reg, a.name); }}
          >
            <span className="dex-ui-choice-body">
              <span className="dex-ui-choice-title">{isDe ? `Nur für „${agendaPoint ? agendaPoint.title : ''}“` : `Only for “${agendaPoint ? agendaPoint.title : ''}”`}</span>
              <span className="dex-ui-choice-desc">
                {isDe
                  ? `Die Person fehlt bei diesem ${agendaTermSingular}. Ihr Event-Status bleibt, beim nächsten ${agendaTermSingular} kann sie wieder erfasst werden.`
                  : `The person is missing at this ${agendaTermSingular.toLowerCase()}. The event status stays; they can be recorded at the next one.`}
              </span>
            </span>
          </button>
          <button
            type="button"
            className="dex-ui-choice"
            onClick={() => { const a = noShowAsk; setNoShowAsk(null); if (a) void applyEventNoShow(a.reg, a.name); }}
          >
            <span className="dex-ui-choice-body">
              <span className="dex-ui-choice-title">{isDe ? 'Für das ganze Event' : 'For the whole event'}</span>
              <span className="dex-ui-choice-desc">
                {isDe
                  ? 'Der Event-Status wird auf „No-Show“ gesetzt — die Person gilt für das gesamte Event als nicht erschienen.'
                  : 'The event status is set to “No-Show” — the person counts as absent for the entire event.'}
              </span>
            </span>
          </button>
          <p className="dex-ui-muted" style={{ margin: 0 }}>
            {isDe ? 'Beides lässt sich unter „Letzte Check-ins“ zurücknehmen.' : 'Both can be reverted under “Recent check-ins”.'}
          </p>
        </div>
      </Modal>

      {/* v31.4: Trikot-Ausgabe aus der Trefferliste. Dieselbe Größenwahl wie in
          der Bestätigungskarte — vorbelegt mit dem Gegenvorschlag, sonst der
          Wunschgröße; „Andere Größe" für den Griff in den falschen Karton. */}
      <Modal
        open={!!shirtAsk}
        onClose={() => setShirtAsk(null)}
        maxWidth={520}
        title={isDe ? `${shirtAsk ? shirtAsk.name : ''} — Trikot ausgeben` : `${shirtAsk ? shirtAsk.name : ''} — hand out shirt`}
        subtitle={isDe ? 'Welche Größe hast du herausgegeben?' : 'Which size did you hand out?'}
        footer={<>
          <button type="button" className="btn btn-secondary" onClick={() => setShirtAsk(null)}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
          <button
            type="button"
            className="btn btn-primary"
            disabled={shirtBusy || !askShirtSize.trim()}
            onClick={() => {
              const a = shirtAsk;
              const size = askShirtSize;
              setShirtAsk(null);
              if (!a) return;
              const ev = events.find(e => e.id === a.eventId);
              void applyShirtIssue(
                { regId: a.reg.Id, name: a.name, eventId: a.eventId, subsiteUrl: (ev && ev.subsiteUrl) || '', prevStatus: a.reg.Status },
                size,
              );
            }}
          >
            {isDe ? 'Ausgabe festhalten' : 'Record handout'}
          </button>
        </>}
      >
        {shirtAsk && (() => {
          const si = shirtDeskInfoFor(shirtAsk.reg, shirtAsk.eventId);
          return (
            <div className="dex-ui-stack">
              <div className="dex-ui-inline">
                <ShirtSizePicker key={shirtAsk.reg.Id} sizes={si ? si.sizes : []} value={askShirtSize} onChange={setAskShirtSize} isDe={isDe} />
              </div>
              {si && (si.wish || si.proposal) && (
                <p className="dex-ui-muted" style={{ margin: 0 }}>
                  {si.wish && <>{isDe ? 'Wunschgröße: ' : 'Wished size: '}<strong>{si.wish}</strong></>}
                  {si.proposal && <>{si.wish ? ' · ' : ''}{isDe ? 'vorgeschlagen: ' : 'proposed: '}<strong>{si.proposal}</strong></>}
                </p>
              )}
              <div className="dex-ui-callout dex-ui-callout--neutral">
                <span>
                  {isDe
                    ? <>Die Größe wird dauerhaft vom Bestand abgezogen — auch wenn die Person später abgemeldet oder als No-Show markiert wird. Zurücknehmen geht über &bdquo;Rücknehmen&ldquo; in der Zeile oder unter &bdquo;Letzte Check-ins&ldquo;.</>
                    : <>The size is permanently deducted from the stock — even if the person is cancelled or marked as a no-show later. You can revert it via &ldquo;Undo&rdquo; in the row or under &ldquo;Recent check-ins&rdquo;.</>}
                </span>
              </div>
            </div>
          );
        })()}
      </Modal>
    </div>
  );
}
