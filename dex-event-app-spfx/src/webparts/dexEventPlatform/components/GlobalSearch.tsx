/**
 * v22.50/v22.51: Globale Such-Leiste im Header (für Admins UND Organizer
 * eigener Events). Durchsucht geclustert: Events, Aktionen (pro verwaltbarem
 * Event), Seiten/Funktionen, Handbuch und Teilnehmer — und springt direkt
 * dorthin.
 *
 * Berechtigungs-Scoping (strikt): Es werden nur Events, Aktionen und
 * Teilnehmer angezeigt, die der eingeloggte Nutzer sehen/verwalten darf
 * (Admin: alle; Organizer: nur eigene/Co-Organizer-Events). Seiten/Handbuch
 * werden rollen-gefiltert.
 *
 * v22.51:
 * - Teilnehmer-Suche läuft über die Teilnehmerlisten der verwaltbaren Events
 *   (Full Control auf der Subsite) statt über die ILS-gesperrte
 *   DEX_Participants-Liste — sonst sieht ein Nicht-Owner-Admin nur die selbst
 *   angelegten Zeilen.
 * - Handbuch-Suche durchsucht den gesamten Schritt-Text (nicht nur
 *   Titel + Kurzbeschreibung), damit z.B. „Excel"/„SharePoint"/„Export" auch
 *   im Fließtext gefunden werden.
 * - Tippfehler-tolerant (Levenshtein) je Suchwort.
 * - Eingeklappt: nur ein Such-Icon; erst beim Klick klappt die Leiste auf.
 */
import * as React from 'react';
import { Search, X } from './Icons';
import { useNavigation, Page } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useLanguage } from '../context/LanguageContext';
import { getManualSections } from './manual/handbookContent';
import { ManualSection } from './manual/types';
import { DeloitteEvent } from '../types';

type Gate = 'all' | 'manage' | 'admin';

interface PageEntry { page: Page; de: string; en: string; kw: string[]; gate: Gate }
const PAGE_CATALOG: PageEntry[] = [
  { page: 'register', de: 'Events finden / Anmelden', en: 'Find events / Register', kw: ['event', 'events', 'anmelden', 'registrieren', 'finden', 'register'], gate: 'all' },
  { page: 'my-events', de: 'Meine Events', en: 'My Events', kw: ['meine events', 'my events', 'anmeldungen', 'mein qr', 'qr-code'], gate: 'all' },
  { page: 'admin', de: 'Organizer Center', en: 'Organizer Center', kw: ['organizer', 'admin', 'verwalten', 'teilnehmer', 'event verwalten', 'admin center'], gate: 'manage' },
  { page: 'create-event', de: 'Neues Event erstellen', en: 'Create new event', kw: ['neu', 'erstellen', 'anlegen', 'create', 'wizard', 'neues event'], gate: 'manage' },
  { page: 'check-in', de: 'Check-in', en: 'Check-in', kw: ['checkin', 'check-in', 'scan', 'scannen', 'einchecken'], gate: 'manage' },
  { page: 'flowcharts', de: 'Prozess-Übersicht', en: 'Process overview', kw: ['prozess', 'flowchart', 'ablauf', 'diagramm', 'process'], gate: 'manage' },
  { page: 'participants', de: 'Teilnehmer-Übersicht', en: 'Participants overview', kw: ['teilnehmer-übersicht', 'participants', 'register', 'cross-event'], gate: 'admin' },
  { page: 'role-matrix', de: 'Rollen-Matrix', en: 'Role matrix', kw: ['rollen', 'matrix', 'rechte', 'berechtigungen', 'permissions'], gate: 'admin' },
  { page: 'settings', de: 'Einstellungen', en: 'Settings', kw: ['einstellungen', 'settings', 'rollen verwalten', 'templates', 'logos', 'reseed'], gate: 'admin' },
  { page: 'manual', de: 'Handbuch', en: 'Manual', kw: ['handbuch', 'hilfe', 'manual', 'anleitung', 'help'], gate: 'all' },
  { page: 'profile', de: 'Profil', en: 'Profile', kw: ['profil', 'profile', 'konto', 'account'], gate: 'all' },
];

interface ActionEntry { key: string; de: string; en: string; kw: string[]; gate: Gate }
const ACTION_CATALOG: ActionEntry[] = [
  { key: 'export', de: 'Teilnehmerliste exportieren (Excel)', en: 'Export attendee list (Excel)', kw: ['export', 'exportieren', 'excel', 'csv', 'sharepoint', 'teilnehmerliste', 'liste', 'download', 'tabelle'], gate: 'manage' },
  { key: 'qr', de: 'QR-Codes versenden', en: 'Send QR codes', kw: ['qr', 'qr-code', 'qr code', 'checkin-code', 'code versenden'], gate: 'manage' },
  { key: 'massmail', de: 'E-Mail versenden', en: 'Send email', kw: ['massenmail', 'mail', 'email', 'e-mail', 'nachricht', 'rundmail', 'anschreiben', 'senden', 'versenden'], gate: 'manage' },
  { key: 'invite', de: 'Einladungsmail verschicken', en: 'Send invitation mail', kw: ['einladung', 'einladen', 'invite', 'anmelde-link', 'invitation'], gate: 'manage' },
  { key: 'audit', de: 'Audit-Log / Änderungsprotokoll', en: 'Audit log / change log', kw: ['audit', 'log', 'protokoll', 'historie', 'änderungen', 'changelog'], gate: 'manage' },
  { key: 'selfcheckin', de: 'Self-Check-in einstellen', en: 'Configure self-check-in', kw: ['self-check-in', 'selfcheckin', 'self check', 'qr-plakat', 'pdf', 'live-anzeige'], gate: 'manage' },
  { key: 'idreorder', de: 'TeilnehmerIDs neu vergeben', en: 'Reassign attendee IDs', kw: ['id', 'ids', 'nummer', 'teilnehmer-id', 'reorder', 'renummerieren'], gate: 'manage' },
  { key: 'overbook', de: 'Überbuchung prüfen', en: 'Check overbooking', kw: ['überbuchung', 'overbook', 'kapazität', 'over capacity'], gate: 'manage' },
  { key: 'accessfix', de: 'Fremd-Anmeldungen: Zugriff reparieren', en: 'Repair proxy-registration access', kw: ['zugriff', 'author', 'reparieren', 'sichtbarkeit', 'meine events', 'access'], gate: 'manage' },
  { key: 'fixcols', de: 'Spalten fixen (Teilnehmerliste)', en: 'Fix columns (attendee list)', kw: ['spalten', 'spalte', 'fix', 'columns', 'felder'], gate: 'admin' },
];

interface SearchHit { id: string; primary: string; secondary?: string; onSelect: () => void }
interface Cluster { key: string; label: string; hits: SearchHit[] }
// Aus den Event-Teilnehmerlisten zusammengetragene Person (berechtigungs-gescoped).
interface RegPart { name: string; email: string; status: string; eventId: string; eventTitle: string }

const norm = (s: string): string => (s || '').toLowerCase().trim();

// Levenshtein-Distanz mit Früh-Abbruch (gibt max+1 zurück, sobald klar ist,
// dass die Distanz das Limit überschreitet — spart Arbeit bei No-Matches).
function lev(a: string, b: string, max: number): number {
  const al = a.length;
  const bl = b.length;
  if (Math.abs(al - bl) > max) return max + 1;
  let prev: number[] = [];
  for (let j = 0; j <= bl; j++) prev[j] = j;
  for (let i = 1; i <= al; i++) {
    const cur: number[] = [i];
    let best = i;
    const ac = a.charCodeAt(i - 1);
    for (let j = 1; j <= bl; j++) {
      const cost = ac === b.charCodeAt(j - 1) ? 0 : 1;
      const v = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
      cur[j] = v;
      if (v < best) best = v;
    }
    if (best > max) return max + 1; // ganze Zeile schon über Limit
    prev = cur;
  }
  return prev[bl];
}

// Erlaubte Tippfehler je Token-Länge: kurze Begriffe streng, lange tolerant.
const fuzzyTol = (len: number): number => (len <= 3 ? 0 : len <= 6 ? 1 : 2);

// Token trifft den Heuhaufen, wenn es als Teilstring vorkommt ODER fuzzy
// (Levenshtein ≤ Toleranz) auf ein Wort bzw. dessen Wort-Anfang passt.
function tokenMatches(token: string, haystack: string, hayWords: string[]): boolean {
  if (!token) return true;
  if (haystack.indexOf(token) >= 0) return true; // exakter Teilstring (schnell)
  const t = fuzzyTol(token.length);
  if (t === 0) return false;
  for (const w of hayWords) {
    if (!w) continue;
    if (lev(token, w, t) <= t) return true;
    // Längere Wörter: Token gegen den Wort-Anfang prüfen (Tippfehler am Ende
    // langer Begriffe, z.B. "teilnehmerlist" → "teilnehmerliste").
    if (w.length > token.length + t) {
      if (lev(token, w.slice(0, token.length + t), t) <= t) return true;
    }
  }
  return false;
}

const matchAll = (tokens: string[], haystack: string, hayWords: string[]): boolean =>
  tokens.every(t => tokenMatches(t, haystack, hayWords));

const words = (haystack: string): string[] => haystack.split(/[^a-z0-9äöüß]+/).filter(Boolean);

// Rekursiver Text-Extraktor für die React-Knoten der Handbuch-Schritte —
// damit der gesamte Fließtext (z.B. „Excel", „SharePoint", „exportiert")
// durchsuchbar wird, nicht nur Titel + Kurzbeschreibung.
function nodeText(node: React.ReactNode): string {
  if (node === null || node === undefined || typeof node === 'boolean') return '';
  if (typeof node === 'string' || typeof node === 'number') return String(node);
  if (Array.isArray(node)) return node.map(nodeText).join(' ');
  if (React.isValidElement(node)) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return nodeText((node.props as any)?.children);
  }
  return '';
}

function manualHaystack(s: ManualSection): string {
  const parts: string[] = [s.title, s.description, s.keywords || ''];
  for (const block of s.perspectives || []) {
    if (block.title) parts.push(block.title);
    parts.push(nodeText(block.intro));
    for (const step of block.steps || []) {
      parts.push(step.title, nodeText(step.description), step.tip || '', step.warning || '', step.visualHint || '');
    }
  }
  return norm(parts.join(' '));
}

export default function GlobalSearch(): React.ReactElement | null {
  const { navigate } = useNavigation();
  const { isAdmin, isOrganizer, originalIsAdmin } = useRoles();
  const { events, getAllRegistrations } = useEvents();
  const { currentUser } = useCurrentUser();
  const { locale } = useLanguage();
  const isDe = locale === 'de';

  const emailLc = (currentUser?.email || '').toLowerCase();
  const adminLike = originalIsAdmin || isAdmin;
  const isOrganizerFor = React.useCallback((e: DeloitteEvent): boolean => {
    if (!emailLc) return false;
    if ((e.organizerEmails || []).some(x => (x || '').toLowerCase() === emailLc)) return true;
    return (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === emailLc);
  }, [emailLc]);

  // Verwaltbare Events (Berechtigungs-Scope): Admin → alle; sonst nur eigene.
  const managedEvents = React.useMemo<DeloitteEvent[]>(() => {
    const list = adminLike ? (events || []) : (events || []).filter(isOrganizerFor);
    return list.filter(e => !e.isDemoShowcase);
  }, [events, adminLike, isOrganizerFor]);

  const canManageAny = adminLike || managedEvents.length > 0 || isOrganizer;
  const visible = canManageAny;
  const role: 'Admin' | 'Organizer' | 'User' = adminLike ? 'Admin' : (canManageAny ? 'Organizer' : 'User');

  const [expanded, setExpanded] = React.useState(false);
  const [query, setQuery] = React.useState('');
  const [open, setOpen] = React.useState(false);
  const [activeIdx, setActiveIdx] = React.useState(0);
  const [regParts, setRegParts] = React.useState<RegPart[] | null>(null);
  const [partLoading, setPartLoading] = React.useState(false);
  const rootRef = React.useRef<HTMLDivElement | null>(null);
  const inputRef = React.useRef<HTMLInputElement | null>(null);

  // Handbuch-Heuhaufen einmal pro Sprache/Rolle vorberechnen (Text-Extraktion
  // ist teuer — nicht bei jedem Tastendruck wiederholen).
  const manualIndex = React.useMemo(() => {
    if (!visible) return [] as Array<{ s: ManualSection; hay: string; words: string[] }>;
    try {
      const sections = getManualSections(isDe ? 'de' : 'en').filter(s =>
        adminLike ? true
          : role === 'Organizer'
            ? (s.visibleFor.indexOf('Organizer') >= 0 || s.visibleFor.indexOf('User') >= 0)
            : s.visibleFor.indexOf('User') >= 0,
      );
      return sections.map(s => { const hay = manualHaystack(s); return { s, hay, words: words(hay) }; });
    } catch { return []; }
  }, [isDe, adminLike, role, visible]);

  // Teilnehmer der verwaltbaren Events lazy laden, sobald getippt wird. Quelle
  // sind die Event-Teilnehmerlisten (Full Control auf der Subsite) — NICHT die
  // ILS-gesperrte DEX_Participants-Liste.
  React.useEffect(() => {
    if (!visible || !expanded) return;
    if (query.trim().length < 2) return;
    if (regParts !== null || partLoading) return;
    setPartLoading(true);
    const evs = managedEvents.slice();
    let idx = 0;
    const out: RegPart[] = [];
    const CONC = 6;
    const worker = async (): Promise<void> => {
      while (idx < evs.length) {
        const e = evs[idx++];
        try {
          const regs = await getAllRegistrations(e.id);
          for (const r of regs) {
            const name = `${r.Vorname || ''} ${r.Nachname || ''}`.trim() || r.ParticipantName || r.ParticipantEmail || r.Title;
            out.push({ name, email: r.ParticipantEmail || r.Title || '', status: r.Status || '', eventId: e.id, eventTitle: e.title });
          }
        } catch { /* einzelnes Event best-effort überspringen */ }
      }
    };
    Promise.all(Array.from({ length: Math.min(CONC, evs.length) }, worker))
      .then(() => setRegParts(out))
      .catch(() => setRegParts([]))
      .finally(() => setPartLoading(false));
  }, [query, visible, expanded, regParts, partLoading, managedEvents, getAllRegistrations]);

  // Outside-Click: Dropdown schließen, bei leerer Suche auch wieder einklappen.
  React.useEffect(() => {
    const onDown = (e: MouseEvent): void => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) {
        setOpen(false);
        if (!query.trim()) setExpanded(false);
      }
    };
    window.addEventListener('mousedown', onDown);
    return () => window.removeEventListener('mousedown', onDown);
  }, [query]);

  // Beim Aufklappen das Eingabefeld fokussieren.
  React.useEffect(() => { if (expanded) inputRef.current?.focus(); }, [expanded]);

  const go = React.useCallback((fn: () => void): void => {
    setOpen(false); setExpanded(false); setQuery('');
    fn();
  }, []);

  const clusters = React.useMemo<Cluster[]>(() => {
    const q = norm(query);
    if (q.length < 2) return [];
    const tokens = q.split(/\s+/).filter(Boolean);
    const out: Cluster[] = [];

    // ---- Events ----
    const evHits: SearchHit[] = [];
    for (const e of managedEvents) {
      const hay = norm(`${e.title} ${e.eventNumber || ''} ${e.location || ''}`);
      if (matchAll(tokens, hay, words(hay))) {
        evHits.push({
          id: `ev-${e.id}`,
          primary: e.title || (isDe ? 'Event ohne Titel' : 'Untitled event'),
          secondary: [e.eventNumber ? `Nr. ${e.eventNumber}` : '', e.location || '', e.isFictive ? (isDe ? 'Entwurf' : 'Draft') : ''].filter(Boolean).join(' · '),
          onSelect: () => navigate('admin', e.id),
        });
      }
      if (evHits.length >= 6) break;
    }
    if (evHits.length) out.push({ key: 'events', label: 'Events', hits: evHits });

    // ---- Aktionen (kombiniert Aktion × verwaltbares Event) ----
    const actHits: SearchHit[] = [];
    const sortedEvents = managedEvents.slice().sort((a, b) => (b.startDate || '').localeCompare(a.startDate || ''));
    outer:
    for (const a of ACTION_CATALOG) {
      if (a.gate === 'admin' && !adminLike) continue;
      const actLabel = isDe ? a.de : a.en;
      const actHay = norm(`${a.de} ${a.en} ${a.kw.join(' ')}`);
      const actWords = words(actHay);
      for (const e of sortedEvents) {
        const evHay = norm(`${e.title} ${e.eventNumber || ''} ${e.location || ''}`);
        const combined = `${actHay} ${evHay}`;
        if (!matchAll(tokens, combined, words(combined))) continue;
        // Mindestens ein Token muss die Aktion treffen (sonst reiner Event-Treffer).
        if (!tokens.some(t => tokenMatches(t, actHay, actWords))) continue;
        actHits.push({
          id: `act-${a.key}-${e.id}`,
          primary: actLabel,
          secondary: e.title,
          onSelect: () => { try { window.localStorage.setItem('dex_search_focus_action', a.key); } catch { /* */ } navigate('admin', e.id); },
        });
        if (actHits.length >= 12) break outer;
      }
    }
    if (actHits.length) out.push({ key: 'actions', label: isDe ? 'Aktionen' : 'Actions', hits: actHits });

    // ---- Seiten & Funktionen ----
    const pageHits: SearchHit[] = [];
    for (const p of PAGE_CATALOG) {
      if (p.gate === 'admin' && !adminLike) continue;
      if (p.gate === 'manage' && !canManageAny) continue;
      const label = isDe ? p.de : p.en;
      const hay = norm(`${p.de} ${p.en} ${p.kw.join(' ')}`);
      if (matchAll(tokens, hay, words(hay))) {
        pageHits.push({ id: `pg-${p.page}`, primary: label, onSelect: () => navigate(p.page) });
      }
    }
    if (pageHits.length) out.push({ key: 'pages', label: isDe ? 'Seiten & Funktionen' : 'Pages & functions', hits: pageHits });

    // ---- Handbuch (gesamter Schritt-Text) ----
    // v24.66: Relevanz-Sortierung. Vorher wurden einfach die ersten 6 Treffer in
    // Array-Reihenfolge genommen — der eigentlich beste Artikel (Suchbegriff im
    // Titel) wurde von früher einsortierten Schwach-Treffern aus den Top 6
    // verdrängt (z.B. „E-Mail an alle Teilnehmer senden" tauchte bei „email
    // senden" gar nicht auf). Jetzt: alle Treffer scoren (Titel > Stichwort >
    // Beschreibung), absteigend sortieren, Top 6.
    const stripHyphen = (s: string): string => s.replace(/-/g, '');
    const manScored: Array<{ m: { s: ManualSection }; score: number }> = [];
    for (const m of manualIndex) {
      if (!matchAll(tokens, m.hay, m.words)) continue;
      const tHay = stripHyphen(norm(m.s.title));
      const kHay = stripHyphen(norm(m.s.keywords || ''));
      const dHay = stripHyphen(norm(m.s.description));
      let score = 1;
      let allInTitle = tokens.length > 0;
      for (const tk of tokens) {
        const t = stripHyphen(tk);
        if (t && tHay.indexOf(t) >= 0) { score += 20; }
        else { allInTitle = false; if (t && kHay.indexOf(t) >= 0) score += 8; else if (t && dHay.indexOf(t) >= 0) score += 4; }
      }
      if (allInTitle) score += 30; // alle Suchbegriffe im Titel = klarer Top-Treffer
      manScored.push({ m, score });
    }
    manScored.sort((a, b) => b.score - a.score);
    const manHits: SearchHit[] = manScored.slice(0, 6).map(({ m }) => ({
      id: `man-${m.s.id}`,
      primary: m.s.title,
      secondary: m.s.description,
      onSelect: () => { try { window.localStorage.setItem('dex_open_manual_section', m.s.id); } catch { /* */ } navigate('manual'); },
    }));
    if (manHits.length) out.push({ key: 'manual', label: isDe ? 'Handbuch' : 'Manual', hits: manHits });

    // ---- Teilnehmer (aus den Teilnehmerlisten der verwaltbaren Events) ----
    // Eine Zeile PRO Person × Event — taucht jemand in mehreren Events auf,
    // stehen alle Events untereinander, damit man gezielt eines anklicken kann.
    if (regParts && regParts.length) {
      const seen = new Set<string>();
      const matched: RegPart[] = [];
      for (const rp of regParts) {
        const hay = norm(`${rp.name} ${rp.email}`);
        if (!matchAll(tokens, hay, words(hay))) continue;
        const key = `${(rp.email || rp.name).toLowerCase()}::${rp.eventId}`;
        if (seen.has(key)) continue;
        seen.add(key);
        matched.push(rp);
      }
      // Gleiche Person zusammen gruppieren (Name, dann Event-Titel).
      matched.sort((a, b) => {
        const pa = (a.name || a.email).toLowerCase();
        const pb = (b.name || b.email).toLowerCase();
        if (pa !== pb) return pa.localeCompare(pb);
        return (a.eventTitle || '').localeCompare(b.eventTitle || '');
      });
      const partHits: SearchHit[] = [];
      for (const rp of matched) {
        partHits.push({
          id: `part-${(rp.email || rp.name).toLowerCase()}-${rp.eventId}`,
          primary: rp.name || rp.email,
          secondary: [rp.email, rp.eventTitle, rp.status].filter(Boolean).join(' · '),
          onSelect: () => navigate('admin', rp.eventId),
        });
        if (partHits.length >= 12) break;
      }
      if (partHits.length) out.push({ key: 'participants', label: isDe ? 'Teilnehmer' : 'Attendees', hits: partHits });
    }

    return out;
  }, [query, managedEvents, adminLike, canManageAny, regParts, isDe, manualIndex, navigate]);

  // Flache Trefferliste für Tastatur-Navigation.
  const flat = React.useMemo(() => clusters.reduce<SearchHit[]>((acc, c) => [...acc, ...c.hits], []), [clusters]);
  React.useEffect(() => { setActiveIdx(0); }, [query]);

  if (!visible) return null;

  // Eingeklappt: nur das Such-Icon.
  if (!expanded) {
    return (
      <div ref={rootRef} style={{ flex: '1 1 420px', maxWidth: 460, minWidth: 0, margin: '0 16px', display: 'flex' }}>
        <button
          type="button"
          onClick={() => { setExpanded(true); setOpen(true); }}
          title={isDe ? 'Suchen' : 'Search'}
          aria-label={isDe ? 'Suchen' : 'Search'}
          style={{
            width: '100%', boxSizing: 'border-box',
            display: 'inline-flex', alignItems: 'center', gap: 10,
            padding: '8px 16px', borderRadius: 999,
            border: '1px solid var(--dex-gray-300)', background: 'var(--dex-gray-100, #eef0f2)',
            color: 'var(--dex-gray-600)', cursor: 'pointer',
            fontFamily: 'inherit', fontSize: '0.85rem', fontWeight: 500, textAlign: 'left',
          }}
        >
          <Search size={16} />
          <span>{isDe ? 'Suche' : 'Search'}</span>
        </button>
      </div>
    );
  }

  const showPanel = open && query.trim().length >= 2;
  const noResults = showPanel && flat.length === 0 && !partLoading;

  return (
    <div ref={rootRef} style={{ position: 'relative', flex: '1 1 420px', maxWidth: 460, minWidth: 0, margin: '0 16px' }}>
      <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
        <span style={{ position: 'absolute', left: 12, display: 'inline-flex', color: 'var(--dex-gray-400)', pointerEvents: 'none' }}>
          <Search size={16} />
        </span>
        <input
          ref={inputRef}
          type="text"
          value={query}
          placeholder={isDe ? 'Suche: Events, Aktionen, Teilnehmer, Handbuch …' : 'Search: events, actions, attendees, manual …'}
          onChange={e => { setQuery(e.target.value); setOpen(true); }}
          onFocus={() => setOpen(true)}
          onKeyDown={e => {
            if (e.key === 'Escape') {
              setOpen(false);
              if (!query.trim()) setExpanded(false);
              (e.currentTarget as HTMLInputElement).blur();
              return;
            }
            if (!showPanel || flat.length === 0) return;
            if (e.key === 'ArrowDown') { e.preventDefault(); setActiveIdx(i => Math.min(i + 1, flat.length - 1)); }
            else if (e.key === 'ArrowUp') { e.preventDefault(); setActiveIdx(i => Math.max(i - 1, 0)); }
            else if (e.key === 'Enter') { e.preventDefault(); const hit = flat[activeIdx]; if (hit) go(hit.onSelect); }
          }}
          style={{
            width: '100%', boxSizing: 'border-box',
            padding: '8px 32px 8px 34px', borderRadius: 999,
            border: '1px solid var(--dex-gray-300)', background: 'var(--dex-gray-50, #f7f7f7)',
            fontSize: '0.85rem', fontFamily: 'inherit', outline: 'none',
          }}
        />
        <button
          type="button"
          onClick={() => { setQuery(''); setOpen(false); setExpanded(false); }}
          aria-label={isDe ? 'Schließen' : 'Close'}
          style={{ position: 'absolute', right: 8, background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-500)', padding: 4, display: 'inline-flex' }}
        >
          <X size={14} />
        </button>
      </div>

      {showPanel && (
        <div style={{
          position: 'absolute', top: '100%', left: 0, right: 0, marginTop: 6, zIndex: 1100,
          background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 12,
          boxShadow: '0 12px 32px rgba(0,0,0,0.18)', maxHeight: '70vh', overflowY: 'auto',
          padding: '6px 0',
        }}>
          {partLoading && flat.length === 0 && (
            <div style={{ padding: '12px 16px', fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>{isDe ? 'Suche läuft …' : 'Searching …'}</div>
          )}
          {noResults && (
            <div style={{ padding: '12px 16px', fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>{isDe ? 'Keine Treffer.' : 'No results.'}</div>
          )}
          {(() => {
            let running = -1;
            return clusters.map(c => (
              <div key={c.key} style={{ paddingBottom: 4 }}>
                <div style={{ padding: '8px 16px 4px', fontSize: '0.68rem', fontWeight: 700, letterSpacing: '0.06em', textTransform: 'uppercase', color: 'var(--dex-green-dark, #4a7c1f)' }}>
                  {c.label} <span style={{ color: 'var(--dex-gray-400)', fontWeight: 600 }}>({c.hits.length})</span>
                </div>
                {c.hits.map(h => {
                  running++;
                  const idx = running;
                  const active = idx === activeIdx;
                  return (
                    <button
                      key={h.id}
                      type="button"
                      onMouseEnter={() => setActiveIdx(idx)}
                      onClick={() => go(h.onSelect)}
                      style={{
                        display: 'block', width: '100%', textAlign: 'left', cursor: 'pointer',
                        border: 'none', background: active ? 'rgba(134,188,37,0.10)' : 'transparent',
                        padding: '8px 16px', fontFamily: 'inherit',
                        borderLeft: active ? '3px solid var(--dex-green, #86bc25)' : '3px solid transparent',
                      }}
                    >
                      <span style={{ display: 'block', fontSize: '0.86rem', fontWeight: 600, color: 'var(--dex-gray-800)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{h.primary}</span>
                      {h.secondary && (
                        <span style={{ display: 'block', fontSize: '0.74rem', color: 'var(--dex-gray-500)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{h.secondary}</span>
                      )}
                    </button>
                  );
                })}
              </div>
            ));
          })()}
        </div>
      )}
    </div>
  );
}
