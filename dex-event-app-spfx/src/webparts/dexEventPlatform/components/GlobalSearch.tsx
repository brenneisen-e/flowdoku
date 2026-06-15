/**
 * v22.50: Globale Such-Leiste im Header (für Admins UND Organizer eigener
 * Events). Durchsucht geclustert: Events, Aktionen (pro verwaltbarem Event),
 * Seiten/Funktionen, Handbuch und Teilnehmer — und springt direkt dorthin.
 *
 * Berechtigungs-Scoping (strikt): Es werden nur Events, Aktionen und
 * Teilnehmer angezeigt, die der eingeloggte Nutzer sehen/verwalten darf
 * (Admin: alle; Organizer: nur eigene/Co-Organizer-Events). Seiten/Handbuch
 * werden rollen-gefiltert.
 */
import * as React from 'react';
import { Search, X } from './Icons';
import { useNavigation, Page } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useLanguage } from '../context/LanguageContext';
import { getManualSections } from './manual/handbookContent';
import { SPParticipant } from '../services/EventService';
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
  { key: 'massmail', de: 'Massenmail an Teilnehmer', en: 'Mass mail to attendees', kw: ['massenmail', 'mail', 'email', 'e-mail', 'nachricht', 'rundmail', 'anschreiben'], gate: 'manage' },
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

const matchAll = (tokens: string[], haystack: string): boolean => {
  const hayWords = haystack.split(/[^a-z0-9äöüß]+/).filter(Boolean);
  return tokens.every(t => tokenMatches(t, haystack, hayWords));
};

export default function GlobalSearch(): React.ReactElement | null {
  const { navigate } = useNavigation();
  const { isAdmin, isOrganizer, originalIsAdmin } = useRoles();
  const { events, getAllParticipants } = useEvents();
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
  // Sichtbarkeit der Leiste: nur wer mindestens ein Event verwalten kann.
  const visible = canManageAny;

  const [query, setQuery] = React.useState('');
  const [open, setOpen] = React.useState(false);
  const [activeIdx, setActiveIdx] = React.useState(0);
  const [participants, setParticipants] = React.useState<SPParticipant[] | null>(null);
  const [partLoading, setPartLoading] = React.useState(false);
  const rootRef = React.useRef<HTMLDivElement | null>(null);

  // Teilnehmer einmal lazy laden, sobald der Nutzer zu tippen beginnt.
  React.useEffect(() => {
    if (!visible) return;
    if (query.trim().length < 2) return;
    if (participants !== null || partLoading) return;
    setPartLoading(true);
    getAllParticipants()
      .then(p => setParticipants(p || []))
      .catch(() => setParticipants([]))
      .finally(() => setPartLoading(false));
  }, [query, visible, participants, partLoading, getAllParticipants]);

  // Outside-Click schließt das Dropdown.
  React.useEffect(() => {
    if (!open) return undefined;
    const onDown = (e: MouseEvent): void => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) setOpen(false);
    };
    window.addEventListener('mousedown', onDown);
    return () => window.removeEventListener('mousedown', onDown);
  }, [open]);

  const go = React.useCallback((fn: () => void): void => {
    setOpen(false); setQuery('');
    fn();
  }, []);

  const role: 'Admin' | 'Organizer' | 'User' = adminLike ? 'Admin' : (canManageAny ? 'Organizer' : 'User');

  const clusters = React.useMemo<Cluster[]>(() => {
    const q = norm(query);
    if (q.length < 2) return [];
    const tokens = q.split(/\s+/).filter(Boolean);
    const out: Cluster[] = [];

    // ---- Events ----
    const evHits: SearchHit[] = [];
    for (const e of managedEvents) {
      const hay = norm(`${e.title} ${e.eventNumber || ''} ${e.location || ''}`);
      if (matchAll(tokens, hay)) {
        evHits.push({
          id: `ev-${e.id}`,
          primary: e.title || (isDe ? 'Event ohne Titel' : 'Untitled event'),
          secondary: [e.eventNumber ? `Nr. ${e.eventNumber}` : '', e.location || '', e.isFictive ? (isDe ? 'Entwurf' : 'Draft') : ''].filter(Boolean).join(' · '),
          onSelect: () => navigate('admin', e.id),
        });
      }
      if (evHits.length >= 6) break;
    }
    if (evHits.length) out.push({ key: 'events', label: isDe ? 'Events' : 'Events', hits: evHits });

    // ---- Aktionen (kombiniert Aktion × verwaltbares Event) ----
    const actHits: SearchHit[] = [];
    const sortedEvents = managedEvents.slice().sort((a, b) => (b.startDate || '').localeCompare(a.startDate || ''));
    outer:
    for (const a of ACTION_CATALOG) {
      if (a.gate === 'admin' && !adminLike) continue;
      const actLabel = isDe ? a.de : a.en;
      const actHay = norm(`${a.de} ${a.en} ${a.kw.join(' ')}`);
      for (const e of sortedEvents) {
        // Treffer, wenn ALLE Tokens entweder in der Aktion ODER im Event stecken.
        const combined = `${actHay} ${norm(`${e.title} ${e.eventNumber || ''} ${e.location || ''}`)}`;
        if (!matchAll(tokens, combined)) continue;
        // Mindestens ein Token muss die Aktion treffen (sonst reiner Event-Treffer).
        if (!tokens.some(t => actHay.indexOf(t) >= 0)) continue;
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
      if (matchAll(tokens, hay)) {
        pageHits.push({ id: `pg-${p.page}`, primary: label, onSelect: () => navigate(p.page) });
      }
    }
    if (pageHits.length) out.push({ key: 'pages', label: isDe ? 'Seiten & Funktionen' : 'Pages & functions', hits: pageHits });

    // ---- Handbuch ----
    const manHits: SearchHit[] = [];
    try {
      const sections = getManualSections(isDe ? 'de' : 'en')
        .filter(s => adminLike ? true : (role === 'Organizer' ? (s.visibleFor.indexOf('Organizer') >= 0 || s.visibleFor.indexOf('User') >= 0) : s.visibleFor.indexOf('User') >= 0));
      for (const s of sections) {
        const hay = norm(`${s.title} ${s.description}`);
        if (matchAll(tokens, hay)) {
          manHits.push({
            id: `man-${s.id}`,
            primary: s.title,
            secondary: s.description,
            onSelect: () => { try { window.localStorage.setItem('dex_open_manual_section', s.id); } catch { /* */ } navigate('manual'); },
          });
        }
        if (manHits.length >= 6) break;
      }
    } catch { /* Handbuch best-effort */ }
    if (manHits.length) out.push({ key: 'manual', label: isDe ? 'Handbuch' : 'Manual', hits: manHits });

    // ---- Teilnehmer (berechtigungs-gescoped über verwaltbare Events) ----
    if (participants && participants.length) {
      const managedNums = new Set(managedEvents.map(e => e.eventNumber).filter(Boolean));
      const numToEvent = new Map<number, DeloitteEvent>();
      for (const e of managedEvents) { if (e.eventNumber) numToEvent.set(e.eventNumber, e); }
      const parseNums = (s?: string): number[] => (s || '').split(',').map(x => parseInt(x.trim(), 10)).filter(n => Number.isFinite(n));
      const partHits: SearchHit[] = [];
      for (const p of participants) {
        const hay = norm(`${p.Vorname || ''} ${p.Nachname || ''} ${p.Email || p.Title || ''}`);
        if (!matchAll(tokens, hay)) continue;
        const nums = Array.from(new Set([...parseNums(p.EventRegistered), ...parseNums(p.EventOnWaitlist)]));
        const myEvents = nums.filter(n => managedNums.has(n)).map(n => numToEvent.get(n)).filter(Boolean) as DeloitteEvent[];
        if (myEvents.length === 0) continue; // keine sichtbaren Events → nicht anzeigen
        const name = `${p.Vorname || ''} ${p.Nachname || ''}`.trim() || p.Email || p.Title;
        partHits.push({
          id: `part-${p.Id}`,
          primary: name,
          secondary: `${p.Email || p.Title} · ${myEvents.map(e => e.title).slice(0, 3).join(', ')}`,
          onSelect: () => navigate('admin', myEvents[0].id),
        });
        if (partHits.length >= 8) break;
      }
      if (partHits.length) out.push({ key: 'participants', label: isDe ? 'Teilnehmer' : 'Attendees', hits: partHits });
    }

    return out;
  }, [query, managedEvents, adminLike, canManageAny, participants, isDe, role, navigate]);

  // Flache Trefferliste für Tastatur-Navigation.
  const flat = React.useMemo(() => clusters.reduce<SearchHit[]>((acc, c) => [...acc, ...c.hits], []), [clusters]);
  React.useEffect(() => { setActiveIdx(0); }, [query]);

  if (!visible) return null;

  const showPanel = open && query.trim().length >= 2;
  const noResults = showPanel && flat.length === 0 && !partLoading;

  return (
    <div ref={rootRef} style={{ position: 'relative', flex: '1 1 420px', maxWidth: 460, minWidth: 0, margin: '0 16px' }}>
      <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
        <span style={{ position: 'absolute', left: 12, display: 'inline-flex', color: 'var(--dex-gray-400)', pointerEvents: 'none' }}>
          <Search size={16} />
        </span>
        <input
          type="text"
          value={query}
          placeholder={isDe ? 'Suche: Events, Aktionen, Teilnehmer, Handbuch …' : 'Search: events, actions, attendees, manual …'}
          onChange={e => { setQuery(e.target.value); setOpen(true); }}
          onFocus={() => setOpen(true)}
          onKeyDown={e => {
            if (e.key === 'Escape') { setOpen(false); (e.currentTarget as HTMLInputElement).blur(); return; }
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
        {query && (
          <button
            type="button"
            onClick={() => { setQuery(''); setOpen(false); }}
            aria-label={isDe ? 'Leeren' : 'Clear'}
            style={{ position: 'absolute', right: 8, background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-500)', padding: 4, display: 'inline-flex' }}
          >
            <X size={14} />
          </button>
        )}
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
