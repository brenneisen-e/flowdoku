/**
 * v31.63: Automatische Archivierung & Löschung beim Admin-Start — mit
 * Fortschritts-Abzeichen oben rechts.
 *
 * Nutzer-Ansage 15.09.2026: „Automatismus mit Archivierung und Löschen beim
 * Aufruf durch Admin (aktuell muss ich das manuell bestätigen) … ich will,
 * dass es einfach im Hintergrund läuft und gemacht wird … gerne auch durch
 * ein spinning progress balken mit % oben rechts in der Ecke, damit ich
 * weiß, dass da gerade ein Job läuft."
 *
 * Was hier passiert:
 *  - Sobald Events und Rollen geladen sind und die Person Admin ist, startet
 *    mit 15 s Verzögerung `runAutoMaintenance` (nach den anderen Boot-Jobs
 *    in DexEventPlatform: 4/6/8/11 s). Gedrosselt 1× je 6 h je Browser —
 *    zwei Admin-Tabs derselben Person laufen sonst gleichzeitig über
 *    dieselben Zeilen, und jede Zeile ist Insert→Delete (Dubletten im
 *    Archiv, 404 beim zweiten Delete).
 *  - Während des Laufs steht oben rechts ein Ring mit Prozentzahl, darunter
 *    die Phase („Archivieren … DEX_Emails"). Nach dem Lauf bleibt kurz die
 *    Bilanz stehen; stand nichts an, verschwindet das Abzeichen wortlos —
 *    ein „nichts zu tun"-Kasten wäre bei jedem Start Lärm.
 *  - Nach dem Lauf feuert `dex-auto-maintenance-done`; die Landing Page
 *    zählt ihre Kästen dann neu, sonst zeigt sie die Zahlen von VOR dem Lauf.
 *
 * Die Komponente rendert für Nicht-Admins nichts und hängt unbedingt im
 * Baum (nicht hinter `isBootLoading`), damit ihr Timer den Boot überlebt.
 */

import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useRoles } from '../context/RoleContext';
import { useLocaleSafe } from '../context/LanguageContext';
import { AutoMaintenanceProgress, AutoMaintenanceResult } from '../context/actions/archiveAndPurge';
import { ensureDexUiStyles } from './dexUi';

export const AUTO_MAINTENANCE_DONE_EVENT = 'dex-auto-maintenance-done';
const THROTTLE_KEY = 'dex_auto_maintenance_v1';
const THROTTLE_MS = 6 * 60 * 60 * 1000;
const START_DELAY_MS = 15000;
const RESULT_VISIBLE_MS = 9000;

type Badge =
  | { kind: 'running'; p: AutoMaintenanceProgress }
  | { kind: 'done'; r: AutoMaintenanceResult };

export default function AdminAutoMaintenance(): React.ReactElement | null {
  const { isAdmin, isRolesLoading } = useRoles();
  const { isEventsLoading, events, runAutoMaintenance } = useEvents();
  const isDe = useLocaleSafe() === 'de';
  const [badge, setBadge] = React.useState<Badge | null>(null);
  const startedRef = React.useRef(false);
  // Das Einblenden nutzt `dexUiFadeIn` aus dem dexUi-Stylesheet.
  ensureDexUiStyles();

  React.useEffect(() => {
    if (startedRef.current) return;
    if (!isAdmin || isRolesLoading || isEventsLoading) return;
    if (!events || events.length === 0) return;
    let due = true;
    try {
      const last = parseInt(window.localStorage.getItem(THROTTLE_KEY) || '0', 10);
      if (last && Date.now() - last < THROTTLE_MS) due = false;
    } catch { /* */ }
    if (!due) { startedRef.current = true; return; }
    startedRef.current = true;
    // Stempel VOR dem Lauf — ein zweiter Tab derselben Person soll nicht
    // parallel starten. Läuft der erste schief, versucht es der nächste
    // Start nach 6 h wieder; die Handwege bleiben jederzeit offen.
    try { window.localStorage.setItem(THROTTLE_KEY, String(Date.now())); } catch { /* */ }
    // Kein clearTimeout-Cleanup (s. EventContext shadowHeal, v30.67): ein
    // Re-Render durch `events` würde den Timer sonst abräumen, bevor er feuert.
    window.setTimeout(() => {
      setBadge({ kind: 'running', p: { phase: 'zaehlen', done: 0, total: 0, pct: 0, label: '' } });
      runAutoMaintenance(p => setBadge({ kind: 'running', p }))
        .then(r => {
          if (r.nothingToDo) { setBadge(null); return; }
          setBadge({ kind: 'done', r });
          try { window.dispatchEvent(new CustomEvent(AUTO_MAINTENANCE_DONE_EVENT)); } catch { /* */ }
          window.setTimeout(() => setBadge(prev => (prev && prev.kind === 'done' ? null : prev)), RESULT_VISIBLE_MS);
        })
        .catch(err => {
          console.warn('[DEX] auto maintenance failed:', err);
          setBadge({ kind: 'done', r: { archived: 0, archiveFailed: 0, deleted: 0, deleteFailed: 0, participantsDeleted: 0, participantsFailed: 0, nothingToDo: false, error: err instanceof Error ? err.message : String(err) } });
          window.setTimeout(() => setBadge(prev => (prev && prev.kind === 'done' ? null : prev)), RESULT_VISIBLE_MS);
        });
    }, START_DELAY_MS);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isAdmin, isRolesLoading, isEventsLoading, events]);

  if (!badge) return null;

  const phaseText = (p: AutoMaintenanceProgress): string => {
    switch (p.phase) {
      case 'zaehlen': return isDe ? 'Prüfe, was ansteht …' : 'Checking what is due …';
      case 'archiv': return (isDe ? 'Archivieren' : 'Archiving') + (p.label ? ` · ${p.label}` : ' …');
      case 'archivloeschen': return isDe ? 'Archiv aufräumen …' : 'Cleaning up archive …';
      case 'teilnehmer': return (isDe ? 'Teilnehmerliste löschen' : 'Deleting attendee list') + (p.label ? ` · ${p.label}` : ' …');
      default: return isDe ? 'Fertig' : 'Done';
    }
  };

  const shell: React.CSSProperties = {
    position: 'fixed', top: 76, right: 16, zIndex: 2400,
    display: 'flex', alignItems: 'center', gap: 12,
    padding: '10px 14px 10px 10px', borderRadius: 14,
    background: '#fff', border: '1px solid var(--dex-gray-200, #e1e1e1)',
    boxShadow: '0 10px 28px rgba(0,0,0,0.14)',
    maxWidth: 'min(360px, calc(100vw - 32px))',
    fontSize: '0.82rem', color: 'var(--dex-gray-700, #333)',
    animation: 'dexUiFadeIn 0.25s ease-out both',
  };

  if (badge.kind === 'running') {
    const p = badge.p;
    const indeterminate = p.phase === 'zaehlen' || p.total === 0;
    return (
      <div role="status" aria-live="polite" style={shell} title={isDe ? 'Archivierung & Löschen laufen im Hintergrund' : 'Archiving & deletion running in the background'}>
        <ProgressRing pct={indeterminate ? null : p.pct} />
        <div style={{ minWidth: 0 }}>
          <div style={{ fontWeight: 700, color: 'var(--dex-black, #000)', whiteSpace: 'nowrap' }}>
            {isDe ? 'Aufräumen läuft' : 'Cleanup running'}
          </div>
          <div style={{ color: 'var(--dex-gray-500, #666)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
            {phaseText(p)}{!indeterminate && ` · ${p.done}/${p.total}`}
          </div>
        </div>
      </div>
    );
  }

  const r = badge.r;
  const hadError = !!r.error || r.archiveFailed > 0 || r.deleteFailed > 0 || r.participantsFailed > 0;
  const parts: string[] = [];
  if (r.archived > 0) parts.push(isDe ? `${r.archived} archiviert` : `${r.archived} archived`);
  if (r.deleted > 0) parts.push(isDe ? `${r.deleted} Archivzeilen entfernt` : `${r.deleted} archive rows removed`);
  if (r.participantsDeleted > 0) parts.push(isDe ? `${r.participantsDeleted} Teilnehmerliste(n) gelöscht` : `${r.participantsDeleted} attendee list(s) deleted`);
  const failed = r.archiveFailed + r.deleteFailed + r.participantsFailed;
  if (failed > 0) parts.push(isDe ? `${failed} fehlgeschlagen` : `${failed} failed`);
  if (r.error) parts.push(r.error);
  return (
    <div role="status" aria-live="polite" style={shell}>
      <span aria-hidden="true" style={{
        width: 40, height: 40, borderRadius: '50%', flexShrink: 0,
        display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
        background: hadError ? 'rgba(218,41,28,0.10)' : 'rgba(134,188,37,0.14)',
        color: hadError ? '#da291c' : '#6b9a1e', fontWeight: 800, fontSize: 18,
      }}>{hadError ? '!' : '✓'}</span>
      <div style={{ minWidth: 0 }}>
        <div style={{ fontWeight: 700, color: 'var(--dex-black, #000)' }}>
          {hadError ? (isDe ? 'Aufräumen mit Fehlern' : 'Cleanup with errors') : (isDe ? 'Aufgeräumt' : 'Cleanup done')}
        </div>
        <div style={{ color: 'var(--dex-gray-500, #666)' }}>
          {parts.join(' · ') || (isDe ? 'Nichts zu tun.' : 'Nothing to do.')}
        </div>
      </div>
      <button
        type="button"
        onClick={() => setBadge(null)}
        aria-label={isDe ? 'Schließen' : 'Close'}
        style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-400, #999)', fontSize: 18, lineHeight: 1, padding: '0 0 0 4px', flexShrink: 0 }}
      >×</button>
    </div>
  );
}

/**
 * Ring mit Prozentzahl in der Mitte. Der grüne Bogen füllt sich mit dem
 * Fortschritt; ein kurzer, heller Bogen dreht darüber, damit man auch bei
 * gleichbleibender Zahl sieht, dass etwas arbeitet (eine große Liste kann
 * zwischen zwei Prozentschritten Sekunden brauchen). `pct === null` = noch
 * nichts gezählt → nur der drehende Bogen.
 */
function ProgressRing(props: { pct: number | null }): React.ReactElement {
  const size = 44, stroke = 4, r = (size - stroke) / 2, c = 2 * Math.PI * r;
  const pct = props.pct;
  const offset = pct === null ? c : c * (1 - pct / 100);
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} aria-hidden="true" style={{ flexShrink: 0 }}>
      <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke="rgba(134,188,37,0.18)" strokeWidth={stroke} />
      <circle
        cx={size / 2} cy={size / 2} r={r} fill="none" stroke="#86bc25" strokeWidth={stroke} strokeLinecap="round"
        strokeDasharray={c} strokeDashoffset={offset}
        transform={`rotate(-90 ${size / 2} ${size / 2})`}
        style={{ transition: 'stroke-dashoffset 0.35s ease-out' }}
      />
      <circle
        cx={size / 2} cy={size / 2} r={r} fill="none" stroke={pct === null ? '#86bc25' : 'rgba(255,255,255,0.75)'} strokeWidth={stroke} strokeLinecap="round"
        strokeDasharray={`${c * 0.16} ${c}`}
      >
        <animateTransform attributeName="transform" type="rotate" from={`0 ${size / 2} ${size / 2}`} to={`360 ${size / 2} ${size / 2}`} dur="1.1s" repeatCount="indefinite" />
      </circle>
      <text x="50%" y="50%" dominantBaseline="central" textAnchor="middle" fontSize={pct === null ? 14 : 11} fontWeight={700} fill="#1a1a1a" fontFamily="inherit">
        {pct === null ? '…' : `${pct}%`}
      </text>
    </svg>
  );
}
