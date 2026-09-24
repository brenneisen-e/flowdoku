/**
 * WalkInPanel (v31.90) — Walk-in am Check-in-Tisch.
 *
 * Nutzer-Ansage 24.09.2026: „bau die Möglichkeit, einen Walk-in zu
 * registrieren … wenn die Person nicht angemeldet ist, dann kann man die
 * Person auch eingeben und sie wird über die Personensuche gesucht, und dann
 * kann man Walk-in machen, und dann wird sie auch auf der Teilnehmerliste
 * eingeschrieben mit Registriert am / Registriert von." Und: „hier dann
 * eigene Kachel" — unter „Weitere Wege zum Einchecken", neben Live-Scanner
 * und Self-Check-in.
 *
 * Ablauf: Person suchen (Tenant-Suche, wie im Assistenten) → bei einer
 * Klammer die Termine wählen → „Walk-in anmelden". Die Anmeldung läuft über
 * `registerForEvent` — derselbe Pfad wie die stellvertretende Anmeldung
 * (Register, Platz, Klammer-Schattenzeile, Log); `RegisteredBy*` ist damit
 * die Person am Tisch, `RegistrationDate` jetzt. Mail und Outlook sind
 * standardmäßig AUS (die Person steht vor einem). Danach ruft der Aufrufer
 * die Liste nach und öffnet die Bestätigungskarte zum Einchecken — der
 * Check-in selbst bleibt EIN Weg (CLAUDE.md).
 */
import * as React from 'react';
import { DeloitteEvent } from '../types';
import InternationalSearchToggle from './InternationalSearchToggle';
import { Check, X } from './Icons';
import { cx } from './dexUi';

/** Foto aus dem Profil, sonst Initiale — wie die Zeilen im Assistenten. */
const PersonAvatar: React.FC<{ email: string; name: string }> = ({ email, name }) => {
  const [broken, setBroken] = React.useState(false);
  const initial = (name || email || '?').trim().charAt(0).toUpperCase();
  return (
    <span style={{ width: 32, height: 32, borderRadius: '50%', flexShrink: 0, overflow: 'hidden', background: 'var(--dex-gray-100, #eee)', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.8rem', fontWeight: 700, color: 'var(--dex-gray-600)' }}>
      {!broken && email
        ? <img src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=S`} alt="" onError={() => setBroken(true)} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
        : initial}
    </span>
  );
};

export interface WalkInHit { email: string; displayName: string; location?: string; jobTitle?: string }

export interface WalkInPanelProps {
  isDe: boolean;
  /** Das unten gewählte Event (Klammer, Hauptevent oder Termin). */
  event: DeloitteEvent | null;
  /** Termine der Klammer (leer bei normalem Event/Termin). */
  termine: DeloitteEvent[];
  /** Gesperrt, weil die QR-Codes für ein anderes Event der Familie rausgingen. */
  gesperrt: boolean;
  sperrHinweis: string;
  searchUsers: (q: string, includeIntl?: boolean) => Promise<WalkInHit[]>;
  registerForEvent: (
    eventId: string,
    customData: Record<string, string>,
    participantFirstName?: string,
    participantLastName?: string,
    participantEmail?: string,
    preferredStarterType?: string,
    opts?: { suppressMail?: boolean; suppressOutlook?: boolean; skipReload?: boolean; proxyConsentConfirmed?: boolean; actorAllowedAsAssistant?: boolean; walkIn?: boolean },
  ) => Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  /** Schon in der geladenen Liste? Dann kein zweiter Walk-in, sondern Hinweis. */
  bereitsAngemeldet: (emailLc: string) => boolean;
  /** Nach erfolgreicher Anmeldung: Liste nachladen und Bestätigungskarte öffnen. */
  onRegistered: (emailLc: string, name: string) => Promise<void>;
}

const splitName = (displayName: string, email: string): { first: string; last: string } => {
  const dn = (displayName || '').trim();
  if (!dn || dn === email) return { first: '', last: email };
  if (dn.indexOf(',') >= 0) {
    const p = dn.split(',').map(s => s.trim());
    return { last: p[0] || dn, first: p[1] || '' };
  }
  const w = dn.split(/\s+/);
  if (w.length >= 2) return { last: w[w.length - 1], first: w.slice(0, -1).join(' ') };
  return { last: dn, first: '' };
};

const reasonText = (reason: string | undefined, isDe: boolean): string => {
  switch (reason) {
    case 'not-allowed': return isDe ? 'Du darfst für diese Person nicht anmelden — bist du im Check-in-Team dieses Events?' : 'You are not allowed to register this person — are you on this event’s check-in team?';
    case 'deadline': return isDe ? 'Anmeldefrist abgelaufen und kein Check-in-Team-Recht auf diesem Termin.' : 'Registration deadline passed and no check-in team right on this date.';
    case 'already-registered': return isDe ? 'Schon angemeldet.' : 'Already registered.';
    case 'dup-check-failed': return isDe ? 'Liste nicht lesbar — Anmeldung nicht möglich.' : 'List not readable — cannot register.';
    default: return reason ? `${isDe ? 'Fehler' : 'Error'}: ${reason}` : (isDe ? 'Fehlgeschlagen.' : 'Failed.');
  }
};

export default function WalkInPanel(props: WalkInPanelProps): React.ReactElement {
  const { isDe, event, termine, gesperrt, sperrHinweis, searchUsers, registerForEvent, bereitsAngemeldet, onRegistered } = props;
  const [query, setQuery] = React.useState('');
  const [intl, setIntl] = React.useState(false);
  const [hits, setHits] = React.useState<WalkInHit[]>([]);
  const [searching, setSearching] = React.useState(false);
  const [person, setPerson] = React.useState<WalkInHit | null>(null);
  const [zielIds, setZielIds] = React.useState<string[]>([]);
  const [sendMail, setSendMail] = React.useState(false);
  const [sendOutlook, setSendOutlook] = React.useState(false);
  const [running, setRunning] = React.useState(false);
  const [report, setReport] = React.useState<Array<{ target: string; ok: boolean; text: string }> | null>(null);
  const timerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  const istKlammer = !!event && !!event.subEventsOnlyMode && termine.length > 0;
  // Bei einer Klammer sind die Termine die Buchung; die Klammer-Zeile stellt
  // registerForEvent selbst sicher (v30.68). Vorgabe: alle Termine.
  React.useEffect(() => {
    setZielIds(istKlammer ? termine.map(t => t.id) : (event ? [event.id] : []));
    setPerson(null); setReport(null); setQuery(''); setHits([]);
  }, [event ? event.id : '', istKlammer, termine.map(t => t.id).join('|')]); // eslint-disable-line react-hooks/exhaustive-deps

  const runSearch = (q: string, includeIntl: boolean): void => {
    if (timerRef.current) clearTimeout(timerRef.current);
    const qq = q.trim();
    if (qq.length < 2) { setHits([]); return; }
    timerRef.current = setTimeout(() => {
      setSearching(true);
      searchUsers(qq, includeIntl)
        .then(r => setHits(r.slice(0, 8)))
        .catch(() => setHits([]))
        .then(() => setSearching(false));
    }, 350);
  };

  const anmelden = async (): Promise<void> => {
    if (!person || !event || running) return;
    const emailLc = person.email.toLowerCase().trim();
    if (bereitsAngemeldet(emailLc)) {
      setReport([{ target: event.title, ok: false, text: isDe ? 'Die Person steht schon in der Liste — bitte unten über den Namen einchecken.' : 'This person is already on the list — please check in by name below.' }]);
      return;
    }
    const ziele = istKlammer ? termine.filter(t => zielIds.indexOf(t.id) >= 0) : [event];
    if (ziele.length === 0) {
      setReport([{ target: '', ok: false, text: isDe ? 'Bitte mindestens einen Termin wählen.' : 'Please pick at least one date.' }]);
      return;
    }
    setRunning(true);
    const { first, last } = splitName(person.displayName, person.email);
    const rows: Array<{ target: string; ok: boolean; text: string }> = [];
    let einer = false;
    for (const z of ziele) {
      try {
        const res = await registerForEvent(z.id, {}, first, last, person.email, undefined, {
          suppressMail: !sendMail, suppressOutlook: !sendOutlook, skipReload: true,
          proxyConsentConfirmed: true, actorAllowedAsAssistant: true, walkIn: true,
        });
        if (res.ok) einer = true;
        rows.push({
          target: z.title, ok: res.ok,
          text: res.ok
            ? (res.status === 'Warteliste' ? (isDe ? 'Warteliste (voll)' : 'Waitlist (full)') : (isDe ? 'Angemeldet' : 'Registered'))
            : reasonText(res.reason, isDe),
        });
      } catch (err) {
        rows.push({ target: z.title, ok: false, text: String((err as Error)?.message || err) });
      }
    }
    setReport(rows);
    setRunning(false);
    if (einer) {
      try { await onRegistered(emailLc, person.displayName || person.email); } catch { /* Aufrufer meldet */ }
    }
  };

  if (!event) {
    return <p className="dex-ui-muted" style={{ margin: '12px 0 0', fontSize: '0.85rem' }}>{isDe ? 'Bitte oben ein Event wählen.' : 'Please pick an event above.'}</p>;
  }
  if (gesperrt) {
    return <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 12 }}>{sperrHinweis}</div>;
  }

  return (
    <div style={{ marginTop: 12 }}>
      <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', margin: '0 0 12px' }}>
        {isDe
          ? 'Jemand steht vor dem Tisch und ist nicht angemeldet? Person suchen, anmelden — die Zeile bekommt „Registriert am" jetzt und „Registriert von" dich —, dann direkt einchecken. Mail und Outlook-Termin sind aus, weil die Person ja schon da ist.'
          : 'Someone is at the desk without a registration? Search the person, register them — the row gets “Registered on” now and “Registered by” you — then check in right away. Email and Outlook are off because the person is already here.'}
      </p>
      {!person ? (
        <>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
            <input
              className="form-input"
              value={query}
              onChange={e => { setQuery(e.target.value); runSearch(e.target.value, intl); }}
              placeholder={isDe ? 'Name oder E-Mail der Person …' : 'Name or email of the person …'}
              style={{ flex: '1 1 240px', padding: '10px 14px', fontSize: '0.95rem' }}
              autoComplete="off"
            />
            <InternationalSearchToggle checked={intl} onChange={v => { setIntl(v); runSearch(query, v); }} isDe={isDe} compact />
          </div>
          {searching && <p className="dex-ui-muted" style={{ margin: '8px 0 0', fontSize: '0.8rem' }}>{isDe ? 'Suche …' : 'Searching …'}</p>}
          {!searching && hits.length > 0 && (
            <div className="dex-ui-card dex-ui-card--soft" style={{ padding: '4px 6px', marginTop: 8 }}>
              {hits.map(h => (
                <button key={h.email} type="button" className="dex-ui-rowbtn dex-ui-row dex-ui-row--bordered" onClick={() => { setPerson(h); setReport(null); }} style={{ width: '100%', textAlign: 'left' }}>
                  <PersonAvatar email={h.email} name={h.displayName} />
                  <div className="dex-ui-row-main">
                    <div className="dex-ui-row-title">{h.displayName}</div>
                    <div className="dex-ui-row-sub">{[h.email, h.jobTitle, h.location].filter(Boolean).join(' · ')}</div>
                  </div>
                  {bereitsAngemeldet(h.email.toLowerCase()) && <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'schon in der Liste' : 'already on the list'}</span>}
                </button>
              ))}
            </div>
          )}
          {!searching && query.trim().length >= 2 && hits.length === 0 && (
            <p className="dex-ui-muted" style={{ margin: '8px 0 0', fontSize: '0.8rem' }}>{isDe ? 'Niemand gefunden — anders schreiben oder „international" einschalten.' : 'Nobody found — try another spelling or enable “international”.'}</p>
          )}
        </>
      ) : (
        <>
          <div className="dex-ui-row dex-ui-row--bordered" style={{ marginBottom: 10 }}>
            <PersonAvatar email={person.email} name={person.displayName} />
            <div className="dex-ui-row-main">
              <div className="dex-ui-row-title">{person.displayName}</div>
              <div className="dex-ui-row-sub">{[person.email, person.jobTitle, person.location].filter(Boolean).join(' · ')}</div>
            </div>
            <div className="dex-ui-row-actions">
              <button type="button" className="dex-ui-iconbtn" onClick={() => { setPerson(null); setReport(null); }} title={isDe ? 'Andere Person' : 'Other person'} aria-label={isDe ? 'Andere Person' : 'Other person'} disabled={running}><X size={16} /></button>
            </div>
          </div>
          {istKlammer && (
            <div style={{ marginBottom: 10 }}>
              <div style={{ fontSize: '0.8rem', fontWeight: 700, marginBottom: 6 }}>{isDe ? 'Für welche Termine?' : 'Which dates?'}</div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                {termine.map(t => {
                  const on = zielIds.indexOf(t.id) >= 0;
                  return (
                    <button key={t.id} type="button" className={cx('dex-ui-chip', on && 'is-active')} aria-pressed={on} disabled={running}
                      onClick={() => setZielIds(prev => on ? prev.filter(x => x !== t.id) : [...prev, t.id])}>
                      {on && <Check size={12} />}{t.title}
                    </button>
                  );
                })}
              </div>
            </div>
          )}
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 14, marginBottom: 12, fontSize: '0.82rem' }}>
            <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: 'pointer' }}>
              <input type="checkbox" checked={sendMail} onChange={e => setSendMail(e.target.checked)} disabled={running} />
              {isDe ? 'Bestätigungsmail schicken' : 'Send confirmation email'}
            </label>
            <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: 'pointer' }}>
              <input type="checkbox" checked={sendOutlook} onChange={e => setSendOutlook(e.target.checked)} disabled={running} />
              {isDe ? 'Outlook-Termin schicken' : 'Send Outlook invitation'}
            </label>
          </div>
          <button type="button" className="btn btn-primary" onClick={() => { void anmelden(); }} disabled={running}>
            {running ? (isDe ? 'Wird angemeldet …' : 'Registering …') : (isDe ? 'Walk-in anmelden und einchecken' : 'Register walk-in and check in')}
          </button>
        </>
      )}
      {report && (
        <div className="dex-ui-card dex-ui-card--soft" style={{ padding: '8px 10px', marginTop: 10, fontSize: '0.82rem' }}>
          {report.map((r, i) => (
            <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '3px 0', color: r.ok ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-red, #c00)' }}>
              {r.ok ? <Check size={14} /> : <X size={14} />}
              <span>{r.target ? `${r.target}: ` : ''}{r.text}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
