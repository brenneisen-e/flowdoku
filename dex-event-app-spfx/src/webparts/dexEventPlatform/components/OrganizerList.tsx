/**
 * Gemeinsame Komponente für die Anzeige von Organisatoren mit Foto.
 *
 * Verwendet die SharePoint-Userphoto-URL:
 *   /_layouts/15/userphoto.aspx?accountname=<email>&size=<S|M|L>
 *
 * - Default: kleines Avatar (28px) plus Name als Chip
 * - Mouse-Over: größeres Bild (96px) plus Name in Tooltip-Box
 * - Fallback: Initialen-Avatar, falls Foto nicht geladen werden kann
 */

import * as React from 'react';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';

export interface OrganizerListProps {
  names: string[];
  emails: string[];
  size?: 'sm' | 'md';
  compact?: boolean;
  /** v23.25: 'card' = Organizer dauerhaft groß (Foto + Name + Mail + Rolle
   *  direkt sichtbar). Default 'chip' = klein mit Hover-Popup. */
  display?: 'chip' | 'card';
}

// v11.95: pro Email einmalig profile-lookup, Ergebnis App-weit gecached
// (nur Tab-Lebensdauer). Verhindert dass jeder Hover einen neuen REST-Call
// triggert wenn der User mehrfach drüberfährt.
const profileCache = new Map<string, { jobTitle: string; location: string }>();

function getInitials(name: string): string {
  const parts = name.includes(',')
    ? name.split(',').reverse().map(s => s.trim())
    : name.split(/\s+/).filter(Boolean);
  return parts.slice(0, 2).map(p => (p[0] || '').toUpperCase()).join('');
}

function photoUrl(email: string, size: 'S' | 'M' | 'L'): string {
  if (!email) return '';
  return `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=${size}`;
}

function OrganizerChip({ name, email, sizeClass, expanded }: { name: string; email: string; sizeClass: 'sm' | 'md'; expanded?: boolean }): React.ReactElement {
  const [hovered, setHovered] = React.useState(false);
  const [failed, setFailed] = React.useState(false);
  // SP /_layouts/15/userphoto.aspx liefert sporadisch transient 404 — z.B. wenn die
  // Mailbox gerade geprovisioned wird oder die Anfrage parallel zu vielen anderen
  // läuft. Ein einmaliger Retry mit Cache-Bust-Suffix recovered das oft. Erst nach
  // dem zweiten Fehlschlag fallen wir auf das Initialen-Avatar zurück.
  const [retryAttempt, setRetryAttempt] = React.useState(0);
  const [coords, setCoords] = React.useState<{ x: number; y: number; above: boolean } | null>(null);
  // v11.95: lazy-loaded JobTitle + Standort. Wird beim ersten Hover über
  // searchUser nachgeladen und im modul-globalen profileCache gemerkt.
  const [profile, setProfile] = React.useState<{ jobTitle: string; location: string } | null>(
    email && profileCache.has(email.toLowerCase()) ? profileCache.get(email.toLowerCase()) as { jobTitle: string; location: string } : null
  );
  const { searchUser } = useRoles();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const wrapperRef = React.useRef<HTMLSpanElement>(null);
  // v23.25: verzögertes Schließen, damit die Maus vom kleinen Chip über die
  // 8px-Lücke in die große Karte wandern kann, ohne dass sie sofort zuklappt.
  const closeTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const cancelClose = (): void => { if (closeTimerRef.current) { clearTimeout(closeTimerRef.current); closeTimerRef.current = null; } };
  const scheduleClose = (): void => {
    cancelClose();
    closeTimerRef.current = setTimeout(() => { setHovered(false); setCoords(null); }, 200);
  };
  React.useEffect(() => () => cancelClose(), []);
  const avatarSize = sizeClass === 'sm' ? 24 : 32;
  const enlargedSize = 120;
  const popoverHeight = 180; // ungefähre Popover-Höhe für Flip-Entscheidung
  const initials = getInitials(name);
  const cacheBust = retryAttempt > 0 ? `&_r=${retryAttempt}` : '';

  const openPopover = (): void => {
    const r = wrapperRef.current?.getBoundingClientRect();
    if (!r) return;
    const spaceBelow = window.innerHeight - r.bottom;
    const above = spaceBelow < popoverHeight + 16;
    const y = above ? r.top - 8 : r.bottom + 8;
    const x = r.left + r.width / 2;
    setCoords({ x, y, above });
    setHovered(true);
    // v11.95: beim ersten Hover JobTitle + Standort lazy nachladen.
    // Cache pro Email — beim wiederholten Hover sofort verfügbar.
    if (email && !profile) {
      const cacheKey = email.toLowerCase();
      const cached = profileCache.get(cacheKey);
      if (cached) {
        setProfile(cached);
      } else {
        searchUser(email).then(res => {
          if (res) {
            const entry = { jobTitle: res.jobTitle || '', location: res.location || '' };
            profileCache.set(cacheKey, entry);
            setProfile(entry);
          }
        }).catch(() => { /* silent — falls Lookup fehlschlägt einfach nur Name+Email zeigen */ });
      }
    }
  };

  // v23.25: Im „card"-Modus das Profil (Rolle + Standort) sofort laden, damit
  // E-Mail + Infos ohne Hover sichtbar sind.
  React.useEffect(() => {
    if (!expanded || !email || profile) return;
    const cacheKey = email.toLowerCase();
    const cached = profileCache.get(cacheKey);
    if (cached) { setProfile(cached); return; }
    searchUser(email).then(res => {
      if (res) {
        const entry = { jobTitle: res.jobTitle || '', location: res.location || '' };
        profileCache.set(cacheKey, entry);
        setProfile(entry);
      }
    }).catch(() => { /* silent */ });
  }, [expanded, email, profile, searchUser]);

  // v23.25: Dauerhaft große Karte (Foto + Name + Mail + Rolle), wenn der
  // Organizer das im Wizard so eingestellt hat.
  if (expanded) {
    return (
      <div
        style={{
          display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 6,
          background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 12,
          padding: '14px 18px', boxShadow: '0 2px 8px rgba(0,0,0,0.06)', textAlign: 'center',
          minWidth: 180,
        }}
      >
        <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
          {isDe ? 'Bei Fragen wende dich gerne an:' : 'If you have any questions, feel free to reach out:'}
        </span>
        {!failed && email ? (
          <img
            src={`${photoUrl(email, 'L')}${cacheBust}`}
            alt={name}
            onError={() => { if (retryAttempt < 1) setRetryAttempt(retryAttempt + 1); else setFailed(true); }}
            style={{ width: enlargedSize, height: enlargedSize, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-200)' }}
          />
        ) : (
          <span style={{ width: enlargedSize, height: enlargedSize, borderRadius: '50%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', background: 'linear-gradient(135deg, #86bc25, #0076a8)', color: '#fff', fontSize: enlargedSize * 0.36, fontWeight: 700 }}>{initials}</span>
        )}
        <span style={{ fontSize: '0.9rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>{name}</span>
        {email && (
          <a
            href={`mailto:${email}`}
            style={{ fontSize: '0.78rem', color: 'var(--dex-green, #86bc25)', textDecoration: 'none', fontWeight: 600 }}
            onMouseEnter={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'underline'; }}
            onMouseLeave={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'none'; }}
          >{email}</a>
        )}
        {profile && (profile.jobTitle || profile.location) && (
          <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
            {[profile.jobTitle, profile.location].filter(Boolean).join(' · ')}
          </span>
        )}
      </div>
    );
  }

  return (
    <span
      ref={wrapperRef}
      style={{
        display: 'inline-flex', alignItems: 'center', gap: 8,
        padding: sizeClass === 'sm' ? '2px 10px 2px 2px' : '3px 12px 3px 3px',
        background: 'var(--dex-gray-100, #f2f2f2)',
        borderRadius: 999,
        fontSize: sizeClass === 'sm' ? '0.78rem' : '0.85rem',
        color: 'var(--dex-gray-800)',
        position: 'relative',
        cursor: 'default',
      }}
      onMouseEnter={() => { cancelClose(); openPopover(); }}
      onMouseLeave={scheduleClose}
    >
      {!failed && email ? (
        <img
          src={`${photoUrl(email, 'S')}${cacheBust}`}
          alt={name}
          onError={() => {
            if (retryAttempt < 1) setRetryAttempt(retryAttempt + 1);
            else setFailed(true);
          }}
          style={{
            width: avatarSize, height: avatarSize, borderRadius: '50%',
            objectFit: 'cover', background: 'var(--dex-gray-200)',
            flexShrink: 0,
          }}
        />
      ) : (
        <span
          style={{
            width: avatarSize, height: avatarSize, borderRadius: '50%',
            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
            background: 'linear-gradient(135deg, #86bc25, #0076a8)',
            color: '#fff', fontSize: avatarSize * 0.42, fontWeight: 700,
            flexShrink: 0,
          }}
        >{initials}</span>
      )}
      <span style={{ whiteSpace: 'nowrap' }}>{name}</span>

      {/* Hover-Vergrößerung: fixed positioning damit Container-Overflow nichts abschneidet */}
      {hovered && email && !failed && coords && (
        <span
          onMouseEnter={cancelClose}
          onMouseLeave={scheduleClose}
          style={{
            position: 'fixed',
            top: coords.above ? undefined : coords.y,
            bottom: coords.above ? window.innerHeight - coords.y : undefined,
            left: coords.x,
            transform: 'translateX(-50%)',
            zIndex: 2000,
            background: '#fff', borderRadius: 10, padding: 12,
            boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
            border: '1px solid var(--dex-gray-200)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6,
            // v23.25: pointerEvents aktiv, damit die Maus in die Karte fahren
            // und auf die Mail-Adresse klicken kann.
          }}
        >
          {/* v23.25: freundlicher Hinweis-Kopf. */}
          <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap', marginBottom: 2 }}>
            {isDe ? 'Bei Fragen wende dich gerne an:' : 'If you have any questions, feel free to reach out:'}
          </span>
          <img
            src={photoUrl(email, 'L')}
            alt={name}
            style={{
              width: enlargedSize, height: enlargedSize, borderRadius: '50%',
              objectFit: 'cover', background: 'var(--dex-gray-200)',
            }}
          />
          <span style={{ fontSize: '0.8rem', fontWeight: 600, whiteSpace: 'nowrap' }}>{name}</span>
          {/* v23.25: klickbarer Mailto-Link. */}
          <a
            href={`mailto:${email}`}
            style={{ fontSize: '0.72rem', color: 'var(--dex-green, #86bc25)', whiteSpace: 'nowrap', textDecoration: 'none', fontWeight: 600 }}
            onMouseEnter={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'underline'; }}
            onMouseLeave={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'none'; }}
          >{email}</a>
          {/* v11.95: JobTitle + Standort aus dem SP-Profil — lazy beim
              ersten Hover geladen, danach gecached. */}
          {profile && (profile.jobTitle || profile.location) && (
            <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap', textAlign: 'center' }}>
              {[profile.jobTitle, profile.location].filter(Boolean).join(' · ')}
            </span>
          )}
        </span>
      )}
    </span>
  );
}

/**
 * Extrahiert Lastname-Tokens aus einem Display-Namen.
 *  "von Albedyll, Benedikt" → ["von", "albedyll"]
 *  "Benedikt von Albedyll" → ["von", "albedyll"]  (letzte 2 Worte als Lastname-Heuristik bei Adelspräfix)
 *  "Heymann, Thorsten"     → ["heymann"]
 *  "Eike Brenneisen"       → ["brenneisen"]
 */
function lastnameTokens(name: string): string[] {
  const trimmed = (name || '').trim();
  if (!trimmed) return [];
  let lastnamePart: string;
  if (trimmed.indexOf(',') >= 0) {
    lastnamePart = trimmed.split(',')[0].trim();
  } else {
    // Heuristik: letztes Wort ist Lastname; falls vorletztes Wort ein Adelspräfix ist,
    // nimm die letzten beiden.
    const words = trimmed.split(/\s+/).filter(Boolean);
    if (words.length === 0) return [];
    if (words.length === 1) return [words[0].toLowerCase()];
    const PREFIXES = new Set(['von', 'van', 'de', 'der', 'den', 'di', 'da', 'du', 'la', 'le']);
    if (words.length >= 2 && PREFIXES.has(words[words.length - 2].toLowerCase())) {
      lastnamePart = words.slice(-2).join(' ');
    } else {
      lastnamePart = words[words.length - 1];
    }
  }
  return lastnamePart.toLowerCase().split(/\s+/).filter(t => t.length >= 3);
}

/**
 * Defensives Pairing: bevorzugt Index-Match wenn die Email-Local-Part den Lastname
 * der Person enthält, sonst sucht den ersten passenden Email-Eintrag aus dem Pool.
 *
 * Hintergrund: bei legacy oder closure-bug-betroffenen Events kann es vorkommen,
 * dass `organizers[]` und `organizerEmails[]` aus dem SP-Storage out-of-sync sind.
 * Reines `emails[i]`-Pairing zeigt dann das falsche Foto neben dem Namen. Mit dem
 * Lastname-Matcher korrigiert sich das visuell, auch wenn die Storage-Reihenfolge
 * gedreht ist.
 */
function pairNamesEmails(names: string[], emails: string[]): Array<{ name: string; email: string }> {
  const used = new Set<number>();
  const result: Array<{ name: string; email: string }> = [];
  const localParts = emails.map(e => (e || '').toLowerCase().split('@')[0]);

  const matchesLastname = (emailIdx: number, tokens: string[]): boolean => {
    if (emailIdx < 0 || emailIdx >= localParts.length) return false;
    const local = localParts[emailIdx];
    if (!local) return false;
    return tokens.some(t => local.indexOf(t) >= 0);
  };

  for (let i = 0; i < names.length; i++) {
    const tokens = lastnameTokens(names[i]);
    let chosen = -1;
    if (!used.has(i) && matchesLastname(i, tokens)) {
      chosen = i;
    } else {
      for (let j = 0; j < emails.length; j++) {
        if (!used.has(j) && matchesLastname(j, tokens)) { chosen = j; break; }
      }
    }
    if (chosen < 0) {
      // Kein Lastname-Match — Index-Fallback wenn Slot frei, sonst nächster freier Slot.
      if (!used.has(i) && i < emails.length) chosen = i;
      else for (let j = 0; j < emails.length; j++) if (!used.has(j)) { chosen = j; break; }
    }
    if (chosen >= 0) used.add(chosen);
    result.push({ name: names[i].trim(), email: chosen >= 0 ? (emails[chosen] || '').trim() : '' });
  }
  return result;
}

export default function OrganizerList({ names, emails, size = 'md', compact = false, display = 'chip' }: OrganizerListProps): React.ReactElement | null {
  const items = pairNamesEmails(names, emails).filter(o => !!o.name);
  if (items.length === 0) return null;
  const isCard = display === 'card';
  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: isCard ? 12 : (compact ? 4 : 6) }}>
      {items.map((o, i) => (
        <OrganizerChip key={`${o.email || o.name}-${i}`} name={o.name} email={o.email} sizeClass={size} expanded={isCard} />
      ))}
    </div>
  );
}
