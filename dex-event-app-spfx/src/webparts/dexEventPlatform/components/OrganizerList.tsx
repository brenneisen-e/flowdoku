/**
 * Gemeinsame Komponente fuer die Anzeige von Organisatoren mit Foto.
 *
 * Verwendet die SharePoint-Userphoto-URL:
 *   /_layouts/15/userphoto.aspx?accountname=<email>&size=<S|M|L>
 *
 * - Default: kleines Avatar (28px) plus Name als Chip
 * - Mouse-Over: groesseres Bild (96px) plus Name in Tooltip-Box
 * - Fallback: Initialen-Avatar, falls Foto nicht geladen werden kann
 */

import * as React from 'react';

export interface OrganizerListProps {
  names: string[];
  emails: string[];
  size?: 'sm' | 'md';
  compact?: boolean;
}

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

function OrganizerChip({ name, email, sizeClass }: { name: string; email: string; sizeClass: 'sm' | 'md' }): React.ReactElement {
  const [hovered, setHovered] = React.useState(false);
  const [failed, setFailed] = React.useState(false);
  const [coords, setCoords] = React.useState<{ x: number; y: number; above: boolean } | null>(null);
  const wrapperRef = React.useRef<HTMLSpanElement>(null);
  const avatarSize = sizeClass === 'sm' ? 24 : 32;
  const enlargedSize = 120;
  const popoverHeight = 180; // ungefaehre Popover-Hoehe fuer Flip-Entscheidung
  const initials = getInitials(name);

  const openPopover = (): void => {
    const r = wrapperRef.current?.getBoundingClientRect();
    if (!r) return;
    const spaceBelow = window.innerHeight - r.bottom;
    const above = spaceBelow < popoverHeight + 16;
    const y = above ? r.top - 8 : r.bottom + 8;
    const x = r.left + r.width / 2;
    setCoords({ x, y, above });
    setHovered(true);
  };

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
      onMouseEnter={openPopover}
      onMouseLeave={() => { setHovered(false); setCoords(null); }}
    >
      {!failed && email ? (
        <img
          src={photoUrl(email, 'S')}
          alt={name}
          onError={() => setFailed(true)}
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

      {/* Hover-Vergroesserung: fixed positioning damit Container-Overflow nichts abschneidet */}
      {hovered && email && !failed && coords && (
        <span
          style={{
            position: 'fixed',
            top: coords.above ? undefined : coords.y,
            bottom: coords.above ? window.innerHeight - coords.y : undefined,
            left: coords.x,
            transform: 'translateX(-50%)',
            zIndex: 2000,
            background: '#fff', borderRadius: 10, padding: 10,
            boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
            border: '1px solid var(--dex-gray-200)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6,
            pointerEvents: 'none',
          }}
        >
          <img
            src={photoUrl(email, 'L')}
            alt={name}
            style={{
              width: enlargedSize, height: enlargedSize, borderRadius: '50%',
              objectFit: 'cover', background: 'var(--dex-gray-200)',
            }}
          />
          <span style={{ fontSize: '0.8rem', fontWeight: 600, whiteSpace: 'nowrap' }}>{name}</span>
          <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>{email}</span>
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

export default function OrganizerList({ names, emails, size = 'md', compact = false }: OrganizerListProps): React.ReactElement | null {
  const items = pairNamesEmails(names, emails).filter(o => !!o.name);
  if (items.length === 0) return null;
  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: compact ? 4 : 6 }}>
      {items.map((o, i) => (
        <OrganizerChip key={`${o.email || o.name}-${i}`} name={o.name} email={o.email} sizeClass={size} />
      ))}
    </div>
  );
}
