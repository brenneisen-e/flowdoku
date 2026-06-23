/**
 * v24.56: Avatar mit Hover-Kontaktkarte (E-Mail + MS-Teams-Chat) — analog zur
 * Organizer-Hover-Karte auf der Anmeldeseite, aber als wiederverwendbare
 * Komponente für die Teilnehmerliste im Organizer Center.
 *
 * Mouse-Over auf das Foto öffnet ein kleines Popover (fixed positioniert, damit
 * Tabellen-Overflow nichts abschneidet) mit großem Foto, Name, klickbarer
 * E-Mail-Adresse und Teams-Chat-Link + (falls vorhanden) Position/Standort/Firma.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';

function getInitials(name: string): string {
  const parts = (name || '').includes(',')
    ? name.split(',').reverse().map(s => s.trim())
    : (name || '').split(/\s+/).filter(Boolean);
  return parts.slice(0, 2).map(p => (p[0] || '').toUpperCase()).join('');
}

function photoUrl(email: string, size: 'S' | 'M' | 'L'): string {
  if (!email) return '';
  return `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=${size}`;
}

function teamsChatUrl(email: string): string {
  return `msteams:/l/chat/0/0?users=${encodeURIComponent(email)}`;
}

export interface PersonContactHoverProps {
  email: string;
  name: string;
  /** Avatar-Durchmesser in px (Default 30). */
  size?: number;
  subline?: string; // "Position • Standort • Firma"
  isDe?: boolean;
}

export function PersonContactHover(props: PersonContactHoverProps): React.ReactElement {
  const { email, name, size = 30, subline, isDe = true } = props;
  const [failed, setFailed] = React.useState(false);
  const [coords, setCoords] = React.useState<{ x: number; y: number; above: boolean } | null>(null);
  const [open, setOpen] = React.useState(false);
  const wrapperRef = React.useRef<HTMLSpanElement>(null);
  const closeTimer = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const initials = getInitials(name);
  const popoverHeight = 200;

  const cancelClose = (): void => { if (closeTimer.current) { clearTimeout(closeTimer.current); closeTimer.current = null; } };
  const scheduleClose = (): void => { cancelClose(); closeTimer.current = setTimeout(() => setOpen(false), 200); };
  React.useEffect(() => () => cancelClose(), []);

  const openPopover = (): void => {
    if (!email) return;
    const r = wrapperRef.current?.getBoundingClientRect();
    if (!r) return;
    const spaceBelow = window.innerHeight - r.bottom;
    const above = spaceBelow < popoverHeight + 16;
    setCoords({ x: r.left + r.width / 2, y: above ? r.top - 8 : r.bottom + 8, above });
    setOpen(true);
  };

  return (
    <span ref={wrapperRef} style={{ display: 'inline-flex', flexShrink: 0 }}
      onMouseEnter={() => { cancelClose(); openPopover(); }}
      onMouseLeave={scheduleClose}
    >
      {!failed && email ? (
        <img
          src={photoUrl(email, 'L')}
          alt={name}
          onError={() => setFailed(true)}
          style={{ width: size, height: size, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0, cursor: 'default' }}
        />
      ) : (
        <span style={{ width: size, height: size, borderRadius: '50%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', background: 'linear-gradient(135deg, #86bc25, #0076a8)', color: '#fff', fontSize: size * 0.4, fontWeight: 700, flexShrink: 0 }}>{initials}</span>
      )}

      {open && email && coords && (
        <span
          onMouseEnter={cancelClose}
          onMouseLeave={scheduleClose}
          style={{
            position: 'fixed',
            top: coords.above ? undefined : coords.y,
            bottom: coords.above ? window.innerHeight - coords.y : undefined,
            left: coords.x, transform: 'translateX(-50%)', zIndex: 3000,
            background: '#fff', borderRadius: 10, padding: 12,
            boxShadow: '0 8px 24px rgba(0,0,0,0.18)', border: '1px solid var(--dex-gray-200)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6, minWidth: 180,
          }}
        >
          {!failed && (
            <img src={photoUrl(email, 'L')} alt={name} onError={() => setFailed(true)}
              style={{ width: 96, height: 96, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-200)' }} />
          )}
          <span style={{ fontSize: '0.82rem', fontWeight: 700, color: 'var(--dex-gray-800)', textAlign: 'center' }}>{name}</span>
          {subline && <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', textAlign: 'center' }}>{subline}</span>}
          <a href={`mailto:${email}`}
            style={{ fontSize: '0.75rem', color: 'var(--dex-green, #86bc25)', textDecoration: 'none', fontWeight: 600, whiteSpace: 'nowrap' }}
            onMouseEnter={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'underline'; }}
            onMouseLeave={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'none'; }}
          >{email}</a>
          <a href={teamsChatUrl(email)}
            style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: '0.72rem', color: '#6264A7', textDecoration: 'none', fontWeight: 600, whiteSpace: 'nowrap' }}
            onMouseEnter={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'underline'; }}
            onMouseLeave={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'none'; }}
          >
            <Icon iconName="TeamsLogo" style={{ fontSize: 13 }} /> {isDe ? 'Teams-Chat' : 'Teams chat'}
          </a>
        </span>
      )}
    </span>
  );
}

export default PersonContactHover;
