/**
 * v31.2: Hinweis im Wizard, dass gerade jemand anderes dasselbe Event
 * bearbeitet — mit Foto und pulsierendem Ring (Nutzer 07.09.2026: „das soll
 * dann pulsieren und die Person auch mit Foto angezeigt werden").
 *
 * Warum ein Warn-Kasten und kein Sperren: Der Wizard kennt kein Zeilen-Lock;
 * wer zuletzt speichert, überschreibt. Die Anzeige soll die beiden Personen
 * zum Absprechen bringen, nicht eine davon aussperren — ein Lock, den ein
 * geschlossener Tab hält, wäre die schlimmere Falle.
 */
import * as React from 'react';
import type { EditPresenceEntry } from '../../services/events/editPresence';

function initials(name: string, email: string): string {
  const src = (name || '').trim() || (email || '').split('@')[0].replace(/[._-]+/g, ' ');
  const parts = src.split(/[\s,]+/).filter(Boolean);
  if (parts.length >= 2) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  return (src.slice(0, 2) || '?').toUpperCase();
}

/** Schreibweise „Nachname, Vorname" (Verzeichnis) in „Vorname Nachname" drehen. */
function displayName(name: string, email: string): string {
  const n = (name || '').trim();
  if (n.indexOf(',') > 0) {
    const [last, first] = n.split(',').map(s => s.trim());
    if (first && last) return `${first} ${last}`;
  }
  return n || email;
}

const PresenceAvatar: React.FC<{ entry: EditPresenceEntry }> = ({ entry }) => {
  const [broken, setBroken] = React.useState(false);
  const src = entry.email ? `/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(entry.email)}` : '';
  return (
    <span
      className="dex-ui-avatar dex-ui-avatar--lg dex-ui-pulse"
      title={displayName(entry.name, entry.email)}
      style={{ border: '2px solid #fff', background: 'var(--dex-gray-200, #e8e8e8)' }}
    >
      {src && !broken
        ? <img src={src} alt="" onError={() => setBroken(true)} style={{ width: '100%', height: '100%', objectFit: 'cover', borderRadius: '50%' }} />
        : <span aria-hidden="true">{initials(entry.name, entry.email)}</span>}
    </span>
  );
};

export const EditPresenceBadge: React.FC<{ others: EditPresenceEntry[]; isDe: boolean }> = ({ others, isDe }) => {
  if (!others.length) return null;
  const names = others.map(o => displayName(o.name, o.email));
  const who = names.length === 1
    ? names[0]
    : names.length === 2
      ? `${names[0]} ${isDe ? 'und' : 'and'} ${names[1]}`
      : `${names.slice(0, -1).join(', ')} ${isDe ? 'und' : 'and'} ${names[names.length - 1]}`;
  return (
    <div
      role="status"
      aria-live="polite"
      className="dex-ui-callout dex-ui-callout--warn dex-ui-fade-in"
      style={{ alignItems: 'center', gap: 14, margin: '0 0 16px', padding: '12px 16px' }}
    >
      <span className="dex-ui-avatar-stack" style={{ flexShrink: 0 }}>
        {others.slice(0, 3).map(o => <PresenceAvatar key={o.email || o.itemId} entry={o} />)}
      </span>
      <span style={{ flex: 1, minWidth: 0 }}>
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontWeight: 700, fontSize: '0.9rem' }}>
          <span aria-hidden="true" style={{ width: 8, height: 8, borderRadius: '50%', background: 'var(--dex-orange, #ed8b00)', boxShadow: '0 0 0 3px rgba(237,139,0,0.25)', flexShrink: 0 }} />
          {isDe
            ? `${who} ${names.length === 1 ? 'bearbeitet' : 'bearbeiten'} dieses Event gerade auch.`
            : `${who} ${names.length === 1 ? 'is' : 'are'} also editing this event right now.`}
        </span>
        <span style={{ display: 'block', fontSize: '0.8rem', marginTop: 3, opacity: 0.9 }}>
          {isDe
            ? 'Sprecht euch kurz ab: Wer zuletzt speichert, überschreibt die Änderungen der anderen.'
            : 'Have a quick word: whoever saves last overwrites the changes of the others.'}
        </span>
      </span>
    </div>
  );
};
