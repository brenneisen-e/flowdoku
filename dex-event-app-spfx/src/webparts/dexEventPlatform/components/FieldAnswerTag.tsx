import * as React from 'react';

/**
 * v27.10 (Refactor): unverändert aus MyEventsPage.tsx extrahiert.
 *
 * v19.34: People-Picker-Antworten (Feldtyp `user`/`roommate`) im „Meine
 * Events"-Antwort-Tag mit Profilfoto statt als Rohtext „Name <email>"
 * anzeigen — analog zum Chip im People-Picker selbst.
 */
const parsePersonAnswer = (v: string): { name: string; email: string } | null => {
  const m = (v || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
  if (!m) return null;
  return { name: m[1].trim(), email: m[2].trim() };
};

export default function FieldAnswerTag(props: { label: string; value: string; type?: string; small?: boolean }): React.ReactElement {
  const { label, value, type, small } = props;
  const person = (type === 'user' || type === 'roommate') ? parsePersonAnswer(value) : null;
  const baseStyle: React.CSSProperties = {
    fontSize: small ? '0.72rem' : '0.78rem',
    padding: small ? '3px 8px' : '4px 10px',
    borderRadius: 4,
    background: 'rgba(134,188,37,0.14)',
    color: 'var(--dex-green-dark, #4a7c1f)',
    border: '1px solid rgba(134,188,37,0.30)',
  };
  if (person) {
    return (
      <span style={{ ...baseStyle, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        {label}:
        <img
          src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(person.email)}&size=L`}
          alt={person.name}
          onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
          style={{ width: 22, height: 22, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', transition: 'transform 0.15s', transformOrigin: 'center' }}
          onMouseEnter={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(2.6)'; (e.currentTarget as HTMLImageElement).style.zIndex = '20'; (e.currentTarget as HTMLImageElement).style.position = 'relative'; (e.currentTarget as HTMLImageElement).style.boxShadow = '0 6px 18px rgba(0,0,0,0.18)'; }}
          onMouseLeave={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(1)'; (e.currentTarget as HTMLImageElement).style.boxShadow = 'none'; }}
        />
        <strong>{person.name}</strong>
      </span>
    );
  }
  return (
    <span style={baseStyle}>
      {label}: <strong>{value}</strong>
    </span>
  );
}
