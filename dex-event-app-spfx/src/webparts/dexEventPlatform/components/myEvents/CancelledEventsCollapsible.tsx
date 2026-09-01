/* CancelledEventsCollapsible — aus MyEventsPage.tsx ausgelagert (Zeilen
 * 3597-3674 des urspruenglichen Stands, v30.65). Einklappbare Liste der
 * abgemeldeten Events. Der Code ist zeichengleich uebernommen.
 */
import * as React from 'react';

// v11.97: Einklappbare Liste der abgemeldeten Events. Default eingeklappt,
// damit lange Cancelled-Listen die Hauptliste nicht überlagern. Header
// zeigt Count + Chevron, Klick togglt die Liste.
export default function CancelledEventsCollapsible(props: {
  count: number;
  title: string;
  locale: string;
  entries: Array<{ event: { id: string; title: string }; registration: { CancellationDate?: string } }>;
  formatDate: (iso: string) => string;
  statusLabel: string;
  // v18.62: Hinweis + Button, dass man sich für eine erneute Teilnahme
  // im Anmelde-Bereich neu registrieren muss (Abmeldung ist endgültig).
  hintText: string;
  reRegisterLabel: string;
  onReRegister: () => void;
}): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  return (
    <div className="mt-24">
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        style={{
          background: 'none', border: 'none', padding: 0, cursor: 'pointer',
          display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12,
          color: 'var(--dex-gray-600)', fontSize: '1rem', fontWeight: 600,
        }}
        aria-expanded={open}
      >
        <span style={{
          display: 'inline-block', width: 14, textAlign: 'center',
          transform: open ? 'rotate(90deg)' : 'rotate(0deg)',
          transition: 'transform 0.15s ease',
          fontSize: '0.85rem',
        }}>▶</span>
        <span>{props.title} ({props.count})</span>
      </button>
      {open && (
        <div className="my-events-list">
          <div style={{
            fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginBottom: 12,
            padding: '10px 14px', borderRadius: 8,
            background: 'var(--dex-gray-50, #fafafa)', border: '1px solid var(--dex-gray-200)',
            display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap',
          }}>
            <span style={{ flex: '1 1 auto' }}>{props.hintText}</span>
            <button
              type="button"
              className="btn btn-outline"
              style={{ fontSize: '0.78rem', padding: '5px 14px', whiteSpace: 'nowrap' }}
              onClick={props.onReRegister}
            >{props.reRegisterLabel}</button>
          </div>
          {props.entries.map(({ event, registration }) => (
            <div
              key={event.id}
              className="card my-event-card"
              style={{
                opacity: 0.6,
                display: 'flex',
                alignItems: 'center',
                gap: 12,
                padding: '8px 16px',
                flexWrap: 'wrap',
              }}
            >
              <strong style={{ flex: '1 1 auto', fontSize: '0.95rem', margin: 0 }}>{event.title}</strong>
              <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)' }}>
                {props.locale}: {registration.CancellationDate ? props.formatDate(registration.CancellationDate) : '-'}
              </span>
              <span className="badge badge-red">{props.statusLabel}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
