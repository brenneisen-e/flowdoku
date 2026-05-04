/**
 * v9.19: PageId — kleines unauffaelliges Label unten links auf jeder Seite.
 *
 * Zweck: User kann mir bei UI-Anfragen praezise sagen, welche Seite er meint
 * ("PAGE-ID: events → Card ABC aendern"). Ohne dieses Label muesste der User
 * Screenshots schicken oder umstaendlich beschreiben.
 *
 * Styling: dezent in der unteren linken Ecke, monospace, niedrige Opacity —
 * stoert nicht beim normalen Use, ist aber jederzeit lesbar.
 */
import * as React from 'react';

export function PageId(props: { id: string }): React.ReactElement {
  return (
    <div
      style={{
        position: 'fixed',
        bottom: 8,
        left: 12,
        fontSize: '0.65rem',
        fontFamily: 'monospace',
        color: 'var(--dex-gray-400, #9ca3af)',
        background: 'rgba(255,255,255,0.85)',
        padding: '2px 8px',
        borderRadius: 4,
        border: '1px solid var(--dex-gray-200, #e5e7eb)',
        opacity: 0.6,
        pointerEvents: 'none',
        zIndex: 100,
        userSelect: 'text',
        letterSpacing: '0.02em',
      }}
      aria-hidden="true"
    >
      PAGE-ID: {props.id}
    </div>
  );
}
