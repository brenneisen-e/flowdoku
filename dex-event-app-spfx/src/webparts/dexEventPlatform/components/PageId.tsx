/**
 * v9.19/v9.20: PageId — kleines unauffaelliges Label am Seitenende mit der
 * aktuellen Seiten-ID.
 *
 * Zweck: User kann mir bei UI-Anfragen praezise sagen, welche Seite er meint
 * ("PAGE-ID: events → Card ABC aendern").
 *
 * v9.20-Fix: position:fixed funktioniert in SPFx-Webparts nicht zuverlaessig
 * (transformierter Parent macht aus fixed effektiv absolute → relativ zum
 * Webpart-Container, und der ist oft hidden/clipped). Statt dessen rendern
 * wir das Label als Inline-Block am unteren linken Rand des main-Containers
 * — sticht nicht heraus, ist aber immer beim Scroll-Ende sichtbar.
 */
import * as React from 'react';

export function PageId(props: { id: string }): React.ReactElement {
  return (
    <div
      style={{
        // sticky innerhalb des main-Containers
        position: 'sticky',
        bottom: 0,
        left: 0,
        // Inline-Position: Tag rutscht mit nach links unten beim Scrollen
        marginTop: 24,
        padding: '4px 10px 4px 12px',
        fontSize: '0.7rem',
        fontFamily: 'monospace',
        color: 'var(--dex-gray-500, #6b7280)',
        background: 'rgba(255,255,255,0.95)',
        borderTop: '1px solid var(--dex-gray-200, #e5e7eb)',
        borderRight: '1px solid var(--dex-gray-200, #e5e7eb)',
        borderTopRightRadius: 6,
        opacity: 0.7,
        userSelect: 'text',
        letterSpacing: '0.02em',
        // selbst-positionierend am linken Rand, nicht volle Breite
        display: 'inline-block',
        alignSelf: 'flex-start',
        zIndex: 5,
      }}
      aria-hidden="true"
      title="Page-ID — bei UI-Anfragen kannst du diese ID nennen, dann finde ich die Seite sofort."
    >
      PAGE-ID: {props.id}
    </div>
  );
}
