/**
 * v28.94: Aus `EventCreationPage` herausgeloest. Die kleine Nummer vor einer
 * Feld-Beschriftung — Support und Tooltips verweisen darauf.
 */
import * as React from 'react';

export function StepBadge({ n }: { n: number }): React.ReactElement {
  // v22.29: Aufräum-Pass — dezenter Outline-Badge statt gefüllter grüner
  // Kreis. Die durchlaufenden Nummern bleiben (Tooltips/Support referenzieren
  // sie), treten optisch aber hinter die eigentlichen Labels zurück.
  return (
    <span style={{
      display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
      width: 20, height: 20, borderRadius: '50%',
      background: '#fff', color: 'var(--dex-green-dark, #4a7c1f)',
      border: '1.5px solid var(--dex-green, #86bc25)',
      // v22.31: lineHeight 1 — die geerbte line-height (1.5) schob die
      // Ziffer aus der vertikalen Mitte des Kreises.
      fontSize: '0.66rem', fontWeight: 700, flexShrink: 0, lineHeight: 1,
      boxSizing: 'border-box',
    }}>{n}</span>
  );
}
