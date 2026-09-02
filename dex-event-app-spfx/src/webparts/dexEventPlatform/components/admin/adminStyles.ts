/**
 * Modul-Ebene aus AdminPage.tsx ausgelagert (v30.66) — geteilte Style-Objekte.
 */
import * as React from 'react';

/**
 * v30.36: Optik der Auswahl-Karten im „QR-Codes und Check-In"-Modal.
 * Als Konstante, weil vier Karten sie teilen — vier inline kopierte
 * Style-Objekte laufen erfahrungsgemaess auseinander.
 */
/** v30.36: Zeilen-Optik der Aufklapper im QR-Versand-Modal. */
export const qrDisclosureStyle: React.CSSProperties = {
  background: 'none', border: 'none', padding: '4px 0', cursor: 'pointer',
  color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.82rem', textAlign: 'left',
  font: 'inherit', fontFamily: 'inherit', fontWeight: 600,
};

export const hubChoiceStyle: React.CSSProperties = {
  display: 'block', width: '100%', textAlign: 'left', cursor: 'pointer',
  padding: '16px 18px', borderRadius: 12,
  border: '1px solid var(--dex-gray-200)', background: '#fff',
  font: 'inherit', fontFamily: 'inherit',
};
