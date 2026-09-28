/* v31.100: Profil-Karte „automatisch übernommene Daten" als EINE Komponente.
 *
 * Bis v31.99 stand sie nur in der Anmeldeseite (PersonalDataSection, v27.13).
 * Seit der Wizard in Schritt 5 dieselbe Karte als Beispiel zeigt (Nutzer-
 * Ansage 28.09.2026: „als Beispiel immer das hier … gleiches Design wie bei
 * Anmeldung"), liegt sie hier — zwei Kopien desselben Kastens laufen
 * erfahrungsgemäß auseinander (CLAUDE.md, „Vor einer neuen Ansicht prüfen …").
 *
 * Rein darstellend: Foto (userphoto.aspx mit Initialen-Fallback), Name,
 * Position, Ort, aufklappbare Zeilen und ein Fußsatz. */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';

export interface ProfileDataCardProps {
  displayName: string;
  initials: string;
  /** E-Mail für das Profilfoto; leer = nur Initialen. */
  photoEmail?: string;
  jobTitle?: string;
  location?: string;
  rows: Array<{ label: string; value: string }>;
  rowsTitle: string;
  notSetLabel: string;
  expanded: boolean;
  /** Ohne Umschalter bleiben die Zeilen fest offen (Beispiel im Wizard). */
  onToggle?: () => void;
  toggleTitleOpen?: string;
  toggleTitleClosed?: string;
  footer?: React.ReactNode;
}

export const ProfileDataCard: React.FC<ProfileDataCardProps> = (p) => (
  <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: '16px 18px', background: '#fff' }}>
    <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
      {/* Foto: userphoto.aspx mit Initialen-Fallback (Bild liegt über dem
          Initialen-Kreis; bei Ladefehler wird es ausgeblendet). */}
      <div style={{ position: 'relative', width: 88, height: 88, borderRadius: '50%', background: 'var(--dex-gray-100)', flexShrink: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: '1.5rem', color: 'var(--dex-gray-500)', overflow: 'hidden' }}>
        {p.initials || '?'}
        {p.photoEmail && (
          <img
            src={`/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(p.photoEmail.trim())}`}
            alt=""
            style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'cover' }}
            onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
          />
        )}
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontWeight: 700, fontSize: '1.18rem', color: 'var(--dex-gray-800)' }}>{p.displayName}</div>
        {p.jobTitle && (
          <div style={{ color: 'var(--dex-gray-600)', marginTop: 2 }}>{p.jobTitle}</div>
        )}
        {p.location && (
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: 5, color: 'var(--dex-gray-500)', fontSize: '0.88rem', marginTop: 3 }}>
            <Icon iconName="POI" style={{ fontSize: 14, color: 'var(--dex-green-dark, #4a7c1f)' }} />
            {p.location}
          </div>
        )}
      </div>
      {p.onToggle && (
        <button
          type="button"
          onClick={p.onToggle}
          title={p.expanded ? p.toggleTitleOpen : p.toggleTitleClosed}
          aria-expanded={p.expanded}
          style={{
            width: 36, height: 36, borderRadius: '50%', flexShrink: 0,
            border: '1px solid var(--dex-gray-300)', background: p.expanded ? 'var(--dex-gray-100)' : '#fff',
            cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center',
            fontSize: '1.25rem', lineHeight: 1, color: 'var(--dex-gray-600)', fontWeight: 600,
          }}
        >
          {p.expanded ? '−' : '+'}
        </button>
      )}
    </div>
    {p.expanded && (
      <div style={{ marginTop: 14, borderTop: '1px solid var(--dex-gray-100)', paddingTop: 10 }}>
        <div style={{ fontSize: '0.72rem', fontWeight: 700, color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 6 }}>
          {p.rowsTitle}
        </div>
        {p.rows.map(row => (
          <div key={row.label} style={{ display: 'flex', gap: 10, padding: '4px 0', fontSize: '0.86rem', borderBottom: '1px solid var(--dex-gray-50, #fafafa)' }}>
            <span style={{ width: 140, flexShrink: 0, color: 'var(--dex-gray-500)' }}>{row.label}</span>
            <span style={{ color: row.value ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)', wordBreak: 'break-word' }}>
              {row.value || `— ${p.notSetLabel}`}
            </span>
          </div>
        ))}
      </div>
    )}
    {p.footer && (
      <div style={{ marginTop: 10, fontSize: '0.68rem', fontStyle: 'italic', color: 'var(--dex-gray-400)', lineHeight: 1.45 }}>
        {p.footer}
      </div>
    )}
  </div>
);
