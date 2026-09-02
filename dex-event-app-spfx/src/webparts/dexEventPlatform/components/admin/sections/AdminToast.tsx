/* AdminToast — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 6783-6842 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';

export interface AdminToastProps {
  adminToast: { kind: "cancelling"; name: string; } | { kind: "promoted"; name: string; email: string; type?: string; } | { kind: "no-promote"; name: string; };
  isMobile: boolean;
  setAdminToast: React.Dispatch<React.SetStateAction<{ kind: "cancelling"; name: string; } | { kind: "promoted"; name: string; email: string; type?: string; } | { kind: "no-promote"; name: string; }>>;
}

export const AdminToast: React.FC<AdminToastProps> = (p) => {
  const { adminToast, isMobile, setAdminToast } = p;
        const accent = adminToast.kind === 'cancelling'
          ? 'var(--dex-orange, #ed8b00)'
          : adminToast.kind === 'promoted'
            ? 'var(--dex-green, #86bc25)'
            : 'var(--dex-gray-400)';
        const accentDark = adminToast.kind === 'cancelling'
          ? 'var(--dex-orange, #ed8b00)'
          : adminToast.kind === 'promoted'
            ? 'var(--dex-green-dark, #6b9a1e)'
            : 'var(--dex-gray-600)';
        const closable = adminToast.kind !== 'cancelling';
        return (
          <div style={{
            position: 'fixed', top: 80, zIndex: 1000,
            ...(isMobile
              ? { left: 12, right: 12, maxWidth: 'min(460px, calc(100vw - 24px))' }
              : { right: 20, maxWidth: 460 }),
            padding: '14px 18px', borderRadius: 'var(--dex-radius, 12px)',
            background: '#fff',
            border: `1px solid ${accent}`,
            borderLeft: `4px solid ${accent}`,
            boxShadow: '0 6px 20px rgba(0,0,0,0.15)',
            display: 'flex', alignItems: 'flex-start', gap: 12,
          }}>
            {adminToast.kind === 'cancelling' && (
              <div style={{
                width: 20, height: 20, marginTop: 2, flexShrink: 0,
                border: `3px solid var(--dex-gray-200)`,
                borderTopColor: accent,
                borderRadius: '50%',
                animation: 'dex-spin 0.8s linear infinite',
              }} />
            )}
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: 700, color: accentDark, marginBottom: 4 }}>
                {adminToast.kind === 'cancelling' && `Abmeldung von ${adminToast.name} wird verarbeitet…`}
                {adminToast.kind === 'promoted' && `Nachgerückt: ${adminToast.name}`}
                {adminToast.kind === 'no-promote' && `Abmeldung von ${adminToast.name} verarbeitet`}
              </div>
              <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                {adminToast.kind === 'cancelling' && 'Teilnehmer wird abgemeldet, Warteliste wird geprüft und ggf. ein Nachrücker informiert.'}
                {adminToast.kind === 'promoted' && (
                  <>
                    <strong>{adminToast.email}</strong> wurde automatisch aus der Warteliste{adminToast.type ? ` (${adminToast.type})` : ''} nachgerückt. Nachrück-Mail + Outlook-Einladung wurden versendet.
                  </>
                )}
                {adminToast.kind === 'no-promote' && 'Aktuell ist niemand auf der Warteliste (bzw. kein passender Starter-Typ). Der Platz bleibt frei.'}
              </div>
            </div>
            {closable && (
              <button
                onClick={() => setAdminToast(null)}
                aria-label="Schließen"
                style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.2rem', color: 'var(--dex-gray-500)', lineHeight: 1, padding: 0 }}
              >×</button>
            )}
          </div>
        );
};

