/**
 * Rollen-Matrix - Uebersicht aller Berechtigungen pro Rolle
 *
 * Zeigt eine Tabelle mit allen Funktionen und welche Rolle Zugriff hat.
 * Nur fuer SuperAdmin zugaenglich.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';

interface PermissionRow {
  category: string;
  feature: string;
  user: boolean | string;
  eventAdmin: boolean | string;
  superAdmin: boolean | string;
}

const PERMISSIONS: PermissionRow[] = [
  // Events
  { category: 'Events', feature: 'Events ansehen', user: true, eventAdmin: true, superAdmin: true },
  { category: 'Events', feature: 'Fuer Events registrieren', user: true, eventAdmin: true, superAdmin: true },
  { category: 'Events', feature: 'Events erstellen', user: false, eventAdmin: true, superAdmin: true },
  { category: 'Events', feature: 'Eigene Events bearbeiten', user: false, eventAdmin: true, superAdmin: true },
  { category: 'Events', feature: 'Alle Events bearbeiten', user: false, eventAdmin: false, superAdmin: true },

  // Registrierungen
  { category: 'Registrierungen', feature: 'Eigene Registrierungen sehen', user: true, eventAdmin: true, superAdmin: true },
  { category: 'Registrierungen', feature: 'Eigene Registrierung stornieren', user: true, eventAdmin: true, superAdmin: true },
  { category: 'Registrierungen', feature: 'Fuer andere registrieren', user: true, eventAdmin: true, superAdmin: true },
  { category: 'Registrierungen', feature: 'Teilnehmerliste eigener Events sehen', user: false, eventAdmin: true, superAdmin: true },
  { category: 'Registrierungen', feature: 'Alle Teilnehmerlisten sehen', user: false, eventAdmin: false, superAdmin: true },

  // Administration
  { category: 'Administration', feature: 'Admin-Bereich oeffnen', user: false, eventAdmin: true, superAdmin: true },
  { category: 'Administration', feature: 'Rollen verwalten', user: false, eventAdmin: false, superAdmin: true },
  { category: 'Administration', feature: 'Rollen-Matrix einsehen', user: false, eventAdmin: false, superAdmin: true },
  { category: 'Administration', feature: 'User suchen', user: false, eventAdmin: false, superAdmin: true },

  // SharePoint
  { category: 'SharePoint', feature: 'DEX_Events Liste: Lesen', user: true, eventAdmin: true, superAdmin: true },
  { category: 'SharePoint', feature: 'DEX_Events Liste: Schreiben', user: false, eventAdmin: true, superAdmin: true },
  { category: 'SharePoint', feature: 'DEX_Roles Liste: Lesen', user: false, eventAdmin: 'Nur eigene', superAdmin: true },
  { category: 'SharePoint', feature: 'DEX_Roles Liste: Schreiben', user: false, eventAdmin: false, superAdmin: true },
  { category: 'SharePoint', feature: 'Event-Subsites: Full Control', user: false, eventAdmin: 'Eigene Events', superAdmin: true },
  { category: 'SharePoint', feature: 'Teilnehmerliste: Eigene Eintraege', user: true, eventAdmin: true, superAdmin: true },
  { category: 'SharePoint', feature: 'Teilnehmerliste: Alle Eintraege', user: false, eventAdmin: 'Eigene Events', superAdmin: true },

  // Profil
  { category: 'Profil', feature: 'Eigenes Profil ansehen', user: true, eventAdmin: true, superAdmin: true },
  { category: 'Profil', feature: 'Settings-Seite oeffnen', user: true, eventAdmin: true, superAdmin: true },
];

function renderCell(value: boolean | string): React.ReactElement {
  if (value === true) {
    return <span style={{ color: '#22c55e', fontWeight: 600, fontSize: '1.1rem' }}>&#10003;</span>;
  }
  if (value === false) {
    return <span style={{ color: '#ef4444', fontWeight: 500, fontSize: '1.1rem' }}>&mdash;</span>;
  }
  return <span style={{ color: '#f59e0b', fontSize: '0.8rem', fontWeight: 500 }}>{value}</span>;
}

export default function RoleMatrixPage(): React.ReactElement {
  const { navigate } = useNavigation();

  // Kategorien gruppieren
  const categories = Array.from(new Set(PERMISSIONS.map(p => p.category)));

  return (
    <div className="page-container">
      <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
        {/* Header */}
        <div style={{
          padding: '24px 28px 16px',
          borderBottom: '1px solid var(--dex-gray-200)',
          background: 'linear-gradient(135deg, rgba(134,239,172,0.08) 0%, rgba(59,130,246,0.06) 100%)',
        }}>
          <h2 style={{ margin: 0, fontSize: '1.2rem', fontWeight: 700 }}>Rollen-Matrix</h2>
          <p style={{ margin: '6px 0 0', color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
            Uebersicht aller Berechtigungen nach Rolle
          </p>
        </div>

        {/* Tabelle */}
        <div style={{ overflowX: 'auto' }}>
          <table style={{
            width: '100%',
            borderCollapse: 'collapse',
            fontSize: '0.85rem',
          }}>
            <thead>
              <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                <th style={{ ...thStyle, width: '40%' }}>Funktion</th>
                <th style={{ ...thStyle, textAlign: 'center', width: '20%' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2 }}>
                    <span style={roleBadgeStyle('#3b82f6')}>User</span>
                  </div>
                </th>
                <th style={{ ...thStyle, textAlign: 'center', width: '20%' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2 }}>
                    <span style={roleBadgeStyle('#f59e0b')}>EventAdmin</span>
                  </div>
                </th>
                <th style={{ ...thStyle, textAlign: 'center', width: '20%' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2 }}>
                    <span style={roleBadgeStyle('#86efac')}>SuperAdmin</span>
                  </div>
                </th>
              </tr>
            </thead>
            <tbody>
              {categories.map(cat => {
                const rows = PERMISSIONS.filter(p => p.category === cat);
                return (
                  <React.Fragment key={cat}>
                    {/* Kategorie-Header */}
                    <tr>
                      <td
                        colSpan={4}
                        style={{
                          padding: '12px 20px 6px',
                          fontWeight: 700,
                          fontSize: '0.8rem',
                          textTransform: 'uppercase' as const,
                          letterSpacing: '0.05em',
                          color: 'var(--dex-gray-500)',
                          borderBottom: '1px solid var(--dex-gray-100)',
                        }}
                      >
                        {cat}
                      </td>
                    </tr>
                    {/* Feature-Zeilen */}
                    {rows.map((row, idx) => (
                      <tr
                        key={row.feature}
                        style={{
                          borderBottom: idx === rows.length - 1 ? '2px solid var(--dex-gray-100)' : '1px solid var(--dex-gray-50)',
                          background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.01)',
                        }}
                      >
                        <td style={{ padding: '10px 20px', color: 'var(--dex-gray-700)' }}>{row.feature}</td>
                        <td style={{ padding: '10px 20px', textAlign: 'center' }}>{renderCell(row.user)}</td>
                        <td style={{ padding: '10px 20px', textAlign: 'center' }}>{renderCell(row.eventAdmin)}</td>
                        <td style={{ padding: '10px 20px', textAlign: 'center' }}>{renderCell(row.superAdmin)}</td>
                      </tr>
                    ))}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Legende */}
        <div style={{
          padding: '16px 28px',
          borderTop: '1px solid var(--dex-gray-200)',
          display: 'flex',
          gap: 24,
          fontSize: '0.8rem',
          color: 'var(--dex-gray-500)',
        }}>
          <span><span style={{ color: '#22c55e', fontWeight: 600 }}>&#10003;</span> = Berechtigt</span>
          <span><span style={{ color: '#ef4444', fontWeight: 500 }}>&mdash;</span> = Kein Zugriff</span>
          <span><span style={{ color: '#f59e0b', fontWeight: 500 }}>Text</span> = Eingeschraenkt</span>
        </div>
      </div>

      <div style={{ marginTop: 16, textAlign: 'center' }}>
        <button className="btn btn-secondary" onClick={() => navigate('settings')}>
          Zurueck zu Settings
        </button>
      </div>
    </div>
  );
}

const thStyle: React.CSSProperties = {
  padding: '14px 20px',
  textAlign: 'left',
  fontWeight: 600,
  fontSize: '0.85rem',
  color: 'var(--dex-gray-700)',
};

function roleBadgeStyle(color: string): React.CSSProperties {
  return {
    display: 'inline-block',
    padding: '3px 10px',
    borderRadius: 12,
    background: `${color}18`,
    color: color,
    fontWeight: 600,
    fontSize: '0.78rem',
  };
}
