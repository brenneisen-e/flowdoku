/**
 * Rollen-Matrix - Übersicht aller Berechtigungen pro Rolle
 *
 * Nur für Admin zugänglich.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';

interface PermissionRow {
  category: string;
  feature: string;
  user: boolean | string;
  organizer: boolean | string;
  admin: boolean | string;
}

const PERMISSIONS: PermissionRow[] = [
  // Events ansehen
  { category: 'Events ansehen', feature: 'Events des eigenen Standorts sehen', user: true, organizer: true, admin: true },
  { category: 'Events ansehen', feature: 'Alle Events sehen', user: false, organizer: true, admin: true },
  { category: 'Events ansehen', feature: 'Prozess-Übersicht (Flowcharts)', user: false, organizer: true, admin: true },

  // Event-Verwaltung
  { category: 'Event-Verwaltung', feature: 'Events erstellen', user: false, organizer: true, admin: true },
  { category: 'Event-Verwaltung', feature: 'Eigene Events bearbeiten', user: false, organizer: true, admin: true },
  { category: 'Event-Verwaltung', feature: 'Alle Events bearbeiten', user: false, organizer: false, admin: true },
  { category: 'Event-Verwaltung', feature: 'Events löschen', user: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Event-Bild hochladen (Item-Attachment)', user: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Event-Dokumente hochladen', user: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Agenda / Transferzeiten / Quiz pflegen', user: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'E-Mail-Templates pro Event anpassen', user: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'E-Mails pro Event deaktivieren', user: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Outlook-Einladungen pro Event deaktivieren', user: false, organizer: 'Eigene', admin: true },

  // Registrierungen
  { category: 'Registrierungen', feature: 'Selbst registrieren', user: true, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Eigene Angaben bearbeiten', user: true, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Eigene Registrierung stornieren', user: true, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Für andere registrieren', user: 'Assistant: nur Partner/Director ¹', organizer: 'Eigene Events ²', admin: true },
  { category: 'Registrierungen', feature: 'Nach Anmeldefrist registrieren', user: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Registrierungen', feature: 'Audit-Trail: RegisteredBy wird automatisch gesetzt', user: true, organizer: true, admin: true },

  // Teilnehmerverwaltung (Admin Center)
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmerliste sehen', user: false, organizer: 'Eigene Events', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmer suchen / sortieren', user: false, organizer: 'Eigene Events', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmer ein-/auschecken', user: false, organizer: 'Eigene Events', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmer abmelden', user: false, organizer: 'Eigene Events', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'QR-Codes versenden', user: false, organizer: 'Eigene Events', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'E-Mail-Adressen kopieren', user: false, organizer: 'Eigene Events', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Massenmail an Teilnehmer (RichText + Template)', user: false, organizer: 'Eigene Events', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'IDs neu vergeben (sequentielle Renummerierung)', user: false, organizer: false, admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Spalten fixen (Schema reparieren)', user: false, organizer: false, admin: true },

  // Administration
  { category: 'Administration', feature: 'Admin-Bereich öffnen', user: false, organizer: true, admin: true },
  { category: 'Administration', feature: 'Rollen verwalten', user: false, organizer: false, admin: true },
  { category: 'Administration', feature: 'Rollen-Matrix einsehen', user: false, organizer: false, admin: true },
  { category: 'Administration', feature: 'User suchen', user: false, organizer: false, admin: true },
  { category: 'Administration', feature: 'Standort-Filter konfigurieren', user: false, organizer: true, admin: true },
  { category: 'Administration', feature: 'Zielgruppen-Filter konfigurieren', user: false, organizer: true, admin: true },
  { category: 'Administration', feature: 'Globale Email-Templates bearbeiten', user: false, organizer: false, admin: true },

  // SharePoint (Visitors = DEALL, Owners = Admins)
  { category: 'SharePoint', feature: 'DEX_Events: Lesen', user: 'Visitors (Read)', organizer: 'Contribute', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Events: Schreiben (inkl. Item-Attachments)', user: false, organizer: 'Contribute', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Roles: Lesen', user: false, organizer: 'Read', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Roles: Schreiben', user: false, organizer: false, admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Emails: Queue (eigene)', user: 'Contribute + ILS', organizer: 'Contribute + ILS', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Outlook: Queue (eigene)', user: 'Contribute + ILS', organizer: 'Contribute + ILS', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_IDReorder: Queue', user: false, organizer: false, admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_EmailTemplates: Schreiben', user: false, organizer: false, admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Participants: eigene Einträge', user: 'Contribute + ILS', organizer: 'Contribute + ILS', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'Event-Subsite', user: 'Visitors (Read)', organizer: 'Full Control', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'Teilnehmerliste: eigener Eintrag', user: 'Contribute + ILS', organizer: 'Full Control', admin: 'Full Control' },

  // Profil
  { category: 'Profil', feature: 'Eigenes Profil ansehen', user: true, organizer: true, admin: true },
  { category: 'Profil', feature: 'Settings-Seite öffnen', user: true, organizer: true, admin: true },
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

  const categories = Array.from(new Set(PERMISSIONS.map(p => p.category)));

  return (
    <div className="page-container">
      <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
        <div style={{
          padding: '24px 28px 16px',
          borderBottom: '1px solid var(--dex-gray-200)',
          background: 'linear-gradient(135deg, rgba(134,239,172,0.08) 0%, rgba(59,130,246,0.06) 100%)',
        }}>
          <h2 style={{ margin: 0, fontSize: '1.2rem', fontWeight: 700 }}>Rollen-Matrix</h2>
          <p style={{ margin: '6px 0 0', color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
            Übersicht aller Berechtigungen nach Rolle
          </p>
        </div>

        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
            <thead>
              <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                <th style={{ ...thStyle, width: '40%' }}>Funktion</th>
                <th style={{ ...thStyle, textAlign: 'center', width: '20%' }}>
                  <span style={roleBadgeStyle('#3b82f6')}>User</span>
                </th>
                <th style={{ ...thStyle, textAlign: 'center', width: '20%' }}>
                  <span style={roleBadgeStyle('#f59e0b')}>Organizer</span>
                </th>
                <th style={{ ...thStyle, textAlign: 'center', width: '20%' }}>
                  <span style={roleBadgeStyle('#86efac')}>Admin</span>
                </th>
              </tr>
            </thead>
            <tbody>
              {categories.map(cat => {
                const rows = PERMISSIONS.filter(p => p.category === cat);
                return (
                  <React.Fragment key={cat}>
                    <tr>
                      <td colSpan={4} style={{
                        padding: '12px 20px 6px', fontWeight: 700, fontSize: '0.8rem',
                        textTransform: 'uppercase' as const, letterSpacing: '0.05em',
                        color: 'var(--dex-gray-500)', borderBottom: '1px solid var(--dex-gray-100)',
                      }}>
                        {cat}
                      </td>
                    </tr>
                    {rows.map((row, idx) => (
                      <tr key={row.feature} style={{
                        borderBottom: idx === rows.length - 1 ? '2px solid var(--dex-gray-100)' : '1px solid var(--dex-gray-50)',
                        background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.01)',
                      }}>
                        <td style={{ padding: '10px 20px', color: 'var(--dex-gray-700)' }}>{row.feature}</td>
                        <td style={{ padding: '10px 20px', textAlign: 'center' }}>{renderCell(row.user)}</td>
                        <td style={{ padding: '10px 20px', textAlign: 'center' }}>{renderCell(row.organizer)}</td>
                        <td style={{ padding: '10px 20px', textAlign: 'center' }}>{renderCell(row.admin)}</td>
                      </tr>
                    ))}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>

        <div style={{
          padding: '16px 28px', borderTop: '1px solid var(--dex-gray-200)',
          display: 'flex', gap: 24, fontSize: '0.8rem', color: 'var(--dex-gray-500)',
          flexWrap: 'wrap',
        }}>
          <span><span style={{ color: '#22c55e', fontWeight: 600 }}>&#10003;</span> = Berechtigt</span>
          <span><span style={{ color: '#ef4444', fontWeight: 500 }}>&mdash;</span> = Kein Zugriff</span>
          <span><span style={{ color: '#f59e0b', fontWeight: 500 }}>Text</span> = Eingeschränkt</span>
        </div>
        <div style={{
          padding: '12px 28px 20px', borderTop: '1px dashed var(--dex-gray-200)',
          fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.6,
        }}>
          <div style={{ marginBottom: 4 }}>
            <strong>¹ Assistant-Ausnahme:</strong> User mit JobTitle <em>&quot;Assistant&quot;</em> oder <em>&quot;Senior Assistant&quot;</em> dürfen &quot;Für andere registrieren&quot; nutzen, aber nur für Personen mit JobTitle <strong>Partner</strong> oder <strong>Director</strong> — und nur solange die Anmeldefrist noch nicht abgelaufen ist.
          </div>
          <div>
            <strong>² Eigene Events:</strong> gilt nur wenn der User in <code>OrganizerEmail</code> des jeweiligen Events steht. Tenant-weiter Organizer-Status reicht nicht — Event A-Organizer können keine Admin-Aktionen für Event B ausführen. Admin darf global alles.
          </div>
        </div>
      </div>

      <div style={{ marginTop: 16, textAlign: 'center' }}>
        <button className="btn btn-secondary" onClick={() => navigate('settings')}>
          Zurück zu Settings
        </button>
      </div>
    </div>
  );
}

const thStyle: React.CSSProperties = {
  padding: '14px 20px', textAlign: 'left', fontWeight: 600, fontSize: '0.85rem', color: 'var(--dex-gray-700)',
};

function roleBadgeStyle(color: string): React.CSSProperties {
  return {
    display: 'inline-block', padding: '3px 10px', borderRadius: 12,
    background: `${color}18`, color: color, fontWeight: 600, fontSize: '0.78rem',
  };
}
