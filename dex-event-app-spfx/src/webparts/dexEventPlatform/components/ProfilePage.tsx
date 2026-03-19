/**
 * Profil-Seite
 *
 * Zeigt alle verfuegbaren Informationen zum eingeloggten Benutzer:
 * Name, E-Mail, Standort, Telefon, Abteilung, Titel etc.
 *
 * - Eike, Maerz 2026
 */

import * as React from 'react';
import { useCurrentUser } from '../context/UserContext';

export default function ProfilePage(): React.ReactElement {
  const { currentUser, isLoading } = useCurrentUser();

  if (isLoading) {
    return (
      <div className="page-container text-center">
        <p style={{ color: 'var(--dex-gray-500)' }}>Loading profile...</p>
      </div>
    );
  }

  // Profilzeilen - nur anzeigen wenn Wert vorhanden
  const profileRows: Array<{ label: string; value: string }> = [
    { label: 'First Name', value: currentUser.firstName },
    { label: 'Surname', value: currentUser.surname },
    { label: 'E-Mail', value: currentUser.email },
    { label: 'Location', value: currentUser.location },
    { label: 'Login', value: currentUser.id },
  ].filter(row => row.value);

  return (
    <div className="page-container">
      <div className="card" style={{ maxWidth: 600, margin: '0 auto' }}>
        {/* Profil-Header mit Avatar */}
        <div style={{
          display: 'flex', alignItems: 'center', gap: 20,
          padding: '24px 24px 20px', borderBottom: '1px solid var(--dex-gray-200, #eee)',
        }}>
          <div style={{
            width: 72, height: 72, borderRadius: '50%',
            background: 'linear-gradient(135deg, #86bc25, #0076a8)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            color: '#fff', fontWeight: 700, fontSize: '1.6rem', flexShrink: 0,
          }}>
            {currentUser.firstName ? currentUser.firstName[0] : ''}{currentUser.surname ? currentUser.surname[0] : ''}
          </div>
          <div>
            <h2 style={{ margin: 0, fontSize: '1.3rem' }}>{currentUser.firstName} {currentUser.surname}</h2>
            <p style={{ margin: '4px 0 0', color: 'var(--dex-gray-500, #888)', fontSize: '0.9rem' }}>
              {currentUser.email}
            </p>
            <span style={{
              display: 'inline-block', marginTop: 8, padding: '2px 12px', borderRadius: 12,
              background: currentUser.isAdmin ? '#e8f5e9' : '#f5f5f5',
              color: currentUser.isAdmin ? '#2e7d32' : '#666',
              fontSize: '0.8rem', fontWeight: 500,
            }}>
              {currentUser.isAdmin ? 'Admin' : 'User'}
            </span>
          </div>
        </div>

        {/* Profil-Details */}
        <div style={{ padding: '20px 24px' }}>
          <h3 style={{ fontSize: '1rem', marginBottom: 16, color: 'var(--dex-gray-700, #444)' }}>Profile Details</h3>
          {profileRows.map(row => (
            <div key={row.label} style={{
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              padding: '10px 0', borderBottom: '1px solid var(--dex-gray-100, #f0f0f0)',
            }}>
              <span style={{ color: 'var(--dex-gray-500, #888)', fontSize: '0.9rem' }}>{row.label}</span>
              <span style={{ fontWeight: 500, fontSize: '0.9rem', textAlign: 'right' }}>{row.value}</span>
            </div>
          ))}
        </div>

        {/* Hinweis */}
        <div style={{
          padding: '16px 24px', background: 'var(--dex-gray-50, #fafafa)',
          borderRadius: '0 0 var(--dex-radius-lg, 12px) var(--dex-radius-lg, 12px)',
          fontSize: '0.8rem', color: 'var(--dex-gray-400, #aaa)',
        }}>
          This data is read from your SharePoint profile. To update your information, please contact your IT department or update your profile in Microsoft 365.
        </div>
      </div>
    </div>
  );
}
