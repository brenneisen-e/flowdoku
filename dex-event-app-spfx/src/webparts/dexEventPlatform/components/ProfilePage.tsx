/**
 * Profil-Seite
 *
 * Zeigt ALLE verfuegbaren Informationen aus dem SharePoint User Profile:
 * Name, Titel, Abteilung, Bereich, Standort, Telefon, Manager etc.
 *
 * - Eike, Maerz 2026
 */

import * as React from 'react';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

// Properties die wir aus dem SP User Profile anzeigen wollen (in dieser Reihenfolge)
const PROFILE_FIELDS: Array<{ key: string; label: string }> = [
  { key: 'FirstName', label: 'Vorname' },
  { key: 'LastName', label: 'Nachname' },
  { key: 'Title', label: 'Titel / Position' },
  { key: 'Department', label: 'Abteilung' },
  { key: 'Office', label: 'Standort' },
  { key: 'WorkEmail', label: 'E-Mail' },
  { key: 'WorkPhone', label: 'Telefon' },
  { key: 'CellPhone', label: 'Mobiltelefon' },
  { key: 'Manager', label: 'Manager' },
];

interface ProfileData {
  displayName: string;
  pictureUrl: string;
  properties: Array<{ key: string; label: string; value: string }>;
}

// Context wird ueber eine separate Provider-Komponente bereitgestellt
// Hier nutzen wir einen kleinen Trick: wir holen den Context ueber das Window-Objekt
// (wurde im DexEventPlatform gesetzt)

export interface ProfilePageProps {
  context?: WebPartContext;
}

export default function ProfilePage(props: ProfilePageProps): React.ReactElement {
  const { currentUser, photoUrl: userPhotoUrl } = useCurrentUser();
  const { currentUserRole } = useRoles();
  const [profile, setProfile] = React.useState<ProfileData | null>(null);
  const [isProfileLoading, setIsProfileLoading] = React.useState(true);

  // Rollen-Farbe
  const roleColors: Record<string, { bg: string; color: string }> = {
    'Admin': { bg: '#e8f5e9', color: '#2e7d32' },
    'Organizer': { bg: '#e3f2fd', color: '#1565c0' },
    'User': { bg: '#f5f5f5', color: '#666' },
  };
  const rc = roleColors[currentUserRole] || roleColors['User'];

  React.useEffect(() => {
    loadFullProfile();
  }, []);

  async function loadFullProfile(): Promise<void> {
    try {
      // SPFx Context aus dem Window-Objekt holen (wird von DexEventPlatform gesetzt)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx: WebPartContext | undefined = props.context || (window as any).__dexSpfxContext;
      if (!ctx) {
        setIsProfileLoading(false);
        return;
      }

      const siteUrl = ctx.pageContext.web.absoluteUrl;
      const response = await ctx.spHttpClient.get(
        `${siteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        const allProps: Array<{ Key: string; Value: string }> = data.UserProfileProperties || [];

        // Properties filtern und in der gewuenschten Reihenfolge anzeigen
        const mapped: Array<{ key: string; label: string; value: string }> = [];

        for (const field of PROFILE_FIELDS) {
          const prop = allProps.find(p => p.Key === field.key);
          if (prop && prop.Value && prop.Value.trim()) {
            mapped.push({ key: field.key, label: field.label, value: prop.Value });
          }
        }

        // Profilfoto aus UserContext verwenden (bereits via Graph geladen)
        const pictureUrl = userPhotoUrl || '';

        setProfile({
          displayName: data.DisplayName || `${currentUser.firstName} ${currentUser.surname}`,
          pictureUrl,
          properties: mapped,
        });
      }
    } catch {
      // Profil konnte nicht geladen werden
    }
    setIsProfileLoading(false);
  }

  return (
    <div className="page-container">
      <div className="card" style={{ maxWidth: 700, margin: '0 auto' }}>
        {/* Profil-Header mit Avatar */}
        <div style={{
          display: 'flex', alignItems: 'center', gap: 20,
          padding: '24px 24px 20px', borderBottom: '1px solid var(--dex-gray-200, #eee)',
        }}>
          {profile && profile.pictureUrl ? (
            <img
              src={profile.pictureUrl}
              alt="Profile"
              style={{ width: 72, height: 72, borderRadius: '50%', objectFit: 'cover', flexShrink: 0 }}
              onError={(e) => { (e.target as HTMLImageElement).style.display = 'none'; }}
            />
          ) : (
            <div style={{
              width: 72, height: 72, borderRadius: '50%',
              background: 'linear-gradient(135deg, #86bc25, #0076a8)',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              color: '#fff', fontWeight: 700, fontSize: '1.6rem', flexShrink: 0,
            }}>
              {currentUser.firstName ? currentUser.firstName[0] : ''}{currentUser.surname ? currentUser.surname[0] : ''}
            </div>
          )}
          <div>
            <h2 style={{ margin: 0, fontSize: '1.3rem' }}>
              {profile ? profile.displayName : `${currentUser.firstName} ${currentUser.surname}`}
            </h2>
            <p style={{ margin: '4px 0 0', color: 'var(--dex-gray-500, #888)', fontSize: '0.9rem' }}>
              {currentUser.email}
            </p>
            <span style={{
              display: 'inline-block', marginTop: 8, padding: '2px 12px', borderRadius: 12,
              background: rc.bg, color: rc.color,
              fontSize: '0.8rem', fontWeight: 500,
            }}>
              {currentUserRole}
            </span>
          </div>
        </div>

        {/* Profil-Details */}
        <div style={{ padding: '20px 24px' }}>
          <h3 style={{ fontSize: '1rem', marginBottom: 16, color: 'var(--dex-gray-700, #444)' }}>Profil Details</h3>

          {isProfileLoading ? (
            <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>Lade Profildaten...</p>
          ) : profile && profile.properties.length > 0 ? (
            profile.properties.map(row => (
              <div key={row.key} style={{
                display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                padding: '10px 0', borderBottom: '1px solid var(--dex-gray-100, #f0f0f0)',
                gap: 16,
              }}>
                <span style={{ color: 'var(--dex-gray-500, #888)', fontSize: '0.85rem', flexShrink: 0 }}>{row.label}</span>
                <span style={{
                  fontWeight: 500, fontSize: '0.85rem', textAlign: 'right',
                  wordBreak: 'break-word', maxWidth: '60%',
                }}>
                  {row.value}
                </span>
              </div>
            ))
          ) : (
            // Fallback: Basisdaten anzeigen
            [
              { label: 'Vorname', value: currentUser.firstName },
              { label: 'Nachname', value: currentUser.surname },
              { label: 'E-Mail', value: currentUser.email },
              { label: 'Standort', value: currentUser.location },
              { label: 'Login', value: currentUser.id },
            ].filter(r => r.value).map(row => (
              <div key={row.label} style={{
                display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                padding: '10px 0', borderBottom: '1px solid var(--dex-gray-100, #f0f0f0)',
              }}>
                <span style={{ color: 'var(--dex-gray-500, #888)', fontSize: '0.85rem' }}>{row.label}</span>
                <span style={{ fontWeight: 500, fontSize: '0.85rem', textAlign: 'right' }}>{row.value}</span>
              </div>
            ))
          )}
        </div>

        {/* Hinweis */}
        <div style={{
          padding: '16px 24px', background: 'var(--dex-gray-50, #fafafa)',
          borderRadius: '0 0 var(--dex-radius-lg, 12px) var(--dex-radius-lg, 12px)',
          fontSize: '0.8rem', color: 'var(--dex-gray-400, #aaa)',
        }}>
          Diese Daten werden aus deinem SharePoint-Profil gelesen. Um sie zu aktualisieren,
          wende dich an deine IT oder aktualisiere dein Profil in Microsoft 365.
        </div>
      </div>
    </div>
  );
}
