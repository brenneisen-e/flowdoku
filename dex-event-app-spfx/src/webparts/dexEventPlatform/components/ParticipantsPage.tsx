/**
 * Teilnehmer-Übersicht (DEX_Participants)
 *
 * Zeigt alle registrierten Teilnehmer mit ihren Event-Zuordnungen.
 * EventNumbers sind klickbar und navigieren zur Teilnehmerliste des Events.
 * Link zur SharePoint-Liste für direkten Zugriff.
 */

import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useRoles } from '../context/RoleContext';
import { useNavigation } from '../context/NavigationContext';
import { EventService, SPParticipant } from '../services/EventService';
import { DeloitteEvent } from '../types';
import { Users } from './Icons';
import { useLanguage } from '../context/LanguageContext';

export default function ParticipantsPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { events } = useEvents();
  const { isAdmin } = useRoles();
  const { t } = useLanguage();
  const [participants, setParticipants] = React.useState<SPParticipant[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [searchTerm, setSearchTerm] = React.useState('');

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const context = (window as any).__dexSpfxContext;
  const eventService = React.useMemo(() => context ? new EventService(context) : null, []);

  // EventNumber -> Event Map für schnelle Lookups
  const eventByNumber = React.useMemo(() => {
    const map: Record<number, DeloitteEvent> = {};
    for (const e of events) {
      if (e.eventNumber) map[e.eventNumber] = e;
    }
    return map;
  }, [events]);

  React.useEffect(() => {
    if (eventService) {
      eventService.getAllParticipants().then(data => {
        setParticipants(data);
        setIsLoading(false);
      }).catch(() => setIsLoading(false));
    }
  }, []);

  if (!isAdmin) {
    return (
      <div className="page-container text-center">
        <p style={{ color: 'var(--dex-gray-400)', padding: 48 }}>Admin only.</p>
      </div>
    );
  }

  // v27.13: Direkt-Link zur SharePoint-Liste entfernt — alle Aktionen laufen
  // über die App (direktes SP-Editieren umgeht Audit + Validierung).

  // Filtern
  const filtered = searchTerm
    ? participants.filter(p =>
        `${p.Vorname} ${p.Nachname} ${p.Email}`.toLowerCase().includes(searchTerm.toLowerCase())
      )
    : participants;

  // EventNumbers als klickbare Badges rendern
  const renderEventNumbers = (numbersStr: string, color: string): React.ReactElement | null => {
    if (!numbersStr) return null;
    const numbers = numbersStr.split(',').map(s => s.trim()).filter(s => s);
    if (numbers.length === 0) return null;

    return (
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
        {numbers.map(n => {
          const num = parseInt(n, 10);
          const event = eventByNumber[num];
          return (
            <span
              key={n}
              onClick={(e) => {
                e.stopPropagation();
                if (event) navigate('admin', event.id);
              }}
              title={event ? event.title : `Event #${n}`}
              style={{
                display: 'inline-block',
                padding: '2px 8px',
                borderRadius: 12,
                fontSize: '0.75rem',
                fontWeight: 600,
                background: color,
                color: '#fff',
                cursor: event ? 'pointer' : 'default',
              }}
            >
              {n}
            </span>
          );
        })}
      </div>
    );
  };

  return (
    <div className="page-container">
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
        <h2 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: 8 }}>
          <Users size={24} /> {t('participants.title')}
        </h2>
        {/* v27.13: „In SharePoint öffnen"-Button entfernt — alle Aktionen in der App. */}
      </div>

      <div style={{ marginBottom: 16 }}>
        <input
          className="form-input"
          placeholder={t('participants.search')}
          value={searchTerm}
          onChange={e => setSearchTerm(e.target.value)}
          style={{ maxWidth: 400 }}
        />
      </div>

      {isLoading ? (
        <p style={{ color: 'var(--dex-gray-400)', padding: 48, textAlign: 'center' }}>{t('participants.loading')}</p>
      ) : filtered.length === 0 ? (
        <div className="card text-center" style={{ padding: 48 }}>
          <p style={{ color: 'var(--dex-gray-400)' }}>
            {searchTerm ? t('participants.noresults') : t('participants.empty')}
          </p>
        </div>
      ) : (
        <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', minWidth: 640, borderCollapse: 'collapse', fontSize: '0.85rem' }}>
              <thead>
                <tr style={{ borderBottom: '2px solid var(--dex-gray-200)', background: 'var(--dex-gray-50)' }}>
                  <th style={{ textAlign: 'left', padding: '10px 12px' }}>{t('profile.firstname')}</th>
                  <th style={{ textAlign: 'left', padding: '10px 12px' }}>{t('profile.lastname')}</th>
                  <th style={{ textAlign: 'left', padding: '10px 12px' }}>{t('profile.email')}</th>
                  <th style={{ textAlign: 'left', padding: '10px 12px' }}>{t('participants.registered')}</th>
                  <th style={{ textAlign: 'left', padding: '10px 12px' }}>{t('participants.waitlist')}</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map(p => (
                  <tr key={p.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                    <td style={{ padding: '8px 12px', fontWeight: 500 }}>{p.Vorname || '-'}</td>
                    <td style={{ padding: '8px 12px', fontWeight: 500 }}>{p.Nachname || '-'}</td>
                    <td style={{ padding: '8px 12px', color: 'var(--dex-gray-600)' }}>{p.Email || p.Title}</td>
                    <td style={{ padding: '8px 12px' }}>{renderEventNumbers(p.EventRegistered, 'var(--dex-green)')}</td>
                    <td style={{ padding: '8px 12px' }}>{renderEventNumbers(p.EventOnWaitlist, 'var(--dex-orange)')}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ padding: '8px 12px', fontSize: '0.8rem', color: 'var(--dex-gray-400)', borderTop: '1px solid var(--dex-gray-100)' }}>
            {filtered.length} Teilnehmer{filtered.length !== participants.length ? ` (von ${participants.length} gefiltert)` : ''}
          </div>
        </div>
      )}
    </div>
  );
}
