/**
 * TicketEventBox (v26.0.0) — einklappbare Box „Offene Fragen (User)" in der
 * Event-Übersicht (AdminPage), zwischen den Event-Infos und den KPI-Kacheln.
 *
 * Zeigt die Fragen normaler User zu DIESEM Event (Audience „Organizer"), die der
 * Organizer hier im Kontext beantwortet. Der Kopf trägt einen Badge mit der
 * Anzahl noch offener Fragen. Über den Deep-Link aus der Benachrichtigungs-Mail
 * (?action=admin&event=<id>&ticket=<id>) klappt das passende Ticket direkt auf.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { useTickets } from '../../context/TicketContext';
import { useLanguage } from '../../context/LanguageContext';
import TicketCard from './TicketCard';

export default function TicketEventBox(props: { eventId: string }): React.ReactElement | null {
  const { ticketsForEvent, openCountForEvent } = useTickets();
  const { locale } = useLanguage();
  const isDe = locale === 'de';

  const list = ticketsForEvent(props.eventId);
  const openCount = openCountForEvent(props.eventId);

  const deepLinkTicketId = React.useMemo<number | null>(() => {
    try { const v = new URLSearchParams(window.location.search).get('ticket'); return v ? Number(v) : null; } catch { return null; }
  }, []);

  const [collapsed, setCollapsed] = React.useState<boolean>(openCount === 0 && deepLinkTicketId == null);

  // Wenn ein Deep-Link auf ein Ticket dieses Events zeigt, aufklappen.
  React.useEffect(() => {
    if (deepLinkTicketId != null && list.some((t) => t.id === deepLinkTicketId)) setCollapsed(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [deepLinkTicketId, list.length]);

  if (list.length === 0) return null;

  const sorted = [...list].sort((a, b) => {
    const rank = (s: string): number => (s === 'Open' ? 0 : s === 'InProgress' ? 1 : 2);
    const r = rank(a.status) - rank(b.status);
    if (r !== 0) return r;
    return (b.created || '').localeCompare(a.created || '');
  });

  return (
    <div style={{ background: '#fff', border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 12, marginBottom: 18, boxShadow: '0 1px 4px rgba(0,0,0,0.04)' }}>
      <button
        type="button"
        onClick={() => setCollapsed((v) => !v)}
        style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 10, background: 'transparent', border: 'none', cursor: 'pointer', padding: '14px 18px', fontFamily: 'inherit', textAlign: 'left' }}
      >
        <Icon iconName="Chat" style={{ fontSize: 18, color: 'var(--dex-green,#86bc25)' }} />
        <span style={{ fontSize: '1.02rem', fontWeight: 700, color: 'var(--dex-gray-800,#333)' }}>
          {isDe ? 'Offene Fragen (User)' : 'Open questions (users)'}
        </span>
        {openCount > 0 && (
          <span style={{ background: '#ed8b00', color: '#fff', borderRadius: 11, minWidth: 22, height: 22, padding: '0 7px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700 }}>
            {openCount}
          </span>
        )}
        <span style={{ marginLeft: 'auto', color: 'var(--dex-gray-400,#a0a0a0)' }}>
          <Icon iconName={collapsed ? 'ChevronDown' : 'ChevronUp'} style={{ fontSize: 13 }} />
        </span>
      </button>
      {!collapsed && (
        <div style={{ padding: '0 18px 16px', display: 'flex', flexDirection: 'column', gap: 10 }}>
          {sorted.map((t) => <TicketCard key={t.id} ticket={t} defaultExpanded={deepLinkTicketId === t.id} />)}
        </div>
      )}
    </div>
  );
}
