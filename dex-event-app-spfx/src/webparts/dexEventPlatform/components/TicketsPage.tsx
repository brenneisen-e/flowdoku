/**
 * TicketsPage (v26.0.0) — Antwort-Inbox für Power-User & Admins.
 *
 * Zeigt die Fragen, die an die Power-User gerichtet sind (Audience „PowerUser":
 * Fragen von Organizern/Admins), gruppiert nach Offen / In Bearbeitung /
 * Beantwortet. Über den Deep-Link aus der Benachrichtigungs-Mail
 * (?action=tickets&id=<id>) wird das passende Ticket direkt aufgeklappt.
 *
 * User-Fragen zu einem konkreten Event werden NICHT hier, sondern in der
 * Event-Übersicht (Box „Offene Fragen (User)") beantwortet.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { useTickets } from '../context/TicketContext';
import { useLanguage } from '../context/LanguageContext';
import { useRoles } from '../context/RoleContext';
import { DexTicket } from '../types';
import TicketCard from './tickets/TicketCard';

export default function TicketsPage(): React.ReactElement {
  const { powerUserQueue, isTicketsLoading, reloadTickets, canAnswerTickets } = useTickets();
  const { locale } = useLanguage();
  const { isAdmin } = useRoles();
  const isDe = locale === 'de';

  const [closedOpen, setClosedOpen] = React.useState(false);
  const deepLinkId = React.useMemo<number | null>(() => {
    try { const v = new URLSearchParams(window.location.search).get('id'); return v ? Number(v) : null; } catch { return null; }
  }, []);

  React.useEffect(() => { reloadTickets().catch(() => { /* */ }); }, [reloadTickets]);

  if (!canAnswerTickets && !isAdmin) {
    return (
      <div className="page-container">
        <div style={{ maxWidth: 640, margin: '40px auto', textAlign: 'center', color: 'var(--dex-gray-500,#808080)' }}>
          <Icon iconName="Chat" style={{ fontSize: 40, color: 'var(--dex-gray-300,#d1d1d1)' }} />
          <p style={{ marginTop: 12 }}>{isDe ? 'Diese Seite ist für Power-User und Admins. Stell deine Frage über den grünen Button „Hast du Fragen?" oben rechts.' : 'This page is for power users and admins. Ask your question via the green „Have a question?" button at the top right.'}</p>
        </div>
      </div>
    );
  }

  const open = powerUserQueue.filter((t) => t.status === 'Open');
  const inProgress = powerUserQueue.filter((t) => t.status === 'InProgress');
  const closed = powerUserQueue.filter((t) => t.status === 'Closed');

  const section = (title: string, list: DexTicket[], tint: string): React.ReactElement | null => {
    if (list.length === 0) return null;
    return (
      <div style={{ marginBottom: 22 }}>
        <h3 style={{ fontSize: '1.02rem', color: 'var(--dex-green-dark,#4a7c1f)', margin: '0 0 10px', display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ width: 8, height: 8, borderRadius: '50%', background: tint, display: 'inline-block' }} />
          {title} <span style={{ color: 'var(--dex-gray-400,#a0a0a0)', fontWeight: 500 }}>({list.length})</span>
        </h3>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {list.map((t) => <TicketCard key={t.id} ticket={t} defaultExpanded={deepLinkId === t.id} />)}
        </div>
      </div>
    );
  };

  return (
    <div className="page-container">
      <div style={{ maxWidth: 900, margin: '0 auto' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 6 }}>
          <Icon iconName="Chat" style={{ fontSize: 24, color: 'var(--dex-green,#86bc25)' }} />
          <h1 style={{ margin: 0, fontSize: '1.5rem' }}>{isDe ? 'Tickets' : 'Tickets'}</h1>
        </div>
        <p style={{ color: 'var(--dex-gray-600,#666)', marginTop: 0, fontSize: '0.9rem' }}>
          {isDe ? 'Fragen aus der App an das Power-User-Team. Übernimm ein Ticket, um es zu beantworten — andere sehen dann „in Bearbeitung".' : 'Questions from the app to the power-user team. Take a ticket to answer it — others then see „in progress".'}
        </p>

        {isTicketsLoading && powerUserQueue.length === 0 && (
          <p style={{ color: 'var(--dex-gray-400,#a0a0a0)' }}>{isDe ? 'Lädt …' : 'Loading …'}</p>
        )}
        {!isTicketsLoading && powerUserQueue.length === 0 && (
          <div style={{ textAlign: 'center', color: 'var(--dex-gray-500,#808080)', padding: '40px 0' }}>
            <Icon iconName="SkypeCircleCheck" style={{ fontSize: 36, color: 'var(--dex-green,#86bc25)' }} />
            <p style={{ marginTop: 10 }}>{isDe ? 'Keine offenen Fragen — alles erledigt.' : 'No open questions — all done.'}</p>
          </div>
        )}

        {section(isDe ? 'Offen' : 'Open', open, '#ed8b00')}
        {section(isDe ? 'In Bearbeitung' : 'In progress', inProgress, '#1565c0')}

        {closed.length > 0 && (
          <div>
            <button type="button" onClick={() => setClosedOpen((v) => !v)}
              style={{ display: 'inline-flex', alignItems: 'center', gap: 7, background: 'transparent', border: 'none', cursor: 'pointer', fontFamily: 'inherit', color: 'var(--dex-gray-600,#666)', fontWeight: 600, fontSize: '0.95rem', padding: '4px 0' }}>
              <Icon iconName={closedOpen ? 'ChevronDown' : 'ChevronRight'} style={{ fontSize: 12 }} />
              {isDe ? 'Beantwortet' : 'Answered'} ({closed.length})
            </button>
            {closedOpen && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 10 }}>
                {closed.map((t) => <TicketCard key={t.id} ticket={t} defaultExpanded={deepLinkId === t.id} />)}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
