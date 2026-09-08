/**
 * TicketEventBox (v26.0.0) — einklappbare Box „Offene Fragen (User)" in der
 * Event-Übersicht (AdminPage), zwischen den Event-Infos und den KPI-Kacheln.
 *
 * Zeigt die Fragen normaler User zu DIESEM Event (Audience „Organizer"), die der
 * Organizer hier im Kontext beantwortet. Der Kopf trägt einen Badge mit der
 * Anzahl noch offener Fragen. Über den Deep-Link aus der Benachrichtigungs-Mail
 * (?action=admin&event=<id>&ticket=<id>) klappt das passende Ticket direkt auf.
 *
 * v31.3: Nach docs/ui-leitfaden.md (5a Punkt 7) umgebaut — Kopf als
 * `dex-ui-card-head` mit Status-Pille und Aufklapper-Chevron. Der Titel nennt
 * jetzt den Inhalt („Fragen zu diesem Event"), eine Zeile darunter die Folge:
 * vorher hieß der Kopf „Offene Fragen (User)", obwohl die Box auch die längst
 * beantworteten zeigt.
 */
import * as React from 'react';
import { useTickets } from '../../context/TicketContext';
import { useLanguage } from '../../context/LanguageContext';
import { deepLinkParams } from '../../utils/deepLink';
import { cx, ensureDexUiStyles } from '../dexUi';
import { ChevronDown, MessageSquare } from '../Icons';
import TicketCard from './TicketCard';

export default function TicketEventBox(props: { eventId: string }): React.ReactElement | null {
  const { ticketsForEvent, openCountForEvent } = useTickets();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  // Idempotent — die Box steht ohne Modal/WizardFormShell in der Event-Seite,
  // ohne diesen Aufruf greifen die dex-ui-Klassen hier nicht.
  ensureDexUiStyles();

  const list = ticketsForEvent(props.eventId);
  const openCount = openCountForEvent(props.eventId);

  const deepLinkTicketId = React.useMemo<number | null>(() => {
    try { const v = deepLinkParams().get('ticket'); return v ? Number(v) : null; } catch { return null; }
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
    <div className="dex-ui-card" style={{ padding: 0, marginBottom: 18 }}>
      {/* v31.3: Die ganze Kopfzeile ist der Aufklapper (Hover über dex-ui-row) —
          vorher war der Knopf zwar auch die ganze Zeile, sah aber wie eine
          Überschrift aus, weil ein Inline-Style kein :hover kann. */}
      <button
        type="button"
        onClick={() => setCollapsed((v) => !v)}
        aria-expanded={!collapsed}
        className="dex-ui-rowbtn dex-ui-row"
        // Der Hover-Grund der Zeile soll in den Ecken der Karte bleiben:
        // zugeklappt rundum, aufgeklappt nur oben.
        style={{ padding: '14px 18px', borderRadius: collapsed ? 13 : '13px 13px 0 0' }}
      >
        <span className="dex-ui-card-head" style={{ width: '100%' }}>
          <span className="dex-ui-card-head-title">
            <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', display: 'inline-flex' }}><MessageSquare size={18} /></span>
            {isDe ? 'Fragen zu diesem Event' : 'Questions about this event'}
          </span>
          <span className={cx('dex-ui-pill', openCount > 0 ? 'dex-ui-pill--orange' : 'dex-ui-pill--green')}>
            {openCount > 0
              ? (isDe ? `${openCount} offen` : `${openCount} open`)
              : (isDe ? 'alle beantwortet' : 'all answered')}
          </span>
          <span className="dex-ui-card-head-meta">
            {isDe
              ? (list.length === 1 ? '1 Frage insgesamt' : `${list.length} Fragen insgesamt`)
              : (list.length === 1 ? '1 question in total' : `${list.length} questions in total`)}
          </span>
          <span className="dex-ui-card-head-actions">
            <span className={cx('dex-ui-disclosure-chevron', !collapsed && 'is-open')}>
              <ChevronDown size={16} />
            </span>
          </span>
        </span>
      </button>
      {!collapsed && (
        <div className="dex-ui-stack" style={{ padding: '0 18px 16px' }}>
          {/* Eine Zeile Folge: wer fragt, und wohin die Antwort geht. */}
          <div className="dex-ui-muted">
            {isDe
              ? 'Nutzer haben dich als Organizer gefragt — deine Antwort geht ihnen per Mail zu.'
              : 'Users asked you as the organizer — your answer is emailed to them.'}
          </div>
          {sorted.map((t) => <TicketCard key={t.id} ticket={t} defaultExpanded={deepLinkTicketId === t.id} />)}
        </div>
      )}
    </div>
  );
}
