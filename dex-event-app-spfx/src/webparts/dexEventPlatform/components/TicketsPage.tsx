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
import { deepLinkParams } from '../utils/deepLink';
import { useDialog } from '../context/DialogContext';
import TicketCard from './tickets/TicketCard';

export default function TicketsPage(): React.ReactElement {
  const { powerUserQueue, isTicketsLoading, reloadTickets, canAnswerTickets, remindPowerUsers, reminderTargets } = useTickets();
  const { locale } = useLanguage();
  const { isAdmin } = useRoles();
  const { confirmDialog, showAlert } = useDialog();
  const isDe = locale === 'de';

  const [closedOpen, setClosedOpen] = React.useState(false);
  // v29.14: Erinnerung an die Power-User. Nach dem Senden bleibt der Knopf
  // gesperrt — nicht als Schutz vor Missbrauch, sondern weil ein zweiter Klick
  // aus Unsicherheit („ist sie raus?") eine zweite Mail erzeugen würde. Die
  // Mail liegt danach in der Warteschlange; der Flow holt sie ab.
  const [reminderBusy, setReminderBusy] = React.useState(false);
  const [reminderSent, setReminderSent] = React.useState(false);
  const deepLinkId = React.useMemo<number | null>(() => {
    try { const v = deepLinkParams().get('id'); return v ? Number(v) : null; } catch { return null; }
  }, []);

  React.useEffect(() => { reloadTickets().catch(() => { /* */ }); }, [reloadTickets]);

  // v26.7: Zeigt der Deep-Link auf ein bereits beantwortetes Ticket, die
  // (sonst eingeklappte) „Beantwortet"-Sektion automatisch aufklappen.
  React.useEffect(() => {
    if (deepLinkId == null) return;
    if (powerUserQueue.some(t => t.id === deepLinkId && t.status === 'Closed')) setClosedOpen(true);
  }, [deepLinkId, powerUserQueue]);

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
  // Unbeantwortet = alles, was noch keine Antwort hat — auch das bereits
  // übernommene „in Bearbeitung". Genau das erinnert die Mail.
  const unanswered = powerUserQueue.filter((t) => t.status !== 'Closed' && !t.answeredAt);

  const handleRemind = async (): Promise<void> => {
    const targets = reminderTargets();
    if (targets.length === 0) {
      showAlert(isDe
        ? 'Es gibt niemanden, den die Erinnerung erreichen würde — außer dir ist niemand als Power-User (oder Admin) eingetragen. Power-User setzt du in den Einstellungen unter Rollen mit dem Stern.'
        : 'There is nobody the reminder could reach — apart from you, nobody is registered as a power user (or admin). You set power users in Settings under roles via the star.');
      return;
    }
    const list = targets.length > 4
      ? `${targets.slice(0, 4).join(', ')} ${isDe ? `und ${targets.length - 4} weitere` : `and ${targets.length - 4} more`}`
      : targets.join(', ');
    const ok = await confirmDialog(
      isDe
        ? `Erinnerung an ${targets.length} ${targets.length === 1 ? 'Person' : 'Personen'} schicken?\n\nEine Mail mit allen ${unanswered.length} unbeantworteten Tickets geht an: ${list}.\n\nDu selbst bekommst sie nicht.`
        : `Send a reminder to ${targets.length} ${targets.length === 1 ? 'person' : 'people'}?\n\nOne mail listing all ${unanswered.length} unanswered tickets goes to: ${list}.\n\nYou will not receive it yourself.`,
      { confirmLabel: isDe ? 'Erinnerung schicken' : 'Send reminder' },
    );
    if (!ok) return;
    setReminderBusy(true);
    try {
      const res = await remindPowerUsers();
      if (res.ok) {
        setReminderSent(true);
        showAlert(isDe
          ? `Erinnerung an ${res.recipients.length} ${res.recipients.length === 1 ? 'Person' : 'Personen'} eingestellt — sie wird innerhalb weniger Minuten verschickt.`
          : `Reminder queued for ${res.recipients.length} ${res.recipients.length === 1 ? 'person' : 'people'} — it will go out within a few minutes.`, { variant: 'success' });
      } else if (res.reason === 'no-tickets') {
        showAlert(isDe ? 'Es gibt gerade kein unbeantwortetes Ticket.' : 'There is no unanswered ticket right now.');
      } else {
        showAlert(isDe
          ? 'Die Erinnerung konnte nicht eingestellt werden. Bitte später noch einmal versuchen.'
          : 'The reminder could not be queued. Please try again later.', { variant: 'error' });
      }
    } finally {
      setReminderBusy(false);
    }
  };

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

        {/* v29.14: Erinnerung auf Knopfdruck. Die Automatik schickt sie erst
            nach zwei Werktagen und nur, wenn ein Beantworter die App öffnet —
            wer sieht, dass ein Ticket liegen bleibt, konnte bisher nur privat
            nachfassen. Der Knopf steht nur da, wenn es auch etwas zu erinnern
            gibt. */}
        {unanswered.length > 0 && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap', marginBottom: 18 }}>
            <button
              type="button"
              className="btn btn-secondary"
              disabled={reminderBusy || reminderSent}
              onClick={() => { handleRemind().catch(() => { /* */ }); }}
              style={{
                display: 'inline-flex', alignItems: 'center', gap: 7,
                fontSize: '0.85rem', padding: '7px 16px',
                opacity: (reminderBusy || reminderSent) ? 0.55 : 1,
                cursor: (reminderBusy || reminderSent) ? 'default' : 'pointer',
              }}
              title={isDe
                ? 'Schickt den Power-Usern eine Mail mit allen unbeantworteten Tickets. Du selbst bekommst sie nicht.'
                : 'Sends the power users one mail listing all unanswered tickets. You will not receive it yourself.'}
            >
              <Icon iconName={reminderSent ? 'CheckMark' : 'Ringer'} style={{ fontSize: 14 }} />
              {reminderBusy
                ? (isDe ? 'Wird eingestellt …' : 'Queueing …')
                : reminderSent
                  ? (isDe ? 'Erinnerung verschickt' : 'Reminder sent')
                  : (isDe ? 'Erinnerung an die Power-User' : 'Remind the power users')}
            </button>
            <span style={{ color: 'var(--dex-gray-500,#808080)', fontSize: '0.82rem' }}>
              {isDe
                ? `${unanswered.length} ${unanswered.length === 1 ? 'Ticket wartet' : 'Tickets warten'} auf eine Antwort.`
                : `${unanswered.length} ${unanswered.length === 1 ? 'ticket is' : 'tickets are'} waiting for an answer.`}
            </span>
          </div>
        )}

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
