/**
 * ticketThread (v26.8.0)
 *
 * Gemeinsame Helfer für die Ticket-Darstellung:
 *  - `contactSubline` baut die „Position · Standort"-Unterzeile für die
 *    Foto-Kontaktkarte (PersonContactHover).
 *  - `renderTicketThread` rendert den Rückfragen-Verlauf (followUps) chronologisch
 *    mit Foto-Kontaktkarte je Sprecher — genutzt sowohl auf der Antwort-Seite
 *    (TicketCard) als auch in „Deine Fragen" (QuestionButton).
 */
import * as React from 'react';
import { DexTicket } from '../../types';
import PersonContactHover from '../PersonContactHover';

/** „Position · Standort" (leere Teile werden ausgelassen). */
export function contactSubline(jobTitle?: string, location?: string): string | undefined {
  const parts = [jobTitle, location].map((x) => (x || '').trim()).filter(Boolean);
  return parts.length ? parts.join(' · ') : undefined;
}

function fmt(at: string, isDe: boolean): string {
  try { return at ? new Date(at).toLocaleString(isDe ? 'de-DE' : 'en-GB') : ''; } catch { return ''; }
}

/** Rückfragen-Verlauf (followUps) chronologisch nach Zeit rendern. */
export function renderTicketThread(ticket: DexTicket, isDe: boolean): React.ReactElement | null {
  const ups = [...(ticket.followUps || [])].sort((a, b) => (a.at || '').localeCompare(b.at || ''));
  if (ups.length === 0) return null;
  return (
    <div style={{ marginTop: 10, display: 'flex', flexDirection: 'column', gap: 8 }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: 'var(--dex-gray-600,#666)' }}>
        {isDe ? 'Verlauf' : 'Conversation'}
      </div>
      {ups.map((u, i) => {
        const isAsker = u.byRole === 'asker';
        return (
          <div key={i} style={{ display: 'flex', gap: 8, alignItems: 'flex-start' }}>
            <PersonContactHover email={u.byEmail} name={u.byName} size={28} isDe={isDe} />
            <div style={{
              flex: 1, background: isAsker ? '#fff' : '#f1f7e8',
              border: `1px solid ${isAsker ? 'var(--dex-gray-200,#e8e8e8)' : 'var(--dex-green,#86bc25)'}`,
              borderRadius: 8, padding: '7px 10px',
            }}>
              <div style={{ display: 'flex', alignItems: 'baseline', gap: 8, flexWrap: 'wrap', marginBottom: 2 }}>
                <strong style={{ fontSize: '0.82rem' }}>{u.byName || u.byEmail}</strong>
                <span style={{ fontSize: 11, color: 'var(--dex-gray-400,#a0a0a0)' }}>
                  {isAsker ? (isDe ? 'Rückfrage' : 'Follow-up') : (isDe ? 'Antwort' : 'Reply')}
                  {u.at ? ` · ${fmt(u.at, isDe)}` : ''}
                </span>
              </div>
              <div style={{ fontSize: '0.86rem', whiteSpace: 'pre-wrap' }}>{u.text}</div>
            </div>
          </div>
        );
      })}
    </div>
  );
}
