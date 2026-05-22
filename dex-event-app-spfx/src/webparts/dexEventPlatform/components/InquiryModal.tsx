/**
 * Geteiltes Inquiry-Modal — "DEX App für dein Event anfragen".
 * Wird aus LandingPage und StartPage (Organizer-Tile-Overlay) verwendet.
 * Sendet eine Nachricht an die DEX-Maintainer; nimmt Name (vorbefüllt mit
 * Display-Name des eingeloggten Users), Event-Name und Freitext-Beschreibung
 * entgegen.
 */
import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useLanguage } from '../context/LanguageContext';

interface InquiryModalProps {
  open: boolean;
  onClose: () => void;
}

export default function InquiryModal({ open, onClose }: InquiryModalProps): React.ReactElement | null {
  const { sendAdminInquiry } = useEvents();
  const { currentUser } = useCurrentUser();
  const { locale } = useLanguage();
  const userFullName = `${currentUser.firstName} ${currentUser.surname}`.trim();
  const [name, setName] = React.useState(userFullName);
  const [eventName, setEventName] = React.useState('');
  const [message, setMessage] = React.useState('');
  const [sending, setSending] = React.useState(false);
  const [status, setStatus] = React.useState<'' | 'success' | 'error'>('');

  React.useEffect(() => {
    if (open && !name && userFullName) setName(userFullName);
  }, [open, userFullName]);

  async function handleSubmit(): Promise<void> {
    if (!eventName.trim() || !message.trim() || sending) return;
    setSending(true);
    setStatus('');
    const ok = await sendAdminInquiry(
      name.trim() || userFullName,
      currentUser.email || '',
      eventName.trim(),
      message.trim(),
    );
    setSending(false);
    if (ok) {
      setStatus('success');
      setEventName('');
      setMessage('');
      setTimeout(() => { onClose(); setStatus(''); }, 1800);
    } else {
      setStatus('error');
    }
  }

  if (!open) return null;
  return (
    <div
      role="dialog"
      aria-modal="true"
      onClick={() => { if (!sending) onClose(); }}
      style={{
        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        zIndex: 9999, padding: 16,
      }}
    >
      <div
        onClick={e => e.stopPropagation()}
        style={{
          background: '#fff', borderRadius: 16, padding: '24px 28px',
          maxWidth: 480, width: '100%', boxShadow: '0 12px 48px rgba(0,0,0,0.18)',
          display: 'flex', flexDirection: 'column', gap: 14,
        }}
      >
        <h2 style={{ margin: 0, fontSize: '1.3rem', color: 'var(--dex-gray-800)' }}>
          {locale === 'de' ? 'DEX App für dein Event anfragen' : 'Request the DEX App for your event'}
        </h2>
        <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
          {locale === 'de'
            ? 'Wir melden uns kurz bei dir und besprechen, wie wir dein Event unterstützen können.'
            : 'We will get back to you and discuss how we can support your event.'}
        </p>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {locale === 'de' ? 'Dein Name' : 'Your name'}
          <input
            type="text"
            value={name}
            onChange={e => setName(e.target.value)}
            disabled={sending}
            style={{ padding: '8px 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.9rem' }}
          />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {locale === 'de' ? 'Event-Name' : 'Event name'}
          <input
            type="text"
            value={eventName}
            onChange={e => setEventName(e.target.value)}
            disabled={sending}
            placeholder={locale === 'de' ? 'z.B. Summer Party 2026' : 'e.g. Summer Party 2026'}
            style={{ padding: '8px 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.9rem' }}
          />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {locale === 'de' ? 'Was brauchst du?' : 'What do you need?'}
          <textarea
            value={message}
            onChange={e => setMessage(e.target.value)}
            disabled={sending}
            rows={5}
            placeholder={locale === 'de'
              ? 'Kurz beschreiben: Anzahl Teilnehmer, Termin, gewünschte Funktionen...'
              : 'Briefly describe: number of participants, date, features needed...'}
            style={{ padding: '8px 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.9rem', fontFamily: 'inherit', resize: 'vertical' }}
          />
        </label>
        {status === 'success' && (
          <div style={{ color: 'var(--dex-green)', fontSize: '0.85rem' }}>
            {locale === 'de' ? 'Anfrage gesendet — wir melden uns!' : 'Request sent — we will get back to you!'}
          </div>
        )}
        {status === 'error' && (
          <div style={{ color: 'var(--dex-red)', fontSize: '0.85rem' }}>
            {locale === 'de' ? 'Senden fehlgeschlagen. Bitte später erneut versuchen.' : 'Sending failed. Please try again later.'}
          </div>
        )}
        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 4 }}>
          <button type="button" className="btn btn-secondary" onClick={onClose} disabled={sending}>
            {locale === 'de' ? 'Abbrechen' : 'Cancel'}
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={handleSubmit}
            disabled={sending || !eventName.trim() || !message.trim()}
          >
            {sending
              ? (locale === 'de' ? 'Wird gesendet...' : 'Sending...')
              : (locale === 'de' ? 'Anfrage senden' : 'Send inquiry')}
          </button>
        </div>
      </div>
    </div>
  );
}
