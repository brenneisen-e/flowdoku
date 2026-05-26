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
import Modal from './Modal';
import LandingInfoModal from './LandingInfoModal';
import { Info } from './Icons';

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
  const [showInfo, setShowInfo] = React.useState(false);

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

  // v13.1: Modal-Wrapper-Komponente — kapselt Backdrop/Escape/Padding.
  return (
    <Modal
      open={open}
      onClose={onClose}
      dismissable={!sending}
      ariaLabel={locale === 'de' ? 'DEX-Anfrage' : 'DEX inquiry'}
    >
        <h2 style={{ margin: 0, fontSize: '1.3rem', color: 'var(--dex-gray-800)' }}>
          {locale === 'de' ? 'DEX App für dein Event anfragen' : 'Request the DEX App for your event'}
        </h2>
        <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
          {locale === 'de'
            ? 'Wir melden uns kurz bei dir und besprechen, wie wir dein Event unterstützen können.'
            : 'We will get back to you and discuss how we can support your event.'}
        </p>
        <button
          type="button"
          onClick={() => setShowInfo(true)}
          style={{
            display: 'flex', alignItems: 'center', gap: 10,
            padding: '12px 14px',
            background: 'linear-gradient(135deg, rgba(134,188,37,0.10), rgba(0,118,168,0.06))',
            border: '1px solid var(--dex-green, #86bc25)',
            borderRadius: 10,
            cursor: 'pointer',
            textAlign: 'left',
            fontFamily: 'inherit',
            transition: 'all 0.15s',
          }}
          onMouseEnter={e => { e.currentTarget.style.background = 'linear-gradient(135deg, rgba(134,188,37,0.18), rgba(0,118,168,0.10))'; }}
          onMouseLeave={e => { e.currentTarget.style.background = 'linear-gradient(135deg, rgba(134,188,37,0.10), rgba(0,118,168,0.06))'; }}
        >
          <span style={{
            width: 32, height: 32, borderRadius: '50%', flexShrink: 0,
            background: 'var(--dex-green, #86bc25)', color: '#fff',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
            <Info size={16} />
          </span>
          <span style={{ fontSize: '0.85rem', lineHeight: 1.35, color: 'var(--dex-gray-700)' }}>
            <strong style={{ display: 'block', color: 'var(--dex-gray-800)', marginBottom: 2 }}>
              {locale === 'de' ? 'Erst mal mehr über die DEX App erfahren?' : 'First learn more about the DEX App?'}
            </strong>
            {locale === 'de'
              ? 'Hier klicken — alle Funktionen, Zielgruppen und Beispiel-Events im Überblick.'
              : 'Click here — all features, audiences and example events at a glance.'}
          </span>
        </button>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {locale === 'de' ? 'Dein Name' : 'Your name'}
          {/* v13.0: form-input statt inline-style — konsistente Höhe / Border. */}
          <input
            type="text"
            className="form-input"
            value={name}
            onChange={e => setName(e.target.value)}
            disabled={sending}
          />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {locale === 'de' ? 'Event-Name' : 'Event name'}
          <input
            type="text"
            className="form-input"
            value={eventName}
            onChange={e => setEventName(e.target.value)}
            disabled={sending}
            placeholder={locale === 'de' ? 'z.B. Summer Party 2026' : 'e.g. Summer Party 2026'}
          />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {locale === 'de' ? 'Was brauchst du?' : 'What do you need?'}
          <textarea
            className="form-textarea"
            value={message}
            onChange={e => setMessage(e.target.value)}
            disabled={sending}
            rows={5}
            placeholder={locale === 'de'
              ? 'Kurz beschreiben: Anzahl Teilnehmer, Termin, gewünschte Funktionen...'
              : 'Briefly describe: number of participants, date, features needed...'}
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
      <LandingInfoModal
        open={showInfo}
        locale={locale === 'de' ? 'de' : 'en'}
        onClose={() => setShowInfo(false)}
      />
    </Modal>
  );
}
