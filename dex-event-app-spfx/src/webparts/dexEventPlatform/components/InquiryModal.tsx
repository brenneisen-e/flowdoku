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
  /** v23.37: „Organizer werden"-Antrag statt allgemeiner DEX-Anfrage.
   *  Legt einen nachverfolgbaren Antrag an (Admins bestätigen ihn in der App). */
  organizerMode?: boolean;
}

export default function InquiryModal({ open, onClose, organizerMode }: InquiryModalProps): React.ReactElement | null {
  const { sendAdminInquiry, requestOrganizerRole } = useEvents();
  const { currentUser, photoUrl } = useCurrentUser();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const userFullName = `${currentUser.firstName} ${currentUser.surname}`.trim();
  // v24.24: Standort + Position des eingeloggten Users (read-only Anzeige +
  // gehen mit in die Anfrage-Mail an die Admins).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const userLocation = ((currentUser as any).location || '').trim();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const userJobTitle = ((currentUser as any).jobTitle || '').trim();
  const userInitials = `${(currentUser.firstName || '')[0] || ''}${(currentUser.surname || '')[0] || ''}`.toUpperCase() || '?';
  const [eventName, setEventName] = React.useState('');
  const [message, setMessage] = React.useState('');
  const [sending, setSending] = React.useState(false);
  const [status, setStatus] = React.useState<'' | 'success' | 'error'>('');
  const [showInfo, setShowInfo] = React.useState(false);
  // v28.41: Art des Events VOR allem anderen abfragen. Bei „extern" ist DEX
  // das falsche Werkzeug — dann statt eines Formulars der Verweis auf die
  // Event-Management-Seite im DeloitteNet, und der Absende-Knopf bleibt gesperrt.
  const [eventScope, setEventScope] = React.useState<'' | 'internal' | 'external'>('');

  // v23.37: im Organizer-Modus reicht der Name (Nachricht optional, kein
  // Event-Name) — die allgemeine Anfrage braucht Event-Name + Nachricht.
  // v28.41: Ohne Angabe der Event-Art laesst sich nichts absenden, und bei
  // „extern" bleibt der Knopf gesperrt — DEX ist dafuer schlicht das falsche
  // Werkzeug, eine Anfrage waere fuer beide Seiten verlorene Zeit.
  const canSubmit = organizerMode
    ? true
    : (!!eventName.trim() && !!message.trim() && eventScope === 'internal');

  async function handleSubmit(): Promise<void> {
    if (!canSubmit || sending) return;
    setSending(true);
    setStatus('');
    let ok = false;
    if (organizerMode) {
      const res = await requestOrganizerRole(currentUser.email || '', userFullName, userLocation, message.trim());
      ok = res.ok;
    } else {
      ok = await sendAdminInquiry(userFullName, currentUser.email || '', eventName.trim(), message.trim(), userLocation, userJobTitle);
    }
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
      ariaLabel={organizerMode ? (isDe ? 'Organizer werden' : 'Become an organizer') : (isDe ? 'DEX-Anfrage' : 'DEX inquiry')}
    >
        <h2 style={{ margin: 0, fontSize: '1.3rem', color: 'var(--dex-gray-800)' }}>
          {organizerMode
            ? (isDe ? 'Organizer werden' : 'Become an organizer')
            : (isDe ? 'DEX App für dein Event anfragen' : 'Request the DEX App for your event')}
        </h2>
        <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
          {organizerMode
            ? (isDe
                ? 'Stelle einen Antrag, Organizer zu werden — dann kannst du eigene Events anlegen und verwalten. Die Admins prüfen deinen Antrag und schalten dich frei.'
                : 'Request to become an organizer — you can then create and manage your own events. The admins review your request and grant access.')
            : (isDe
                ? 'Wir melden uns kurz bei dir und besprechen, wie wir dein Event unterstützen können.'
                : 'We will get back to you and discuss how we can support your event.')}
        </p>
        {!organizerMode && <button
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
        </button>}
        {/* v24.24: „Dein Name" ist nicht mehr frei editierbar — stattdessen eine
            read-only Personen-Karte des eingeloggten Users (Foto, Name, Position,
            Standort), analog zur Teilnehmerliste. Diese Infos gehen mit in die
            Anfrage-Mail an die Admins. */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {locale === 'de' ? 'Du fragst an als' : 'Requesting as'}
          <div style={{
            display: 'flex', alignItems: 'center', gap: 12,
            padding: '10px 12px', borderRadius: 10,
            border: '1px solid var(--dex-gray-200)', background: 'var(--dex-gray-50, #fafafa)',
          }}>
            {photoUrl
              ? <img src={photoUrl} alt="" style={{ width: 44, height: 44, borderRadius: '50%', objectFit: 'cover', flexShrink: 0 }} />
              : <span style={{
                  width: 44, height: 44, borderRadius: '50%', flexShrink: 0,
                  background: 'var(--dex-green, #86bc25)', color: '#fff',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  fontWeight: 700, fontSize: '0.95rem',
                }}>{userInitials}</span>}
            <span style={{ minWidth: 0 }}>
              <span style={{ display: 'block', fontWeight: 600, color: 'var(--dex-gray-800)', fontSize: '0.92rem' }}>
                {userFullName || (currentUser.email || '')}
              </span>
              {(userJobTitle || userLocation) && (
                <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                  {[userJobTitle, userLocation].filter(Boolean).join(' · ')}
                </span>
              )}
              {currentUser.email && (
                <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-400)' }}>
                  {currentUser.email}
                </span>
              )}
            </span>
          </div>
        </div>
        {/* v28.40: Einsatzbereich klarstellen. Bis hierhin stand in der
            Anfrage-Strecke nirgends, fuer welche Art von Events DEX gedacht
            ist — die einzige Erwaehnung von „extern" war die technische
            Aussage „keine externen APIs" in der Info-Box, die man sogar
            falsch herum lesen kann. */}
        {!organizerMode && (
          <div style={{
            padding: '12px 14px', borderRadius: 8,
            background: 'var(--dex-gray-50, #f7f7f5)',
            border: '1px solid var(--dex-gray-200)',
            fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5,
          }}>
            <div style={{ fontWeight: 700, marginBottom: 8 }}>
              {isDe ? 'Um was für ein Event geht es?' : 'What kind of event is it?'}
            </div>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, cursor: 'pointer' }}>
              <input type="radio" name="dexEventScope" checked={eventScope === 'internal'}
                onChange={() => setEventScope('internal')} disabled={sending} style={{ marginTop: 3 }} />
              <span>
                <strong>{isDe ? 'Internes Event' : 'Internal event'}</strong>
                <span style={{ display: 'block', fontSize: '0.76rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Ein Deloitte-internes Event — oder die Koordination der Deloitte-Teilnahme an einer externen Veranstaltung (z.B. B2Run).'
                    : 'A Deloitte-internal event — or coordinating Deloitte participation in an external event (e.g. B2Run).'}
                </span>
              </span>
            </label>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer' }}>
              <input type="radio" name="dexEventScope" checked={eventScope === 'external'}
                onChange={() => setEventScope('external')} disabled={sending} style={{ marginTop: 3 }} />
              <span>
                <strong>{isDe ? 'Externes Event mit externen Teilnehmern' : 'External event with external attendees'}</strong>
                <span style={{ display: 'block', fontSize: '0.76rem', color: 'var(--dex-gray-600)' }}>
                  {isDe ? 'Gäste außerhalb von Deloitte melden sich selbst an.' : 'Guests from outside Deloitte register themselves.'}
                </span>
              </span>
            </label>
            {eventScope === 'external' && (
              <div style={{
                marginTop: 10, padding: '10px 12px', borderRadius: 8,
                background: '#fef3f2', border: '1px solid var(--dex-red, #c00)',
                color: '#7a1f1c', fontSize: '0.8rem', lineHeight: 1.55,
              }}>
                {isDe
                  ? <><strong>Dafür ist DEX nicht nutzbar.</strong> Die Plattform ist auf Deloitte-interne Events ausgelegt; externe Gäste bekommen keinen Zugang und können sich nicht selbst anmelden.<br /><br />Alles zu externen Veranstaltungen findest du hier:<br /><a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer" style={{ color: '#7a1f1c', fontWeight: 700 }}>Event Management im DeloitteNet</a></>
                  : <><strong>DEX cannot be used for this.</strong> The platform is built for Deloitte-internal events; external guests get no access and cannot register themselves.<br /><br />Everything about external events can be found here:<br /><a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer" style={{ color: '#7a1f1c', fontWeight: 700 }}>Event Management on DeloitteNet</a></>}
              </div>
            )}
          </div>
        )}
        {!organizerMode && (
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
        )}
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {organizerMode
            ? (isDe ? 'Warum möchtest du Organizer werden? (optional)' : 'Why do you want to become an organizer? (optional)')
            : (isDe ? 'Was brauchst du?' : 'What do you need?')}
          <textarea
            className="form-textarea"
            value={message}
            onChange={e => setMessage(e.target.value)}
            disabled={sending}
            rows={organizerMode ? 3 : 5}
            placeholder={organizerMode
              ? (isDe ? 'Optional: kurz, worum es geht …' : 'Optional: briefly what it is about …')
              : (isDe
                ? 'Kurz beschreiben: Anzahl Teilnehmer, Termin, gewünschte Funktionen...'
                : 'Briefly describe: number of participants, date, features needed...')}
          />
        </label>
        {status === 'success' && (
          <div style={{ color: 'var(--dex-green)', fontSize: '0.85rem' }}>
            {organizerMode
              ? (isDe ? 'Antrag gesendet — die Admins prüfen ihn und schalten dich frei.' : 'Request sent — the admins will review and grant access.')
              : (isDe ? 'Anfrage gesendet — wir melden uns!' : 'Request sent — we will get back to you!')}
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
            disabled={sending || !canSubmit}
          >
            {sending
              ? (isDe ? 'Wird gesendet...' : 'Sending...')
              : organizerMode
                ? (isDe ? 'Antrag senden' : 'Send request')
                : (isDe ? 'Anfrage senden' : 'Send inquiry')}
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
