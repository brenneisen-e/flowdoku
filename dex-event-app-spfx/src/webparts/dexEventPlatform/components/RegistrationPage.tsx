/**
 * Registrierungsseite fuer ein einzelnes Event
 *
 * Drei-Spalten-Layout: Event-Info | persoenliche Daten | eventspezifische Felder
 * Speichert die Registrierung in der SharePoint-Teilnehmerliste des Events.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { Salutation } from '../types';
import { Info, Trash2, Send } from './Icons';

function formatDate(iso: string): string {
  const d = new Date(iso);
  return (
    d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
    ' ' +
    d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })
  );
}

export default function RegistrationPage(): React.ReactElement {
  const { selectedEventId, navigate } = useNavigation();
  const { events, registerForEvent } = useEvents();
  const { currentUser } = useCurrentUser();
  const { canCreateEvents } = useRoles();
  const event = events.find(e => e.id === selectedEventId);

  // Sichtbarkeits-Check: Würde dieses Event dem User als normaler User angezeigt werden?
  const showLocationBanner = canCreateEvents && event && (() => {
    const locFilters = event.locationAudience;
    const audFilters = event.audienceFilter || [];
    const hasLoc = locFilters.length > 0;
    const hasAud = audFilters.length > 0;
    if (!hasLoc && !hasAud) return false; // kein Filter = alle sehen es

    const loc = (currentUser.location || '').toLowerCase();
    const email = currentUser.email.toLowerCase();

    const locMatch = !hasLoc || locFilters.some(f => {
      const fl = f.trim().toLowerCase();
      if (fl === 'all') return true;
      const norm = fl.replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ä/g, 'ae');
      return loc.indexOf(fl) >= 0 || loc.indexOf(norm) >= 0;
    });

    const audMatch = !hasAud || audFilters.some(f => {
      const fl = f.trim().toLowerCase();
      if (fl.indexOf('@') >= 0) return email === fl;
      if (fl === 'deall') return true;
      if (fl.startsWith('de')) return loc.indexOf(fl.substring(2)) >= 0;
      return false;
    });

    const mode = event.filterMode || 'OR';
    const visible = mode === 'AND'
      ? (hasLoc && hasAud ? locMatch && audMatch : hasLoc ? locMatch : audMatch)
      : (locMatch || audMatch);
    return !visible;
  })();

  const [salutation, setSalutation] = React.useState<Salutation | ''>('');
  const [firstName, setFirstName] = React.useState(currentUser.firstName);
  const [surname, setSurname] = React.useState(currentUser.surname);
  const [email, setEmail] = React.useState(currentUser.email);
  const [registerForOther, setRegisterForOther] = React.useState(false);
  const [eventSpecific, setEventSpecific] = React.useState<Record<string, string>>({});
  const [submitted, setSubmitted] = React.useState(false);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [error, setError] = React.useState('');
  const [showErrors, setShowErrors] = React.useState(false);
  const [showDescription, setShowDescription] = React.useState(false);

  if (!event) {
    return (
      <div className="page-container text-center">
        <h2>Event not found</h2>
        <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>
          Back to Events
        </button>
      </div>
    );
  }

  const isFull = event.currentParticipants >= event.maxParticipants;
  const errorBorder = { border: '2px solid var(--dex-red)' };

  const handleSubmit = async (): Promise<void> => {
    // Validierung Pflichtfelder
    setShowErrors(true);
    if (!salutation || !firstName.trim() || !surname.trim() || !email.trim()) {
      setError('Bitte alle Pflichtfelder ausfüllen.');
      return;
    }

    // Pflicht-Custom-Fields validieren
    const missingRequired = event.eventSpecificFields
      .filter(f => f.required && !eventSpecific[f.id]?.trim());
    if (missingRequired.length > 0) {
      setError(`Bitte ausfüllen: ${missingRequired.map(f => f.label).join(', ')}`);
      return;
    }

    setError('');
    setIsSubmitting(true);

    try {
      const customData: Record<string, string> = {
        salutation,
        ...eventSpecific,
      };

      const participantEmail = email.trim();

      const success = await registerForEvent(
        selectedEventId!,
        customData,
        firstName.trim(),
        surname.trim(),
        participantEmail
      );

      if (success) {
        setSubmitted(true);
      } else {
        setError('Registrierung fehlgeschlagen. Möglicherweise bist du bereits angemeldet.');
      }
    } catch {
      setError('Ein Fehler ist aufgetreten. Bitte versuche es erneut.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleClear = (): void => {
    setSalutation('');
    setFirstName('');
    setSurname('');
    setEmail('');
    setEventSpecific({});
    setRegisterForOther(false);
  };

  if (submitted) {
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '64px 32px' }}>
          <h2>{isFull ? 'Auf die Warteliste gesetzt' : 'Registrierung erfolgreich!'}</h2>
          <p className="mt-8" style={{ color: 'var(--dex-gray-600)' }}>
            {isFull
              ? `Du wurdest auf die Warteliste für "${event.title}" gesetzt. Du wirst benachrichtigt, sobald ein Platz frei wird.`
              : `Du wurdest erfolgreich für "${event.title}" registriert. Eine Bestätigung wurde an ${email} gesendet.`}
          </p>
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center' }}>
            <button className="btn btn-primary" onClick={() => navigate('my-events')}>
              My Events
            </button>
            <button className="btn btn-secondary" onClick={() => navigate('register')}>
              Register for another event
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="page-container">
      {showLocationBanner && (
        <div style={{
          padding: '10px 16px', marginBottom: 16, borderRadius: 'var(--dex-radius-md)',
          background: 'rgba(237,139,0,0.1)', border: '1px solid var(--dex-orange)',
          color: 'var(--dex-orange)', fontSize: '0.85rem',
        }}>
          Hinweis: Du würdest dieses Event als normaler User nicht sehen.
          {event && event.locationAudience.length > 0 && <> Standort-Filter: <strong>{event.locationAudience.join(', ')}</strong>.</>}
          {event && event.audienceFilter && event.audienceFilter.length > 0 && <> Zielgruppe: <strong>{event.audienceFilter.join(', ')}</strong>.</>}
          {event && event.filterMode === 'AND' && <> (UND-Verknüpfung)</>}
          {' '}Dein Standort: {currentUser.location || 'unbekannt'}.
        </div>
      )}
      <div className="registration-layout">
        {/* Event-Info links */}
        <div className="registration-event">
          <div className="section-header section-header--red">Selected Event</div>
          <div className="registration-event__card">
            <div
              className="registration-event__image"
              style={{
                background: event.imageUrl
                  ? `url(${event.imageUrl}) center/cover no-repeat`
                  : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
              }}
            >
              <button className="event-card__info-btn" aria-label="Event info" onClick={() => setShowDescription(!showDescription)}>
                <Info size={16} />
              </button>
              <div className="registration-event__overlay">
                <h4>{event.title}</h4>
                <p>
                  {formatDate(event.startDate)} until<br />
                  {formatDate(event.endDate)}
                </p>
              </div>
            </div>
            {showDescription && event.description && (
              <div style={{
                padding: '12px 16px', fontSize: '0.85rem', color: 'var(--dex-gray-700)',
                background: 'var(--dex-gray-50)', borderRadius: '0 0 var(--dex-radius) var(--dex-radius)',
                borderTop: '1px solid var(--dex-gray-200)',
              }}>
                {event.description}
              </div>
            )}
            {isFull && (
              <p className="text-red text-center mt-8" style={{ padding: '0 12px 12px', fontWeight: 600, fontSize: '0.85rem' }}>
                All places are taken. There are currently {event.waitlistCount} people on the waiting list
              </p>
            )}
          </div>
        </div>

        {/* Persoenliche Daten */}
        <div className="registration-form">
          <div className="section-header">Personal Information</div>
          <div style={{ padding: '24px 20px' }}>
            <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 12 }}>
              <span className="required">*</span> = Pflichtfeld
            </p>

            {canCreateEvents && (
              <button
                className="btn btn-outline"
                style={{ marginBottom: 20, fontSize: '0.85rem' }}
                onClick={() => {
                  setRegisterForOther(!registerForOther);
                  if (!registerForOther) { setFirstName(''); setSurname(''); setEmail(''); }
                  else { setFirstName(currentUser.firstName); setSurname(currentUser.surname); setEmail(currentUser.email); }
                }}
              >
                {registerForOther ? 'Für mich selbst registrieren' : 'Für jemand anderen registrieren'}
              </button>
            )}

            <div className="form-group">
              <label className="form-label"><span className="required">*</span> Salutation</label>
              <select className="form-select" value={salutation} onChange={e => setSalutation(e.target.value as Salutation)} style={showErrors && !salutation ? errorBorder : {}}>
                <option value="">Please select</option>
                <option value="Herr">Herr</option>
                <option value="Frau">Frau</option>
                <option value="Divers">Divers</option>
              </select>
            </div>

            <div className="form-group">
              <label className="form-label"><span className="required">*</span> First Name</label>
              <input className="form-input" value={firstName} onChange={e => setFirstName(e.target.value)} placeholder="First Name" disabled={!registerForOther} style={{ ...(!registerForOther ? { background: 'var(--dex-gray-100)' } : {}), ...(showErrors && !firstName.trim() ? errorBorder : {}) }} />
            </div>

            <div className="form-group">
              <label className="form-label"><span className="required">*</span> Surname</label>
              <input className="form-input" value={surname} onChange={e => setSurname(e.target.value)} placeholder="Surname" disabled={!registerForOther} style={{ ...(!registerForOther ? { background: 'var(--dex-gray-100)' } : {}), ...(showErrors && !surname.trim() ? errorBorder : {}) }} />
            </div>

            <div className="form-group">
              <label className="form-label"><span className="required">*</span> E-Mail</label>
              <input className="form-input" type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="email@deloitte.de" disabled={!registerForOther} style={{ ...(!registerForOther ? { background: 'var(--dex-gray-100)' } : {}), ...(showErrors && !email.trim() ? errorBorder : {}) }} />
            </div>
          </div>
        </div>

        {/* Eventspezifische Felder */}
        <div className="registration-specific">
          <div className="section-header">Event specific Information</div>
          <div style={{ padding: '24px 20px' }}>
            {event.eventSpecificFields.length === 0 ? (
              <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>No additional information required.</p>
            ) : (
              event.eventSpecificFields.map(field => (
                <div className="form-group" key={field.id}>
                  <label className="form-label">
                    {field.required && <span className="required">*</span>}
                    {field.label}
                    {field.helpText && <span className="info-icon" title={field.helpText} style={{ marginLeft: 8 }}>i</span>}
                  </label>
                  {field.type === 'select' ? (
                    <select className="form-select" value={eventSpecific[field.id] || ''} onChange={e => setEventSpecific({ ...eventSpecific, [field.id]: e.target.value })} style={showErrors && field.required && !eventSpecific[field.id]?.trim() ? errorBorder : {}}>
                      <option value="">Please select</option>
                      {field.options && field.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                    </select>
                  ) : (
                    <input className="form-input" value={eventSpecific[field.id] || ''} onChange={e => setEventSpecific({ ...eventSpecific, [field.id]: e.target.value })} placeholder={field.label} style={showErrors && field.required && !eventSpecific[field.id]?.trim() ? errorBorder : {}} />
                  )}
                </div>
              ))
            )}
          </div>
        </div>
      </div>

      {/* Datenschutz-Hinweis */}
      <div className="footer-disclaimer mt-24" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
        <p>
          As part of the event {event.title}, audio, image, and video recordings will be made for
          Deloitte. By the clear confirming action, e.g., posing or smiling at the camera, you
          consent to the creation of audio, image, and video recordings. The recordings may be used
          for Deloitte internal publication purposes (e.g., DeloitteNet). The consent can be revoked
          at any time with future effect and without any negative consequences by sending an email to
          privacy@deloitte.de.
        </p>
      </div>

      {/* Fehlermeldung */}
      {error && (
        <div className="mt-16" style={{ padding: '12px 16px', background: 'rgba(239,68,68,0.1)', border: '1px solid var(--dex-red)', borderRadius: 'var(--dex-radius-md)', color: 'var(--dex-red)', fontSize: '0.9rem' }}>
          {error}
        </div>
      )}

      {/* Buttons */}
      <div className="registration-actions mt-24">
        <button className="btn btn-danger" onClick={handleClear} disabled={isSubmitting}><Trash2 size={16} /> Delete</button>
        <button className="btn btn-primary" onClick={handleSubmit} disabled={isSubmitting}>
          <Send size={16} /> {isSubmitting ? 'Wird gesendet...' : 'Register'}
        </button>
      </div>
    </div>
  );
}
