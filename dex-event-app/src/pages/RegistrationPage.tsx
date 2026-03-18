import { useState } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { Info, Trash2, Send } from 'lucide-react';
import { events, currentUser } from '../data/mockData';
import type { Salutation } from '../types';

function formatDate(iso: string) {
  const d = new Date(iso);
  return (
    d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
    ' ' +
    d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })
  );
}

export default function RegistrationPage() {
  const { eventId } = useParams();
  const navigate = useNavigate();
  const event = events.find((e) => e.id === eventId);

  const [salutation, setSalutation] = useState<Salutation | ''>('');
  const [firstName, setFirstName] = useState(currentUser.firstName);
  const [surname, setSurname] = useState(currentUser.surname);
  const [email, setEmail] = useState(currentUser.email);
  const [registerForOther, setRegisterForOther] = useState(false);
  const [eventSpecific, setEventSpecific] = useState<Record<string, string>>({});
  const [submitted, setSubmitted] = useState(false);

  if (!event) {
    return (
      <div className="page-container text-center">
        <h2>Event not found</h2>
        <button className="btn btn-primary mt-24" onClick={() => navigate('/register')}>
          Back to Events
        </button>
      </div>
    );
  }

  const isFull = event.currentParticipants >= event.maxParticipants;

  const handleSubmit = () => {
    setSubmitted(true);
  };

  const handleClear = () => {
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
          <div style={{ fontSize: '3rem', marginBottom: 16 }}>
            {isFull ? '⏳' : '✅'}
          </div>
          <h2>{isFull ? 'Added to Waitlist' : 'Registration Successful!'}</h2>
          <p className="mt-8" style={{ color: 'var(--dex-gray-600)' }}>
            {isFull
              ? `You have been added to the waiting list for "${event.title}". You will be notified if a spot becomes available.`
              : `You have been successfully registered for "${event.title}". A confirmation email has been sent to ${email}.`}
          </p>
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center' }}>
            <button className="btn btn-primary" onClick={() => navigate('/my-events')}>
              My Events
            </button>
            <button className="btn btn-secondary" onClick={() => navigate('/register')}>
              Register for another event
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="page-container">
      <div className="registration-layout">
        {/* Left: Selected Event */}
        <div className="registration-event">
          <div className="section-header section-header--red">Selected Event</div>
          <div className="registration-event__card">
            <div
              className="registration-event__image"
              style={{
                background: 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
              }}
            >
              <button className="event-card__info-btn" aria-label="Event info">
                <Info size={16} />
              </button>
              <div className="registration-event__overlay">
                <h4>{event.title}</h4>
                <p>
                  {formatDate(event.startDate)} until
                  <br />
                  {formatDate(event.endDate)}
                </p>
              </div>
            </div>
            {isFull && (
              <p className="text-red text-center mt-8" style={{ padding: '0 12px 12px', fontWeight: 600, fontSize: '0.85rem' }}>
                All places are taken. There are currently {event.waitlistCount} people on the waiting
                list
              </p>
            )}
          </div>
        </div>

        {/* Center: Personal Information */}
        <div className="registration-form">
          <div className="section-header">Personal Information</div>
          <div style={{ padding: '24px 20px' }}>
            <button
              className="btn btn-outline"
              style={{ marginBottom: 20, fontSize: '0.85rem' }}
              onClick={() => {
                setRegisterForOther(!registerForOther);
                if (!registerForOther) {
                  setFirstName('');
                  setSurname('');
                  setEmail('');
                } else {
                  setFirstName(currentUser.firstName);
                  setSurname(currentUser.surname);
                  setEmail(currentUser.email);
                }
              }}
            >
              {registerForOther ? 'Register for myself' : 'Click here to register for someone else'}
            </button>

            <div className="form-group">
              <label className="form-label">
                <span className="required">*</span> Salutation
              </label>
              <select
                className="form-select"
                value={salutation}
                onChange={(e) => setSalutation(e.target.value as Salutation)}
              >
                <option value="">Please select</option>
                <option value="Herr">Herr</option>
                <option value="Frau">Frau</option>
                <option value="Divers">Divers</option>
              </select>
            </div>

            <div className="form-group">
              <label className="form-label">
                <span className="required">*</span> First Name
              </label>
              <input
                className="form-input"
                value={firstName}
                onChange={(e) => setFirstName(e.target.value)}
                placeholder="First Name"
              />
            </div>

            <div className="form-group">
              <label className="form-label">
                <span className="required">*</span> Surname
              </label>
              <input
                className="form-input"
                value={surname}
                onChange={(e) => setSurname(e.target.value)}
                placeholder="Surname"
              />
            </div>

            <div className="form-group">
              <label className="form-label">
                <span className="required">*</span> E-Mail
              </label>
              <input
                className="form-input"
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="email@deloitte.de"
              />
            </div>
          </div>
        </div>

        {/* Right: Event specific */}
        <div className="registration-specific">
          <div className="section-header">Event specific Information</div>
          <div style={{ padding: '24px 20px' }}>
            {event.eventSpecificFields.length === 0 ? (
              <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
                No additional information required.
              </p>
            ) : (
              event.eventSpecificFields.map((field) => (
                <div className="form-group" key={field.id}>
                  <label className="form-label">
                    {field.required && <span className="required">*</span>}
                    {field.label}
                    {field.helpText && (
                      <span className="info-icon" title={field.helpText} style={{ marginLeft: 8 }}>
                        i
                      </span>
                    )}
                  </label>
                  {field.type === 'select' ? (
                    <select
                      className="form-select"
                      value={eventSpecific[field.id] || ''}
                      onChange={(e) =>
                        setEventSpecific({ ...eventSpecific, [field.id]: e.target.value })
                      }
                    >
                      <option value="">Please select</option>
                      {field.options?.map((opt) => (
                        <option key={opt} value={opt}>
                          {opt}
                        </option>
                      ))}
                    </select>
                  ) : (
                    <input
                      className="form-input"
                      value={eventSpecific[field.id] || ''}
                      onChange={(e) =>
                        setEventSpecific({ ...eventSpecific, [field.id]: e.target.value })
                      }
                      placeholder={field.label}
                    />
                  )}
                </div>
              ))
            )}
          </div>
        </div>
      </div>

      {/* Disclaimer */}
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

      {/* Action buttons */}
      <div className="registration-actions mt-24">
        <button className="btn btn-danger" onClick={handleClear}>
          <Trash2 size={16} /> Delete
        </button>
        <button className="btn btn-primary" onClick={handleSubmit}>
          <Send size={16} /> Register
        </button>
      </div>
    </div>
  );
}
