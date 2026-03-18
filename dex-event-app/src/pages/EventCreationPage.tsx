import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Trash2, Send, Plus, X } from 'lucide-react';
import type { EventType } from '../types';

interface CustomField {
  id: string;
  label: string;
  type: 'text' | 'select';
  required: boolean;
  options: string;
}

export default function EventCreationPage() {
  const navigate = useNavigate();
  const [title, setTitle] = useState('');
  const [organizers, setOrganizers] = useState('');
  const [location, setLocation] = useState('');
  const [locationAudience, setLocationAudience] = useState<string[]>([]);
  const [mailAudience, setMailAudience] = useState('');
  const [description, setDescription] = useState('');
  const [eventType, setEventType] = useState<EventType>('Other');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [registrationDeadline, setRegistrationDeadline] = useState('');
  const [maxParticipants, setMaxParticipants] = useState('');
  const [customFields, setCustomFields] = useState<CustomField[]>([]);
  const [submitted, setSubmitted] = useState(false);

  const locationOptions = ['Berlin', 'Düsseldorf', 'Frankfurt', 'Hamburg', 'Hannover', 'Köln', 'München', 'Stuttgart', 'All'];

  const addCustomField = () => {
    setCustomFields([
      ...customFields,
      { id: `cf-${Date.now()}`, label: '', type: 'text', required: false, options: '' },
    ]);
  };

  const removeCustomField = (id: string) => {
    setCustomFields(customFields.filter((f) => f.id !== id));
  };

  const updateCustomField = (id: string, updates: Partial<CustomField>) => {
    setCustomFields(customFields.map((f) => (f.id === id ? { ...f, ...updates } : f)));
  };

  if (submitted) {
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '64px 32px' }}>
          <div style={{ fontSize: '3rem', marginBottom: 16 }}>✅</div>
          <h2>Event Created Successfully!</h2>
          <p className="mt-8" style={{ color: 'var(--dex-gray-600)' }}>
            "{title}" has been created. A confirmation email has been sent to the organizer(s).
          </p>
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center' }}>
            <button className="btn btn-primary" onClick={() => navigate('/register')}>
              View Events
            </button>
            <button className="btn btn-secondary" onClick={() => { setSubmitted(false); setTitle(''); }}>
              Create Another
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="page-container">
      <div className="card">
        <div className="creation-form">
          <div className="form-group">
            <label className="form-label"><span className="required">*</span> Event Title</label>
            <input className="form-input" value={title} onChange={(e) => setTitle(e.target.value)} />
          </div>

          <div className="form-group">
            <label className="form-label"><span className="required">*</span> Event Type</label>
            <select className="form-select" value={eventType} onChange={(e) => setEventType(e.target.value as EventType)}>
              <option value="Other">Other Deloitte Event</option>
              <option value="B2Run">B2Run</option>
              <option value="JPMorgan">JP Morgan Run</option>
            </select>
          </div>

          <div className="form-group">
            <label className="form-label"><span className="required">*</span> Organizer of Event | Multiple persons possible</label>
            <input className="form-input" value={organizers} onChange={(e) => setOrganizers(e.target.value)} placeholder="Please select" />
          </div>

          <div className="form-group">
            <label className="form-label">Location of Event</label>
            <input className="form-input" value={location} onChange={(e) => setLocation(e.target.value)} />
          </div>

          <div className="form-group">
            <label className="form-label">
              Location Audience
              <span className="info-icon" title="Select which office locations can see and register for this event" style={{ marginLeft: 8 }}>i</span>
            </label>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
              {locationOptions.map((loc) => (
                <label key={loc} style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: '0.9rem' }}>
                  <input
                    type="checkbox"
                    checked={locationAudience.includes(loc)}
                    onChange={(e) => {
                      if (e.target.checked) setLocationAudience([...locationAudience, loc]);
                      else setLocationAudience(locationAudience.filter((l) => l !== loc));
                    }}
                  />
                  {loc}
                </label>
              ))}
            </div>
          </div>

          <div className="form-group">
            <label className="form-label">
              Mail Audience
              <span className="info-icon" title="Enter email addresses to notify specific users about this event" style={{ marginLeft: 8 }}>i</span>
            </label>
            <input className="form-input" value={mailAudience} onChange={(e) => setMailAudience(e.target.value)} placeholder="Add Users" />
          </div>

          <div className="form-group">
            <label className="form-label"><span className="required">*</span> Event description</label>
            <textarea
              className="form-textarea"
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              style={{ minHeight: 150 }}
            />
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
            <div className="form-group">
              <label className="form-label"><span className="required">*</span> Start Date & Time</label>
              <input className="form-input" type="datetime-local" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
            </div>
            <div className="form-group">
              <label className="form-label"><span className="required">*</span> End Date & Time</label>
              <input className="form-input" type="datetime-local" value={endDate} onChange={(e) => setEndDate(e.target.value)} />
            </div>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
            <div className="form-group">
              <label className="form-label"><span className="required">*</span> Registration Deadline</label>
              <input className="form-input" type="datetime-local" value={registrationDeadline} onChange={(e) => setRegistrationDeadline(e.target.value)} />
            </div>
            <div className="form-group">
              <label className="form-label"><span className="required">*</span> Max Participants</label>
              <input className="form-input" type="number" min="1" value={maxParticipants} onChange={(e) => setMaxParticipants(e.target.value)} />
            </div>
          </div>

          {/* Custom fields section */}
          <div style={{ borderTop: '1px solid var(--dex-gray-200)', paddingTop: 24, marginTop: 8 }}>
            <div className="flex-between mb-16">
              <label className="form-label" style={{ marginBottom: 0 }}>Event-specific Registration Fields</label>
              <button className="btn btn-outline" onClick={addCustomField} style={{ fontSize: '0.85rem', padding: '6px 14px' }}>
                <Plus size={14} /> Add Field
              </button>
            </div>
            {customFields.map((field) => (
              <div key={field.id} className="custom-field-row">
                <input
                  className="form-input"
                  placeholder="Field Label"
                  value={field.label}
                  onChange={(e) => updateCustomField(field.id, { label: e.target.value })}
                  style={{ flex: 2 }}
                />
                <select
                  className="form-select"
                  value={field.type}
                  onChange={(e) => updateCustomField(field.id, { type: e.target.value as 'text' | 'select' })}
                  style={{ flex: 1 }}
                >
                  <option value="text">Text</option>
                  <option value="select">Dropdown</option>
                </select>
                {field.type === 'select' && (
                  <input
                    className="form-input"
                    placeholder="Options (comma-separated)"
                    value={field.options}
                    onChange={(e) => updateCustomField(field.id, { options: e.target.value })}
                    style={{ flex: 2 }}
                  />
                )}
                <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: '0.85rem', whiteSpace: 'nowrap' }}>
                  <input
                    type="checkbox"
                    checked={field.required}
                    onChange={(e) => updateCustomField(field.id, { required: e.target.checked })}
                  />
                  Required
                </label>
                <button
                  onClick={() => removeCustomField(field.id)}
                  style={{ background: 'none', border: 'none', color: 'var(--dex-red)', padding: 4 }}
                >
                  <X size={18} />
                </button>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Action buttons */}
      <div className="registration-actions mt-24">
        <button className="btn btn-danger" onClick={() => navigate(-1)}>
          <Trash2 size={16} /> Löschen
        </button>
        <button
          className="btn btn-primary"
          disabled={!title || !description}
          onClick={() => setSubmitted(true)}
          style={{ opacity: !title || !description ? 0.5 : 1 }}
        >
          <Send size={16} /> Submit
        </button>
      </div>
    </div>
  );
}
