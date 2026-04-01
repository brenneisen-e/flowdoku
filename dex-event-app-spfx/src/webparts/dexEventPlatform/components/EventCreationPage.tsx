/**
 * Event-Erstellung (nur fuer Organizer / SuperAdmin)
 *
 * Erstellt ein Event in der DEX_Events-Liste und eine
 * separate Teilnehmerliste mit Item-Level Permissions.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { EventType } from '../types';
import { Trash2, Send, Plus, X, ChevronUp, ChevronDown } from './Icons';

interface CustomFieldInput {
  id: string;
  label: string;
  type: 'text' | 'select' | 'number' | 'checkbox';
  required: boolean;
  options: string;
  visible: boolean;
}

export default function EventCreationPage(): React.ReactElement {
  const { navigate, goBack } = useNavigation();
  const { createEvent } = useEvents();
  const { currentUser } = useCurrentUser();
  const [title, setTitle] = React.useState('');
  const [organizer, setOrganizer] = React.useState(
    `${currentUser.firstName} ${currentUser.surname}`
  );
  const [location, setLocation] = React.useState('');
  const [locationFilter, setLocationFilter] = React.useState('');
  const [audience, setAudience] = React.useState('');
  const [description, setDescription] = React.useState('');
  const [eventType, setEventType] = React.useState<EventType>('Other');
  const [startDate, setStartDate] = React.useState('');
  const [endDate, setEndDate] = React.useState('');
  const [registrationDeadline, setRegistrationDeadline] = React.useState('');
  const [lastDeregisterDate, setLastDeregisterDate] = React.useState('');
  const [maxParticipants, setMaxParticipants] = React.useState('');
  const [waitlistEnabled, setWaitlistEnabled] = React.useState(true);
  const [eventImageUrl, setEventImageUrl] = React.useState('');
  const [customFields, setCustomFields] = React.useState<CustomFieldInput[]>([]);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [submitted, setSubmitted] = React.useState(false);
  const [error, setError] = React.useState('');

  const locationOptions = ['Berlin', 'Düsseldorf', 'Frankfurt', 'Hamburg', 'Hannover', 'Köln', 'München', 'Stuttgart', 'All'];

  const addCustomField = (): void => {
    setCustomFields([...customFields, {
      id: `cf-${Date.now()}`, label: '', type: 'text',
      required: false, options: '', visible: true,
    }]);
  };

  const removeCustomField = (id: string): void => {
    setCustomFields(customFields.filter(f => f.id !== id));
  };

  const updateCustomField = (id: string, updates: Partial<CustomFieldInput>): void => {
    setCustomFields(customFields.map(f => f.id === id ? { ...f, ...updates } : f));
  };

  const fillDemo = (): void => {
    const now = new Date();
    const nextWeek = new Date(now);
    nextWeek.setDate(now.getDate() + 7);
    const nextWeekEnd = new Date(nextWeek);
    nextWeekEnd.setHours(nextWeek.getHours() + 4);
    const deadline = new Date(nextWeek);
    deadline.setDate(nextWeek.getDate() - 2);
    const lastDereg = new Date(nextWeek);
    lastDereg.setDate(nextWeek.getDate() - 1);

    const toLocal = (d: Date): string => {
      const pad = (n: number): string => (n < 10 ? '0' : '') + n;
      return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
    };

    const dateStr = toLocal(nextWeek).replace(/[-T:]/g, '').substring(0, 8);
    setTitle(`Test_${dateStr}`);
    setEventType('Other');
    setDescription('Testbeschreibung für ein Demo-Event.');
    setLocation('Köln, Testort');
    setLocationFilter('');
    setAudience('All');
    setStartDate(toLocal(nextWeek));
    setEndDate(toLocal(nextWeekEnd));
    setRegistrationDeadline(toLocal(deadline));
    setLastDeregisterDate(toLocal(lastDereg));
    setMaxParticipants('50');
    setWaitlistEnabled(true);
    setEventImageUrl('');
    setCustomFields([
      { id: `cf-${Date.now()}`, label: 'T-Shirt Größe', type: 'select', required: true, options: 'XS, S, M, L, XL, XXL', visible: true },
      { id: `cf-${Date.now() + 1}`, label: 'Notfallkontakt (Name & Telefon)', type: 'text', required: true, options: '', visible: true },
      { id: `cf-${Date.now() + 2}`, label: 'Ernährungsbesonderheiten', type: 'text', required: false, options: '', visible: true },
    ]);
  };

  const moveCustomField = (id: string, direction: 'up' | 'down'): void => {
    const idx = customFields.findIndex(f => f.id === id);
    if (idx < 0) return;
    const newIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (newIdx < 0 || newIdx >= customFields.length) return;
    const updated = [...customFields];
    [updated[idx], updated[newIdx]] = [updated[newIdx], updated[idx]];
    setCustomFields(updated);
  };

  const handleSubmit = async (): Promise<void> => {
    if (!title || !description) return;
    setIsSubmitting(true);
    setError('');

    const eventId = await createEvent({
      title,
      type: eventType,
      status: 'Active',
      description,
      location,
      locationFilter,
      audience,
      startDate: startDate ? new Date(startDate).toISOString() : '',
      endDate: endDate ? new Date(endDate).toISOString() : '',
      registrationDeadline: registrationDeadline ? new Date(registrationDeadline).toISOString() : '',
      lastDeregisterDate: lastDeregisterDate ? new Date(lastDeregisterDate).toISOString() : '',
      maxParticipants: Number(maxParticipants) || 0,
      waitlistEnabled,
      eventImageUrl,
      organizer,
      organizerEmail: currentUser.email,
      outlookEventId: '',
      customFields: customFields.map(f => ({
        id: f.id,
        label: f.label,
        type: f.type,
        required: f.required,
        visible: f.visible,
        ...(f.type === 'select' ? { options: f.options.split(',').map(o => o.trim()).filter(Boolean) } : {}),
      })),
    });

    setIsSubmitting(false);

    if (eventId) {
      setSubmitted(true);
    } else {
      setError('Event konnte nicht erstellt werden. Bitte versuche es erneut.');
    }
  };

  if (submitted) {
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '64px 32px' }}>
          <h2>Event erfolgreich erstellt!</h2>
          <p className="mt-8" style={{ color: 'var(--dex-gray-600)' }}>
            "{title}" wurde angelegt. Die Teilnehmerliste wurde auf SharePoint erstellt.
          </p>
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center' }}>
            <button className="btn btn-primary" onClick={() => navigate('register')}>Events anzeigen</button>
            <button className="btn btn-secondary" onClick={() => { setSubmitted(false); setTitle(''); }}>Weiteres Event erstellen</button>
          </div>
        </div>
      </div>
    );
  }

  // Hilfsfunktion fuer die Vorschau
  const formatPreviewDate = (val: string): string => {
    if (!val) return '--';
    const d = new Date(val);
    return d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
      ' ' + d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  };

  return (
    <div className="page-container">
      <div className="creation-layout">
        {/* ===== Linke Seite: Formular ===== */}
        <div>
          <div className="card">
            <div className="creation-form">
              {error && (
                <div style={{ padding: '10px 16px', background: '#fce4ec', color: '#c62828', borderRadius: 8, marginBottom: 16, fontSize: '0.85rem' }}>
                  {error}
                </div>
              )}

              <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 16 }}>
                <button
                  className="btn btn-outline"
                  onClick={fillDemo}
                  style={{ fontSize: '0.8rem', padding: '4px 12px' }}
                >
                  Demo
                </button>
              </div>

              <div className="form-group">
                <label className="form-label"><span className="required">*</span> Event Titel</label>
                <input className="form-input" value={title} onChange={e => setTitle(e.target.value)} />
              </div>

              <div className="form-group">
                <label className="form-label"><span className="required">*</span> Event Typ</label>
                <select className="form-select" value={eventType} onChange={e => setEventType(e.target.value as EventType)}>
                  <option value="Other">Sonstiges Deloitte Event</option>
                  <option value="B2Run">B2Run</option>
                  <option value="JPMorgan">JP Morgan Run</option>
                </select>
              </div>

              <div className="form-group">
                <label className="form-label"><span className="required">*</span> Organisator</label>
                <input className="form-input" value={organizer} onChange={e => setOrganizer(e.target.value)} />
              </div>

              <div className="form-group">
                <label className="form-label">Veranstaltungsort</label>
                <input className="form-input" value={location} onChange={e => setLocation(e.target.value)} />
              </div>

              <div className="form-group">
                <label className="form-label">
                  Standort-Filter
                  <span className="info-icon" title="Welche Standorte sollen das Event sehen und sich registrieren können?" style={{ marginLeft: 8 }}>i</span>
                </label>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                  {locationOptions.map(loc => (
                    <label key={loc} style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: '0.9rem' }}>
                      <input
                        type="checkbox"
                        checked={locationFilter.split(',').map(s => s.trim()).indexOf(loc) >= 0}
                        onChange={e => {
                          const current = locationFilter.split(',').map(s => s.trim()).filter(Boolean);
                          if (e.target.checked) setLocationFilter([...current, loc].join(', '));
                          else setLocationFilter(current.filter(l => l !== loc).join(', '));
                        }}
                      />
                      {loc}
                    </label>
                  ))}
                </div>
              </div>

              <div className="form-group">
                <label className="form-label">Audience / Zielgruppe</label>
                <input className="form-input" value={audience} onChange={e => setAudience(e.target.value)} placeholder="z.B. Technology & Transformation, All" />
              </div>

              <div className="form-group">
                <label className="form-label"><span className="required">*</span> Beschreibung</label>
                <textarea className="form-textarea" value={description} onChange={e => setDescription(e.target.value)} style={{ minHeight: 120 }} />
              </div>

              <div className="form-group">
                <label className="form-label">Event-Bild URL</label>
                <input className="form-input" value={eventImageUrl} onChange={e => setEventImageUrl(e.target.value)} placeholder="https://..." />
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                <div className="form-group">
                  <label className="form-label"><span className="required">*</span> Startdatum</label>
                  <input className="form-input" type="datetime-local" value={startDate} onChange={e => setStartDate(e.target.value)} />
                </div>
                <div className="form-group">
                  <label className="form-label"><span className="required">*</span> Enddatum</label>
                  <input className="form-input" type="datetime-local" value={endDate} onChange={e => setEndDate(e.target.value)} />
                </div>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                <div className="form-group">
                  <label className="form-label">Anmelde-Deadline</label>
                  <input className="form-input" type="datetime-local" value={registrationDeadline} onChange={e => setRegistrationDeadline(e.target.value)} />
                </div>
                <div className="form-group">
                  <label className="form-label">Letzte Abmeldemöglichkeit</label>
                  <input className="form-input" type="datetime-local" value={lastDeregisterDate} onChange={e => setLastDeregisterDate(e.target.value)} />
                </div>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                <div className="form-group">
                  <label className="form-label">Max. Teilnehmer</label>
                  <input className="form-input" type="number" min={0} value={maxParticipants} onChange={e => setMaxParticipants(e.target.value)} placeholder="0 = unbegrenzt" />
                </div>
                <div className="form-group">
                  <label className="form-label">Warteliste</label>
                  <div className="toggle-wrapper" style={{ marginTop: 8 }}>
                    <label className="toggle">
                      <input type="checkbox" checked={waitlistEnabled} onChange={e => setWaitlistEnabled(e.target.checked)} />
                      <span className="toggle-slider" />
                    </label>
                    <span style={{ fontSize: '0.9rem' }}>{waitlistEnabled ? 'Aktiviert' : 'Deaktiviert'}</span>
                  </div>
                </div>
              </div>

              {/* Dynamische Felder */}
              <div style={{ borderTop: '1px solid var(--dex-gray-200)', paddingTop: 24, marginTop: 8 }}>
                <div className="flex-between mb-16">
                  <label className="form-label" style={{ marginBottom: 0 }}>Zusätzliche Registrierungsfelder</label>
                  <button className="btn btn-outline" onClick={addCustomField} style={{ fontSize: '0.85rem', padding: '6px 14px' }}>
                    <Plus size={14} /> Feld hinzufügen
                  </button>
                </div>
                {customFields.map((field, idx) => (
                  <div key={field.id} className="custom-field-row">
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                      <button
                        onClick={() => moveCustomField(field.id, 'up')}
                        disabled={idx === 0}
                        style={{ background: 'none', border: 'none', padding: 2, color: idx === 0 ? 'var(--dex-gray-300)' : 'var(--dex-gray-600)', cursor: idx === 0 ? 'default' : 'pointer' }}
                        title="Nach oben"
                      >
                        <ChevronUp size={14} />
                      </button>
                      <button
                        onClick={() => moveCustomField(field.id, 'down')}
                        disabled={idx === customFields.length - 1}
                        style={{ background: 'none', border: 'none', padding: 2, color: idx === customFields.length - 1 ? 'var(--dex-gray-300)' : 'var(--dex-gray-600)', cursor: idx === customFields.length - 1 ? 'default' : 'pointer' }}
                        title="Nach unten"
                      >
                        <ChevronDown size={14} />
                      </button>
                    </div>
                    <input className="form-input" placeholder="Feldname" value={field.label} onChange={e => updateCustomField(field.id, { label: e.target.value })} style={{ flex: 2 }} />
                    <select className="form-select" value={field.type} onChange={e => updateCustomField(field.id, { type: e.target.value as CustomFieldInput['type'] })} style={{ flex: 1 }}>
                      <option value="text">Text</option>
                      <option value="select">Dropdown</option>
                      <option value="number">Zahl</option>
                      <option value="checkbox">Checkbox</option>
                    </select>
                    {field.type === 'select' && (
                      <input className="form-input" placeholder="Optionen (kommasepariert)" value={field.options} onChange={e => updateCustomField(field.id, { options: e.target.value })} style={{ flex: 2 }} />
                    )}
                    <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: '0.85rem', whiteSpace: 'nowrap' }}>
                      <input type="checkbox" checked={field.required} onChange={e => updateCustomField(field.id, { required: e.target.checked })} />
                      Pflicht
                    </label>
                    <button onClick={() => removeCustomField(field.id)} style={{ background: 'none', border: 'none', color: 'var(--dex-red)', padding: 4 }}>
                      <X size={18} />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          </div>

          <div className="registration-actions mt-24">
            <button className="btn btn-danger" onClick={() => goBack()}><Trash2 size={16} /> Abbrechen</button>
            <button
              className="btn btn-primary"
              disabled={!title || !description || isSubmitting}
              onClick={handleSubmit}
              style={{ opacity: !title || !description || isSubmitting ? 0.5 : 1 }}
            >
              <Send size={16} /> {isSubmitting ? 'Wird erstellt...' : 'Event erstellen'}
            </button>
          </div>
        </div>

        {/* ===== Rechte Seite: Live-Vorschau ===== */}
        <div style={{ position: 'sticky', top: 16 }}>
          <div style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-400)', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 12 }}>
            Vorschau: Registrierungsseite
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16, transform: 'scale(0.92)', transformOrigin: 'top left', width: 'calc(100% / 0.92)' }}>
            {/* Event-Karte */}
            <div className="registration-event" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
              <div className="section-header section-header--red">Selected Event</div>
              <div className="registration-event__card">
                <div
                  className="registration-event__image"
                  style={{
                    background: eventImageUrl
                      ? `url(${eventImageUrl}) center/cover`
                      : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
                  }}
                >
                  <div className="registration-event__overlay">
                    <h4>{title || 'Event Titel'}</h4>
                    <p>
                      {formatPreviewDate(startDate)} until<br />
                      {formatPreviewDate(endDate)}
                    </p>
                  </div>
                </div>
              </div>
            </div>

            {/* Personal Information */}
            <div className="registration-form" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
              <div className="section-header">Personal Information</div>
              <div style={{ padding: '16px 20px' }}>
                <div className="form-group">
                  <label className="form-label"><span className="required">*</span> Salutation</label>
                  <select className="form-select" disabled><option>Please select</option></select>
                </div>
                <div className="form-group">
                  <label className="form-label"><span className="required">*</span> First Name</label>
                  <input className="form-input" disabled placeholder="First Name" />
                </div>
                <div className="form-group">
                  <label className="form-label"><span className="required">*</span> Surname</label>
                  <input className="form-input" disabled placeholder="Surname" />
                </div>
                <div className="form-group">
                  <label className="form-label"><span className="required">*</span> E-Mail</label>
                  <input className="form-input" disabled placeholder="email@deloitte.de" />
                </div>
              </div>
            </div>

            {/* Event specific Information */}
            <div className="registration-specific" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
              <div className="section-header">Event specific Information</div>
              <div style={{ padding: '16px 20px' }}>
                {customFields.filter(f => f.label).length === 0 ? (
                  <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', fontSize: '0.9rem' }}>No additional information required.</p>
                ) : (
                  customFields.filter(f => f.label).map(field => (
                    <div className="form-group" key={field.id}>
                      <label className="form-label">
                        {field.required && <span className="required">*</span>}
                        {field.label}
                      </label>
                      {field.type === 'select' ? (
                        <select className="form-select" disabled>
                          <option>Please select</option>
                          {field.options.split(',').map(o => o.trim()).filter(Boolean).map(opt => (
                            <option key={opt}>{opt}</option>
                          ))}
                        </select>
                      ) : field.type === 'checkbox' ? (
                        <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.9rem' }}>
                          <input type="checkbox" disabled /> {field.label}
                        </label>
                      ) : (
                        <input className="form-input" disabled placeholder={field.label} type={field.type === 'number' ? 'number' : 'text'} />
                      )}
                    </div>
                  ))
                )}
              </div>
            </div>

            {/* Vorschau-Buttons (deaktiviert) */}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 16 }}>
              <button className="btn btn-danger" disabled style={{ opacity: 0.5 }}><Trash2 size={16} /> Delete</button>
              <button className="btn btn-primary" disabled style={{ opacity: 0.5 }}><Send size={16} /> Register</button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
