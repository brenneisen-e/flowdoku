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
import { Icon } from '@fluentui/react/lib/Icon';
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
  const { canCreateEvents, searchUsers, isOrganizer, isAdmin } = useRoles();
  const { t } = useLanguage();
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
  // Deloitte-Mitarbeitersuche
  const [userSearch, setUserSearch] = React.useState('');
  const [userResults, setUserResults] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const [isSearchingUser, setIsSearchingUser] = React.useState(false);
  const searchTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  if (!event) {
    return (
      <div className="page-container text-center">
        <h2>{t('reg.eventnotfound')}</h2>
        <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>
          {t('reg.backtoevents')}
        </button>
      </div>
    );
  }

  // Registrierungs-Deadline pruefen
  const isDeadlinePassed = event.registrationDeadline && new Date(event.registrationDeadline) < new Date();

  if (isDeadlinePassed && !isOrganizer && !isAdmin) {
    return (
      <div className="page-container">
        <div className="card" style={{ position: 'relative', overflow: 'hidden' }}>
          <div style={{
            height: 200,
            background: event.imageUrl
              ? `url(${event.imageUrl}) center/cover no-repeat`
              : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
            borderRadius: '16px 16px 0 0',
          }} />
          <div style={{ padding: 32, textAlign: 'center' }}>
            <Icon iconName="Clock" style={{ fontSize: 48, color: 'var(--dex-orange)', marginBottom: 16 }} />
            <h2 style={{ marginBottom: 8 }}>{t('reg.deadlinepassed.title')}</h2>
            <p style={{ color: 'var(--dex-gray-600)', marginBottom: 8 }}>
              {t('reg.deadlinepassed.text')}
            </p>
            <p style={{ color: 'var(--dex-gray-400)', fontSize: '0.85rem' }}>
              {t('reg.deadlinepassed.date')}: {new Date(event.registrationDeadline).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}
            </p>
            <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>
              {t('reg.backtoevents')}
            </button>
          </div>
        </div>
      </div>
    );
  }

  const isFull = event.maxParticipants > 0 && event.currentParticipants >= event.maxParticipants;
  const errorBorder = { border: '2px solid var(--dex-red)' };

  const handleSubmit = async (): Promise<void> => {
    // Validierung Pflichtfelder
    setShowErrors(true);
    if (!salutation || !firstName.trim() || !surname.trim() || !email.trim()) {
      setError(t('reg.requiredfields'));
      return;
    }

    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) {
      setError(t('reg.invalidemail') || 'Ungültige E-Mail-Adresse');
      return;
    }

    // Pflicht-Custom-Fields validieren
    const missingRequired = event.eventSpecificFields
      .filter(f => f.required && !eventSpecific[f.id]?.trim());
    if (missingRequired.length > 0) {
      setError(`${t('reg.requiredcustom')}: ${missingRequired.map(f => f.label).join(', ')}`);
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
        setError(t('reg.error'));
      }
    } catch {
      setError(t('reg.genericerror'));
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
          <h2>{isFull ? t('reg.waitlisttitle') : t('reg.success')}</h2>
          <p className="mt-8" style={{ color: 'var(--dex-gray-600)' }}>
            {isFull
              ? t('reg.waitlistmsg').replace('{title}', event.title)
              : t('reg.successmsg').replace('{title}', event.title).replace('{email}', email)}
          </p>
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center' }}>
            <button className="btn btn-primary" onClick={() => navigate('my-events')}>
              {t('myevents.title')}
            </button>
            <button className="btn btn-secondary" onClick={() => navigate('register')}>
              {t('reg.registeranother')}
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
          {t('reg.locationnotice')}
          {event && event.locationAudience.length > 0 && <> {t('reg.locationfilter')}: <strong>{event.locationAudience.join(', ')}</strong>.</>}
          {event && event.audienceFilter && event.audienceFilter.length > 0 && <> {t('reg.audience')}: <strong>{event.audienceFilter.join(', ')}</strong>.</>}
          {event && event.filterMode === 'AND' && <> ({t('reg.andmode')})</>}
          {' '}{t('reg.yourlocation')}: {currentUser.location || t('reg.unknown')}.
        </div>
      )}
      <div className="registration-layout">
        {/* Event-Info links */}
        <div className="registration-event">
          <div className="section-header section-header--red">{t('reg.selectedevent')}</div>
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
                  {formatDate(event.startDate)} {t('reg.until')}<br />
                  {formatDate(event.endDate)}
                </p>
                <div style={{ fontSize: '0.78rem', opacity: 0.8, marginTop: 4 }}>
                  Organizer:
                  <ul style={{ margin: '2px 0 0 16px', padding: 0, listStyle: 'disc' }}>
                    {event.organizers.reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map((o, i) => {
                      const trimmed = o.trim();
                      const parts = trimmed.split(',').map(s => s.trim());
                      const name = parts.length === 2 ? `${parts[1]} ${parts[0]}` : trimmed;
                      return <li key={i}>{name}</li>;
                    })}
                  </ul>
                </div>
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
                {t('reg.allplaces').replace('{count}', String(event.waitlistCount))}
              </p>
            )}
          </div>
        </div>

        {/* Persoenliche Daten */}
        <div className="registration-form">
          <div className="section-header">{t('reg.personalinfo')}</div>
          <div style={{ padding: '24px 20px' }}>
            <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 12 }}>
              <span className="required">*</span> = {t('reg.requiredfield')}
            </p>

            {canCreateEvents && (
              <>
                <button
                  className="btn btn-outline"
                  style={{ marginBottom: 20, fontSize: '0.85rem' }}
                  onClick={() => {
                    setRegisterForOther(!registerForOther);
                    if (!registerForOther) { setFirstName(''); setSurname(''); setEmail(''); setUserSearch(''); setUserResults([]); }
                    else { setFirstName(currentUser.firstName); setSurname(currentUser.surname); setEmail(currentUser.email); setUserSearch(''); setUserResults([]); }
                  }}
                >
                  {registerForOther ? t('reg.registerself') : t('reg.registerother')}
                </button>
                {registerForOther && (
                  <div className="form-group" style={{ position: 'relative', marginBottom: 20 }}>
                    <label className="form-label">{t('reg.searchemployee') || 'Deloitte Mitarbeiter suchen'}</label>
                    <input
                      className="form-input"
                      value={userSearch}
                      onChange={e => {
                        const val = e.target.value;
                        setUserSearch(val);
                        if (searchTimerRef.current) clearTimeout(searchTimerRef.current);
                        if (val.length >= 2) {
                          searchTimerRef.current = setTimeout(async () => {
                            setIsSearchingUser(true);
                            const results = await searchUsers(val);
                            setUserResults(results);
                            setIsSearchingUser(false);
                          }, 300);
                        } else {
                          setUserResults([]);
                        }
                      }}
                      placeholder={t('reg.searchplaceholder') || 'Name oder E-Mail eingeben...'}
                    />
                    {isSearchingUser && (
                      <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>Suche...</div>
                    )}
                    {userResults.length > 0 && (
                      <div style={{
                        position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
                        background: '#fff', border: '1px solid var(--dex-gray-200)',
                        borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                        maxHeight: 200, overflowY: 'auto',
                      }}>
                        {userResults.map(u => {
                          // "Nachname, Vorname" oder "Vorname Nachname" Format
                          let uFirstName = '';
                          let uSurname = '';
                          if (u.displayName.includes(',')) {
                            const parts = u.displayName.split(',').map(s => s.trim());
                            uSurname = parts[0] || '';
                            uFirstName = parts[1] || '';
                          } else {
                            const parts = u.displayName.split(' ');
                            uFirstName = parts[0] || '';
                            uSurname = parts.slice(1).join(' ') || '';
                          }
                          return (
                            <div
                              key={u.email}
                              style={{
                                padding: '8px 12px', cursor: 'pointer', fontSize: '0.85rem',
                                borderBottom: '1px solid var(--dex-gray-100)',
                              }}
                              onMouseEnter={e => { (e.currentTarget as HTMLElement).style.background = 'var(--dex-gray-50)'; }}
                              onMouseLeave={e => { (e.currentTarget as HTMLElement).style.background = '#fff'; }}
                              onMouseDown={() => {
                                setFirstName(uFirstName);
                                setSurname(uSurname);
                                setEmail(u.email);
                                setUserSearch(u.displayName);
                                setUserResults([]);
                              }}
                            >
                              <strong>{u.displayName}</strong>
                              <span style={{ color: 'var(--dex-gray-400)', marginLeft: 8 }}>{u.email}</span>
                              {u.location && <span style={{ color: 'var(--dex-gray-400)', marginLeft: 8, fontSize: '0.8rem' }}>({u.location})</span>}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                )}
              </>
            )}

            <div className="form-group">
              <label className="form-label"><span className="required">*</span> {t('reg.salutation')}</label>
              <select className="form-select" value={salutation} onChange={e => setSalutation(e.target.value as Salutation)} style={showErrors && !salutation ? errorBorder : {}}>
                <option value="">{t('reg.pleaseselect')}</option>
                <option value="Herr">Herr</option>
                <option value="Frau">Frau</option>
                <option value="Divers">Divers</option>
              </select>
            </div>

            <div className="form-group">
              <label className="form-label"><span className="required">*</span> {t('reg.firstname')}</label>
              <input className="form-input" value={firstName} onChange={e => { if (registerForOther) setFirstName(e.target.value); }} placeholder={t('reg.firstname')} disabled={!registerForOther} style={{ background: 'var(--dex-gray-100)', ...(showErrors && !firstName.trim() ? errorBorder : {}) }} />
            </div>

            <div className="form-group">
              <label className="form-label"><span className="required">*</span> {t('reg.surname')}</label>
              <input className="form-input" value={surname} onChange={e => { if (registerForOther) setSurname(e.target.value); }} placeholder={t('reg.surname')} disabled={!registerForOther} style={{ background: 'var(--dex-gray-100)', ...(showErrors && !surname.trim() ? errorBorder : {}) }} />
            </div>

            <div className="form-group">
              <label className="form-label"><span className="required">*</span> {t('reg.email')}</label>
              <input className="form-input" type="email" value={email} onChange={e => { if (registerForOther) setEmail(e.target.value); }} placeholder="email@deloitte.de" disabled={!registerForOther} style={{ background: 'var(--dex-gray-100)', ...(showErrors && !email.trim() ? errorBorder : {}) }} />
            </div>
          </div>
        </div>

        {/* Eventspezifische Felder */}
        <div className="registration-specific">
          <div className="section-header">{t('reg.eventinfo')}</div>
          <div style={{ padding: '24px 20px' }}>
            {event.eventSpecificFields.length === 0 ? (
              <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{t('reg.noadditional')}</p>
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
                      <option value="">{t('reg.pleaseselect')}</option>
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
          {t('reg.privacy').replace('{title}', event.title)}
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
        <button className="btn btn-danger" onClick={handleClear} disabled={isSubmitting}><Trash2 size={16} /> {t('reg.delete')}</button>
        <button className="btn btn-primary" onClick={handleSubmit} disabled={isSubmitting}>
          <Send size={16} /> {isSubmitting ? t('reg.submitting') : t('reg.register')}
        </button>
      </div>
    </div>
  );
}
