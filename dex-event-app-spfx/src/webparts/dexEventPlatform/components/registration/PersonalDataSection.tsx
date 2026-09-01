/* PersonalDataSection — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Station 2: die persoenlichen Daten aus M365, der Umschalter „fuer andere
 * anmelden" (oeffnet den Proxy-Wizard) und der Team-Modus-Schalter.
 * Inhalt zeichengleich uebernommen. */
import * as React from 'react';
import { CollapsibleSection } from './regHelpers';
import { Icon } from '@fluentui/react/lib/Icon';
import { DeloitteEvent, Salutation } from '../../types';
import { Locale } from '../../context/LanguageContext';

/** Station 2 — Deine Daten: Profil aus M365, Stellvertreter-Einstieg, Team-Schalter. */
export interface PersonalDataSectionProps {
  canCreateEvents: boolean;
  canRegisterForOther: boolean;
  ccSelfDecidedRef: React.MutableRefObject<boolean>;
  ccSelfRef: React.MutableRefObject<boolean>;
  currentUser: import("../../types/index").User;
  email: string;
  errorBorder: { border: string; };
  event: DeloitteEvent;
  externalEmailConfirmedRef: React.MutableRefObject<boolean>;
  externalPerson: boolean;
  firstName: string;
  isAssistant: boolean;
  isMobile: boolean;
  isTeamCapable: boolean;
  isTeamMode: boolean;
  locale: Locale;
  parentAlreadyRegistered: boolean;
  pickedUserProfile: { jobTitle?: string; department?: string; location?: string; mobilePhone?: string; company?: string; };
  profileCardExpanded: boolean;
  proxyStep: 0 | 2 | 1;
  registerForOther: boolean;
  salutation: "" | Salutation;
  setEmail: React.Dispatch<React.SetStateAction<string>>;
  setExternalPerson: React.Dispatch<React.SetStateAction<boolean>>;
  setFirstName: React.Dispatch<React.SetStateAction<string>>;
  setIsTeamMode: React.Dispatch<React.SetStateAction<boolean>>;
  setMassImportOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setMassImportResult: React.Dispatch<React.SetStateAction<{ ok: number; failed: string[]; }>>;
  setMassImportRows: React.Dispatch<React.SetStateAction<{ email: string; firstName: string; lastName: string; jobTitle: string; location: string; status: "ok" | "duplicate" | "notfound"; raw: string; }[]>>;
  setMassImportStep: React.Dispatch<React.SetStateAction<"input" | "preview">>;
  setOtherConsentConfirmed: React.Dispatch<React.SetStateAction<boolean>>;
  setPendingJoinTeam: React.Dispatch<React.SetStateAction<{ teamId: string; teamName: string; }>>;
  setPickedUserProfile: React.Dispatch<React.SetStateAction<{ jobTitle?: string; department?: string; location?: string; mobilePhone?: string; company?: string; }>>;
  setProfileCardExpanded: React.Dispatch<React.SetStateAction<boolean>>;
  setProxyStep: React.Dispatch<React.SetStateAction<0 | 2 | 1>>;
  setRegisterForOther: React.Dispatch<React.SetStateAction<boolean>>;
  setSalutation: React.Dispatch<React.SetStateAction<"" | Salutation>>;
  setSurname: React.Dispatch<React.SetStateAction<string>>;
  setThirdPartyCheck: React.Dispatch<React.SetStateAction<{ alreadyRegistered: boolean; notInAudience: boolean; registeredName?: string; registeredDate?: string; }>>;
  setUserResults: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; location: string; jobTitle: string; }[]>>;
  setUserSearch: React.Dispatch<React.SetStateAction<string>>;
  showErrors: boolean;
  surname: string;
  t: (key: string) => string;
  teamSize: number;
  thirdPartyCheck: { alreadyRegistered: boolean; notInAudience: boolean; registeredName?: string; registeredDate?: string; };
}
export const PersonalDataSection: React.FC<PersonalDataSectionProps> = (p) => {
  const { canCreateEvents, canRegisterForOther, ccSelfDecidedRef, ccSelfRef, currentUser, email, errorBorder, event, externalEmailConfirmedRef, externalPerson, firstName, isAssistant, isMobile, isTeamCapable, isTeamMode, locale, parentAlreadyRegistered, pickedUserProfile, profileCardExpanded, proxyStep, registerForOther, salutation, setEmail, setExternalPerson, setFirstName, setIsTeamMode, setMassImportOpen, setMassImportResult, setMassImportRows, setMassImportStep, setOtherConsentConfirmed, setPendingJoinTeam, setPickedUserProfile, setProfileCardExpanded, setProxyStep, setRegisterForOther, setSalutation, setSurname, setThirdPartyCheck, setUserResults, setUserSearch, showErrors, surname, t, teamSize, thirdPartyCheck } = p;
  return (
        <div className="registration-form">
          {/* v11.97: Section-Header + Register-for-other-Toggle in einer
              Zeile (grünes Section-Header-Pill links, Toggle als Link
              rechts daneben). Vorher saß der Toggle unter dem Header
              im Body — wenig auffällig. „* = Required field"-Legende
              ist hier weg und sitzt jetzt am Event-Specific-Header. */}
          <CollapsibleSection
            isMobile={isMobile}
            icon="ContactInfo"
            title={t('reg.personalinfo')}
            headerExtra={(canRegisterForOther || (registerForOther && canCreateEvents)) ? (
            <>
            {canRegisterForOther && (
              <button
                type="button"
                onClick={() => {
                  setRegisterForOther(!registerForOther);
                  setThirdPartyCheck(null);
                  setPickedUserProfile(null);
                  setOtherConsentConfirmed(false);
                  setExternalPerson(false); // v18.74: Extern-Modus beim Wechsel zurücksetzen
                  // v19.6: CC-Frage-Entscheidung beim Moduswechsel zurücksetzen.
                  ccSelfDecidedRef.current = false;
                  ccSelfRef.current = false;
                  if (!registerForOther) {
                    setFirstName(''); setSurname(''); setEmail(''); setUserSearch(''); setUserResults([]);
                    // v26.76: geführten Wizard öffnen (Person suchen → Zustimmung).
                    setProxyStep(1);
                  } else {
                    setFirstName(currentUser.firstName); setSurname(currentUser.surname); setEmail(currentUser.email); setUserSearch(''); setUserResults([]);
                    setProxyStep(0);
                  }
                }}
                style={{
                  // v26.82: Als „angedockte" Tab-Optik neben dem grünen
                  // „Persönliche Informationen"-Header. Standard (für andere
                  // anmelden) = GRAU (inaktiver Tab, klarer Farbunterschied);
                  // aktiv (im Fremd-Modus, „zurück zur Selbst-Anmeldung") = grün.
                  // v26.89: Der Tab dockt jetzt SPIEGELBILDLICH zum grünen
                  // „Persönliche Informationen"-Header in die obere RECHTE Ecke
                  // — bündig an Ober- und Rechtskante (alignSelf: stretch +
                  // oben abgerundete Ecken wie der grüne Tab, unten eckig).
                  display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                  marginLeft: 'auto', // an die rechte Ecke schieben
                  alignSelf: 'stretch', boxSizing: 'border-box',
                  padding: '7px 18px', borderRadius: 'var(--dex-radius) var(--dex-radius) 0 0',
                  fontSize: '0.85rem', fontWeight: 600, cursor: 'pointer',
                  transition: 'background 0.15s ease, color 0.15s ease',
                  ...(registerForOther
                    ? { background: 'var(--dex-green, #86bc25)', border: '1.5px solid var(--dex-green, #86bc25)', color: '#fff' }
                    : { background: 'var(--dex-gray-100, #eef0f2)', border: '1.5px solid var(--dex-gray-300, #cfd4d9)', color: 'var(--dex-gray-600, #5a6470)' }),
                }}
              >
                <Icon iconName={registerForOther ? 'Contact' : 'AddFriend'} style={{ fontSize: 14 }} />
                {registerForOther ? t('reg.registerself') : t('reg.registerother')}
              </button>
            )}
            {/* v18.13: Massenimport — nur Organizer/Admin im „Für andere"-Modus. */}
            {registerForOther && canCreateEvents && (
              <button
                type="button"
                onClick={() => { setMassImportResult(null); setMassImportRows([]); setMassImportStep('input'); setMassImportOpen(true); }}
                style={{
                  background: 'none', border: 'none', padding: '4px 12px',
                  color: 'var(--dex-blue, #0076a8)', fontSize: '0.85rem',
                  textDecoration: 'underline', cursor: 'pointer', fontWeight: 600,
                }}
              >
                {locale === 'de' ? 'Massenimport' : 'Bulk import'}
              </button>
            )}
            </>
            ) : undefined}
          >
          <div style={{ padding: '24px 20px' }}>
            {canRegisterForOther && (
              <>
                {registerForOther && isAssistant && !canCreateEvents && (
                  <div style={{
                    padding: '8px 12px', marginBottom: 12, borderRadius: 'var(--dex-radius-md)',
                    background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                    color: 'var(--dex-orange)', fontSize: '0.8rem',
                  }}>
                    As an Assistant you can only register <strong>Partners</strong> or <strong>Directors</strong> for this event.
                  </div>
                )}
                {/* v23.4: Der „Person außerhalb Deloitte"-Umschalter ist nach
                    UNTEN gewandert (direkt unter die „@deloitte.com"-Such-Zeile,
                    siehe weiter unten) und in der Schriftgröße an diese Zeile
                    angeglichen. */}
                {/* v27.6: Kompakte Zusammenfassung der stellvertretend
                    anzumeldenden Person. Auswahl/Extern/Zustimmung passieren
                    komplett im Wizard-Modal — „Ändern" öffnet ihn wieder. */}
                {registerForOther && proxyStep === 0 && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20, padding: '10px 12px', borderRadius: 'var(--dex-radius-md)', background: 'rgba(134,188,37,0.08)', border: '1px solid var(--dex-green, #86bc25)' }}>
                    <Icon iconName="Contact" style={{ fontSize: 18, color: 'var(--dex-green-dark, #4a7c1f)', flexShrink: 0 }} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.03em' }}>{locale === 'de' ? 'Du meldest an' : 'You are registering'}</div>
                      <div style={{ fontWeight: 700, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{`${firstName} ${surname}`.trim() || email || (locale === 'de' ? '— keine Person gewählt —' : '— no person selected —')}</div>
                      {!!email && <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email}</div>}
                    </div>
                    <button type="button" onClick={() => setProxyStep(1)} style={{ background: 'none', border: 'none', padding: 0, color: 'var(--dex-blue, #0076a8)', textDecoration: 'underline', cursor: 'pointer', fontSize: '0.8rem', fontWeight: 600, flexShrink: 0 }}>{locale === 'de' ? 'Ändern' : 'Change'}</button>
                  </div>
                )}
                {registerForOther && thirdPartyCheck && (thirdPartyCheck.alreadyRegistered || thirdPartyCheck.notInAudience) && (
                  <div style={{
                    padding: '10px 14px', marginBottom: 16, borderRadius: 'var(--dex-radius-md)',
                    background: thirdPartyCheck.alreadyRegistered ? 'rgba(200,30,30,0.07)' : 'rgba(237,139,0,0.08)',
                    border: `1px solid ${thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)'}`,
                    color: thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)',
                    fontSize: '0.85rem',
                  }}>
                    {thirdPartyCheck.alreadyRegistered && (() => {
                      // v19.8: Konkrete Meldung mit Name + Anmeldedatum statt
                      // generischem „Diese Person ist bereits angemeldet".
                      const nm = (`${firstName} ${surname}`.trim()) || thirdPartyCheck.registeredName || (locale === 'de' ? 'Diese Person' : 'This person');
                      const d = thirdPartyCheck.registeredDate ? new Date(thirdPartyCheck.registeredDate) : null;
                      const dateStr = d && !isNaN(d.getTime())
                        ? d.toLocaleDateString(locale === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' })
                        : '';
                      return (
                        <div>
                          <strong>
                            {nm}{locale === 'de' ? ' ist bereits für das Event angemeldet' : ' is already registered for this event'}
                            {dateStr ? ` (${locale === 'de' ? 'Anmeldedatum' : 'Registered'}: ${dateStr})` : ''}.
                          </strong>
                          <div style={{ marginTop: 4, fontWeight: 400 }}>
                            {locale === 'de'
                              ? 'Eine erneute Anmeldung ist nicht möglich. Bitte wähle eine andere Person.'
                              : 'Registering this person again is not possible. Please pick a different person.'}
                          </div>
                        </div>
                      );
                    })()}
                    {thirdPartyCheck.notInAudience && (
                      <div style={{ marginTop: thirdPartyCheck.alreadyRegistered ? 6 : 0 }}>
                        <strong>{t('reg.thirdparty.notinaudience')}</strong>
                      </div>
                    )}
                  </div>
                )}
              </>
            )}

            {/* v11.80: Anrede-Dropdown nur rendern, wenn das Event das Feld
                explizit abfragt (event.askSalutation === true). Default: aus —
                viele Events brauchen die Anrede nicht. Wenn nicht gerendert,
                bleibt salutation '' und wird so in die Teilnehmer-Zeile
                geschrieben. */}
            {event.askSalutation && (
              <div className="form-group">
                <label className="form-label"><span className="required">*</span> {t('reg.salutation')}</label>
                <select className="form-select" value={salutation} onChange={e => setSalutation(e.target.value as Salutation)} style={showErrors && !salutation ? errorBorder : {}}>
                  <option value="">{t('reg.pleaseselect')}</option>
                  <option value="Herr">{locale === 'de' ? 'Herr' : 'Mr'}</option>
                  <option value="Frau">{locale === 'de' ? 'Frau' : 'Mrs'}</option>
                  <option value="Divers">{locale === 'de' ? 'Divers' : 'Diverse'}</option>
                  <option value="Keine Angabe">{locale === 'de' ? 'Keine Angabe' : 'Prefer not to say'}</option>
                </select>
              </div>
            )}

            {/* v27.13 (Feedback E. Brenneisen): Statt der grauen Feldliste eine
                Profil-KARTE mit großem Foto, Name, Position und Standort. Ein
                Plus-Toggle klappt die vollständige Liste der automatisch aus
                dem M365-Profil übernommenen Daten auf; darunter der Hinweis auf
                den automatischen Abgleich + Ticket-Verweis bei falschen Daten.
                Gilt für die EIGENE Anmeldung UND die Anmeldung Dritter mit
                Deloitte-Profil. Externe Personen (kein M365-Profil) und der
                „noch niemand gewählt"-Zustand behalten die klassischen Felder. */}
            {(() => {
              const profile = registerForOther ? pickedUserProfile : currentUser;
              const jt = profile ? ((profile as { jobTitle?: string }).jobTitle || '') : '';
              const dept = profile ? ((profile as { department?: string }).department || '') : '';
              const loc = profile ? ((profile as { location?: string }).location || '') : '';
              // v24.29: Unternehmenszugehörigkeit / Rechtsträger read-only.
              const comp = profile ? ((profile as { company?: string }).company || '') : '';
              const displayName = `${firstName} ${surname}`.trim();
              const showProfileCard = !externalPerson && !!email.trim() && !!displayName;
              if (showProfileCard) {
                const notSet = locale === 'de' ? 'nicht hinterlegt' : 'not set';
                const initials = `${(firstName.trim()[0] || '')}${(surname.trim()[0] || '')}`.toUpperCase();
                const detailRows: Array<{ label: string; value: string }> = [
                  { label: locale === 'de' ? 'E-Mail' : 'Email', value: email },
                  { label: 'Position', value: jt },
                  { label: locale === 'de' ? 'Geschäftsbereich' : 'Business Area', value: dept },
                  { label: locale === 'de' ? 'Unternehmen' : 'Company', value: comp },
                  { label: locale === 'de' ? 'Büro' : 'Office', value: loc },
                ];
                return (
                  <div className="form-group">
                    <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: '16px 18px', background: '#fff' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
                        {/* Foto: userphoto.aspx mit Initialen-Fallback (Bild
                            liegt über dem Initialen-Kreis; bei Ladefehler
                            wird es ausgeblendet und die Initialen bleiben). */}
                        <div style={{ position: 'relative', width: 88, height: 88, borderRadius: '50%', background: 'var(--dex-gray-100)', flexShrink: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: '1.5rem', color: 'var(--dex-gray-500)', overflow: 'hidden' }}>
                          {initials || '?'}
                          <img
                            src={`/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email.trim())}`}
                            alt=""
                            style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'cover' }}
                            onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
                          />
                        </div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 700, fontSize: '1.18rem', color: 'var(--dex-gray-800)' }}>{displayName}</div>
                          {jt && (
                            <div style={{ color: 'var(--dex-gray-600)', marginTop: 2 }}>{jt}</div>
                          )}
                          {loc && (
                            <div style={{ display: 'inline-flex', alignItems: 'center', gap: 5, color: 'var(--dex-gray-500)', fontSize: '0.88rem', marginTop: 3 }}>
                              <Icon iconName="POI" style={{ fontSize: 14, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                              {loc}
                            </div>
                          )}
                        </div>
                        {/* Plus-Toggle: zeigt ALLE automatisch übernommenen Daten. */}
                        <button
                          type="button"
                          onClick={() => setProfileCardExpanded(o => !o)}
                          title={profileCardExpanded
                            ? (locale === 'de' ? 'Details einklappen' : 'Collapse details')
                            : (locale === 'de' ? 'Alle automatisch übernommenen Daten anzeigen' : 'Show all automatically applied data')}
                          aria-expanded={profileCardExpanded}
                          style={{
                            width: 36, height: 36, borderRadius: '50%', flexShrink: 0,
                            border: '1px solid var(--dex-gray-300)', background: profileCardExpanded ? 'var(--dex-gray-100)' : '#fff',
                            cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center',
                            fontSize: '1.25rem', lineHeight: 1, color: 'var(--dex-gray-600)', fontWeight: 600,
                          }}
                        >
                          {profileCardExpanded ? '−' : '+'}
                        </button>
                      </div>
                      {profileCardExpanded && (
                        <div style={{ marginTop: 14, borderTop: '1px solid var(--dex-gray-100)', paddingTop: 10 }}>
                          <div style={{ fontSize: '0.72rem', fontWeight: 700, color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 6 }}>
                            {locale === 'de' ? 'Automatisch übernommene Daten' : 'Automatically applied data'}
                          </div>
                          {detailRows.map(row => (
                            <div key={row.label} style={{ display: 'flex', gap: 10, padding: '4px 0', fontSize: '0.86rem', borderBottom: '1px solid var(--dex-gray-50, #fafafa)' }}>
                              <span style={{ width: 140, flexShrink: 0, color: 'var(--dex-gray-500)' }}>{row.label}</span>
                              <span style={{ color: row.value ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)', wordBreak: 'break-word' }}>
                                {row.value || `— ${notSet}`}
                              </span>
                            </div>
                          ))}
                        </div>
                      )}
                      {/* v28.1: Hinweis bewusst klein + kursiv, ohne
                          ServiceNow-Verweis (Profildaten-Fehler sind selten;
                          der Weg zur IT ist den Kolleg:innen bekannt). */}
                      <div style={{ marginTop: 10, fontSize: '0.68rem', fontStyle: 'italic', color: 'var(--dex-gray-400)', lineHeight: 1.45 }}>
                        {locale === 'de'
                          ? <>Diese Angaben werden automatisch mit {registerForOther ? 'dem Microsoft-Profil (M365) der ausgewählten Person' : 'deinen Microsoft-Anmeldedaten (M365-Profil)'} abgeglichen und können hier nicht bearbeitet werden.</>
                          : <>These details are automatically synced with {registerForOther ? 'the selected person’s Microsoft profile (M365)' : 'your Microsoft sign-in data (M365 profile)'} and cannot be edited here.</>}
                      </div>
                    </div>
                  </div>
                );
              }
              // v28.6: „Für andere" ohne gewählte Person → KEINE leeren
              // Alt-Felder mehr (der Wizard ist ohnehin offen); stattdessen
              // ein kompakter Hinweis. Die Profil-Karte erscheint, sobald
              // eine Person gewählt wurde.
              if (registerForOther && !externalPerson) {
                return (
                  <div style={{ padding: '16px 18px', border: '1px dashed var(--dex-gray-300)', borderRadius: 10, background: 'var(--dex-gray-50, #fafafa)', color: 'var(--dex-gray-500)', fontSize: '0.85rem', lineHeight: 1.5 }}>
                    {locale === 'de'
                      ? 'Wähle zuerst eine Person aus (Fenster „Für eine andere Person anmelden") — ihre Daten erscheinen dann hier als Profil-Karte.'
                      : 'First pick a person (window “Register another person”) — their details then appear here as a profile card.'}
                  </div>
                );
              }
              // Klassische Felder: nur noch für EXTERNE Personen.
              // v28.68: Sicherheitsnetz. Konnte der Name nicht aus M365
              // aufgelöst werden (z.B. weil in der versteckten Benutzerliste
              // der Site das Claims-Login-Token statt des Namens steht und
              // auch das Benutzerprofil nichts hergibt), waren diese Felder
              // fest deaktiviert, die Pflichtprüfung verlangte aber Vor- UND
              // Nachnamen: „Bitte alle Pflichtfelder ausfüllen" ohne ein
              // einziges ausfüllbares Feld — die Anmeldung war unmöglich.
              // Fehlt einer der beiden Namen, sind sie jetzt editierbar.
              const nameUnresolved = !firstName.trim() || !surname.trim();
              return (
                <>
                  <div className="form-group">
                    <label className="form-label">{t('reg.firstname')}</label>
                    <input className="form-input" value={firstName} onChange={e => { if (externalPerson || nameUnresolved) setFirstName(e.target.value); }} placeholder={t('reg.firstname')} disabled={!externalPerson && !nameUnresolved} style={{ background: nameUnresolved ? 'var(--dex-white, #fff)' : 'var(--dex-gray-100)', ...(showErrors && !firstName.trim() ? errorBorder : {}) }} />
                  </div>
                  <div className="form-group">
                    <label className="form-label">{t('reg.surname')}</label>
                    <input className="form-input" value={surname} onChange={e => { if (externalPerson || nameUnresolved) setSurname(e.target.value); }} placeholder={t('reg.surname')} disabled={!externalPerson && !nameUnresolved} style={{ background: nameUnresolved ? 'var(--dex-white, #fff)' : 'var(--dex-gray-100)', ...(showErrors && !surname.trim() ? errorBorder : {}) }} />
                  </div>
                  {nameUnresolved && !externalPerson && (
                    <div style={{
                      marginTop: -4, marginBottom: 12, padding: '8px 10px', borderRadius: 6,
                      fontSize: '0.78rem', lineHeight: 1.5,
                      background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
                    }}>
                      {locale === 'de'
                        ? 'Dein Name konnte nicht aus deinem M365-Profil gelesen werden. Bitte trage Vor- und Nachnamen einmal von Hand ein — danach kannst du dich ganz normal anmelden.'
                        : 'We could not read your name from your M365 profile. Please enter your first and last name once — after that you can register as usual.'}
                    </div>
                  )}
                  <div className="form-group">
                    <label className="form-label">{t('reg.email')}</label>
                    <input className="form-input" type="email" value={email} onChange={e => { if (externalPerson) { setEmail(e.target.value); externalEmailConfirmedRef.current = false; /* v18.74: Tippfehler-Check bei Änderung erneut erzwingen */ } }} placeholder={externalPerson ? 'name@firma.de' : 'email@deloitte.de'} disabled style={{ background: 'var(--dex-gray-100)', ...(showErrors && !email.trim() ? errorBorder : {}) }} />
                  </div>
                </>
              );
            })()}

            {/* v11.82: Team-Anmeldung-Toggle. Nur sichtbar wenn der Organizer
                in Schritt 4 die Team-Anmeldung aktiviert hat UND der User sich
                NICHT für eine andere Person registriert (Team-für-Andere wird
                nicht unterstützt — der Stellvertreter-Pfad ist auf eine
                Einzel-Person ausgelegt). */}
            {isTeamCapable && !registerForOther && !parentAlreadyRegistered && (
              <div className="form-group" style={{ marginTop: 16, marginBottom: 0 }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={isTeamMode}
                    onChange={e => { setIsTeamMode(e.target.checked); if (e.target.checked) setPendingJoinTeam(null); }}
                    style={{ marginTop: 3 }}
                  />
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)' }}>
                      {locale === 'de'
                        ? `Ich melde mich + mein Team an (Team-Anmeldung)`
                        : 'Register me + my team (team registration)'}
                    </div>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4, lineHeight: 1.4 }}>
                      {locale === 'de'
                        ? `Belegt bis zu ${teamSize} Plätze auf einmal. Jedes Mitglied bekommt automatisch Bestätigungsmail, Outlook-Termin und sieht das Event in „Meine Events".`
                        : `Books up to ${teamSize} seats at once. Each member automatically receives a confirmation email, an Outlook invite, and sees the event in „My Events".`}
                    </div>
                  </div>
                </label>
              </div>
            )}
          </div>
          </CollapsibleSection>
        </div>
  );
};
