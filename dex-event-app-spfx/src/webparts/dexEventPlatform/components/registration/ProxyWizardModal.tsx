/* ProxyWizardModal — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Der gefuehrte Wizard fuer die stellvertretende Anmeldung (v26.76): Schritt 1
 * Person suchen, Schritt 2 Zustimmung. Inhalt zeichengleich uebernommen; die
 * Anzeige-Bedingung (`proxyStep > 0`) ist beim Aufrufer geblieben. */
import * as React from 'react';
import Modal from '../Modal';
import InternationalSearchToggle from '../InternationalSearchToggle';
import { Locale } from '../../context/LanguageContext';
import { DeloitteEvent } from '../../types';

/** Gefuehrter Wizard fuer die stellvertretende Anmeldung (v26.76). */
export interface ProxyWizardModalProps {
  canCreateEvents: boolean;
  checkRegistrationByEmail: (eventId: string, email: string) => Promise<import("../../services/EventService").SPRegistration>;
  currentUser: import("../../types/index").User;
  email: string;
  event: DeloitteEvent;
  externalEmailConfirmedRef: React.MutableRefObject<boolean>;
  externalPerson: boolean;
  firstName: string;
  isAllowedTargetForAssistant: (jt: string) => boolean;
  isAssistant: boolean;
  isSearchingUser: boolean;
  locale: Locale;
  otherConsentConfirmed: boolean;
  pickedUserProfile: { jobTitle?: string; department?: string; location?: string; mobilePhone?: string; company?: string; };
  pickProxyUser: (u: { email: string; displayName: string; location?: string; jobTitle?: string; }) => void;
  proxyStep: 0 | 2 | 1;
  searchTimerRef: React.MutableRefObject<NodeJS.Timeout>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  setEmail: React.Dispatch<React.SetStateAction<string>>;
  setExternalPerson: React.Dispatch<React.SetStateAction<boolean>>;
  setFirstName: React.Dispatch<React.SetStateAction<string>>;
  setIsSearchingUser: React.Dispatch<React.SetStateAction<boolean>>;
  setMassImportOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setMassImportResult: React.Dispatch<React.SetStateAction<{ ok: number; failed: string[]; }>>;
  setMassImportRows: React.Dispatch<React.SetStateAction<{ email: string; firstName: string; lastName: string; jobTitle: string; location: string; status: "ok" | "duplicate" | "notfound"; raw: string; }[]>>;
  setMassImportStep: React.Dispatch<React.SetStateAction<"input" | "preview">>;
  setOtherConsentConfirmed: React.Dispatch<React.SetStateAction<boolean>>;
  setPickedUserProfile: React.Dispatch<React.SetStateAction<{ jobTitle?: string; department?: string; location?: string; mobilePhone?: string; company?: string; }>>;
  setProxyStep: React.Dispatch<React.SetStateAction<0 | 2 | 1>>;
  setRegisterForOther: React.Dispatch<React.SetStateAction<boolean>>;
  setSurname: React.Dispatch<React.SetStateAction<string>>;
  setThirdPartyCheck: React.Dispatch<React.SetStateAction<{ alreadyRegistered: boolean; notInAudience: boolean; registeredName?: string; registeredDate?: string; }>>;
  setUserResults: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; location: string; jobTitle: string; }[]>>;
  setUserSearch: React.Dispatch<React.SetStateAction<string>>;
  setUserSearchIncludeIntl: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../context/DialogContext").AlertOptions) => void;
  surname: string;
  t: (key: string) => string;
  thirdPartyCheck: { alreadyRegistered: boolean; notInAudience: boolean; registeredName?: string; registeredDate?: string; };
  userResults: { email: string; displayName: string; location: string; jobTitle: string; }[];
  userSearch: string;
  userSearchIncludeIntl: boolean;
}
export const ProxyWizardModal: React.FC<ProxyWizardModalProps> = (p) => {
  const { canCreateEvents, checkRegistrationByEmail, currentUser, email, event, externalEmailConfirmedRef, externalPerson, firstName, isAllowedTargetForAssistant, isAssistant, isSearchingUser, locale, otherConsentConfirmed, pickedUserProfile, pickProxyUser, proxyStep, searchTimerRef, searchUsers, setEmail, setExternalPerson, setFirstName, setIsSearchingUser, setMassImportOpen, setMassImportResult, setMassImportRows, setMassImportStep, setOtherConsentConfirmed, setPickedUserProfile, setProxyStep, setRegisterForOther, setSurname, setThirdPartyCheck, setUserResults, setUserSearch, setUserSearchIncludeIntl, showAlert, surname, t, thirdPartyCheck, userResults, userSearch, userSearchIncludeIntl } = p;
        const cancelWizard = (): void => {
          setRegisterForOther(false);
          setProxyStep(0);
          setFirstName(currentUser.firstName); setSurname(currentUser.surname); setEmail(currentUser.email);
          setUserSearch(''); setUserResults([]); setPickedUserProfile(null);
          setThirdPartyCheck(null); setOtherConsentConfirmed(false); setExternalPerson(false);
        };
        const clearPick = (): void => {
          setFirstName(''); setSurname(''); setEmail(''); setUserSearch(''); setUserResults([]);
          setThirdPartyCheck(null); setPickedUserProfile(null);
        };
        const linkBtn: React.CSSProperties = { background: 'none', border: 'none', padding: 0, color: 'var(--dex-blue, #0076a8)', textDecoration: 'underline', cursor: 'pointer', fontSize: '0.78rem', fontWeight: 600 };
        const picked = !!email.trim();
        const blocked = !!(thirdPartyCheck && thirdPartyCheck.alreadyRegistered);
        const pName = `${firstName} ${surname}`.trim() || email;
        return (
          <Modal
            open={proxyStep > 0}
            onClose={cancelWizard}
            maxWidth={560}
            padding={24}
            ariaLabel={locale === 'de' ? 'Für eine andere Person anmelden' : 'Register another person'}
          >
            <h3 style={{ margin: '0 0 2px', fontSize: '1.1rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {locale === 'de' ? 'Für eine andere Person anmelden' : 'Register another person'}
            </h3>
            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginBottom: 14 }}>
              {locale === 'de' ? `Schritt ${proxyStep} von 2 — ${proxyStep === 1 ? 'Person suchen' : 'Zustimmung'}` : `Step ${proxyStep} of 2 — ${proxyStep === 1 ? 'find person' : 'consent'}`}
            </div>

            {proxyStep === 1 && (
              <>
                {!externalPerson && !picked && (
                  <div style={{ position: 'relative' }}>
                    <input
                      className="form-input"
                      autoFocus
                      value={userSearch}
                      onChange={e => {
                        const val = e.target.value;
                        setUserSearch(val);
                        if (searchTimerRef.current) clearTimeout(searchTimerRef.current);
                        if (val.length >= 2) {
                          searchTimerRef.current = setTimeout(async () => {
                            setIsSearchingUser(true);
                            const results = await searchUsers(val, userSearchIncludeIntl);
                            setUserResults(results);
                            setIsSearchingUser(false);
                          }, 300);
                        } else { setUserResults([]); }
                      }}
                      placeholder={t('reg.searchplaceholder') || 'Name oder E-Mail eingeben...'}
                    />
                    <InternationalSearchToggle
                      query={userSearch}
                      checked={userSearchIncludeIntl}
                      onChange={async next => {
                        setUserSearchIncludeIntl(next);
                        const val = userSearch.trim();
                        if (val.length >= 2) { setIsSearchingUser(true); try { setUserResults(await searchUsers(val, next)); } catch { /* */ } setIsSearchingUser(false); }
                      }}
                    />
                    {isSearchingUser && <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 8 }}>{locale === 'de' ? 'Wird gesucht…' : 'Searching…'}</p>}
                    {userResults.length > 0 && (
                      <div style={{ maxHeight: 260, overflowY: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 8, marginTop: 8 }}>
                        {userResults.map(u => {
                          const assistantOnly = isAssistant && !canCreateEvents;
                          const targetAllowed = !assistantOnly || isAllowedTargetForAssistant(u.jobTitle);
                          return (
                            <div
                              key={u.email}
                              onClick={() => { if (targetAllowed) pickProxyUser(u); }}
                              title={targetAllowed ? '' : 'Assistants can only register Partners or Directors for events.'}
                              style={{ padding: '8px 12px', cursor: targetAllowed ? 'pointer' : 'not-allowed', opacity: targetAllowed ? 1 : 0.45, borderBottom: '1px solid var(--dex-gray-100)', display: 'flex', alignItems: 'center', gap: 10 }}
                            >
                              <img src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(u.email)}&size=S`} alt={u.displayName} onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }} style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }} />
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontWeight: 600, fontSize: '0.85rem' }}>{u.displayName}</div>
                                <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.email}{u.jobTitle ? ` · ${u.jobTitle}` : ''}{u.location ? ` · ${u.location}` : ''}</div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                    {canCreateEvents && (
                      <div style={{ marginTop: 14, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                        {locale === 'de' ? 'Person außerhalb Deloitte oder mehrere auf einmal? ' : 'External person or several at once? '}
                        <button type="button" style={linkBtn} onClick={() => { setExternalPerson(true); clearPick(); setOtherConsentConfirmed(false); }}>{locale === 'de' ? 'Externe Person' : 'External person'}</button>
                        {' · '}
                        <button type="button" style={linkBtn} onClick={() => { setProxyStep(0); setMassImportResult(null); setMassImportRows([]); setMassImportStep('input'); setMassImportOpen(true); }}>{locale === 'de' ? 'Massenimport' : 'Bulk import'}</button>
                      </div>
                    )}
                  </div>
                )}
                {!externalPerson && picked && (
                  <>
                    <div style={{ padding: '10px 12px', border: '1px solid var(--dex-green, #86bc25)', borderRadius: 8, background: 'rgba(134,188,37,0.06)', display: 'flex', alignItems: 'center', gap: 10 }}>
                      <img src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=S`} alt={pName} onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }} style={{ width: 40, height: 40, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }} />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 700 }}>{pName}</div>
                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email}{pickedUserProfile?.jobTitle ? ` · ${pickedUserProfile.jobTitle}` : ''}</div>
                      </div>
                      <button type="button" style={linkBtn} onClick={clearPick}>{locale === 'de' ? 'Ändern' : 'Change'}</button>
                    </div>
                    {thirdPartyCheck && (thirdPartyCheck.alreadyRegistered || thirdPartyCheck.notInAudience) && (
                      <div style={{ marginTop: 8, padding: '8px 12px', borderRadius: 8, fontSize: '0.8rem', background: thirdPartyCheck.alreadyRegistered ? 'rgba(200,30,30,0.07)' : 'rgba(237,139,0,0.08)', border: `1px solid ${thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)'}`, color: thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)' }}>
                        {thirdPartyCheck.alreadyRegistered
                          ? (locale === 'de' ? 'Diese Person ist bereits für dieses Event angemeldet.' : 'This person is already registered for this event.')
                          : (locale === 'de' ? 'Hinweis: Diese Person ist nicht im Gästekreis dieses Events — die Anmeldung ist trotzdem möglich.' : 'Note: this person is not in this event’s audience — registration is still possible.')}
                      </div>
                    )}
                  </>
                )}
                {/* v26.85: Externe Person direkt IM Wizard erfassen (statt unten
                    im Formular). Vor-/Nachname + E-Mail hier eingeben, „Weiter"
                    führt zur Zustimmung. */}
                {externalPerson && (
                  <div>
                    <div style={{ padding: '10px 12px', marginBottom: 12, borderRadius: 8, background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange, #ed8b00)', fontSize: '0.82rem', color: 'var(--dex-orange-dark, #b35a00)', lineHeight: 1.5 }}>
                      {/* v27.12: Wording-Feinschliff (Feedback Datenschutz-Review). */}
                      {locale === 'de'
                        ? 'Person außerhalb von Deloitte (externe E-Mail-Adresse). Trage Vorname, Nachname und E-Mail-Adresse ein. Nach der Zustimmung meldest du die Person stellvertretend an — die Einladung und die Datenschutz-Rückmeldung laufen anschließend über dich.'
                        : 'Person outside Deloitte (external email address). Enter first name, last name and email address. After consent you register the person on their behalf — the invitation and the privacy confirmation are then handled through you.'}
                    </div>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 10 }}>
                      <div>
                        <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, marginBottom: 3 }}>{t('reg.firstname') || 'Vorname'}</label>
                        <input className="form-input" autoFocus value={firstName} onChange={e => setFirstName(e.target.value)} placeholder={t('reg.firstname') || 'Vorname'} />
                      </div>
                      <div>
                        <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, marginBottom: 3 }}>{t('reg.surname') || 'Nachname'}</label>
                        <input className="form-input" value={surname} onChange={e => setSurname(e.target.value)} placeholder={t('reg.surname') || 'Nachname'} />
                      </div>
                    </div>
                    <div style={{ marginBottom: 6 }}>
                      <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, marginBottom: 3 }}>E-Mail</label>
                      <input className="form-input" type="email" value={email} onChange={e => { setEmail(e.target.value); externalEmailConfirmedRef.current = false; setThirdPartyCheck(null); /* v27.11: Duplikat-Check bei Adress-Änderung zurücksetzen */ }} placeholder="name@firma.de" />
                    </div>
                    <button type="button" style={linkBtn} onClick={() => { setExternalPerson(false); clearPick(); }}>{locale === 'de' ? '← Zurück zur Personensuche' : '← Back to search'}</button>
                  </div>
                )}
                <div style={{ display: 'flex', justifyContent: 'space-between', gap: 10, marginTop: 18 }}>
                  <button type="button" className="btn btn-secondary" onClick={cancelWizard}>{locale === 'de' ? 'Abbrechen' : 'Cancel'}</button>
                  {externalPerson ? (
                    <button
                      type="button"
                      className="btn btn-primary"
                      disabled={!(firstName.trim() && surname.trim() && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email.trim()))}
                      onClick={() => {
                        // v27.11 (Bug „Externe können mehrfach angemeldet
                        // werden"): Duplikat-Check jetzt auch für externe
                        // Personen — vorher lief er NUR beim Personen-Picker
                        // (interne), Externe rutschten ungeprüft durch.
                        // thirdPartyCheck aktiviert zugleich die bestehende
                        // Submit-Sperre + den Button-Disable am Formular.
                        (async () => {
                          const existing = await checkRegistrationByEmail(event.id, email.trim()).catch(() => null);
                          const alreadyRegistered = !!existing && existing.Status !== 'Abgemeldet';
                          setThirdPartyCheck({
                            alreadyRegistered,
                            notInAudience: false,
                            registeredName: (existing && (existing.ParticipantName || `${existing.Vorname || ''} ${existing.Nachname || ''}`.trim())) || `${firstName} ${surname}`.trim(),
                            registeredDate: (existing && existing.RegistrationDate) || '',
                          });
                          if (alreadyRegistered) {
                            showAlert(locale === 'de'
                              ? `${email.trim()} ist bereits für dieses Event angemeldet — eine erneute Anmeldung ist nicht möglich.`
                              : `${email.trim()} is already registered for this event — registering again is not possible.`, { variant: 'error' });
                            return;
                          }
                          setProxyStep(2);
                        })().catch(() => setProxyStep(2));
                      }}
                    >{locale === 'de' ? 'Weiter' : 'Next'}</button>
                  ) : (
                    <button type="button" className="btn btn-primary" disabled={!picked || blocked} onClick={() => setProxyStep(2)}>{locale === 'de' ? 'Weiter' : 'Next'}</button>
                  )}
                </div>
              </>
            )}

            {proxyStep === 2 && (
              <>
                {/* v26.98: Die ausführliche Ablauf-Erklärung lebt jetzt HIER im
                    Wizard-Schritt „Zustimmung" (statt als große Box auf der
                    Anmeldeseite). Auf der Anmeldeseite bleibt danach nur ein
                    kurzer Hinweis + der Pflicht-Haken. */}
                <div style={{ padding: '12px 14px', background: 'rgba(237,139,0,0.10)', border: '2px solid var(--dex-orange, #ed8b00)', borderRadius: 8, color: '#7a4a00', fontSize: '0.86rem', lineHeight: 1.55 }}>
                  {locale === 'de'
                    ? <>Mit dem Absenden meldest du <strong>{pName}</strong> stellvertretend an. Bitte stelle sicher, dass die Person ihrer Anmeldung <strong>vorher zugestimmt</strong> hat — eine Anmeldung ohne Einverständnis ist nicht erlaubt.</>
                    : <>By submitting you register <strong>{pName}</strong> on their behalf. Please make sure the person has <strong>consented up front</strong> — registering people without their consent is not allowed.</>}
                  <div style={{ marginTop: 8 }}>
                    {/* v27.12: Wording-Feinschliff (Feedback Datenschutz-Review). */}
                    {locale === 'de'
                      ? <>Die Person erscheint anschließend regulär in der Teilnehmerliste. Falls sie doch nicht teilnehmen kann, lässt sich die Anmeldung jederzeit stornieren — bitte gib in dem Fall kurz Bescheid, damit Wartelisten-Plätze nachrücken können.</>
                      : <>The person then appears in the participant list as usual. If they are unable to attend after all, the registration can be cancelled at any time — please let us know in that case so waitlist spots can be filled.</>}
                  </div>
                </div>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginTop: 12, cursor: 'pointer' }}>
                  <input type="checkbox" checked={otherConsentConfirmed} onChange={e => setOtherConsentConfirmed(e.target.checked)} style={{ marginTop: 3 }} />
                  <span style={{ flex: 1, color: 'var(--dex-gray-800)' }}>
                    <span style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                    {locale === 'de'
                      ? 'Ich bestätige, dass die Person ihrer stellvertretenden Anmeldung zugestimmt hat.'
                      : 'I confirm that the person has consented to this registration on their behalf.'}
                  </span>
                </label>
                <div style={{ display: 'flex', justifyContent: 'space-between', gap: 10, marginTop: 18 }}>
                  <button type="button" className="btn btn-secondary" onClick={() => setProxyStep(1)}>{locale === 'de' ? 'Zurück' : 'Back'}</button>
                  <button type="button" className="btn btn-primary" disabled={!otherConsentConfirmed} onClick={() => setProxyStep(0)}>{locale === 'de' ? 'OK, Person übernehmen' : 'OK, take over person'}</button>
                </div>
              </>
            )}
          </Modal>
        );
};
