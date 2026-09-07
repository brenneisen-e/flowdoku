/* ProxyWizardModal — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Der gefuehrte Wizard fuer die stellvertretende Anmeldung (v26.76): Schritt 1
 * Person suchen, Schritt 2 Zustimmung. Inhalt zeichengleich uebernommen; die
 * Anzeige-Bedingung (`proxyStep > 0`) ist beim Aufrufer geblieben. */
import * as React from 'react';
import Modal from '../Modal';
import InternationalSearchToggle from '../InternationalSearchToggle';
import { Locale } from '../../context/LanguageContext';
import { DeloitteEvent } from '../../types';
// v31.2: Gemeinsame Klassen (Zeilen, Kacheln, Callouts) statt Inline-Stapel —
// Hover geht nur über Klassen, und die Treffer-Liste hatte bis dahin keinen.
import { cx } from '../dexUi';
import { AlertCircle, Check, ChevronLeft, Pencil, Search, Users } from '../Icons';

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
        const isDe = locale === 'de';
        const picked = !!email.trim();
        const blocked = !!(thirdPartyCheck && thirdPartyCheck.alreadyRegistered);
        const pName = `${firstName} ${surname}`.trim() || email;
        const emailValid = /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email.trim());
        const ellipsis: React.CSSProperties = { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' };
        // v31.2: Ein Foto-Helfer für Treffer, Auswahl und Zusammenfassung — vorher stand derselbe <img> dreimal mit je eigenen Maßen.
        const photoOf = (mail: string, alt: string, lg?: boolean): React.ReactElement => (
          <img className={cx('dex-ui-avatar', lg && 'dex-ui-avatar--lg')} src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(mail)}&size=S`} alt={alt} onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }} />
        );
        // v31.2: Alternative Wege als Kacheln mit einer Zeile Folge statt zweier
        // Textlinks — so sieht man, was der andere Weg bedeutet, bevor man klickt.
        const choice = (title: string, desc: string, onClick: () => void): React.ReactElement => (
          <button type="button" className="dex-ui-choice" onClick={onClick}>
            <span className="dex-ui-choice-body"><span className="dex-ui-choice-title">{title}</span><span className="dex-ui-choice-desc" style={{ display: 'block' }}>{desc}</span></span>
          </button>
        );
        const continueExternal = (): void => {
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
              showAlert(isDe
                ? `${email.trim()} ist bereits für dieses Event angemeldet — eine erneute Anmeldung ist nicht möglich.`
                : `${email.trim()} is already registered for this event — registering again is not possible.`, { variant: 'error' });
              return;
            }
            setProxyStep(2);
          })().catch(() => setProxyStep(2));
        };
        // v31.2: Die Knöpfe sitzen im Modal-Fuß (eine Knopfzeile für beide
        // Schritte, Primär rechts) — vorher hatte jeder Schritt seine eigene.
        const footer = proxyStep === 1 ? (<>
          <button type="button" className="btn btn-secondary" onClick={cancelWizard}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
          <button type="button" className="btn btn-primary" disabled={externalPerson ? !(firstName.trim() && surname.trim() && emailValid) : (!picked || blocked)} onClick={externalPerson ? continueExternal : () => setProxyStep(2)}>{isDe ? 'Weiter' : 'Next'}</button>
        </>) : (<>
          <button type="button" className="btn btn-secondary" onClick={() => setProxyStep(1)}>{isDe ? 'Zurück' : 'Back'}</button>
          <button type="button" className="btn btn-primary" disabled={!otherConsentConfirmed} onClick={() => setProxyStep(0)}><Check size={16} />{isDe ? 'Person übernehmen' : 'Use this person'}</button>
        </>);
        return (
          <Modal
            open={proxyStep > 0}
            onClose={cancelWizard}
            maxWidth={560}
            ariaLabel={isDe ? 'Für eine andere Person anmelden' : 'Register another person'}
            title={isDe ? 'Für eine andere Person anmelden' : 'Register another person'}
            subtitle={isDe
              ? `Schritt ${proxyStep} von 2 · ${proxyStep === 1 ? 'Wen meldest du an?' : 'Hat die Person zugestimmt?'}`
              : `Step ${proxyStep} of 2 · ${proxyStep === 1 ? 'Who are you registering?' : 'Has the person consented?'}`}
            icon={<Users size={20} />}
            footer={footer}
          >
            {/* v31.2: Fortschritt als zwei Schritt-Nummern — erledigt grün mit Haken, offen grau. */}
            <div className="dex-ui-inline" aria-hidden="true" style={{ gap: 18 }}>
              {([1, 2] as const).map(n => (
                <span key={n} style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.8rem', fontWeight: 600, color: proxyStep === n ? 'var(--dex-gray-800)' : 'var(--dex-gray-500)' }}>
                  <span className="dex-ui-step-num" style={proxyStep > n ? { background: 'var(--dex-green-darker, #4a7c1f)' } : proxyStep < n ? { background: 'var(--dex-gray-300)' } : undefined}>{proxyStep > n ? <Check size={14} /> : n}</span>
                  {n === 1 ? (isDe ? 'Person wählen' : 'Pick person') : (isDe ? 'Zustimmung' : 'Consent')}
                </span>
              ))}
            </div>

            {proxyStep === 1 && (
              <>
                {!externalPerson && !picked && (
                  <div>
                    <div className="dex-ui-field">
                      <label className="dex-ui-label" htmlFor="dex-proxy-search">{isDe ? 'Wen möchtest du anmelden?' : 'Who do you want to register?'}</label>
                      <div style={{ position: 'relative' }}>
                        <span aria-hidden="true" style={{ position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', display: 'inline-flex', color: 'var(--dex-gray-400)', pointerEvents: 'none' }}><Search size={16} /></span>
                        <input
                          id="dex-proxy-search"
                          className="dex-ui-input"
                          style={{ paddingLeft: 38 }}
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
                      </div>
                      <InternationalSearchToggle
                        query={userSearch}
                        checked={userSearchIncludeIntl}
                        onChange={async next => {
                          setUserSearchIncludeIntl(next);
                          const val = userSearch.trim();
                          if (val.length >= 2) { setIsSearchingUser(true); try { setUserResults(await searchUsers(val, next)); } catch { /* */ } setIsSearchingUser(false); }
                        }}
                      />
                      <div className="dex-ui-help">{isSearchingUser ? (isDe ? 'Wird gesucht…' : 'Searching…') : (isDe ? 'Name oder E-Mail eingeben — ab zwei Zeichen suchen wir im Deloitte-Verzeichnis.' : 'Type a name or email — from two characters on we search the Deloitte directory.')}</div>
                    </div>
                    {userResults.length > 0 && (
                      <div className="dex-ui-card" style={{ padding: 4, maxHeight: 260, overflowY: 'auto' }}>
                        {userResults.map(u => {
                          const assistantOnly = isAssistant && !canCreateEvents;
                          const targetAllowed = !assistantOnly || isAllowedTargetForAssistant(u.jobTitle);
                          return (
                            <div
                              key={u.email}
                              role="button"
                              tabIndex={targetAllowed ? 0 : -1}
                              aria-disabled={!targetAllowed}
                              className="dex-ui-row"
                              onClick={() => { if (targetAllowed) pickProxyUser(u); }}
                              onKeyDown={e => { if (targetAllowed && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); pickProxyUser(u); } }}
                              title={targetAllowed ? '' : (isDe ? 'Assistenzen dürfen nur Partner oder Directors für Events anmelden.' : 'Assistants can only register Partners or Directors for events.')}
                              style={{ cursor: targetAllowed ? 'pointer' : 'not-allowed', opacity: targetAllowed ? 1 : 0.45 }}
                            >
                              {photoOf(u.email, u.displayName)}
                              <div className="dex-ui-row-main">
                                <div className="dex-ui-row-title">{u.displayName}</div>
                                <div className="dex-ui-row-sub" style={ellipsis}>{u.email}{u.jobTitle ? ` · ${u.jobTitle}` : ''}{u.location ? ` · ${u.location}` : ''}</div>
                              </div>
                              {targetAllowed && <span className="dex-ui-row-actions"><span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'Auswählen' : 'Select'}</span></span>}
                            </div>
                          );
                        })}
                      </div>
                    )}
                    {canCreateEvents && (
                      <div className="dex-ui-section">
                        <div className="dex-ui-section-title">{isDe ? 'Nicht im Verzeichnis?' : 'Not in the directory?'}</div>
                        <div className="dex-ui-grid-2">
                          {choice(isDe ? 'Person außerhalb Deloitte' : 'Person outside Deloitte', isDe ? 'Externe E-Mail-Adresse — du trägst Name und Adresse selbst ein.' : 'External email address — you enter name and address yourself.', () => { setExternalPerson(true); clearPick(); setOtherConsentConfirmed(false); })}
                          {choice(isDe ? 'Mehrere auf einmal' : 'Several at once', isDe ? 'Massenimport: Liste einfügen, jede Zeile wird geprüft.' : 'Bulk import: paste a list, every row gets checked.', () => { setProxyStep(0); setMassImportResult(null); setMassImportRows([]); setMassImportStep('input'); setMassImportOpen(true); })}
                        </div>
                      </div>
                    )}
                  </div>
                )}
                {!externalPerson && picked && (
                  <div>
                    <div className="dex-ui-card dex-ui-card--accent" style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                      {photoOf(email, pName, true)}
                      <div className="dex-ui-row-main">
                        <div className="dex-ui-row-title">{pName}</div>
                        <div className="dex-ui-row-sub" style={ellipsis}>{email}{pickedUserProfile?.jobTitle ? ` · ${pickedUserProfile.jobTitle}` : ''}</div>
                      </div>
                      <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={clearPick}><Pencil size={14} />{isDe ? 'Ändern' : 'Change'}</button>
                    </div>
                    {thirdPartyCheck && (thirdPartyCheck.alreadyRegistered || thirdPartyCheck.notInAudience) && (
                      <div className={cx('dex-ui-callout', thirdPartyCheck.alreadyRegistered ? 'dex-ui-callout--danger' : 'dex-ui-callout--warn')} style={{ marginTop: 10 }}>
                        <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                        <span>{thirdPartyCheck.alreadyRegistered
                          ? (isDe ? 'Diese Person ist bereits für dieses Event angemeldet — ein zweites Mal geht nicht.' : 'This person is already registered for this event — registering twice is not possible.')
                          : (isDe ? 'Diese Person ist nicht im Gästekreis dieses Events — die Anmeldung ist trotzdem möglich.' : 'This person is not in this event’s audience — registration is still possible.')}</span>
                      </div>
                    )}
                  </div>
                )}
                {/* v26.85: Externe Person direkt IM Wizard erfassen (statt unten
                    im Formular). Vor-/Nachname + E-Mail hier eingeben, „Weiter"
                    führt zur Zustimmung. */}
                {externalPerson && (
                  <div className="dex-ui-section">
                    <div className="dex-ui-section-title">{isDe ? 'Person außerhalb von Deloitte' : 'Person outside Deloitte'}</div>
                    {/* v27.12: Wording-Feinschliff (Feedback Datenschutz-Review). */}
                    <div className="dex-ui-section-desc">{isDe
                      ? 'Trage Vorname, Nachname und die externe E-Mail-Adresse ein. Einladung und Datenschutz-Rückmeldung laufen anschließend über dich.'
                      : 'Enter first name, last name and the external email address. The invitation and the privacy confirmation are then handled through you.'}</div>
                    <div className="dex-ui-grid-2">
                      <div className="dex-ui-field" style={{ marginBottom: 0 }}>
                        <label className="dex-ui-label" htmlFor="dex-proxy-first">{t('reg.firstname') || 'Vorname'}</label>
                        <input id="dex-proxy-first" className="dex-ui-input" autoFocus value={firstName} onChange={e => setFirstName(e.target.value)} placeholder={t('reg.firstname') || 'Vorname'} />
                      </div>
                      <div className="dex-ui-field" style={{ marginBottom: 0 }}>
                        <label className="dex-ui-label" htmlFor="dex-proxy-last">{t('reg.surname') || 'Nachname'}</label>
                        <input id="dex-proxy-last" className="dex-ui-input" value={surname} onChange={e => setSurname(e.target.value)} placeholder={t('reg.surname') || 'Nachname'} />
                      </div>
                    </div>
                    <div className="dex-ui-field" style={{ marginTop: 14 }}>
                      <label className="dex-ui-label" htmlFor="dex-proxy-mail">{isDe ? 'E-Mail-Adresse' : 'Email address'}</label>
                      <input id="dex-proxy-mail" className="dex-ui-input" type="email" value={email} onChange={e => { setEmail(e.target.value); externalEmailConfirmedRef.current = false; setThirdPartyCheck(null); /* v27.11: Duplikat-Check bei Adress-Änderung zurücksetzen */ }} placeholder="name@firma.de" />
                      <div className="dex-ui-help">{isDe ? 'An diese Adresse geht die Einladung.' : 'The invitation goes to this address.'}</div>
                    </div>
                    <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" style={{ marginTop: 10 }} onClick={() => { setExternalPerson(false); clearPick(); }}><ChevronLeft size={14} />{isDe ? 'Zurück zur Personensuche' : 'Back to search'}</button>
                  </div>
                )}
              </>
            )}

            {proxyStep === 2 && (
              <>
                {/* v26.98: Die ausführliche Ablauf-Erklärung lebt jetzt HIER im
                    Wizard-Schritt „Zustimmung" (statt als große Box auf der
                    Anmeldeseite). Auf der Anmeldeseite bleibt danach nur ein
                    kurzer Hinweis + der Pflicht-Haken. */}
                {/* v31.2: Zusammenfassung zuerst — wer hier landet, soll sehen, um WEN es geht,
                    bevor er die Zustimmung bestätigt. Externe haben kein Verzeichnis-Foto → Initialen. */}
                <div className="dex-ui-card dex-ui-card--soft" style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                  {externalPerson ? <span className="dex-ui-avatar dex-ui-avatar--lg" aria-hidden="true">{`${firstName.trim().charAt(0)}${surname.trim().charAt(0)}`.toUpperCase() || '?'}</span> : photoOf(email, pName, true)}
                  <div className="dex-ui-row-main">
                    <div className="dex-ui-row-title">{pName}</div>
                    <div className="dex-ui-row-sub" style={ellipsis}>{email}{externalPerson ? (isDe ? ' · extern' : ' · external') : (pickedUserProfile?.jobTitle ? ` · ${pickedUserProfile.jobTitle}` : '')}</div>
                  </div>
                </div>
                <div className="dex-ui-callout dex-ui-callout--warn">
                  <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                  <span>{isDe
                    ? <>Du meldest <strong>{pName}</strong> stellvertretend an. Die Person muss <strong>vorher zugestimmt</strong> haben — eine Anmeldung ohne Einverständnis ist nicht erlaubt.</>
                    : <>You register <strong>{pName}</strong> on their behalf. The person must have <strong>consented up front</strong> — registering people without their consent is not allowed.</>}</span>
                </div>
                <label className={cx('dex-ui-toggle-row', otherConsentConfirmed && 'is-active')}>
                  <input type="checkbox" checked={otherConsentConfirmed} onChange={e => setOtherConsentConfirmed(e.target.checked)} />
                  <span className="dex-ui-toggle-row-body">
                    <span className="dex-ui-toggle-row-title">{isDe ? 'Die Person hat zugestimmt' : 'The person has consented'}<span style={{ color: 'var(--dex-red)' }} aria-hidden="true">*</span></span>
                    <span className="dex-ui-toggle-row-desc">{isDe ? 'Ich bestätige, dass die Person ihrer stellvertretenden Anmeldung zugestimmt hat.' : 'I confirm that the person has consented to this registration on their behalf.'}</span>
                  </span>
                </label>
                {/* v27.12: Wording-Feinschliff (Feedback Datenschutz-Review). */}
                <p className="dex-ui-muted" style={{ margin: 0, lineHeight: 1.5 }}>{isDe
                  ? <>Mit <strong>Person übernehmen</strong> landet sie im Anmeldeformular; abgeschickt wird erst dort. Danach erscheint sie regulär in der Teilnehmerliste. Kann sie doch nicht teilnehmen, lässt sich die Anmeldung jederzeit stornieren — gib dann kurz Bescheid, damit Wartelisten-Plätze nachrücken können.</>
                  : <><strong>Use this person</strong> puts them into the registration form; nothing is sent before you submit there. They then appear in the participant list as usual. If they cannot attend after all, the registration can be cancelled at any time — please let us know in that case so waitlist spots can be filled.</>}</p>
              </>
            )}
          </Modal>
        );
};
