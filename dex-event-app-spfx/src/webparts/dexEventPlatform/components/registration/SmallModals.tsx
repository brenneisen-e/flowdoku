/* Kleine Dialoge der Anmeldeseite — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Vier in sich geschlossene Modals, die nichts miteinander teilen ausser ihrer
 * Groesse: Starter-Typ-Fallback (v6.5), Warnung vor externer Adresse (v9.22),
 * CC-Frage bei stellvertretender Anmeldung (v19.6) und die Assistenz-Abfrage
 * (v24.48). Inhalt zeichengleich uebernommen; die Anzeige-Bedingungen sind beim
 * Aufrufer geblieben. */
import * as React from 'react';
import Modal from '../Modal';
import { UserFieldPicker } from '../UserFieldPicker';
import { Locale } from '../../context/LanguageContext';

/** Wunsch-Starter-Typ voll, Alternative frei — der Teilnehmer entscheidet zwischen Umsteigen und Warteliste (v6.5). */
export interface FallbackDialogModalProps {
  fallbackDialog: { wunsch: string; alt: string; altFree: number; };
  locale: Locale;
  performRegistration: (starterTypeToUse: string) => Promise<void>;
  setFallbackDialog: React.Dispatch<React.SetStateAction<{ wunsch: string; alt: string; altFree: number; }>>;
  setPreferredStarterType: React.Dispatch<React.SetStateAction<string>>;
  splitLabelA: string;
  splitLabelB: string;
}
export const FallbackDialogModal: React.FC<FallbackDialogModalProps> = (p) => {
  const { fallbackDialog, locale, performRegistration, setFallbackDialog, setPreferredStarterType, splitLabelA, splitLabelB } = p;
  return (
        <Modal
          open={true}
          onClose={() => setFallbackDialog(null)}
          maxWidth={480}
          padding={24}
          ariaLabel="Plätze voll"
        >
            {(() => {
              // v10.20: Label-Mapping für die freie Bezeichnung — wunsch/alt
              // sind interne IDs ('Durchstarter' / 'Funstarter'); die Anzeige
              // nimmt splitLabelA / splitLabelB.
              const wunschLabel = fallbackDialog.wunsch === 'Durchstarter' ? splitLabelA : splitLabelB;
              const altLabel = fallbackDialog.alt === 'Durchstarter' ? splitLabelA : splitLabelB;
              return (
                <>
                  {/* v17.22: Attendee-facing → bilingual. Vorher war dieser
                      Fallback-Dialog (Wunsch-Gruppe voll) rein deutsch. */}
                  <h3 style={{ margin: 0, marginBottom: 10 }}>
                    {locale === 'de' ? `${wunschLabel}-Plätze sind voll` : `${wunschLabel} is full`}
                  </h3>
                  <p style={{ color: 'var(--dex-gray-700)', lineHeight: 1.5, marginBottom: 8 }}>
                    {locale === 'de'
                      ? <>Für <strong>{wunschLabel}</strong> gibt es aktuell keine freien Plätze mehr.</>
                      : <>There are currently no free spots left for <strong>{wunschLabel}</strong>.</>}
                  </p>
                  <p style={{ color: 'var(--dex-gray-700)', lineHeight: 1.5, marginBottom: 20 }}>
                    {locale === 'de'
                      ? <>Es sind allerdings noch <strong>{fallbackDialog.altFree}</strong> Plätze als <strong>{altLabel}</strong> frei. Möchtest du stattdessen als <strong>{altLabel}</strong> starten?</>
                      : <>However, there are still <strong>{fallbackDialog.altFree}</strong> spots available as <strong>{altLabel}</strong>. Would you like to join as <strong>{altLabel}</strong> instead?</>}
                  </p>
                  <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                    <button
                      className="btn btn-secondary"
                      style={{ fontSize: '0.9rem' }}
                      onClick={async () => {
                        const wunsch = fallbackDialog.wunsch;
                        setFallbackDialog(null);
                        // Wunsch beibehalten → landet auf Warteliste für den Wunsch-Typ.
                        await performRegistration(wunsch);
                      }}
                    >
                      {locale === 'de' ? `Auf ${wunschLabel}-Warteliste` : `Join ${wunschLabel} waitlist`}
                    </button>
                    <button
                      className="btn btn-primary"
                      style={{ fontSize: '0.9rem' }}
                      onClick={async () => {
                        const alt = fallbackDialog.alt;
                        setFallbackDialog(null);
                        // Preferred auf den Alt-Typ setzen, damit sowohl Anzeige
                        // als auch das Register-Payload den neuen Wunsch nutzen.
                        setPreferredStarterType(alt);
                        await performRegistration(alt);
                      }}
                    >
                      {locale === 'de' ? `Als ${altLabel} starten` : `Join as ${altLabel}`}
                    </button>
                  </div>
                </>
              );
            })()}
        </Modal>
  );
};

/** Warnung vor einer externen E-Mail-Adresse (v9.22). */
export interface ExternalEmailWarningModalProps {
  email: string;
  externalEmailConfirmedRef: React.MutableRefObject<boolean>;
  externalEmailWarning: boolean;
  handleSubmit: () => Promise<void>;
  locale: Locale;
  setExternalEmailWarning: React.Dispatch<React.SetStateAction<boolean>>;
}
export const ExternalEmailWarningModal: React.FC<ExternalEmailWarningModalProps> = (p) => {
  const { email, externalEmailConfirmedRef, externalEmailWarning, handleSubmit, locale, setExternalEmailWarning } = p;
  return (
        <Modal
          open={externalEmailWarning}
          onClose={() => setExternalEmailWarning(false)}
          maxWidth={540}
          padding={24}
          ariaLabel={locale === 'de' ? 'E-Mail-Adresse prüfen' : 'Check the email address'}
        >
            {/* v18.74: Tippfehler-Gegenlesen — die externe Adresse groß
                anzeigen und zur Bestätigung auffordern. */}
            <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem', color: 'var(--dex-orange-dark, #b35a00)' }}>
              {locale === 'de' ? 'E-Mail-Adresse prüfen' : 'Check the email address'}
            </h3>
            <p style={{ margin: '0 0 10px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
              {locale === 'de'
                ? <>Du meldest eine <strong>externe Person</strong> an. Bitte lies die Adresse genau gegen — an externe Adressen lässt sich ein <strong>Tippfehler nachträglich nicht korrigieren</strong>:</>
                : <>You are registering an <strong>external person</strong>. Please read the address carefully — a <strong>typo cannot be corrected afterwards</strong> for external addresses:</>}
            </p>
            <div style={{
              margin: '0 0 12px', padding: '12px 14px', textAlign: 'center',
              background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)',
              borderRadius: 8, fontSize: '1.05rem', fontWeight: 700, wordBreak: 'break-all',
              color: 'var(--dex-gray-900, #222)',
            }}>
              {email}
            </div>
            <p style={{ margin: '0 0 12px', fontSize: '0.82rem', lineHeight: 1.55, color: 'var(--dex-gray-600)' }}>
              {locale === 'de'
                ? <>Die <strong>Anmeldebestätigung</strong> geht direkt an diese Adresse, mit den <strong>Organizern auf CC</strong>. Ein <strong>Outlook-Termin</strong> wird an externe Adressen nicht versendet.</>
                : <>The <strong>confirmation email</strong> is sent directly to this address, with the <strong>organizers on CC</strong>. An <strong>Outlook invite</strong> is not sent to external addresses.</>}
            </p>
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setExternalEmailWarning(false)}
                style={{ fontSize: '0.85rem' }}
              >
                {locale === 'de' ? 'Zurück, korrigieren' : 'Back, edit'}
              </button>
              <button
                className="btn btn-primary"
                onClick={() => {
                  externalEmailConfirmedRef.current = true;
                  setExternalEmailWarning(false);
                  // Re-trigger submit via short timeout
                  setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
                }}
                style={{ fontSize: '0.85rem' }}
              >
                {locale === 'de' ? 'Adresse ist korrekt' : 'Address is correct'}
              </button>
            </div>
        </Modal>
  );
};

/** CC-Frage bei stellvertretender interner Anmeldung (v19.6). */
export interface CcSelfModalProps {
  ccSelfDecidedRef: React.MutableRefObject<boolean>;
  ccSelfModalOpen: boolean;
  ccSelfRef: React.MutableRefObject<boolean>;
  firstName: string;
  handleSubmit: () => Promise<void>;
  locale: Locale;
  setCcSelfModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  surname: string;
}
export const CcSelfModal: React.FC<CcSelfModalProps> = (p) => {
  const { ccSelfDecidedRef, ccSelfModalOpen, ccSelfRef, firstName, handleSubmit, locale, setCcSelfModalOpen, surname } = p;
  return (
        <Modal
          open={ccSelfModalOpen}
          onClose={() => setCcSelfModalOpen(false)}
          maxWidth={520}
          padding={24}
          ariaLabel={locale === 'de' ? 'Auf Kopie der Bestätigung?' : 'Copy on the confirmation?'}
        >
          <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
            {locale === 'de' ? 'Möchtest du eine Kopie der Bestätigung?' : 'Do you want a copy of the confirmation?'}
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
            {locale === 'de'
              ? <>Du meldest {`${firstName} ${surname}`.trim() ? <strong>{`${firstName} ${surname}`.trim()}</strong> : <>die ausgewählte Person</>} stellvertretend an. Möchtest du selbst auf <strong>CC der Bestätigungs-Mail</strong> gesetzt werden? Du bekommst dann eine Kopie der Anmeldebestätigung.<br /><br />Der <strong>Outlook-Termin</strong> wird davon nicht berührt — die CC gilt nur für die Bestätigungs-Mail.</>
              : <>You are registering {`${firstName} ${surname}`.trim() ? <strong>{`${firstName} ${surname}`.trim()}</strong> : <>the selected person</>} on their behalf. Would you like to be added to the <strong>CC of the confirmation email</strong>? You will then receive a copy of the confirmation.<br /><br />The <strong>Outlook invite</strong> is not affected — the CC only applies to the confirmation email.</>}
          </p>
          <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
            <button
              className="btn btn-secondary"
              onClick={() => {
                ccSelfRef.current = false;
                ccSelfDecidedRef.current = true;
                setCcSelfModalOpen(false);
                setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
              }}
              style={{ fontSize: '0.85rem' }}
            >
              {locale === 'de' ? 'Nein, ohne CC' : 'No, without CC'}
            </button>
            <button
              className="btn btn-primary"
              onClick={() => {
                ccSelfRef.current = true;
                ccSelfDecidedRef.current = true;
                setCcSelfModalOpen(false);
                setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
              }}
              style={{ fontSize: '0.85rem' }}
            >
              {locale === 'de' ? 'Ja, mich auf CC setzen' : 'Yes, add me to CC'}
            </button>
          </div>
        </Modal>
  );
};

/** Assistenz-Abfrage nach dem Register-Klick fuer Partner/Director (v24.48). */
export interface AssistantModalProps {
  assistantModalDecidedRef: React.MutableRefObject<boolean>;
  assistantModalOpen: boolean;
  delegateAssistValue: string;
  delegateChoiceRef: React.MutableRefObject<{ enabled: boolean; value: string; }>;
  handleSubmit: () => Promise<void>;
  locale: Locale;
  parsedDelegateAssist: { name: string; email: string; };
  searchUser: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string; company?: string; }>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  setAssistantModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setDelegateAssistEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setDelegateAssistValue: React.Dispatch<React.SetStateAction<string>>;
}
export const AssistantModal: React.FC<AssistantModalProps> = (p) => {
  const { assistantModalDecidedRef, assistantModalOpen, delegateAssistValue, delegateChoiceRef, handleSubmit, locale, parsedDelegateAssist, searchUser, searchUsers, setAssistantModalOpen, setDelegateAssistEnabled, setDelegateAssistValue } = p;
  return (
        <Modal
          open={assistantModalOpen}
          onClose={() => setAssistantModalOpen(false)}
          maxWidth={560}
          padding={24}
          ariaLabel={locale === 'de' ? 'Assistenz informieren?' : 'Inform assistant?'}
        >
          <div style={{ display: 'inline-block', background: 'var(--dex-green, #86bc25)', color: '#fff', fontSize: '0.72rem', fontWeight: 700, padding: '4px 12px', borderRadius: 999, marginBottom: 12, letterSpacing: 0.4 }}>
            {locale === 'de' ? 'Für Partner & Directoren' : 'For Partners & Directors'}
          </div>
          <h3 style={{ margin: '0 0 10px', fontSize: '1.1rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
            {locale === 'de' ? 'Möchtest du deine Assistenz informieren?' : 'Do you want to inform your assistant?'}
          </h3>
          <p style={{ margin: '0 0 14px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
            {locale === 'de'
              ? 'Deine Assistenz bekommt eine Kopie der Bestätigung und sieht deine Anmeldung in der App — so bleibt sie auf dem Laufenden.'
              : 'Your assistant gets a copy of the confirmation and can see your registration in the app — so they stay in the loop.'}
          </p>
          <label style={{ display: 'block', fontWeight: 600, fontSize: '0.85rem', marginBottom: 6 }}>
            {locale === 'de' ? 'Assistenz auswählen (Name oder E-Mail)' : 'Select assistant (name or email)'}
          </label>
          <UserFieldPicker
            value={delegateAssistValue}
            onChange={setDelegateAssistValue}
            searchUsers={searchUsers}
            searchUserByEmail={searchUser}
            placeholder={locale === 'de' ? 'Vorname Nachname, Nachname Vorname oder E-Mail…' : 'First last, last first or email…'}
            errorStyle={{}}
            forcedIsDe={locale === 'de'}
          />
          <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap', marginTop: 18 }}>
            <button
              className="btn btn-secondary"
              onClick={() => {
                // Ohne Assistenz weiter.
                delegateChoiceRef.current = { enabled: false, value: '' };
                setDelegateAssistEnabled(false);
                setDelegateAssistValue('');
                assistantModalDecidedRef.current = true;
                setAssistantModalOpen(false);
                setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
              }}
              style={{ fontSize: '0.85rem' }}
            >
              {locale === 'de' ? 'Ohne Assistenz anmelden' : 'Register without assistant'}
            </button>
            <button
              className="btn btn-primary"
              disabled={!parsedDelegateAssist}
              title={!parsedDelegateAssist ? (locale === 'de' ? 'Bitte zuerst eine Assistenz auswählen.' : 'Please select an assistant first.') : ''}
              onClick={() => {
                delegateChoiceRef.current = { enabled: true, value: delegateAssistValue };
                setDelegateAssistEnabled(true);
                assistantModalDecidedRef.current = true;
                setAssistantModalOpen(false);
                setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
              }}
              style={{ fontSize: '0.85rem' }}
            >
              {locale === 'de' ? 'Mit Assistenz anmelden' : 'Register with assistant'}
            </button>
          </div>
        </Modal>
  );
};
