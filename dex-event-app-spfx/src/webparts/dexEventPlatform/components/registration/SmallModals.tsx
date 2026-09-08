/* Kleine Dialoge der Anmeldeseite — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Vier in sich geschlossene Modals, die nichts miteinander teilen ausser ihrer
 * Groesse: Starter-Typ-Fallback (v6.5), Warnung vor externer Adresse (v9.22),
 * CC-Frage bei stellvertretender Anmeldung (v19.6) und die Assistenz-Abfrage
 * (v24.48). Inhalt zeichengleich uebernommen; die Anzeige-Bedingungen sind beim
 * Aufrufer geblieben.
 *
 * v31.2: Alle vier nutzen Kopf (title/subtitle/icon) und Fuß des `Modal` statt
 * eigener <h3>- und Knopfzeilen — die Frage ist jeweils der Titel, der Anlass
 * der Untertitel, die Folge ein Hinweiskasten. Handler und Texte unverändert. */
import * as React from 'react';
import Modal from '../Modal';
import { UserFieldPicker } from '../UserFieldPicker';
import { Locale } from '../../context/LanguageContext';
import { AlertCircle, Info, Mail, Users } from '../Icons';

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
  const isDe = locale === 'de';
  // v10.20: Label-Mapping für die freie Bezeichnung — wunsch/alt sind interne
  // IDs ('Durchstarter' / 'Funstarter'); die Anzeige nimmt splitLabelA / splitLabelB.
  const wunschLabel = fallbackDialog.wunsch === 'Durchstarter' ? splitLabelA : splitLabelB;
  const altLabel = fallbackDialog.alt === 'Durchstarter' ? splitLabelA : splitLabelB;
  // v17.22: Attendee-facing → bilingual. Vorher war dieser Fallback-Dialog
  // (Wunsch-Gruppe voll) rein deutsch.
  // v31.2: Zwei Wege mit unterschiedlicher Folge (sofort dabei vs. warten) —
  // deshalb Auswahl-Kacheln mit je einer Zeile Konsequenz statt zweier Knöpfe,
  // bei denen man die Folge erraten muss. Der Klick löst weiterhin sofort aus.
  return (
    <Modal
      open={true}
      onClose={() => setFallbackDialog(null)}
      maxWidth={520}
      ariaLabel="Plätze voll"
      title={isDe ? `${wunschLabel}-Plätze sind voll` : `${wunschLabel} is full`}
      subtitle={isDe
        ? <>Für <strong>{wunschLabel}</strong> gibt es aktuell keine freien Plätze mehr — als <strong>{altLabel}</strong> sind noch <strong>{fallbackDialog.altFree}</strong> frei. Wie möchtest du weitermachen?</>
        : <>There are currently no free spots left for <strong>{wunschLabel}</strong> — <strong>{fallbackDialog.altFree}</strong> are still available as <strong>{altLabel}</strong>. How do you want to continue?</>}
      icon={<Users size={20} />}
    >
      <div className="dex-ui-grid-2">
        <button
          type="button"
          className="dex-ui-choice"
          onClick={async () => {
            const alt = fallbackDialog.alt;
            setFallbackDialog(null);
            // Preferred auf den Alt-Typ setzen, damit sowohl Anzeige
            // als auch das Register-Payload den neuen Wunsch nutzen.
            setPreferredStarterType(alt);
            await performRegistration(alt);
          }}
        >
          <span className="dex-ui-choice-body">
            <span className="dex-ui-choice-title">{isDe ? `Als ${altLabel} starten` : `Join as ${altLabel}`}</span>
            <span className="dex-ui-choice-desc" style={{ display: 'block' }}>
              {isDe ? 'Du bist sofort fest angemeldet — auf einem der freien Plätze.' : 'You are registered right away — on one of the free spots.'}
            </span>
          </span>
          <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'Sofort dabei' : 'In right away'}</span>
        </button>
        <button
          type="button"
          className="dex-ui-choice"
          onClick={async () => {
            const wunsch = fallbackDialog.wunsch;
            setFallbackDialog(null);
            // Wunsch beibehalten → landet auf Warteliste für den Wunsch-Typ.
            await performRegistration(wunsch);
          }}
        >
          <span className="dex-ui-choice-body">
            <span className="dex-ui-choice-title">{isDe ? `Auf ${wunschLabel}-Warteliste` : `Join ${wunschLabel} waitlist`}</span>
            <span className="dex-ui-choice-desc" style={{ display: 'block' }}>
              {isDe ? `Du bleibst bei ${wunschLabel} und rückst nach, sobald dort ein Platz frei wird.` : `You stay with ${wunschLabel} and move up as soon as a spot opens there.`}
            </span>
          </span>
          <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'Warten' : 'Wait'}</span>
        </button>
      </div>
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
  const isDe = locale === 'de';
  // v18.74: Tippfehler-Gegenlesen — die externe Adresse groß anzeigen und zur
  // Bestätigung auffordern. v31.2: Kopf/Fuß aus `Modal`; die Warnung (nicht
  // korrigierbar) steht im Untertitel, die Adresse als einzige große Fläche.
  return (
    <Modal
      open={externalEmailWarning}
      onClose={() => setExternalEmailWarning(false)}
      maxWidth={540}
      ariaLabel={isDe ? 'E-Mail-Adresse prüfen' : 'Check the email address'}
      title={isDe ? 'E-Mail-Adresse prüfen' : 'Check the email address'}
      subtitle={isDe
        ? <>Du meldest eine <strong>externe Person</strong> an. Lies die Adresse genau gegen — an externe Adressen lässt sich ein <strong>Tippfehler nachträglich nicht korrigieren</strong>.</>
        : <>You are registering an <strong>external person</strong>. Read the address carefully — a <strong>typo cannot be corrected afterwards</strong> for external addresses.</>}
      icon={<AlertCircle size={20} />}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={() => setExternalEmailWarning(false)}>
          {isDe ? 'Zurück, korrigieren' : 'Back, edit'}
        </button>
        <button
          type="button"
          className="btn btn-primary"
          onClick={() => {
            externalEmailConfirmedRef.current = true;
            setExternalEmailWarning(false);
            // Re-trigger submit via short timeout
            setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
          }}
        >
          {isDe ? 'Adresse ist korrekt' : 'Address is correct'}
        </button>
      </>}
    >
      <div className="dex-ui-card dex-ui-card--soft" style={{ textAlign: 'center', fontSize: '1.05rem', fontWeight: 700, wordBreak: 'break-all', color: 'var(--dex-gray-800, #333)' }}>
        {email}
      </div>
      <div className="dex-ui-callout dex-ui-callout--neutral">
        <span className="dex-ui-callout-icon"><Mail size={16} /></span>
        <span>
          {isDe
            ? <>Die <strong>Anmeldebestätigung</strong> geht direkt an diese Adresse, mit den <strong>Organizern auf CC</strong>. Ein <strong>Outlook-Termin</strong> wird an externe Adressen nicht versendet.</>
            : <>The <strong>confirmation email</strong> is sent directly to this address, with the <strong>organizers on CC</strong>. An <strong>Outlook invite</strong> is not sent to external addresses.</>}
        </span>
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
  const isDe = locale === 'de';
  const personName = `${firstName} ${surname}`.trim();
  const person = personName ? <strong>{personName}</strong> : <>{isDe ? 'die ausgewählte Person' : 'the selected person'}</>;
  // v31.2: Die Frage ist der Titel, der Anlass (stellvertretende Anmeldung) der
  // Untertitel; die Folge („Outlook bleibt unberührt") steht als Hinweiskasten.
  return (
    <Modal
      open={ccSelfModalOpen}
      onClose={() => setCcSelfModalOpen(false)}
      maxWidth={520}
      ariaLabel={isDe ? 'Auf Kopie der Bestätigung?' : 'Copy on the confirmation?'}
      title={isDe ? 'Möchtest du eine Kopie der Bestätigung?' : 'Do you want a copy of the confirmation?'}
      subtitle={isDe
        ? <>Du meldest {person} stellvertretend an.</>
        : <>You are registering {person} on their behalf.</>}
      icon={<Mail size={20} />}
      footer={<>
        <button
          type="button"
          className="btn btn-secondary"
          onClick={() => {
            ccSelfRef.current = false;
            ccSelfDecidedRef.current = true;
            setCcSelfModalOpen(false);
            setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
          }}
        >
          {isDe ? 'Nein, ohne CC' : 'No, without CC'}
        </button>
        <button
          type="button"
          className="btn btn-primary"
          onClick={() => {
            ccSelfRef.current = true;
            ccSelfDecidedRef.current = true;
            setCcSelfModalOpen(false);
            setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
          }}
        >
          {isDe ? 'Ja, mich auf CC setzen' : 'Yes, add me to CC'}
        </button>
      </>}
    >
      <p style={{ margin: 0, fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
        {isDe
          ? <>Mit <strong>CC der Bestätigungs-Mail</strong> bekommst du selbst eine Kopie der Anmeldebestätigung.</>
          : <>On the <strong>CC of the confirmation email</strong> you receive a copy of the confirmation yourself.</>}
      </p>
      <div className="dex-ui-callout dex-ui-callout--neutral">
        <span className="dex-ui-callout-icon"><Info size={16} /></span>
        <span>
          {isDe
            ? <>Der <strong>Outlook-Termin</strong> wird davon nicht berührt — die CC gilt nur für die Bestätigungs-Mail.</>
            : <>The <strong>Outlook invite</strong> is not affected — the CC only applies to the confirmation email.</>}
        </span>
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
  const isDe = locale === 'de';
  // v31.2: Kopf und Fuß kommen aus `Modal`; die Zielgruppen-Pille steht als
  // Kontext im Untertitel, damit die Frage selbst die Überschrift bleibt.
  return (
    <Modal
      open={assistantModalOpen}
      onClose={() => setAssistantModalOpen(false)}
      maxWidth={560}
      ariaLabel={isDe ? 'Assistenz informieren?' : 'Inform assistant?'}
      title={isDe ? 'Möchtest du deine Assistenz informieren?' : 'Do you want to inform your assistant?'}
      subtitle={<>
        <span className="dex-ui-pill dex-ui-pill--green" style={{ marginRight: 8, verticalAlign: 'middle' }}>
          {isDe ? 'Für Partner & Directoren' : 'For Partners & Directors'}
        </span>
        {isDe
          ? 'Deine Assistenz bekommt eine Kopie der Bestätigung und sieht deine Anmeldung in der App — so bleibt sie auf dem Laufenden.'
          : 'Your assistant gets a copy of the confirmation and can see your registration in the app — so they stay in the loop.'}
      </>}
      icon={<Users size={20} />}
      footer={<>
        <button
          type="button"
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
        >
          {isDe ? 'Ohne Assistenz anmelden' : 'Register without assistant'}
        </button>
        <button
          type="button"
          className="btn btn-primary"
          disabled={!parsedDelegateAssist}
          title={!parsedDelegateAssist ? (isDe ? 'Bitte zuerst eine Assistenz auswählen.' : 'Please select an assistant first.') : ''}
          onClick={() => {
            delegateChoiceRef.current = { enabled: true, value: delegateAssistValue };
            setDelegateAssistEnabled(true);
            assistantModalDecidedRef.current = true;
            setAssistantModalOpen(false);
            setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
          }}
        >
          {isDe ? 'Mit Assistenz anmelden' : 'Register with assistant'}
        </button>
      </>}
    >
      <div className="dex-ui-field">
        <label className="dex-ui-label">{isDe ? 'Wer ist deine Assistenz?' : 'Who is your assistant?'}</label>
        <UserFieldPicker
          value={delegateAssistValue}
          onChange={setDelegateAssistValue}
          searchUsers={searchUsers}
          searchUserByEmail={searchUser}
          placeholder={isDe ? 'Vorname Nachname, Nachname Vorname oder E-Mail…' : 'First last, last first or email…'}
          errorStyle={{}}
          forcedIsDe={isDe}
        />
        <div className="dex-ui-help">
          {isDe
            ? 'Name oder E-Mail eingeben. Ohne Auswahl meldest du dich einfach ohne Assistenz an.'
            : 'Enter a name or email. Without a selection you simply register without an assistant.'}
        </div>
      </div>
    </Modal>
  );
};
