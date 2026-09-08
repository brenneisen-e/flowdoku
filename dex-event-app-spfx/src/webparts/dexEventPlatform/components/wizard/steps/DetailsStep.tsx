/* DetailsStep — aus EventCreationPage.tsx ausgelagert (Zeilen 10111-11152 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; einzige
 * Aenderung ist die Anzeige-Bedingung: aus `currentStep === 1` wurde das Prop `visible`.
 * `visible` schaltet display:none statt unmount — Eingaben ueberleben den
 * Schrittwechsel genauso wie vorher.
 *
 * v31.2: Neu gegliedert nach docs/ui-leitfaden.md. Der Schritt stellt jetzt
 * vier Fragen in der Reihenfolge, in der ein Organizer sie beantwortet:
 * Wer verantwortet das Event? (Pflicht) → Wer hilft mit? (Test-Team,
 * Check-in-Team) → An wen wenden sich Teilnehmer bei Fragen? → Weitere
 * Einstellungen (Anzeige der Organizer, Aufklapper). Die drei fast gleichen
 * Personen-Picker (Suche, Debounce, International-Schalter, Trefferliste)
 * liegen in EINER Hilfskomponente `PersonPicker` — vorher standen sie dreimal
 * im JSX und liefen bei jeder Korrektur auseinander. State, Setter und
 * Handler-Logik sind unverändert; nur Anordnung und Optik sind neu. */
import * as React from 'react';
import { StepBadge } from '../../wizard/StepBadge';
import { InfoTooltip } from '../../InfoTooltip';
import { Users, X, ChevronUp, ChevronDown, Check } from '../../Icons';
import WizardHint from '../../WizardHint';
import { Icon } from '@fluentui/react/lib/Icon';
import InternationalSearchToggle from '../../InternationalSearchToggle';
import OrganizerList from '../../OrganizerList';
import { cx } from '../../dexUi';
export interface DetailsStepProps {
  visible: boolean;
  contactEmail: string;
  contactExpanded: boolean;
  contactInfo: string;
  contactName: string;
  contactOrganizerEmail: string;
  errorBorderStyle: (fieldName: string) => React.CSSProperties;
  hiddenOrganizerEmails: string[];
  hideOrganizer: boolean;
  hideOrganizerIndividualOnly: boolean;
  isDe: boolean;
  isSearchingOrganizer: boolean;
  location: string;
  organizer: string;
  organizerDisplayLarge: boolean;
  organizerEmails: string[];
  organizerIncludeIntl: boolean;
  organizerResults: { email: string; displayName: string; location: string; }[];
  organizerSearch: string;
  organizerTimerRef: React.MutableRefObject<NodeJS.Timeout>;
  qrScannerEmails: string[];
  qrScannerIncludeIntl: boolean;
  qrScannerNames: string[];
  qrScannerResults: { email: string; displayName: string; location: string; }[];
  qrScannerSearch: string;
  qrScannerTimerRef: React.MutableRefObject<NodeJS.Timeout>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  setBulkOrganizerOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setBulkQrScannerOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setBulkTestTeamOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setContactEmail: React.Dispatch<React.SetStateAction<string>>;
  setContactExpanded: React.Dispatch<React.SetStateAction<boolean>>;
  setContactInfo: React.Dispatch<React.SetStateAction<string>>;
  setContactName: React.Dispatch<React.SetStateAction<string>>;
  setContactOrganizerEmail: React.Dispatch<React.SetStateAction<string>>;
  setHideOrganizer: React.Dispatch<React.SetStateAction<boolean>>;
  setHideOrganizerIndividualOnly: React.Dispatch<React.SetStateAction<boolean>>;
  setOrganizer: React.Dispatch<React.SetStateAction<string>>;
  setOrganizerDisplayLarge: React.Dispatch<React.SetStateAction<boolean>>;
  setOrganizerEmails: React.Dispatch<React.SetStateAction<string[]>>;
  setOrganizerIncludeIntl: React.Dispatch<React.SetStateAction<boolean>>;
  setOrganizerResults: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; location: string; }[]>>;
  setOrganizerSearch: React.Dispatch<React.SetStateAction<string>>;
  setQrScannerEmails: React.Dispatch<React.SetStateAction<string[]>>;
  setQrScannerIncludeIntl: React.Dispatch<React.SetStateAction<boolean>>;
  setQrScannerNames: React.Dispatch<React.SetStateAction<string[]>>;
  setQrScannerResults: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; location: string; }[]>>;
  setQrScannerSearch: React.Dispatch<React.SetStateAction<string>>;
  setTestTeamEmails: React.Dispatch<React.SetStateAction<string[]>>;
  setTestTeamIncludeIntl: React.Dispatch<React.SetStateAction<boolean>>;
  setTestTeamNames: React.Dispatch<React.SetStateAction<string[]>>;
  setTestTeamResults: React.Dispatch<React.SetStateAction<{ email: string; displayName: string; location: string; }[]>>;
  setTestTeamSearch: React.Dispatch<React.SetStateAction<string>>;
  startDate: string;
  t: (key: string) => string;
  testTeamEmails: string[];
  testTeamIncludeIntl: boolean;
  testTeamNames: string[];
  testTeamResults: { email: string; displayName: string; location: string; }[];
  testTeamSearch: string;
  testTeamTimerRef: React.MutableRefObject<NodeJS.Timeout>;
  title: string;
  toggleOrganizerHidden: (email: string) => void;
}

type PersonHit = { email: string; displayName: string; location: string; };

/* v31.2: Runde Foto-Kachel mit Initialen dahinter — lädt das SharePoint-Foto
 * nicht (externe Adresse, kein Foto), bleiben die Initialen stehen statt
 * eines leeren Kreises. Auf Modul-Ebene, damit React die Kachel beim
 * Tippen im Suchfeld nicht neu aufbaut. */
const PersonAvatar: React.FC<{ email: string; name: string; muted?: boolean }> = ({ email, name, muted }) => {
  const initials = name.split(/\s+/).filter(Boolean).slice(0, 2).map(s => s.charAt(0)).join('').toUpperCase();
  return (
    <span className="dex-ui-avatar" aria-hidden="true" style={{ position: 'relative', filter: muted ? 'grayscale(1)' : undefined, opacity: muted ? 0.7 : 1 }}>
      {initials}
      {email ? (
        <img
          src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=S`}
          alt=""
          onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
          style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', objectFit: 'cover' }}
        />
      ) : null}
    </span>
  );
};

/* v31.2: „Massenimport" als Textknopf in der Beschriftungszeile — dreimal
 * gleich, deshalb einmal definiert. Bisher stand er nur deutsch da.
 * Leitfaden 2a′: Der Knopf folgt der Beschriftung direkt (Label-gap 6 + 4 px
 * = 10 px), statt mit margin-left:auto allein am rechten Rand zu hängen —
 * „dann ist nur der Button rechts" war die Nutzer-Kritik. */
const BulkImportButton: React.FC<{ isDe: boolean; onClick: () => void }> = ({ isDe, onClick }) => (
  <button
    type="button"
    className="dex-ui-textbtn"
    style={{ marginLeft: 4, whiteSpace: 'nowrap' }}
    onClick={onClick}
    title={isDe ? 'Mehrere Personen auf einmal einfügen (Liste von E-Mail-Adressen)' : 'Add several people at once (list of email addresses)'}
  >
    <Users size={14} /> {isDe ? 'Massenimport' : 'Bulk import'}
  </button>
);

interface PersonPickerProps {
  isDe: boolean;
  value: string;
  setValue: React.Dispatch<React.SetStateAction<string>>;
  timerRef: React.MutableRefObject<NodeJS.Timeout>;
  results: PersonHit[];
  setResults: React.Dispatch<React.SetStateAction<PersonHit[]>>;
  includeIntl: boolean;
  setIncludeIntl: React.Dispatch<React.SetStateAction<boolean>>;
  searchUsers: DetailsStepProps['searchUsers'];
  placeholder: string;
  inputStyle?: React.CSSProperties;
  searching?: boolean;
  /** Treffer, der schon in der Liste steht — wird gedämpft gezeigt und ist nicht wählbar. */
  isAdded: (hit: PersonHit) => boolean;
  onPick: (hit: PersonHit) => void;
}

/* v31.2: EIN Personen-Picker für Organizer, Test-Team und Check-in-Team.
 * Verhalten wie bisher an allen drei Stellen: 350 ms Debounce auf die
 * Graph-Suche (v9.18/v9.20 — jeder Deloitte-User ist wählbar), Trefferliste
 * schließt 150 ms nach dem Verlassen des Feldes (damit der onMouseDown auf
 * einen Treffer noch greift), der International-Schalter sucht sofort neu.
 * Keine Hooks — der Timer kommt vom Aufrufer, wie vorher. */
const PersonPicker: React.FC<PersonPickerProps> = ({ isDe, value, setValue, timerRef, results, setResults, includeIntl, setIncludeIntl, searchUsers, placeholder, inputStyle, searching, isAdded, onPick }) => (
  <div style={{ position: 'relative' }}>
    {/* v31.2: Suchfeld als dex-ui-input in voller Breite — es ist eine
        Personensuche in einer Karte, kein 48-px-Hauptfeld des Formulars. */}
    <input
      className="dex-ui-input"
      value={value}
      onChange={e => {
        const val = e.target.value;
        setValue(val);
        if (timerRef.current) clearTimeout(timerRef.current);
        const q = val.trim();
        if (!q) { setResults([]); return; }
        timerRef.current = setTimeout(async () => {
          try {
            const found = await searchUsers(q, includeIntl);
            setResults(found.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
          } catch { setResults([]); }
        }, 350);
      }}
      onBlur={() => {
        setTimeout(() => { setValue(''); setResults([]); }, 150);
      }}
      placeholder={placeholder}
      style={inputStyle}
    />
    <InternationalSearchToggle
      query={value}
      checked={includeIntl}
      onChange={async next => {
        setIncludeIntl(next);
        const q = value.trim();
        if (q.length >= 1) {
          try {
            const found = await searchUsers(q, next);
            setResults(found.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
          } catch { setResults([]); }
        }
      }}
      isDe={isDe}
    />
    {searching && (
      <div className="dex-ui-muted" style={{ marginTop: 4 }}>{isDe ? 'Suche …' : 'Searching …'}</div>
    )}
    {results.length > 0 && (
      <div
        className="dex-ui-card"
        style={{ position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100, padding: 6, boxShadow: '0 8px 24px rgba(0,0,0,0.10)', maxHeight: 280, overflowY: 'auto' }}
      >
        {results.map(u => {
          const added = isAdded(u);
          return (
            <div
              key={u.email}
              className="dex-ui-row"
              style={{ cursor: added ? 'not-allowed' : 'pointer', opacity: added ? 0.45 : 1 }}
              onMouseDown={() => { if (added) return; onPick(u); }}
            >
              <PersonAvatar email={u.email} name={u.displayName} />
              <div className="dex-ui-row-main">
                <div className="dex-ui-row-title">{u.displayName}</div>
                <div className="dex-ui-row-sub" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {u.email}{u.location ? ` · ${u.location}` : ''}
                </div>
              </div>
              {added && (
                <span className="dex-ui-pill dex-ui-pill--green"><Check size={12} /> {isDe ? 'schon dabei' : 'already added'}</span>
              )}
            </div>
          );
        })}
      </div>
    )}
  </div>
);

export const DetailsStep: React.FC<DetailsStepProps> = (p) => {
  const { visible } = p;
  const { contactEmail, contactExpanded, contactInfo, contactName, contactOrganizerEmail, errorBorderStyle, hiddenOrganizerEmails, hideOrganizer, hideOrganizerIndividualOnly, isDe, isSearchingOrganizer, location, organizer, organizerDisplayLarge, organizerEmails, organizerIncludeIntl, organizerResults, organizerSearch, organizerTimerRef, qrScannerEmails, qrScannerIncludeIntl, qrScannerNames, qrScannerResults, qrScannerSearch, qrScannerTimerRef, searchUsers, setBulkOrganizerOpen, setBulkQrScannerOpen, setBulkTestTeamOpen, setContactEmail, setContactExpanded, setContactInfo, setContactName, setContactOrganizerEmail, setHideOrganizer, setHideOrganizerIndividualOnly, setOrganizer, setOrganizerDisplayLarge, setOrganizerEmails, setOrganizerIncludeIntl, setOrganizerResults, setOrganizerSearch, setQrScannerEmails, setQrScannerIncludeIntl, setQrScannerNames, setQrScannerResults, setQrScannerSearch, setTestTeamEmails, setTestTeamIncludeIntl, setTestTeamNames, setTestTeamResults, setTestTeamSearch, startDate, t, testTeamEmails, testTeamIncludeIntl, testTeamNames, testTeamResults, testTeamSearch, testTeamTimerRef, title, toggleOrganizerHidden } = p;
  // v31.2: Der Aufklapper „Weitere Einstellungen" (Anzeige der Organizer) ist
  // reine Ansichts-Sache dieses Schritts und wird nicht gespeichert — deshalb
  // lokaler State und kein Prop. Standard zu (Leitfaden 1.6); der Zähler am
  // Knopf zeigt, ob darin etwas vom Standard abweicht.
  const [moreOpen, setMoreOpen] = React.useState<boolean>(false);
  const orgList = organizer.split(';').map(s => s.trim()).filter(Boolean);
  const contactLc = (contactOrganizerEmail || '').toLowerCase();
  const changedMore = (hideOrganizer ? 1 : 0) + (organizerDisplayLarge ? 1 : 0);
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
              <h2 className="dex-step-head-title">
                <span className="dex-step-eyebrow">{isDe ? 'Schritt 2 von 9' : 'Step 2 of 9'}</span>
                {isDe ? 'Organizer & Team' : 'Organizers & team'}
              </h2>
              <p className="dex-step-head-lead">
                {isDe
                  ? 'Wer verantwortet das Event, wer hilft mit — und an wen wenden sich Teilnehmer bei Fragen? Pflicht ist nur der Organizer.'
                  : 'Who runs the event, who helps out — and who do attendees contact with questions? Only the organizer is required.'}
              </p>

              {/* ---- Frage 1: Wer verantwortet das Event? (Pflicht) ------------- */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Wer verantwortet das Event?' : 'Who is responsible for the event?'}</div>
                <p className="dex-ui-section-desc">
                  {isDe
                    ? 'Organizer können das Event bearbeiten, die Teilnehmerliste einsehen und Mails versenden. Die erste Person in der Liste ist Haupt-Organizer und steht in Mails als Absender-Name.'
                    : 'Organizers can edit the event, see the attendee list and send mails. The first person in the list is the main organizer and appears as the sender name in mails.'}
                </p>
                {/* v31.2: flexWrap — bei schmaler Breite bricht der Massenimport-Knopf
                    linksbündig unter die Beschriftung, statt sie zu quetschen. */}
                <div className="dex-ui-label" style={{ flexWrap: 'wrap' }}>
                  <StepBadge n={11} />
                  <span className="required">*</span> {t('create.organizer')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> die <strong>verantwortlichen Personen</strong> für dieses Event — beliebige Deloitte-User per Graph-Suche. Du selbst bist standardmäßig vorbefüllt, kannst aber Co-Organizer hinzunehmen oder dich selbst rauslöschen.<br /><br />
                      <strong>Anzeige in der App:</strong> Organizer dürfen das Event <strong>bearbeiten, deaktivieren, löschen</strong>, die <strong>Teilnehmerliste</strong> einsehen, <strong>QR-Codes versenden</strong> und <strong>Massenmails</strong> verschicken. Sie tauchen auf der Anmelde-Seite und in Meine Events als <strong>Ansprechpartner</strong> mit Foto + Mail-Adresse auf.<br /><br />
                      <strong>Automatismen:</strong> Organizer bekommen je nach Einstellung in <strong>Schritt 6 (Kommunikation)</strong> eine BCC-Kopie der Anmelde-/Abmelde-Mails. Late-Cancel- und Roommate-Mails gehen ebenfalls an alle Organizer. Wenn ein Teilnehmer die Outlook-Einladung weiterleitet und der Empfänger nicht angemeldet ist, bekommen die Organizer eine FYI-Mail.<br /><br />
                      <strong>Reihenfolge zählt:</strong> der erste Organizer ist der Haupt-Organizer und wird in Mails als Absender-Name verwendet.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>responsible people</strong> for this event — any Deloitte user via Graph search. You are pre-filled by default, but you can add co-organizers or remove yourself.<br /><br />
                      <strong>Shown in the app:</strong> organizers can <strong>edit, deactivate, delete</strong> the event, see the <strong>attendee list</strong>, <strong>send QR codes</strong> and <strong>mass emails</strong>. They appear on the registration page and in My Events as <strong>contacts</strong> with photo + email.<br /><br />
                      <strong>Automation:</strong> depending on the setting in <strong>step 5 (Communication)</strong>, organizers receive BCC copies of registration / cancellation mails. Late-cancel and roommate mails go to all organizers. If an attendee forwards the Outlook invite to someone unregistered, organizers receive an FYI mail.<br /><br />
                      <strong>Order matters:</strong> the first organizer is the main organizer and is used as the sender name in mails.
                    </>
                  )} />
                  <BulkImportButton isDe={isDe} onClick={() => setBulkOrganizerOpen(true)} />
                </div>
                {/* Mismatch-Warning: bei Legacy-Korruption aus v10.0–v10.2-Closure-Bug
                    haben Events mehr Namen als Emails (oder umgekehrt). Auto-Heal padded
                    inzwischen statt zu truncaten — die Chips zeigen alle Namen, aber bei
                    welchen die Email fehlt, kann der User es hier nachpflegen. */}
                {(() => {
                  const missingEmailCount = orgList.reduce((acc, _, i) => {
                    return acc + ((organizerEmails[i] || '').trim() === '' ? 1 : 0);
                  }, 0);
                  if (missingEmailCount === 0) return null;
                  return (
                    <WizardHint
                      isDe={isDe}
                      title={isDe
                        ? `${missingEmailCount} Organizer ohne hinterlegte E-Mail-Adresse`
                        : `${missingEmailCount} organizer(s) without a stored email address`}
                      style={{ marginBottom: 10 }}
                    >
                      {isDe
                        ? <>Bei diesen Personen fehlen die Mails fürs <strong>BCC</strong> der Anmelde-/Abmelde-Mails, die <strong>Outlook-Einladung</strong> und die <strong>Decline-/Forward-Notifications</strong>. Bitte entferne die betroffenen Einträge (X) und füge sie über die Suche oder den Massenimport neu ein. <em>(Ursache: Legacy-Daten aus einer früheren App-Version — wird beim nächsten Speichern geheilt.)</em></>
                        : <>For these people the emails for the <strong>BCC</strong> of registration/cancellation mails, the <strong>Outlook invitation</strong> and the <strong>decline/forward notifications</strong> are missing. Please remove the affected entries (X) and re-add them via search or bulk import. <em>(Cause: legacy data from an earlier app version — healed on the next save.)</em></>}
                    </WizardHint>
                  );
                })()}
                {/* v31.2: Organizer als Zeilen mit Foto statt grüner Chips — Name,
                    E-Mail, Rolle und die Aktionen (Reihenfolge, Ausblenden, Entfernen)
                    sind so auf einen Blick lesbar. Die Markierung als Rückfragen-
                    Kontakt steht jetzt in Frage 3 (dort gehört sie thematisch hin);
                    hier zeigt nur noch ein Pill, wer markiert ist. */}
                {orgList.length === 0 ? (
                  <div className="dex-ui-empty" style={{ padding: '16px 14px', marginBottom: 10 }}>
                    <div className="dex-ui-empty-title">{isDe ? 'Noch kein Organizer' : 'No organizer yet'}</div>
                    {isDe ? 'Suche unten eine Person und wähle sie aus der Liste.' : 'Search for a person below and pick them from the list.'}
                  </div>
                ) : (() => {
                  const move = (idx: number, dir: -1 | 1): void => {
                    const nextNames = [...orgList];
                    const target = idx + dir;
                    if (target < 0 || target >= nextNames.length) return;
                    [nextNames[idx], nextNames[target]] = [nextNames[target], nextNames[idx]];
                    setOrganizer(nextNames.join('; '));
                    setOrganizerEmails(prev => {
                      if (idx >= prev.length || target >= prev.length) return prev;
                      const nextEmails = [...prev];
                      [nextEmails[idx], nextEmails[target]] = [nextEmails[target], nextEmails[idx]];
                      return nextEmails;
                    });
                  };
                  const remove = (idx: number): void => {
                    // Email-aware Remove: bei State-Korruption (z.B. Events aus
                    // v10.0–v10.2 wo der Closure-Bug emails ohne Namen schrieb)
                    // kann die gleiche Email mehrfach in organizerEmails stehen,
                    // während orgList nur einen Eintrag hat. Ein reiner Index-Filter
                    // würde dann nur EINEN Email-Eintrag killen, der Rest bleibt
                    // drin → die Person bleibt für den Picker „bekannt" und ist
                    // ausgegraut. Deshalb: emailToRemove ermitteln und ALLE
                    // Vorkommen aus organizerEmails entfernen.
                    const emailToRemove = (organizerEmails[idx] || '').toLowerCase();
                    const nextNames = orgList.filter((_, i) => i !== idx);
                    // v28.5: War der Entfernte der Rückfragen-Kontakt → Markierung löschen.
                    const removedEmail = (organizerEmails[idx] || '').toLowerCase();
                    if (removedEmail && (contactOrganizerEmail || '').toLowerCase() === removedEmail) setContactOrganizerEmail('');
                    setOrganizer(nextNames.join('; '));
                    setOrganizerEmails(prev => {
                      if (!emailToRemove) return prev.filter((_, i) => i !== idx);
                      return prev.filter(e => (e || '').toLowerCase() !== emailToRemove);
                    });
                  };
                  return (
                    <div className="dex-ui-card" style={{ padding: '4px 6px', marginBottom: 10 }}>
                      {orgList.map((name, i) => {
                        const email = organizerEmails[i] || '';
                        // v24.12 (J-Gate): Einzel-Ausblenden nur wenn „Organizer
                        // einzeln ausblenden" aktiv ist; sonst kein Auge-Knopf.
                        const canHideToggle = hideOrganizer && hideOrganizerIndividualOnly && !!email;
                        const orgHidden = hideOrganizer && hideOrganizerIndividualOnly && !!email && hiddenOrganizerEmails.indexOf(email.toLowerCase()) >= 0;
                        const isContact = !!email && contactLc === email.toLowerCase();
                        return (
                          <div key={`${name}-${i}`} className="dex-ui-row dex-ui-row--bordered">
                            <PersonAvatar email={email} name={name} muted={orgHidden} />
                            <div className="dex-ui-row-main">
                              <div className="dex-ui-row-title" style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap', whiteSpace: 'normal' }}>
                                <span style={{ textDecoration: orgHidden ? 'line-through' : 'none' }}>{name}</span>
                                {i === 0 && <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'Haupt-Organizer' : 'Main organizer'}</span>}
                                {isContact && (
                                  <span className="dex-ui-pill dex-ui-pill--orange"><Icon iconName="Chat" style={{ fontSize: 10 }} /> {isDe ? 'Rückfragen-Kontakt' : 'Contact for questions'}</span>
                                )}
                                {orgHidden && <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'auf der Anmeldeseite ausgeblendet' : 'hidden on the registration page'}</span>}
                              </div>
                              <div className="dex-ui-row-sub">{email || (isDe ? 'Keine E-Mail-Adresse hinterlegt' : 'No email address stored')}</div>
                            </div>
                            <div className="dex-ui-row-actions">
                              {/* v24.8 (J): Ausblenden auf der Anmeldeseite — Rechte bleiben.
                                  v31.2: eigener Auge-Knopf statt Klick auf den Namen. */}
                              {canHideToggle && (
                                <button
                                  type="button"
                                  className="dex-ui-iconbtn"
                                  onClick={() => toggleOrganizerHidden(email)}
                                  title={isDe
                                    ? (orgHidden ? 'Wird auf der Anmeldeseite NICHT angezeigt — klicken zum Einblenden (Rechte bleiben)' : 'Auf der Anmeldeseite ausblenden (Rechte bleiben)')
                                    : (orgHidden ? 'Hidden on the registration page — click to show (rights stay)' : 'Hide on the registration page (rights stay)')}
                                  aria-pressed={orgHidden}
                                >
                                  <Icon iconName={orgHidden ? 'Hide3' : 'RedEye'} style={{ fontSize: 14 }} />
                                </button>
                              )}
                              {orgList.length > 1 && i > 0 && (
                                <button type="button" className="dex-ui-iconbtn" onClick={() => move(i, -1)} title={isDe ? 'Nach oben' : 'Move up'} aria-label={isDe ? 'Nach oben' : 'Move up'}><ChevronUp size={16} /></button>
                              )}
                              {orgList.length > 1 && i < orgList.length - 1 && (
                                <button type="button" className="dex-ui-iconbtn" onClick={() => move(i, 1)} title={isDe ? 'Nach unten' : 'Move down'} aria-label={isDe ? 'Nach unten' : 'Move down'}><ChevronDown size={16} /></button>
                              )}
                              <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => remove(i)} title={isDe ? 'Entfernen' : 'Remove'} aria-label={isDe ? 'Entfernen' : 'Remove'}><X size={16} /></button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}
                <PersonPicker
                  isDe={isDe}
                  value={organizerSearch}
                  setValue={setOrganizerSearch}
                  timerRef={organizerTimerRef}
                  results={organizerResults}
                  setResults={setOrganizerResults}
                  includeIntl={organizerIncludeIntl}
                  setIncludeIntl={setOrganizerIncludeIntl}
                  searchUsers={searchUsers}
                  placeholder={t('create.organizer.placeholder')}
                  inputStyle={errorBorderStyle('organizer')}
                  searching={isSearchingOrganizer}
                  isAdded={u => organizerEmails.indexOf(u.email) >= 0}
                  onPick={u => {
                    const existing = organizer.split(';').map(s => s.trim()).filter(Boolean);
                    const nextNames = [...existing, u.displayName];
                    setOrganizer(nextNames.join('; '));
                    setOrganizerEmails(prev => [...prev, u.email]);
                    setOrganizerSearch('');
                    setOrganizerResults([]);
                  }}
                />
                {/* v24.8 (J) / v24.15 (Gate): Stand der Anzeige unter der Liste. */}
                {orgList.length > 0 && (
                  <div className="dex-ui-help">
                    {isDe
                      ? (!hideOrganizer
                          ? 'Alle Organizer werden auf der Anmeldeseite als Ansprechpartner angezeigt. Einzelne ausblenden: unten unter „Weitere Einstellungen“ beide Optionen einschalten — dann erscheint hier je Person ein Auge-Symbol.'
                          : (hideOrganizerIndividualOnly
                              ? 'Tipp: Das Auge-Symbol an einer Person blendet sie auf der Anmeldeseite aus (sie behält alle Rechte). Erneut klicken zeigt sie wieder.'
                              : 'Alle Organizer sind auf der Anmeldeseite ausgeblendet. Bei mehreren Organizern kannst du unten unter „Weitere Einstellungen“ stattdessen „Nur einzelne Organizer ausblenden“ wählen.'))
                      : (!hideOrganizer
                          ? 'All organizers are shown on the registration page as contacts. To hide individuals: enable both options under „More settings“ below — an eye icon then appears here for each person.'
                          : (hideOrganizerIndividualOnly
                              ? 'Tip: the eye icon on a person hides them on the registration page (they keep all rights). Click again to show.'
                              : 'All organizers are hidden on the registration page. With several organizers you can choose „Hide individual organizers only“ under „More settings“ below instead.'))}
                  </div>
                )}
              </div>

              {/* ---- Frage 2: Wer hilft mit? (optional) --------------------------- */}
              {/* v9.21: Test-Team pro Event — sieht das Event im Entwurfsmodus
                  und kann sich anmelden, ohne globale Organizer-Rolle.
                  v6.19: QR-Code-Scanner pro Event — eingeschränkter Admin-Zugriff
                  (nur QR-Tool + KPIs), erscheinen NICHT in Organizer-Listen auf
                  MyEvents/RegistrationPage und bekommen KEINE Organizer-Mails.
                  Beide Picker via Graph-Search, beliebige Deloitte-User. */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Wer hilft mit?' : 'Who helps out?'} <span className="dex-ui-label-optional" style={{ textTransform: 'none', letterSpacing: 0 }}>(optional)</span></div>
                <p className="dex-ui-section-desc">
                  {isDe
                    ? 'Diese Personen unterstützen die Organizer, ändern aber nichts am Event. Die Teilnehmerliste sehen sie nur eingeschränkt — das Check-in-Team beim Check-in.'
                    : 'These people support the organizers but cannot change the event. They see the attendee list only to a limited extent — the check-in team during check-in.'}
                </p>
                <div className="dex-ui-stack">
                  <div className="dex-ui-card">
                    <div className="dex-ui-label" style={{ flexWrap: 'wrap' }}>
                      <StepBadge n={12} />
                      {isDe ? 'Test-Team' : 'Test team'}
                      <InfoTooltip text={isDe ? (
                        <>
                          <strong>Was du hier einstellst:</strong> eine kleine Gruppe von Personen, die das Event <strong>schon im Entwurfsmodus</strong> sieht und sich testweise anmelden darf — bevor du es für die echte Zielgruppe freigibst.<br /><br />
                          <strong>Anzeige in der App:</strong> Test-Team-Mitglieder sehen das Event in ihrer Liste, können auf die Anmelde-Seite, sich registrieren, abmelden, eigene Daten ändern. Sie haben <strong>keine Admin-Rechte</strong> — kein Bearbeiten, keine Teilnehmerliste, keine Massenmails. Reguläre User sehen das Event weiterhin nicht, solange der Entwurf-Haken gesetzt ist.<br /><br />
                          <strong>Automatismen:</strong> Test-Anmeldungen lösen ganz normal <strong>Bestätigungs-Mails</strong> und <strong>Outlook-Termine</strong> aus — perfekt um den kompletten Anmelde-Ablauf zu testen. Bei externen Mails (nicht @deloitte.de) greift die normale Umleitung an dich als Organizer.<br /><br />
                          <strong>Empfehlung:</strong> 1–3 Personen reichen typischerweise — ein Co-Organizer und ein naiver Tester, der noch nichts vom Event weiß.
                        </>
                      ) : (
                        <>
                          <strong>What you set here:</strong> a small group of people who can see the event <strong>already in draft mode</strong> and register as a test — before you publish it to the real audience.<br /><br />
                          <strong>Shown in the app:</strong> test-team members see the event in their list, can open the registration page, register, cancel, edit their own data. They have <strong>no admin rights</strong> — no edit, no attendee list, no mass mails. Regular users still do not see the event while the draft toggle is on.<br /><br />
                          <strong>Automation:</strong> test registrations trigger normal <strong>confirmation mails</strong> and <strong>Outlook events</strong> — perfect for testing the full flow. External mails (non-@deloitte.de) follow the standard organizer-redirect rule.<br /><br />
                          <strong>Tip:</strong> 1–3 people are usually enough — a co-organizer and one naive tester who has not seen the event yet.
                        </>
                      )} />
                      <BulkImportButton isDe={isDe} onClick={() => setBulkTestTeamOpen(true)} />
                    </div>
                    {/* v24.4 (L): Kernaussage inline (ohne Mouse-over), Rest im Info-Symbol. */}
                    <p className="dex-ui-help" style={{ margin: '0 0 10px' }}>
                      {isDe
                        ? 'Sieht das Event schon im Entwurf und darf sich testweise anmelden — mit echten Bestätigungs-Mails und Outlook-Terminen, bevor die Zielgruppe es sieht.'
                        : 'Sees the event while it is still a draft and may register as a test — with real confirmation mails and Outlook invites, before the audience sees it.'}
                    </p>
                    {testTeamNames.length > 0 && (() => {
                      const remove = (idx: number): void => {
                        setTestTeamNames(testTeamNames.filter((_, i) => i !== idx));
                        setTestTeamEmails(testTeamEmails.filter((_, i) => i !== idx));
                      };
                      return (
                        <div className="dex-ui-card dex-ui-card--soft" style={{ padding: '4px 6px', marginBottom: 10 }}>
                          {testTeamNames.map((name, i) => {
                            const email = testTeamEmails[i] || '';
                            return (
                              <div key={`${email}-${i}`} className="dex-ui-row dex-ui-row--bordered">
                                <PersonAvatar email={email} name={name} />
                                <div className="dex-ui-row-main">
                                  <div className="dex-ui-row-title">{name}</div>
                                  {email && <div className="dex-ui-row-sub">{email}</div>}
                                </div>
                                <div className="dex-ui-row-actions">
                                  <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => remove(i)} title={isDe ? 'Entfernen' : 'Remove'} aria-label={isDe ? 'Entfernen' : 'Remove'}><X size={16} /></button>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      );
                    })()}
                    <PersonPicker
                      isDe={isDe}
                      value={testTeamSearch}
                      setValue={setTestTeamSearch}
                      timerRef={testTeamTimerRef}
                      results={testTeamResults}
                      setResults={setTestTeamResults}
                      includeIntl={testTeamIncludeIntl}
                      setIncludeIntl={setTestTeamIncludeIntl}
                      searchUsers={searchUsers}
                      placeholder={isDe ? 'Name oder E-Mail eingeben (alle Deloitte-User)' : 'Type name or email (any Deloitte user)'}
                      isAdded={u => testTeamEmails.indexOf(u.email) >= 0}
                      onPick={u => {
                        if (!u.email) return;
                        setTestTeamNames(prev => [...prev, u.displayName]);
                        setTestTeamEmails(prev => [...prev, u.email]);
                        setTestTeamSearch('');
                        setTestTeamResults([]);
                      }}
                    />
                  </div>

                  <div className="dex-ui-card">
                    <div className="dex-ui-label" style={{ flexWrap: 'wrap' }}>
                      <StepBadge n={13} />
                      {t('create.qrscanners') || 'QR-Code-Scanner'}
                      <InfoTooltip text={isDe ? (
                        <>
                          <strong>Was du hier einstellst:</strong> Personen, die am Event-Tag <strong>nur das Check-In-Tool</strong> bedienen dürfen — z.B. Helfer am Empfangstresen oder am Stadioneingang. Beliebige Deloitte-User per Graph-Suche.<br /><br />
                          <strong>Anzeige in der App:</strong> Check-In-Team-Mitglieder sehen oben im Header das <strong>QR-Scanner-Icon</strong> und können den <strong>Check-In-Modus</strong> öffnen — QR-Codes scannen, Teilnehmer manuell ein-/auschecken, Check-In-KPIs sehen. Sie haben <strong>keine weiteren Rechte</strong>: kein Edit, keine Teilnehmerliste, keine Mails.<br /><br />
                          <strong>Automatismen:</strong> Check-In-Team taucht <strong>nicht in der Organizer-Liste</strong> auf der Anmelde-Seite auf und bekommt <strong>keine Organizer-Mails</strong> (BCC, Late-Cancel etc.).<br /><br />
                          <strong>Empfehlung:</strong> für jedes Event genau die Personen eintragen, die am Veranstaltungstag wirklich am Empfang stehen.
                        </>
                      ) : (
                        <>
                          <strong>What you set here:</strong> people who may operate <strong>only the check-in tool</strong> on the event day — e.g. helpers at the welcome desk or stadium entrance. Any Deloitte user via Graph search.<br /><br />
                          <strong>Shown in the app:</strong> check-in team members see the <strong>QR scanner icon</strong> in the header and can open the <strong>check-in mode</strong> — scan QR codes, manually check attendees in/out, view check-in KPIs. They have <strong>no further rights</strong>: no edit, no attendee list, no emails.<br /><br />
                          <strong>Automation:</strong> check-in team does <strong>not appear in the organizer list</strong> on the registration page and does <strong>not receive organizer emails</strong> (BCC, late-cancel etc.).<br /><br />
                          <strong>Tip:</strong> for each event, list exactly the people who will actually staff the welcome desk.
                        </>
                      )} />
                      <BulkImportButton isDe={isDe} onClick={() => setBulkQrScannerOpen(true)} />
                    </div>
                    {/* v24.4 (L): Kernaussage inline (ohne Mouse-over), Rest im Info-Symbol. */}
                    <p className="dex-ui-help" style={{ margin: '0 0 10px' }}>
                      {isDe
                        ? 'Bedient am Event-Tag nur das Check-in-Tool (QR-Codes scannen, Teilnehmer ein-/auschecken) — sonst keine weiteren Rechte, keine Organizer-Mails.'
                        : 'Operates only the check-in tool on the event day (scan QR codes, check attendees in/out) — no other rights, no organizer mails.'}
                    </p>
                    {qrScannerNames.length > 0 && (() => {
                      const move = (idx: number, dir: -1 | 1): void => {
                        const target = idx + dir;
                        if (target < 0 || target >= qrScannerNames.length) return;
                        const nextNames = [...qrScannerNames];
                        const nextEmails = [...qrScannerEmails];
                        [nextNames[idx], nextNames[target]] = [nextNames[target], nextNames[idx]];
                        [nextEmails[idx], nextEmails[target]] = [nextEmails[target], nextEmails[idx]];
                        setQrScannerNames(nextNames);
                        setQrScannerEmails(nextEmails);
                      };
                      const remove = (idx: number): void => {
                        setQrScannerNames(qrScannerNames.filter((_, i) => i !== idx));
                        setQrScannerEmails(qrScannerEmails.filter((_, i) => i !== idx));
                      };
                      return (
                        <div className="dex-ui-card dex-ui-card--soft" style={{ padding: '4px 6px', marginBottom: 10 }}>
                          {qrScannerNames.map((name, i) => {
                            const email = qrScannerEmails[i] || '';
                            return (
                              <div key={`${email}-${i}`} className="dex-ui-row dex-ui-row--bordered">
                                <PersonAvatar email={email} name={name} />
                                <div className="dex-ui-row-main">
                                  <div className="dex-ui-row-title">{name}</div>
                                  {email && <div className="dex-ui-row-sub">{email}</div>}
                                </div>
                                <div className="dex-ui-row-actions">
                                  {qrScannerNames.length > 1 && i > 0 && (
                                    <button type="button" className="dex-ui-iconbtn" onClick={() => move(i, -1)} title={isDe ? 'Nach oben' : 'Move up'} aria-label={isDe ? 'Nach oben' : 'Move up'}><ChevronUp size={16} /></button>
                                  )}
                                  {qrScannerNames.length > 1 && i < qrScannerNames.length - 1 && (
                                    <button type="button" className="dex-ui-iconbtn" onClick={() => move(i, 1)} title={isDe ? 'Nach unten' : 'Move down'} aria-label={isDe ? 'Nach unten' : 'Move down'}><ChevronDown size={16} /></button>
                                  )}
                                  <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => remove(i)} title={isDe ? 'Entfernen' : 'Remove'} aria-label={isDe ? 'Entfernen' : 'Remove'}><X size={16} /></button>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      );
                    })()}
                    <PersonPicker
                      isDe={isDe}
                      value={qrScannerSearch}
                      setValue={setQrScannerSearch}
                      timerRef={qrScannerTimerRef}
                      results={qrScannerResults}
                      setResults={setQrScannerResults}
                      includeIntl={qrScannerIncludeIntl}
                      setIncludeIntl={setQrScannerIncludeIntl}
                      searchUsers={searchUsers}
                      placeholder={t('create.qrscanners.placeholder') || (isDe ? 'Name oder E-Mail eingeben (alle Deloitte-User)' : 'Type name or email (any Deloitte user)')}
                      isAdded={u => qrScannerEmails.indexOf(u.email) >= 0}
                      onPick={u => {
                        if (!u.email) return;
                        setQrScannerNames(prev => [...prev, u.displayName]);
                        setQrScannerEmails(prev => [...prev, u.email]);
                        setQrScannerSearch('');
                        setQrScannerResults([]);
                      }}
                    />
                  </div>
                </div>

                {/* Duplikat-Hinweis: gleiche Person in mehreren Team-Listen.
                    Co-Organizer haben automatisch Check-In- und Test-Team-Rechte
                    (Header.canCheckIn-Logik, Drafts-Sichtbarkeit für Organizer),
                    doppelte Einträge sind redundant. Test-Team und Check-In allein
                    sind orthogonale Rollen — die warnen wir nicht.

                    Pattern: nach jedem Add (Massenimport oder Einzel-Pick) updated
                    sich die Memo automatisch und die Warnung erscheint inline. Pro
                    Eintrag ein Ein-Klick-Button um die Person aus der überflüssigen
                    Liste zu entfernen. */}
                {(() => {
                  const orgSet = new Set(organizerEmails.map(e => (e || '').toLowerCase()));
                  const ttSet = new Set(testTeamEmails.map(e => (e || '').toLowerCase()));
                  const qrSet = new Set(qrScannerEmails.map(e => (e || '').toLowerCase()));
                  const allEmails = new Set<string>();
                  organizerEmails.forEach(e => allEmails.add((e || '').toLowerCase()));
                  testTeamEmails.forEach(e => allEmails.add((e || '').toLowerCase()));
                  qrScannerEmails.forEach(e => allEmails.add((e || '').toLowerCase()));

                  const orgNames = organizer.split(';').map(s => s.trim()).filter(Boolean);
                  const dups: Array<{ email: string; name: string; inOrg: boolean; inTt: boolean; inQr: boolean }> = [];
                  // Array.from() statt `for ... of Set` — TS-Target ES5 erlaubt kein
                  // direktes Set-Iterieren ohne --downlevelIteration.
                  for (const e of Array.from(allEmails)) {
                    if (!e) continue;
                    const inOrg = orgSet.has(e);
                    const inTt = ttSet.has(e);
                    const inQr = qrSet.has(e);
                    // Nur Co-Organizer + (Test|Check-In) ist redundant. Test+Check-In
                    // ohne Co-Organizer sind unterschiedliche Funktionen → nicht warnen.
                    if (!inOrg) continue;
                    if (!inTt && !inQr) continue;
                    // Display-Name aus dem ersten Treffer ziehen (Org bevorzugt).
                    let name = e;
                    const idxOrg = organizerEmails.findIndex(x => (x || '').toLowerCase() === e);
                    if (idxOrg >= 0 && orgNames[idxOrg]) name = orgNames[idxOrg];
                    else {
                      const idxTt = testTeamEmails.findIndex(x => (x || '').toLowerCase() === e);
                      if (idxTt >= 0 && testTeamNames[idxTt]) name = testTeamNames[idxTt];
                      else {
                        const idxQr = qrScannerEmails.findIndex(x => (x || '').toLowerCase() === e);
                        if (idxQr >= 0 && qrScannerNames[idxQr]) name = qrScannerNames[idxQr];
                      }
                    }
                    dups.push({ email: e, name, inOrg, inTt, inQr });
                  }
                  if (dups.length === 0) return null;

                  const removeFromTestTeam = (emailLc: string): void => {
                    const idx = testTeamEmails.findIndex(x => (x || '').toLowerCase() === emailLc);
                    if (idx < 0) return;
                    setTestTeamNames(testTeamNames.filter((_, i) => i !== idx));
                    setTestTeamEmails(testTeamEmails.filter((_, i) => i !== idx));
                  };
                  const removeFromQr = (emailLc: string): void => {
                    const idx = qrScannerEmails.findIndex(x => (x || '').toLowerCase() === emailLc);
                    if (idx < 0) return;
                    setQrScannerNames(qrScannerNames.filter((_, i) => i !== idx));
                    setQrScannerEmails(qrScannerEmails.filter((_, i) => i !== idx));
                  };
                  const removeAllOverlap = (emailLc: string): void => {
                    removeFromTestTeam(emailLc);
                    removeFromQr(emailLc);
                  };

                  return (
                    <WizardHint
                      isDe={isDe}
                      title={dups.length === 1
                        ? (isDe ? '1 Person ist mehrfach gelistet' : '1 person is listed multiple times')
                        : (isDe ? `${dups.length} Personen sind mehrfach gelistet` : `${dups.length} people are listed multiple times`)}
                      style={{ marginTop: 12 }}
                      defaultOpen
                    >
                      {/* v31.2: bisher nur deutsch — jetzt beide Sprachen. */}
                      <p style={{ margin: '0 0 10px' }}>
                        {isDe
                          ? <>Co-Organizer dürfen automatisch das <strong>Check-in-Tool</strong> nutzen und sehen Events auch im <strong>Entwurfsmodus</strong> — ein zusätzlicher Eintrag im Test-Team oder Check-in-Team ist nicht nötig. Entferne die überflüssigen Einträge hier mit einem Klick.</>
                          : <>Co-organizers can automatically use the <strong>check-in tool</strong> and see events in <strong>draft mode</strong> — an extra entry in the test team or check-in team is not needed. Remove the redundant entries here with one click.</>}
                      </p>
                      <div>
                        {dups.map(d => {
                          const teamsLabel: string[] = [];
                          if (d.inOrg) teamsLabel.push(isDe ? 'Co-Organizer' : 'Co-organizer');
                          if (d.inTt) teamsLabel.push(isDe ? 'Test-Team' : 'Test team');
                          if (d.inQr) teamsLabel.push(isDe ? 'Check-in-Team' : 'Check-in team');
                          // v31.2 (Leitfaden 2a′): Die Entfernen-Knöpfe folgen dem Namen
                          // direkt (gap 12, umbrechend) — vorher schob `flex: 1` auf dem
                          // Textblock sie allein an den rechten Rand.
                          return (
                            <div key={d.email} className="dex-ui-row dex-ui-row--bordered" style={{ flexWrap: 'wrap', padding: '8px 0' }}>
                              <div className="dex-ui-row-main" style={{ flex: '0 1 auto' }}>
                                <div className="dex-ui-row-title" style={{ whiteSpace: 'normal' }}>{d.name} <span style={{ color: 'var(--dex-gray-500)', fontWeight: 400, fontSize: '0.8rem' }}>{d.email}</span></div>
                                <div className="dex-ui-row-sub">{isDe ? 'Aktuell in' : 'Currently in'}: {teamsLabel.join(', ')}</div>
                              </div>
                              <div className="dex-ui-inline">
                                {d.inTt && (
                                  <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => removeFromTestTeam(d.email)}>
                                    {isDe ? 'Aus Test-Team entfernen' : 'Remove from test team'}
                                  </button>
                                )}
                                {d.inQr && (
                                  <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => removeFromQr(d.email)}>
                                    {isDe ? 'Aus Check-in-Team entfernen' : 'Remove from check-in team'}
                                  </button>
                                )}
                                {d.inTt && d.inQr && (
                                  <button type="button" className="btn btn-outline dex-ui-btn-sm" onClick={() => removeAllOverlap(d.email)}>
                                    {isDe ? 'Aus beiden entfernen' : 'Remove from both'}
                                  </button>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </WizardHint>
                  );
                })()}
              </div>

              {/* ---- Frage 3: An wen wenden sich Teilnehmer bei Fragen? ---------- */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'An wen wenden sich Teilnehmer bei Fragen?' : 'Who do attendees contact with questions?'}</div>
                {/* v28.5: Rückfragen-Kontakt — genau EINER; erneuter Klick entfernt
                    die Markierung. v28.10: Sprechblase statt „?". v31.2: als
                    Chip-Reihe hier statt als Knopf am Organizer-Chip — die Frage
                    „wen sollen Teilnehmer ansprechen" gehört zu den Kontaktdaten,
                    nicht zur Rechtevergabe. Der Setter ist derselbe. */}
                <div className="dex-ui-field">
                  <div className="dex-ui-label">
                    {isDe ? 'Welcher Organizer beantwortet Rückfragen?' : 'Which organizer answers questions?'}
                    <InfoTooltip text={isDe
                      ? <>Die markierte Person bekommt auf der Anmeldeseite einen <strong>orangen Sprechblasen-Badge</strong> mit dem Hinweis, sich bei Fragen an sie zu wenden. Ohne Markierung stehen alle Organizer gleichberechtigt als Ansprechpartner da. Wird die Person aus den Organizern entfernt, verschwindet die Markierung mit.</>
                      : <>The marked person gets an <strong>orange chat-bubble badge</strong> on the registration page asking attendees to contact them with questions. Without a mark, all organizers appear as equal contacts. Removing the person from the organizers also removes the mark.</>} />
                  </div>
                  {(() => {
                    const candidates = orgList.map((name, i) => ({ name, email: organizerEmails[i] || '' })).filter(c => !!c.email);
                    if (candidates.length === 0) {
                      return (
                        <div className="dex-ui-help">
                          {isDe ? 'Füge oben zuerst einen Organizer mit E-Mail-Adresse hinzu — dann kannst du hier eine Person markieren.' : 'Add an organizer with an email address above first — then you can mark a person here.'}
                        </div>
                      );
                    }
                    return (
                      <>
                        <div className="dex-ui-inline">
                          <button
                            type="button"
                            className={cx('dex-ui-chip', !contactLc.trim() && 'is-active')}
                            onClick={() => setContactOrganizerEmail('')}
                            aria-pressed={!contactLc.trim()}
                          >
                            {isDe ? 'Kein besonderer Kontakt' : 'No specific contact'}
                          </button>
                          {candidates.map((c, i) => {
                            const active = contactLc === c.email.toLowerCase();
                            return (
                              // v31.2: Index im Key — bei Legacy-Korruption (s. remove oben)
                              // kann dieselbe E-Mail mehrfach stehen; ohne Index kollidieren
                              // die Keys und React ordnet die Chips beim Re-Render falsch zu.
                              <button
                                key={`${c.email}-${i}`}
                                type="button"
                                className={cx('dex-ui-chip', active && 'is-active')}
                                onClick={() => setContactOrganizerEmail(prev => (prev || '').toLowerCase() === c.email.toLowerCase() ? '' : c.email)}
                                aria-pressed={active}
                                title={isDe
                                  ? (active ? 'Rückfragen-Kontakt — Teilnehmer werden gebeten, diese Person bei Fragen zu kontaktieren. Klicken zum Entfernen.' : 'Als Rückfragen-Kontakt markieren (oranger Badge auf der Anmeldeseite)')
                                  : (active ? 'Contact for questions — attendees are asked to reach out to this person. Click to remove.' : 'Mark as contact for questions (orange badge on the registration page)')}
                              >
                                <Icon iconName="Chat" style={{ fontSize: 11 }} /> {c.name}
                              </button>
                            );
                          })}
                        </div>
                        <div className="dex-ui-help">
                          {contactLc.trim()
                            ? (isDe
                              ? 'Diese Person bekommt auf der Anmeldeseite den orangen Sprechblasen-Badge mit dem Hinweis, sich bei Fragen an sie zu wenden.'
                              : 'This person gets the orange chat-bubble badge on the registration page asking attendees to contact them with questions.')
                            : (isDe
                              ? 'Ohne Markierung stehen alle Organizer gleichberechtigt als Ansprechpartner auf der Anmeldeseite.'
                              : 'Without a mark, all organizers appear as equal contacts on the registration page.')}
                        </div>
                      </>
                    );
                  })()}
                </div>

                {/* v10.16: Optionaler Ansprechpartner. Reines Anzeige-Feld
                    (kein Login, keine SP-Permissions) — z.B. die Person vor Ort
                    oder eine Hotline-Mail die Teilnehmer bei Fragen anschreiben
                    sollen. Wird auf Register-/MyEvents-Page zusätzlich zu den
                    Organizern gezeigt. Alles drei optional, Freitext.
                    v24.10 (Q2): einklappbar, default zu. */}
                <button
                  type="button"
                  className={cx('dex-ui-disclosure', contactExpanded && 'is-open')}
                  onClick={() => setContactExpanded(v => !v)}
                  aria-expanded={contactExpanded}
                >
                  {/* v31.2: ChevronDown ohne eigene Drehung — die Klasse dreht ihn
                      geöffnet um 180°; ein ChevronRight zeigte dann nach links. */}
                  <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                  {isDe ? 'Zusätzlichen Ansprechpartner ohne App-Zugang angeben' : 'Add an extra contact without app access'}
                  <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                  {!!(contactName.trim() || contactEmail.trim() || contactInfo.trim()) && (
                    <span className="dex-ui-disclosure-count">{isDe ? 'ausgefüllt' : 'filled in'}</span>
                  )}
                </button>
                {contactExpanded && (
                <div className="dex-ui-disclosure-body">
                {/* v24.10 (Q3): Organizer ist Standard-Ansprechpartner; hier nur Externe.
                    v31.2: zwei Sätze sichtbar (Leitfaden 2c); WO die Person erscheint
                    und dass sie keine Rechte bekommt, steht als Hilfezeile unter dem Namen. */}
                <p className="dex-ui-help" style={{ margin: '0 0 12px' }}>
                  {isDe
                    ? 'Standardmäßig sind die Organizer die Ansprechpartner. Trage hier jemanden ein, der keinen App-Zugang braucht — z.B. eine Service-Mailadresse oder die Person vor Ort.'
                    : 'By default the organizers are the contacts. Enter someone here who does not need app access — e.g. a service mailbox or the on-site contact.'}
                </p>
                <div className="dex-ui-grid-2">
                  <div className="dex-ui-field">
                    <label className="dex-ui-label" htmlFor="dex-details-contact-name">{isDe ? 'Name' : 'Name'}</label>
                    <input
                      id="dex-details-contact-name"
                      type="text"
                      className="form-input"
                      value={contactName}
                      onChange={e => setContactName(e.target.value)}
                      placeholder={isDe ? 'z.B. Anna Schmitt' : 'e.g. Anna Schmitt'}
                    />
                    <div className="dex-ui-help">
                      {isDe
                        ? 'Erscheint auf der Anmeldeseite und in „Meine Events“ zusätzlich zu den Organizern — reines Anzeige-Feld, keine Rechte in der App.'
                        : 'Appears on the registration page and in „My Events“ in addition to the organizers — display only, no app permissions.'}
                    </div>
                  </div>
                  <div className="dex-ui-field">
                    <label className="dex-ui-label" htmlFor="dex-details-contact-email">{isDe ? 'E-Mail' : 'Email'}</label>
                    <input
                      id="dex-details-contact-email"
                      type="email"
                      className="form-input"
                      value={contactEmail}
                      onChange={e => setContactEmail(e.target.value)}
                      placeholder={isDe ? 'z.B. event-helpdesk@example.de' : 'e.g. event-helpdesk@example.com'}
                    />
                  </div>
                </div>
                <div className="dex-ui-field" style={{ marginTop: 14 }}>
                  <label className="dex-ui-label" htmlFor="dex-details-contact-info">
                    {isDe ? 'Wie und wann ist die Person erreichbar?' : 'How and when can this person be reached?'}
                    <span className="dex-ui-label-optional">{isDe ? '(Freitext)' : '(free text)'}</span>
                  </label>
                  {/* v31.2: dex-ui-textarea in voller Breite statt form-input —
                      die Klasse ist für Eingabefelder gedacht und ließ das
                      Textfeld schmal in der Zeile stehen. */}
                  <textarea
                    id="dex-details-contact-info"
                    className="dex-ui-textarea"
                    value={contactInfo}
                    onChange={e => setContactInfo(e.target.value)}
                    rows={3}
                    placeholder={isDe
                      ? 'z.B. „Vor Ort am Eventtag ab 7:30 Uhr, mobil unter +49 151 123 456" oder „Bei Fragen vor dem Event direkt per Mail."'
                      : 'e.g. „On-site from 7:30 am on event day, mobile +49 151 123 456" or „For questions before the event, email directly."'}
                  />
                  <div className="dex-ui-help">
                    {isDe ? 'Steht auf der Anmeldeseite und in „Meine Events“ direkt beim Ansprechpartner.' : 'Shown next to the contact on the registration page and in „My Events“.'}
                  </div>
                  {/* v23.18: Hinweis, wenn der Ansprechpartner-Freitext den
                      Event-Titel/Datum/Ort wiederholt — die stehen bereits
                      separat auf der Anmelde-Seite. Gleiche Logik wie der
                      Beschreibungs-Hinweis. */}
                  {(() => {
                    const plain = (contactInfo || '').replace(/\s+/g, ' ').toLowerCase();
                    if (plain.trim().length < 4) return null;
                    const hits: string[] = [];
                    const tl = title.trim().toLowerCase();
                    if (tl.length >= 4 && plain.indexOf(tl) >= 0) hits.push(isDe ? 'der Event-Name' : 'the event name');
                    const locl = location.trim().toLowerCase();
                    if (locl.length >= 4 && plain.indexOf(locl) >= 0) hits.push(isDe ? 'der Ort' : 'the location');
                    if (startDate) {
                      const d = new Date(startDate);
                      if (!isNaN(d.getTime())) {
                        const dd = String(d.getDate()).padStart(2, '0');
                        const mm = String(d.getMonth() + 1).padStart(2, '0');
                        const monthsDe = ['januar', 'februar', 'märz', 'april', 'mai', 'juni', 'juli', 'august', 'september', 'oktober', 'november', 'dezember'];
                        const monthsEn = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
                        const mn = (isDe ? monthsDe : monthsEn)[d.getMonth()];
                        const pats = [`${dd}.${mm}.${d.getFullYear()}`, `${dd}.${mm}.`, `${d.getDate()}. ${mn}`, `${d.getDate()} ${mn}`];
                        if (pats.some(p => plain.indexOf(p) >= 0)) hits.push(isDe ? 'das Datum' : 'the date');
                      }
                    }
                    if (hits.length === 0) return null;
                    const joined = hits.length === 1 ? hits[0] : hits.slice(0, -1).join(', ') + (isDe ? ' und ' : ' and ') + hits[hits.length - 1];
                    return (
                      <WizardHint
                        isDe={isDe}
                        title={isDe ? 'Ansprechpartner-Text wiederholt Basis-Infos' : 'Contact text repeats basic info'}
                        style={{ marginTop: 8 }}
                      >
                        {isDe
                          ? <>Hier steht offenbar <strong>{joined}</strong>. <strong>Event-Titel, Datum und Ort</strong> werden bereits <strong>separat</strong> auf der Anmelde-Seite angezeigt — du musst sie beim Ansprechpartner nicht wiederholen. Nutze dieses Feld nur für die <strong>Erreichbarkeit</strong> (z.B. Telefon/„ab wann vor Ort“).</>
                          : <>This appears to contain <strong>{joined}</strong>. The <strong>event title, date and location</strong> are already shown <strong>separately</strong> on the registration page — no need to repeat them in the contact field. Use it only for <strong>availability</strong> (e.g. phone / „on-site from …“).</>}
                      </WizardHint>
                    );
                  })()}
                </div>
                </div>
                )}
              </div>

              {/* ---- Weitere Einstellungen: Anzeige der Organizer (Aufklapper) --- */}
              {/* v18.9: Organizer-Anzeige ausblenden. Rein visuell — die
                  Organizer behalten alle Rechte + Mail-Benachrichtigungen,
                  werden aber auf der Anmelde-Seite und in „Meine Events"
                  nicht als Ansprechpartner-Chips gezeigt.
                  v23.25/v24.15: Darstellungs-Größe der Organizer. Nur relevant,
                  wenn überhaupt Organizer angezeigt werden — also NICHT, wenn
                  ALLE ausgeblendet sind (hideOrganizer ohne Einzel-Modus). */}
              <div className="dex-ui-section">
                <button
                  type="button"
                  className={cx('dex-ui-disclosure', moreOpen && 'is-open')}
                  onClick={() => setMoreOpen(v => !v)}
                  aria-expanded={moreOpen}
                >
                  {/* v31.2: ChevronDown ohne eigene Drehung (s. oben). Der Zähler
                      steht immer da — „Standard" sagt, dass sich das Öffnen sparen
                      kann, wer nichts geändert hat. */}
                  <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                  {isDe ? 'Weitere Einstellungen — wie erscheinen die Organizer auf der Anmeldeseite?' : 'More settings — how do organizers appear on the registration page?'}
                  <span className="dex-ui-disclosure-count">
                    {changedMore > 0
                      ? (isDe ? `${changedMore} angepasst` : `${changedMore} customized`)
                      : (isDe ? 'Standard' : 'Default')}
                  </span>
                </button>
                {moreOpen && (
                  <div className="dex-ui-disclosure-body dex-ui-stack">
                    <label className={cx('dex-ui-toggle-row', hideOrganizer && 'is-active')}>
                      <input
                        type="checkbox"
                        checked={hideOrganizer}
                        onChange={e => setHideOrganizer(e.target.checked)}
                      />
                      <span className="dex-ui-toggle-row-body">
                        <span className="dex-ui-toggle-row-title">
                          {isDe ? 'Organizer auf der Anmeldeseite ausblenden' : 'Hide organizers on the registration page'}
                          <InfoTooltip text={isDe
                            ? <>
                                <strong>Was du hier einstellst:</strong> ob die <strong>Organizer</strong> auf der <strong>Anmelde-Seite</strong> (und &bdquo;Meine Events&ldquo;) als Ansprechpartner angezeigt werden.<br /><br />
                                <strong>Anzeige in der App:</strong> aktiviert blendet alle Organizer aus. Gibt es mehrere Organizer, erscheint darunter zusätzlich die Option, stattdessen <strong>nur einzelne</strong> auszublenden (per Auge-Symbol in der Organizer-Liste).<br /><br />
                                <strong>Wichtig:</strong> das ist rein optisch — die Organizer behalten alle <strong>Rechte</strong> (bearbeiten, Teilnehmer verwalten) und ihre <strong>Mail-Benachrichtigungen</strong>. Ein zusätzlicher Ansprechpartner (Frage oben) bleibt sichtbar.
                              </>
                            : <>
                                <strong>What this controls:</strong> whether the <strong>organizers</strong> are shown as contacts on the <strong>registration page</strong> (and &bdquo;My Events&ldquo;).<br /><br />
                                <strong>Where you see it:</strong> when enabled, all organizers are hidden. With several organizers, an option appears below to hide <strong>only individual ones</strong> instead (eye icon in the organizer list).<br /><br />
                                <strong>Note:</strong> this is purely visual — organizers keep all <strong>permissions</strong> and their <strong>email notifications</strong>. An extra contact (question above) stays visible.
                              </>
                          } />
                        </span>
                        <span className="dex-ui-toggle-row-desc">
                          {isDe
                            ? 'Dann steht kein Organizer als Ansprechpartner auf der Anmeldeseite und in „Meine Events“; Rechte und Mails bleiben. Standard: aus.'
                            : 'Then no organizer is shown as a contact on the registration page and in „My Events“; rights and mails stay. Default: off.'}
                        </span>
                      </span>
                    </label>
                    {/* v24.15: zweite Option — nur bei „Organizer ausblenden" UND
                        mehreren Organizern: stattdessen nur einzelne ausblenden. */}
                    {hideOrganizer && organizerEmails.length >= 2 && (
                      <label className={cx('dex-ui-toggle-row', hideOrganizerIndividualOnly && 'is-active')} style={{ marginLeft: 28 }}>
                        <input
                          type="checkbox"
                          checked={hideOrganizerIndividualOnly}
                          onChange={e => setHideOrganizerIndividualOnly(e.target.checked)}
                        />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">{isDe ? 'Nur einzelne Organizer ausblenden' : 'Hide only individual organizers'}</span>
                          <span className="dex-ui-toggle-row-desc">
                            {isDe
                              ? 'Dann bleiben alle sichtbar — bis du oben in der Organizer-Liste bei einer Person das Auge-Symbol anklickst. Genau diese wird ausgeblendet, die übrigen bleiben.'
                              : 'Then everyone stays visible — until you click the eye icon on a person in the organizer list above. Exactly that person is hidden, the rest stay.'}
                          </span>
                        </span>
                      </label>
                    )}
                    {(!hideOrganizer || hideOrganizerIndividualOnly) && (
                      <div>
                        {/* v24.4 (I) / v24.10 (Q): Größe + Live-Vorschau. v31.2: zwei
                            Auswahl-Kacheln statt Radio-Paar (Leitfaden 2b). */}
                        <div className="dex-ui-label">
                          {isDe ? 'Wie groß erscheinen die Organizer?' : 'How large do organizers appear?'}
                          <InfoTooltip text={isDe
                            ? <>
                                <strong>Was du hier einstellst:</strong> wie die <strong>Organizer</strong> auf der Anmelde-Seite dargestellt werden.<br /><br />
                                <strong>Klein (Standard):</strong> kleiner <strong>Chip</strong> mit Foto + Name — die Mail-Adresse und Rolle erscheinen erst beim <strong>Drüberfahren</strong> mit der Maus.<br /><br />
                                <strong>Groß:</strong> die Organizer werden <strong>dauerhaft groß</strong> gezeigt (großes Foto, Name, klickbare <strong>E-Mail-Adresse</strong>, Rolle &amp; Standort) — Teilnehmer sehen die Kontaktdaten <strong>sofort</strong>, ohne die Maus darüber zu bewegen.
                              </>
                            : <>
                                <strong>What this controls:</strong> how the <strong>organizers</strong> are displayed on the registration page.<br /><br />
                                <strong>Small (default):</strong> small <strong>chip</strong> with photo + name — email and role only appear on <strong>hover</strong>.<br /><br />
                                <strong>Large:</strong> organizers are shown <strong>large permanently</strong> (big photo, name, clickable <strong>email</strong>, role &amp; location).
                              </>
                          } />
                        </div>
                        <div className="dex-ui-grid-2">
                          <button
                            type="button"
                            className={cx('dex-ui-choice', !organizerDisplayLarge && 'is-active')}
                            onClick={() => setOrganizerDisplayLarge(false)}
                            aria-pressed={!organizerDisplayLarge}
                          >
                            <span className="dex-ui-choice-body">
                              <span className="dex-ui-choice-title">{isDe ? 'Klein (Chip) — Standard' : 'Small (chip) — default'}</span>
                              <span className="dex-ui-choice-desc">
                                {isDe
                                  ? 'Kompakter Chip mit Foto + Name. Mail-Adresse und Rolle erscheinen groß, wenn Teilnehmer mit der Maus darüberfahren.'
                                  : 'Compact chip with photo + name. Email and role appear large when attendees hover over it.'}
                              </span>
                            </span>
                            <span className="dex-ui-choice-check">{!organizerDisplayLarge && <Check size={12} />}</span>
                          </button>
                          <button
                            type="button"
                            className={cx('dex-ui-choice', organizerDisplayLarge && 'is-active')}
                            onClick={() => setOrganizerDisplayLarge(true)}
                            aria-pressed={organizerDisplayLarge}
                          >
                            <span className="dex-ui-choice-body">
                              <span className="dex-ui-choice-title">{isDe ? 'Groß' : 'Large'}</span>
                              <span className="dex-ui-choice-desc">
                                {isDe
                                  ? 'Großes Foto, Name, klickbare E-Mail-Adresse und Rolle sofort sichtbar — ohne Mouse-over.'
                                  : 'Large photo, name, clickable email and role visible right away — no mouse-over.'}
                              </span>
                            </span>
                            <span className="dex-ui-choice-check">{organizerDisplayLarge && <Check size={12} />}</span>
                          </button>
                        </div>
                        {/* v24.10 (Q): Live-Vorschau — so sieht der Teilnehmer die Organizer. */}
                        {orgList.length > 0 && (
                          <div className="dex-ui-card dex-ui-card--soft" style={{ marginTop: 12 }}>
                            <div className="dex-ui-section-title" style={{ marginBottom: 8 }}>
                              {isDe ? 'Vorschau — so sehen es die Teilnehmer' : 'Preview — what attendees see'}
                            </div>
                            <OrganizerList
                              names={orgList}
                              emails={organizerEmails}
                              hiddenEmails={(hideOrganizer && hideOrganizerIndividualOnly) ? hiddenOrganizerEmails : []}
                              size="md"
                              display={organizerDisplayLarge ? 'card' : 'chip'}
                              contactEmail={contactOrganizerEmail || undefined}
                            />
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}
              </div>

              </div>
  );
};
