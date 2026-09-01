/* DetailsStep — aus EventCreationPage.tsx ausgelagert (Zeilen 10111-11152 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; einzige
 * Aenderung ist die Anzeige-Bedingung: aus `currentStep === 1` wurde das Prop `visible`.
 * `visible` schaltet display:none statt unmount — Eingaben ueberleben den
 * Schrittwechsel genauso wie vorher. */
import * as React from 'react';
import { StepBadge } from '../../wizard/StepBadge';
import { InfoTooltip } from '../../InfoTooltip';
import { Users } from '../../Icons';
import WizardHint from '../../WizardHint';
import { Icon } from '@fluentui/react/lib/Icon';
import InternationalSearchToggle from '../../InternationalSearchToggle';
import OrganizerList from '../../OrganizerList';
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
export const DetailsStep: React.FC<DetailsStepProps> = (p) => {
  const { visible } = p;
  const { contactEmail, contactExpanded, contactInfo, contactName, contactOrganizerEmail, errorBorderStyle, hiddenOrganizerEmails, hideOrganizer, hideOrganizerIndividualOnly, isDe, isSearchingOrganizer, location, organizer, organizerDisplayLarge, organizerEmails, organizerIncludeIntl, organizerResults, organizerSearch, organizerTimerRef, qrScannerEmails, qrScannerIncludeIntl, qrScannerNames, qrScannerResults, qrScannerSearch, qrScannerTimerRef, searchUsers, setBulkOrganizerOpen, setBulkQrScannerOpen, setBulkTestTeamOpen, setContactEmail, setContactExpanded, setContactInfo, setContactName, setContactOrganizerEmail, setHideOrganizer, setHideOrganizerIndividualOnly, setOrganizer, setOrganizerDisplayLarge, setOrganizerEmails, setOrganizerIncludeIntl, setOrganizerResults, setOrganizerSearch, setQrScannerEmails, setQrScannerIncludeIntl, setQrScannerNames, setQrScannerResults, setQrScannerSearch, setTestTeamEmails, setTestTeamIncludeIntl, setTestTeamNames, setTestTeamResults, setTestTeamSearch, startDate, t, testTeamEmails, testTeamIncludeIntl, testTeamNames, testTeamResults, testTeamSearch, testTeamTimerRef, title, toggleOrganizerHidden } = p;
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
              <h2 className="dex-step-head-title">
                {isDe ? 'Schritt 2 — Organizer & Team' : 'Step 2 — Organizers & Team'}
              </h2>
              <p className="dex-step-head-lead">
                {isDe
                  ? 'Wer verantwortet das Event (Organizer), wer ist Ansprechpartner und welches erweiterte Team (Test-Team, Check-in-Team) ist beteiligt.'
                  : 'Who runs the event (organizers), who is the contact, and which extended team (test team, check-in team) is involved.'}
              </p>

              {/* v24.4 (G): Zwischenüberschrift „Organizer" mit grünem Balken. */}
              <div style={{ borderLeft: '4px solid var(--dex-green)', padding: '4px 0 4px 12px', margin: '4px 0 18px' }}>
                <div style={{ fontWeight: 800, fontSize: '1.02rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'Organizer' : 'Organizers'}</div>
                <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                  {isDe
                    ? 'Diese Personen verantworten die Eventkoordination, können das Event ändern und die Teilnehmerliste einsehen.'
                    : 'These people are responsible for coordinating the event, can edit it and view the attendee list.'}
                </div>
              </div>
              <div className="form-group" style={{ position: 'relative', paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
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
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ marginLeft: 'auto', fontSize: '0.75rem', padding: '4px 10px', whiteSpace: 'nowrap' }}
                    onClick={() => setBulkOrganizerOpen(true)}
                  >
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                      <Users size={12} /> Massenimport
                    </span>
                  </button>
                </label>
                {/* Mismatch-Warning: bei Legacy-Korruption aus v10.0–v10.2-Closure-Bug
                    haben Events mehr Namen als Emails (oder umgekehrt). Auto-Heal padded
                    inzwischen statt zu truncaten — die Chips zeigen alle Namen, aber bei
                    welchen die Email fehlt, kann der User es hier nachpflegen. */}
                {(() => {
                  const orgList = organizer.split(';').map(s => s.trim()).filter(Boolean);
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
                        ? <>Bei diesen Personen fehlen die Mails fürs <strong>BCC</strong> der Anmelde-/Abmelde-Mails, die <strong>Outlook-Einladung</strong> und die <strong>Decline-/Forward-Notifications</strong>. Bitte entferne die betroffenen Chips (X) und füge sie über den Picker oder Massenimport neu ein. <em>(Ursache: Legacy-Daten aus einer früheren App-Version — wird beim nächsten Speichern geheilt.)</em></>
                        : <>For these people the emails for the <strong>BCC</strong> of registration/cancellation mails, the <strong>Outlook invitation</strong> and the <strong>decline/forward notifications</strong> are missing. Please remove the affected chips (X) and re-add them via the picker or bulk import. <em>(Cause: legacy data from an earlier app version — healed on the next save.)</em></>}
                    </WizardHint>
                  );
                })()}
                {/* Organizer-Chips (immer sichtbar wenn 1+ Organizer) */}
                {(() => {
                  const orgList = organizer.split(';').map(s => s.trim()).filter(Boolean);
                  if (orgList.length === 0) return null;
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
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 10 }}>
                      {orgList.map((name, i) => {
                        const email = organizerEmails[i] || '';
                        // v24.12 (J-Gate): Einzel-Ausblenden nur wenn „Organizer
                        // einzeln ausblenden" aktiv ist; sonst Chips nicht klickbar.
                        const canHideToggle = hideOrganizer && hideOrganizerIndividualOnly && !!email;
                        const orgHidden = hideOrganizer && hideOrganizerIndividualOnly && !!email && hiddenOrganizerEmails.indexOf(email.toLowerCase()) >= 0;
                        return (
                        <span
                          key={`${name}-${i}`}
                          style={{
                            display: 'inline-flex', alignItems: 'center', gap: 6,
                            padding: '3px 6px 3px 4px',
                            background: orgHidden ? 'var(--dex-gray-400, #9aa0a6)' : 'var(--dex-green)', color: '#fff',
                            borderRadius: 999, fontSize: '0.85rem', fontWeight: 500,
                            opacity: orgHidden ? 0.85 : 1,
                          }}
                        >
                          {email ? (
                            <img
                              src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=S`}
                              alt={name}
                              onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
                              style={{ width: 24, height: 24, borderRadius: '50%', objectFit: 'cover', background: 'rgba(255,255,255,0.25)', filter: orgHidden ? 'grayscale(1)' : 'none' }}
                            />
                          ) : null}
                          {/* v24.8 (J): Klick auf den Namen blendet diesen Organizer
                              auf der Anmelde-Seite aus/ein (Rechte bleiben). */}
                          <span
                            onClick={() => { if (canHideToggle) toggleOrganizerHidden(email); }}
                            title={canHideToggle ? (isDe ? (orgHidden ? 'Wird auf der Anmeldeseite NICHT angezeigt — klicken zum Einblenden (Rechte bleiben)' : 'Klicken, um diesen Organizer auf der Anmeldeseite auszublenden (Rechte bleiben)') : (orgHidden ? 'Hidden on the registration page — click to show' : 'Click to hide this organizer on the registration page')) : (isDe ? 'Zum einzelnen Ausblenden „Organizer ausblenden" und „Nur einzelne Organizer ausblenden" aktivieren' : 'Enable „Hide organizers" and „Hide only individual organizers" first')}
                            style={{ cursor: canHideToggle ? 'pointer' : 'default', textDecoration: orgHidden ? 'line-through' : 'none' }}
                          >{name}</span>
                          {orgHidden && <span style={{ fontSize: '0.68rem', fontStyle: 'italic', opacity: 0.95 }}>{isDe ? '(ausgeblendet)' : '(hidden)'}</span>}
                          {/* v28.5: Rückfragen-Kontakt markieren — oranger
                              Kreis. Genau EINER; erneuter Klick entfernt die
                              Markierung. v28.10: Sprechblase statt „?". */}
                          {!!email && (
                            <button
                              type="button"
                              onClick={() => setContactOrganizerEmail(prev => (prev || '').toLowerCase() === email.toLowerCase() ? '' : email)}
                              style={{
                                background: (contactOrganizerEmail || '').toLowerCase() === email.toLowerCase() ? 'var(--dex-orange, #ed8b00)' : 'rgba(255,255,255,0.2)',
                                border: (contactOrganizerEmail || '').toLowerCase() === email.toLowerCase() ? '1.5px solid #fff' : 'none',
                                color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer',
                                lineHeight: 1,
                                display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                              }}
                              title={isDe
                                ? ((contactOrganizerEmail || '').toLowerCase() === email.toLowerCase()
                                  ? 'Rückfragen-Kontakt — Teilnehmer werden gebeten, diese Person bei Fragen zu kontaktieren. Klicken zum Entfernen.'
                                  : 'Als Rückfragen-Kontakt markieren (oranger Badge auf der Anmeldeseite)')
                                : ((contactOrganizerEmail || '').toLowerCase() === email.toLowerCase()
                                  ? 'Contact for questions — attendees are asked to reach out to this person. Click to remove.'
                                  : 'Mark as contact for questions (orange badge on the registration page)')}
                            ><Icon iconName="Chat" style={{ fontSize: 11 }} /></button>
                          )}
                          {orgList.length > 1 && i > 0 && (
                            <button
                              type="button"
                              onClick={() => move(i, -1)}
                              style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.75rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
                              title="Nach vorne"
                            >◀</button>
                          )}
                          {orgList.length > 1 && i < orgList.length - 1 && (
                            <button
                              type="button"
                              onClick={() => move(i, 1)}
                              style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.75rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
                              title="Nach hinten"
                            >▶</button>
                          )}
                          <button
                            type="button"
                            onClick={() => remove(i)}
                            style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.9rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
                            title="Entfernen"
                          >×</button>
                        </span>
                      );
                      })}
                      {/* v28.5: Legende zum Rückfragen-Kontakt. v28.10:
                          Sprechblasen-Icon statt „?". */}
                      <span style={{ flexBasis: '100%', display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.74rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                        <span style={{ width: 16, height: 16, borderRadius: '50%', background: (contactOrganizerEmail || '').trim() ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-400, #9aa0a6)', color: '#fff', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                          <Icon iconName="Chat" style={{ fontSize: 8 }} />
                        </span>
                        <span>
                          {(contactOrganizerEmail || '').trim()
                            ? (isDe
                              ? 'Rückfragen-Kontakt markiert — diese Person bekommt auf der Anmeldeseite den orangen Sprechblasen-Badge. Klick auf ihre Sprechblase entfernt die Markierung.'
                              : 'Contact for questions marked — this person gets the orange chat-bubble badge on the registration page. Click their bubble again to remove it.')
                            : (isDe
                              ? 'Tipp: Klicke die Sprechblase an einem Organizer, um ihn als Rückfragen-Kontakt zu markieren — er bekommt auf der Anmeldeseite einen orangen Badge mit Hinweis.'
                              : 'Tip: click the chat bubble on an organizer to mark them as the contact for questions — they get an orange badge with a note on the registration page.')}
                        </span>
                      </span>
                      {/* v24.8 (J) / v24.15 (Gate): Tipp unter den Organizer-Namen. */}
                      <span style={{ flexBasis: '100%', fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                        {isDe
                          ? (!hideOrganizer
                              ? 'Alle Organizer werden auf der Anmeldeseite angezeigt.'
                              : (hideOrganizerIndividualOnly
                                  ? 'Tipp: Auf einen Namen klicken blendet diese Person auf der Anmeldeseite aus (sie behält alle Rechte). Erneut klicken zeigt sie wieder.'
                                  : 'Alle Organizer sind ausgeblendet. Bei mehreren Organizern kannst du oben „Nur einzelne Organizer ausblenden" wählen, um stattdessen einzelne anzuklicken.'))
                          : (!hideOrganizer
                              ? 'All organizers are shown on the registration page.'
                              : (hideOrganizerIndividualOnly
                                  ? 'Tip: click a name to hide that person on the registration page (they keep all rights). Click again to show.'
                                  : 'All organizers are hidden. With several organizers you can choose „Hide individual organizers only" above to pick individuals instead.'))}
                      </span>
                    </div>
                  );
                })()}
                <input
                  className="form-input"
                  value={organizerSearch}
                  onChange={e => {
                    const val = e.target.value;
                    setOrganizerSearch(val);
                    if (organizerTimerRef.current) clearTimeout(organizerTimerRef.current);
                    const q = val.trim();
                    if (!q) { setOrganizerResults([]); return; }
                    // v9.20: Graph-Search statt Role-Filter — jeder Deloitte-User
                    // kann als Organizer hinzugefügt werden. Damit gibt es nur
                    // einen Picker (kein extra Co-Organizer); der erste in der
                    // Liste ist der "Hauptorganizer", weitere sind gleichwertig.
                    organizerTimerRef.current = setTimeout(async () => {
                      try {
                        const results = await searchUsers(q, organizerIncludeIntl);
                        setOrganizerResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setOrganizerResults([]); }
                    }, 350);
                  }}
                  onBlur={() => {
                    setTimeout(() => { setOrganizerSearch(''); setOrganizerResults([]); }, 150);
                  }}
                  placeholder={t('create.organizer.placeholder')}
                  style={errorBorderStyle('organizer')}
                />
                <InternationalSearchToggle
                  query={organizerSearch}
                  checked={organizerIncludeIntl}
                  onChange={async next => {
                    setOrganizerIncludeIntl(next);
                    const q = organizerSearch.trim();
                    if (q.length >= 1) {
                      try {
                        const results = await searchUsers(q, next);
                        setOrganizerResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setOrganizerResults([]); }
                    }
                  }}
                  isDe={isDe}
                />
                {isSearchingOrganizer && (
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>Suche...</div>
                )}
                {organizerResults.length > 0 && (
                  <div style={{
                    position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
                    background: '#fff', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                    maxHeight: 280, overflowY: 'auto',
                  }}>
                    {organizerResults.map(u => {
                      const alreadyAdded = organizerEmails.indexOf(u.email) >= 0;
                      return (
                        <div
                          key={u.email}
                          style={{
                            padding: '8px 12px', cursor: alreadyAdded ? 'not-allowed' : 'pointer', fontSize: '0.85rem',
                            borderBottom: '1px solid var(--dex-gray-100)',
                            opacity: alreadyAdded ? 0.45 : 1,
                            display: 'flex', alignItems: 'center', gap: 10,
                          }}
                          onMouseDown={() => {
                            if (alreadyAdded) return;
                            const existing = organizer.split(';').map(s => s.trim()).filter(Boolean);
                            const nextNames = [...existing, u.displayName];
                            setOrganizer(nextNames.join('; '));
                            setOrganizerEmails(prev => [...prev, u.email]);
                            setOrganizerSearch('');
                            setOrganizerResults([]);
                          }}
                        >
                          <img
                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(u.email)}&size=S`}
                            alt={u.displayName}
                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                            style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                          />
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontWeight: 600 }}>{u.displayName}</div>
                            <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                              {u.email}{u.location ? ` · ${u.location}` : ''}
                            </div>
                          </div>
                          {alreadyAdded && <span style={{ color: 'var(--dex-green)', fontSize: '0.85rem', flexShrink: 0 }}>✓</span>}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* v18.9: Organizer-Anzeige ausblenden. Rein visuell — die
                  Organizer behalten alle Rechte + Mail-Benachrichtigungen,
                  werden aber auf der Anmelde-Seite und in „Meine Events"
                  nicht als Ansprechpartner-Chips gezeigt. */}
              <div className="form-group" style={{ paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={hideOrganizer}
                    onChange={e => setHideOrganizer(e.target.checked)}
                    style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 2 }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Organizer ausblenden' : 'Hide organizers'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> ob die <strong>Organizer</strong> auf der <strong>Anmelde-Seite</strong> (und &bdquo;Meine Events&ldquo;) als Ansprechpartner angezeigt werden.<br /><br />
                          <strong>Anzeige in der App:</strong> aktiviert blendet alle Organizer aus. Gibt es mehrere Organizer, erscheint darunter zusätzlich die Option, stattdessen <strong>nur einzelne</strong> auszublenden (per Klick auf den Namen).<br /><br />
                          <strong>Wichtig:</strong> das ist rein optisch — die Organizer behalten alle <strong>Rechte</strong> (bearbeiten, Teilnehmer verwalten) und ihre <strong>Mail-Benachrichtigungen</strong>.
                        </>
                      : <>
                          <strong>What this controls:</strong> whether the <strong>organizers</strong> are shown as contacts on the <strong>registration page</strong> (and &bdquo;My Events&ldquo;).<br /><br />
                          <strong>Where you see it:</strong> when enabled, all organizers are hidden. With several organizers, an option appears below to hide <strong>only individual ones</strong> instead (click a name).<br /><br />
                          <strong>Note:</strong> this is purely visual — organizers keep all <strong>permissions</strong> and their <strong>email notifications</strong>.
                        </>
                    } />
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? 'Default: aus — alle Organizer werden angezeigt. An: keine Organizer auf der Anmeldeseite (der optionale Ansprechpartner bleibt).'
                        : 'Default: off — all organizers are shown. On: no organizers on the registration page (the optional contact stays).'}
                    </span>
                  </span>
                </label>
                {/* v24.15: zweite Checkbox — nur bei „Organizer ausblenden" UND
                    mehreren Organizern: stattdessen nur einzelne ausblenden. */}
                {hideOrganizer && organizerEmails.length >= 2 && (
                  <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', marginTop: 12, marginLeft: 28 }}>
                    <input
                      type="checkbox"
                      checked={hideOrganizerIndividualOnly}
                      onChange={e => setHideOrganizerIndividualOnly(e.target.checked)}
                      style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 2 }}
                    />
                    <span style={{ flex: 1 }}>
                      <strong>{isDe ? 'Nur einzelne Organizer ausblenden' : 'Hide only individual organizers'}</strong>
                      <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                        {isDe
                          ? 'An: nicht alle ausblenden — klicke unten auf einzelne Organizer-Namen, um genau diese auszublenden (die übrigen bleiben sichtbar).'
                          : 'On: don’t hide all — click individual organizer names below to hide exactly those (the rest stay visible).'}
                      </span>
                    </span>
                  </label>
                )}
              </div>

              {/* v23.25/v24.15: Darstellungs-Größe der Organizer. Nur relevant,
                  wenn überhaupt Organizer angezeigt werden — also NICHT, wenn
                  ALLE ausgeblendet sind (hideOrganizer ohne Einzel-Modus). */}
              {(!hideOrganizer || hideOrganizerIndividualOnly) && (
                <div className="form-group" style={{ paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                  {/* v24.4 (I) / v24.10 (Q): grüne Zwischenüberschrift + Live-Vorschau. */}
                  <div className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, fontWeight: 800, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                    {isDe ? 'Anzeige auf dem Anmeldeformular' : 'Display on the registration form'}
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
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginTop: 4 }}>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                      <input
                        type="radio"
                        name="organizerDisplaySize"
                        checked={!organizerDisplayLarge}
                        onChange={() => setOrganizerDisplayLarge(false)}
                        style={{ marginTop: 3, cursor: 'pointer' }}
                      />
                      <span style={{ flex: 1 }}>
                        <strong>{isDe ? 'Klein (Chip)' : 'Small (chip)'}</strong>
                        <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                          {isDe
                            ? 'Kompakter Chip mit Foto + Name. Mail-Adresse und Rolle erscheinen beim Drüberfahren mit der Maus (Mouse-over) groß.'
                            : 'Compact chip with photo + name. Email and role appear large on mouse-over.'}
                        </span>
                      </span>
                    </label>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                      <input
                        type="radio"
                        name="organizerDisplaySize"
                        checked={organizerDisplayLarge}
                        onChange={() => setOrganizerDisplayLarge(true)}
                        style={{ marginTop: 3, cursor: 'pointer' }}
                      />
                      <span style={{ flex: 1 }}>
                        <strong>{isDe ? 'Groß' : 'Large'}</strong>
                        <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                          {isDe
                            ? 'Großes Foto, Name, klickbare E-Mail-Adresse und Rolle direkt sichtbar — ohne Mouse-over.'
                            : 'Large photo, name, clickable email and role visible right away — no mouse-over.'}
                        </span>
                      </span>
                    </label>
                  </div>
                  {/* v24.10 (Q): Live-Vorschau — so sieht der Teilnehmer die Organizer. */}
                  {(() => {
                    const _orgNames = organizer.split(';').map(s => s.trim()).filter(Boolean);
                    if (_orgNames.length === 0) return null;
                    return (
                      <div style={{ marginTop: 14, padding: 14, background: 'var(--dex-gray-50, #f7f8f9)', border: '1px dashed var(--dex-gray-300)', borderRadius: 10 }}>
                        <div style={{ fontSize: '0.7rem', fontWeight: 700, textTransform: 'uppercase', letterSpacing: 0.3, color: 'var(--dex-gray-500)', marginBottom: 8 }}>
                          {isDe ? 'Vorschau (so sehen es die Teilnehmer)' : 'Preview (what attendees see)'}
                        </div>
                        <OrganizerList
                          names={_orgNames}
                          emails={organizerEmails}
                          hiddenEmails={(hideOrganizer && hideOrganizerIndividualOnly) ? hiddenOrganizerEmails : []}
                          size="md"
                          display={organizerDisplayLarge ? 'card' : 'chip'}
                          contactEmail={contactOrganizerEmail || undefined}
                        />
                      </div>
                    );
                  })()}
                </div>
              )}

              {/* v10.16: Optionaler Ansprechpartner. Reines Anzeige-Feld
                  (kein Login, keine SP-Permissions) — z.B. die Person vor Ort
                  oder eine Hotline-Mail die Teilnehmer bei Fragen anschreiben
                  sollen. Wird auf Register-/MyEvents-Page zusätzlich zu den
                  Organizern gezeigt. Alles drei optional, Freitext. */}
              <div className="form-group" style={{ paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                {/* v24.10 (Q2): einklappbar, default zu. */}
                <label
                  className="form-label"
                  onClick={() => setContactExpanded(v => !v)}
                  style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', userSelect: 'none' }}
                >
                  <span style={{
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    width: 22, height: 22, borderRadius: '50%',
                    background: 'var(--dex-gray-300)', color: '#fff',
                    fontSize: '0.75rem', fontWeight: 700,
                  }}>{contactExpanded ? '–' : '+'}</span>
                  {isDe ? 'Ansprechpartner (optional)' : 'Contact person (optional)'}
                </label>
                {contactExpanded && (
                <>
                {/* v24.10 (Q3): Organizer ist Standard-Ansprechpartner; hier nur Externe. */}
                <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', margin: '0 0 10px', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Standardmäßig wird der Organizer für Rückfragen aus dem Team angezeigt. Wenn du stattdessen (oder zusätzlich) jemand Externen angeben willst — z.B. eine Service-Mailadresse oder eine Kontaktperson vor Ort — dann hier eintragen. Erscheint auf der Anmelde-Seite und in „Meine Events" zusätzlich zu den Organizern. Hat KEINE App-Berechtigung, ist nur ein Anzeige-Feld.'
                    : 'By default the organizer is shown as the contact for questions from the team. If you want to add an external contact instead (or in addition) — e.g. a service email or an on-site contact — enter it here. Appears on the registration page and in „My Events" in addition to the organizers. Has NO app permissions, display-only.'}
                </p>
                <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 10 }}>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                      {isDe ? 'Name' : 'Name'}
                    </label>
                    <input
                      type="text"
                      className="form-input"
                      value={contactName}
                      onChange={e => setContactName(e.target.value)}
                      placeholder={isDe ? 'z.B. Anna Schmitt' : 'e.g. Anna Schmitt'}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                      {isDe ? 'E-Mail' : 'Email'}
                    </label>
                    <input
                      type="email"
                      className="form-input"
                      value={contactEmail}
                      onChange={e => setContactEmail(e.target.value)}
                      placeholder={isDe ? 'z.B. event-helpdesk@example.de' : 'e.g. event-helpdesk@example.com'}
                    />
                  </div>
                </div>
                <div>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                    {isDe ? 'Zusatz-Info / Erreichbarkeit (Freitext)' : 'Additional info / availability (free text)'}
                  </label>
                  <textarea
                    className="form-input"
                    value={contactInfo}
                    onChange={e => setContactInfo(e.target.value)}
                    rows={3}
                    placeholder={isDe
                      ? 'z.B. „Vor Ort am Eventtag ab 7:30 Uhr, mobil unter +49 151 123 456" oder „Bei Fragen vor dem Event direkt per Mail."'
                      : 'e.g. „On-site from 7:30 am on event day, mobile +49 151 123 456" or „For questions before the event, email directly."'}
                  />
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
                </>
                )}
              </div>

              {/* v9.21: Test-Team pro Event — sieht das Event im Entwurfsmodus
                  und kann sich anmelden, ohne globale Organizer-Rolle. Picker
                  via Graph-Search, beliebige Deloitte-User. */}
              {/* v24.4 (G): Zwischenüberschrift „Erweitertes Organisations-Team". */}
              <div style={{ borderLeft: '4px solid var(--dex-green)', padding: '4px 0 4px 12px', margin: '4px 0 18px' }}>
                <div style={{ fontWeight: 800, fontSize: '1.02rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'Erweitertes Organisations-Team' : 'Extended organization team'}</div>
                <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                  {isDe
                    ? 'Diese Personen unterstützen die Organizer. Sie können das Event nicht anpassen und die Teilnehmerliste nur bedingt einsehen (Check-in-Team beim Check-in).'
                    : 'These people support the organizers. They cannot edit the event and can only view the attendee list to a limited extent (check-in team during check-in).'}
                </div>
              </div>
              <div className="form-group" style={{ position: 'relative', paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={12} />
                  Test-Team
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
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ marginLeft: 'auto', fontSize: '0.75rem', padding: '4px 10px', whiteSpace: 'nowrap' }}
                    onClick={() => setBulkTestTeamOpen(true)}
                  >
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                      <Users size={12} /> Massenimport
                    </span>
                  </button>
                </label>
                {/* v24.4 (L): Kernaussage inline (ohne Mouse-over), Rest im Info-Symbol. */}
                <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Eine kleine Gruppe, die das Event schon im Entwurfsmodus sieht und sich testweise anmelden darf — bevor du es für die echte Zielgruppe freigibst.'
                    : 'A small group that can see the event already in draft mode and register as a test — before you publish it to the real audience.'}
                </p>
                {testTeamNames.length > 0 && (() => {
                  const remove = (idx: number): void => {
                    setTestTeamNames(testTeamNames.filter((_, i) => i !== idx));
                    setTestTeamEmails(testTeamEmails.filter((_, i) => i !== idx));
                  };
                  return (
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 10 }}>
                      {testTeamNames.map((name, i) => {
                        const email = testTeamEmails[i] || '';
                        return (
                          <span key={`${email}-${i}`} style={{
                            display: 'inline-flex', alignItems: 'center', gap: 6,
                            padding: '3px 6px 3px 4px',
                            background: '#0ea5e9', color: '#fff',
                            borderRadius: 999, fontSize: '0.85rem', fontWeight: 500,
                          }}>
                            {email && (
                              <img
                                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=S`}
                                alt={name}
                                onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
                                style={{ width: 24, height: 24, borderRadius: '50%', objectFit: 'cover', background: 'rgba(255,255,255,0.25)' }}
                              />
                            )}
                            <span>{name}</span>
                            <button type="button" onClick={() => remove(i)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.9rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }} title="Entfernen">×</button>
                          </span>
                        );
                      })}
                    </div>
                  );
                })()}
                <input
                  className="form-input"
                  value={testTeamSearch}
                  onChange={e => {
                    const val = e.target.value;
                    setTestTeamSearch(val);
                    if (testTeamTimerRef.current) clearTimeout(testTeamTimerRef.current);
                    const q = val.trim();
                    if (!q) { setTestTeamResults([]); return; }
                    testTeamTimerRef.current = setTimeout(async () => {
                      try {
                        const results = await searchUsers(q, testTeamIncludeIntl);
                        setTestTeamResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setTestTeamResults([]); }
                    }, 350);
                  }}
                  onBlur={() => {
                    setTimeout(() => { setTestTeamSearch(''); setTestTeamResults([]); }, 150);
                  }}
                  placeholder="Name oder E-Mail eingeben (alle Deloitte-User)"
                />
                <InternationalSearchToggle
                  query={testTeamSearch}
                  checked={testTeamIncludeIntl}
                  onChange={async next => {
                    setTestTeamIncludeIntl(next);
                    const q = testTeamSearch.trim();
                    if (q.length >= 1) {
                      try {
                        const results = await searchUsers(q, next);
                        setTestTeamResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setTestTeamResults([]); }
                    }
                  }}
                  isDe={isDe}
                />
                {testTeamResults.length > 0 && (
                  <div style={{
                    position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
                    background: '#fff', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                    maxHeight: 280, overflowY: 'auto',
                  }}>
                    {testTeamResults.map(u => {
                      const alreadyAdded = testTeamEmails.indexOf(u.email) >= 0;
                      return (
                        <div key={u.email} style={{
                          padding: '8px 12px', cursor: alreadyAdded ? 'not-allowed' : 'pointer', fontSize: '0.85rem',
                          borderBottom: '1px solid var(--dex-gray-100)',
                          opacity: alreadyAdded ? 0.45 : 1,
                          display: 'flex', alignItems: 'center', gap: 10,
                        }} onMouseDown={() => {
                          if (alreadyAdded || !u.email) return;
                          setTestTeamNames(prev => [...prev, u.displayName]);
                          setTestTeamEmails(prev => [...prev, u.email]);
                          setTestTeamSearch('');
                          setTestTeamResults([]);
                        }}>
                          <img
                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(u.email)}&size=S`}
                            alt={u.displayName}
                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                            style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                          />
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontWeight: 600 }}>{u.displayName}</div>
                            <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                              {u.email}{u.location ? ` · ${u.location}` : ''}
                            </div>
                          </div>
                          {alreadyAdded && <span style={{ color: 'var(--dex-green)', fontSize: '0.85rem', flexShrink: 0 }}>✓</span>}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* v6.19: QR-Code-Scanner pro Event. Separater Picker im gleichen Stil wie
                  Organizer, aber via Graph-Search (jeder Deloitte-User kann Scanner sein).
                  QR-Scanner haben eingeschränkten Admin-Zugriff (nur QR-Tool + KPIs),
                  erscheinen NICHT in Organizer-Listen auf MyEvents/RegistrationPage und
                  bekommen KEINE Organizer-Mails. */}
              <div className="form-group" style={{ position: 'relative', paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
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
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ marginLeft: 'auto', fontSize: '0.75rem', padding: '4px 10px', whiteSpace: 'nowrap' }}
                    onClick={() => setBulkQrScannerOpen(true)}
                  >
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                      <Users size={12} /> Massenimport
                    </span>
                  </button>
                </label>
                {/* v24.4 (L): Kernaussage inline (ohne Mouse-over), Rest im Info-Symbol. */}
                <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Personen, die am Event-Tag nur das Check-in-Tool bedienen dürfen (QR-Codes scannen, Teilnehmer ein-/auschecken) — sonst keine weiteren Rechte.'
                    : 'People who may only operate the check-in tool on the event day (scan QR codes, check attendees in/out) — no other rights.'}
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
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 10 }}>
                      {qrScannerNames.map((name, i) => {
                        const email = qrScannerEmails[i] || '';
                        return (
                          <span
                            key={`${email}-${i}`}
                            style={{
                              display: 'inline-flex', alignItems: 'center', gap: 6,
                              padding: '3px 6px 3px 4px',
                              background: 'var(--dex-orange, #ed8b00)', color: '#fff',
                              borderRadius: 999, fontSize: '0.85rem', fontWeight: 500,
                            }}
                          >
                            {email ? (
                              <img
                                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=S`}
                                alt={name}
                                onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
                                style={{ width: 24, height: 24, borderRadius: '50%', objectFit: 'cover', background: 'rgba(255,255,255,0.25)' }}
                              />
                            ) : null}
                            <span>{name}</span>
                            {qrScannerNames.length > 1 && i > 0 && (
                              <button type="button" onClick={() => move(i, -1)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.75rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }} title="Nach vorne">◀</button>
                            )}
                            {qrScannerNames.length > 1 && i < qrScannerNames.length - 1 && (
                              <button type="button" onClick={() => move(i, 1)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.75rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }} title="Nach hinten">▶</button>
                            )}
                            <button type="button" onClick={() => remove(i)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.9rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }} title="Entfernen">×</button>
                          </span>
                        );
                      })}
                    </div>
                  );
                })()}
                <input
                  className="form-input"
                  value={qrScannerSearch}
                  onChange={e => {
                    const val = e.target.value;
                    setQrScannerSearch(val);
                    if (qrScannerTimerRef.current) clearTimeout(qrScannerTimerRef.current);
                    const q = val.trim();
                    if (!q) { setQrScannerResults([]); return; }
                    // v9.18: Graph-Search statt Role-Filter — jeder Deloitte-User
                    // kann QR-Scanner sein. Debounce 350ms.
                    qrScannerTimerRef.current = setTimeout(async () => {
                      try {
                        const results = await searchUsers(q, qrScannerIncludeIntl);
                        setQrScannerResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setQrScannerResults([]); }
                    }, 350);
                  }}
                  onBlur={() => {
                    setTimeout(() => { setQrScannerSearch(''); setQrScannerResults([]); }, 150);
                  }}
                  placeholder={t('create.qrscanners.placeholder') || 'Name oder E-Mail eingeben (alle Deloitte-User)'}
                />
                <InternationalSearchToggle
                  query={qrScannerSearch}
                  checked={qrScannerIncludeIntl}
                  onChange={async next => {
                    setQrScannerIncludeIntl(next);
                    const q = qrScannerSearch.trim();
                    if (q.length >= 1) {
                      try {
                        const results = await searchUsers(q, next);
                        setQrScannerResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setQrScannerResults([]); }
                    }
                  }}
                  isDe={isDe}
                />
                {qrScannerResults.length > 0 && (
                  <div style={{
                    position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
                    background: '#fff', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                    maxHeight: 280, overflowY: 'auto',
                  }}>
                    {qrScannerResults.map(u => {
                      const alreadyAdded = qrScannerEmails.indexOf(u.email) >= 0;
                      return (
                        <div
                          key={u.email}
                          style={{
                            padding: '8px 12px', cursor: alreadyAdded ? 'not-allowed' : 'pointer', fontSize: '0.85rem',
                            borderBottom: '1px solid var(--dex-gray-100)',
                            opacity: alreadyAdded ? 0.45 : 1,
                            display: 'flex', alignItems: 'center', gap: 10,
                          }}
                          onMouseDown={() => {
                            if (alreadyAdded || !u.email) return;
                            setQrScannerNames(prev => [...prev, u.displayName]);
                            setQrScannerEmails(prev => [...prev, u.email]);
                            setQrScannerSearch('');
                            setQrScannerResults([]);
                          }}
                        >
                          <img
                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(u.email)}&size=S`}
                            alt={u.displayName}
                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                            style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                          />
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontWeight: 600 }}>{u.displayName}</div>
                            <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                              {u.email}{u.location ? ` · ${u.location}` : ''}
                            </div>
                          </div>
                          {alreadyAdded && <span style={{ color: 'var(--dex-green)', fontSize: '0.85rem', flexShrink: 0 }}>✓</span>}
                        </div>
                      );
                    })}
                  </div>
                )}
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
                    style={{ marginBottom: 20 }}
                  >
                    <p style={{ margin: '0 0 10px' }}>
                      Co-Organizer dürfen automatisch das <strong>Check-In-Tool</strong> nutzen und sehen
                      Events auch im <strong>Entwurfsmodus</strong> — ein zusätzlicher Eintrag im Test-Team oder
                      Check-In-Team ist daher nicht nötig. Du kannst die überflüssigen Einträge hier auf
                      einen Klick entfernen.
                    </p>
                    <ul style={{ margin: 0, padding: 0, listStyle: 'none' }}>
                      {dups.map(d => {
                        const teamsLabel: string[] = [];
                        if (d.inOrg) teamsLabel.push('Co-Organizer');
                        if (d.inTt) teamsLabel.push('Test-Team');
                        if (d.inQr) teamsLabel.push('Check-In-Team');
                        return (
                          <li
                            key={d.email}
                            style={{
                              padding: '8px 0',
                              borderTop: '1px solid var(--dex-orange, #ed8b00)',
                              display: 'flex',
                              alignItems: 'center',
                              gap: 10,
                              flexWrap: 'wrap',
                            }}
                          >
                            <div style={{ flex: '1 1 200px', minWidth: 0 }}>
                              <strong>{d.name}</strong>{' '}
                              <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>{d.email}</span>
                              <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
                                Aktuell in: {teamsLabel.join(', ')}
                              </div>
                            </div>
                            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                              {d.inTt && (
                                <button
                                  type="button"
                                  className="btn btn-secondary"
                                  style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                  onClick={() => removeFromTestTeam(d.email)}
                                >
                                  Aus Test-Team entfernen
                                </button>
                              )}
                              {d.inQr && (
                                <button
                                  type="button"
                                  className="btn btn-secondary"
                                  style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                  onClick={() => removeFromQr(d.email)}
                                >
                                  Aus Check-In-Team entfernen
                                </button>
                              )}
                              {d.inTt && d.inQr && (
                                <button
                                  type="button"
                                  className="btn btn-primary"
                                  style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                  onClick={() => removeAllOverlap(d.email)}
                                >
                                  Aus beiden entfernen
                                </button>
                              )}
                            </div>
                          </li>
                        );
                      })}
                    </ul>
                  </WizardHint>
                );
              })()}

              </div>
  );
};
