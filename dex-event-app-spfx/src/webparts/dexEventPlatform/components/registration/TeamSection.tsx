/* Team-Bloecke der Anmeldeseite — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Die Karte „Ich melde mich + mein Team an" (v11.82) und darunter die Liste der
 * offenen Teams (v18.73). Inhalt zeichengleich uebernommen; die
 * Anzeige-Bedingungen sind beim Aufrufer geblieben.
 *
 * v31.9: Beide Wege nach dem UI-Leitfaden umgebaut — und BEWUSST in
 * derselben Form, obwohl sie getrennt bleiben. „Eigenes Team anmelden" und
 * „offenem Team beitreten" sind sachlich eine Frage (Leitfaden 6a), ihre
 * Sichtbarkeitsbedingungen beim Aufrufer sind aber NICHT deckungsgleich;
 * eine Zusammenlegung braucht zwei Dateien und eine neue gemeinsame
 * Bedingung und gehoert in ein eigenes Ticket. Gleiche Optik heisst: die
 * beiden passen spaeter zusammen, ohne dass jemand sie neu erfindet. */
import * as React from 'react';
import { CollapsibleSection } from './regHelpers';
import { Icon } from '@fluentui/react/lib/Icon';
import { Locale } from '../../context/LanguageContext';
import { DeloitteEvent, EventSpecificField } from '../../types';
import { UserFieldPicker } from '../UserFieldPicker';
import { cx } from '../dexUi';
import { AlertCircle, Users } from '../Icons';

/** Karte „Ich melde mich + mein Team an" (v11.82). */
export interface TeamRegistrationCardProps {
  errorBorder: { border: string; };
  event: DeloitteEvent;
  isMobile: boolean;
  isTeamMode: boolean;
  locale: Locale;
  parseTeamMember: (v: string) => { displayName: string; email: string; };
  renderRegField: (fRaw: EventSpecificField, store?: Record<string, string>, setStore?: (next: Record<string, string>) => void, rowIndex?: number, rowList?: EventSpecificField[]) => React.ReactElement;
  searchUser: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string; company?: string; }>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  setTeamConsentConfirmed: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamMemberFields: React.Dispatch<React.SetStateAction<Record<number, Record<string, string>>>>;
  setTeamMembers: React.Dispatch<React.SetStateAction<string[]>>;
  setTeamName: React.Dispatch<React.SetStateAction<string>>;
  showErrors: boolean;
  teamConsentConfirmed: boolean;
  teamMemberApplicableFields: EventSpecificField[];
  teamMemberFields: Record<number, Record<string, string>>;
  teamMembers: string[];
  teamName: string;
  teamPartialAllowed: boolean;
}
export const TeamRegistrationCard: React.FC<TeamRegistrationCardProps> = (p) => {
  const { errorBorder, event, isMobile, isTeamMode, locale, parseTeamMember, renderRegField, searchUser, searchUsers, setTeamConsentConfirmed, setTeamMemberFields, setTeamMembers, setTeamName, showErrors, teamConsentConfirmed, teamMemberApplicableFields, teamMemberFields, teamMembers, teamName, teamPartialAllowed } = p;
  return (
          <div className="registration-form" style={{ marginTop: 24 }}>
            <CollapsibleSection
              isMobile={isMobile}
              icon="People"
              title={locale === 'de' ? 'Team-Anmeldung' : 'Team registration'}
            >
            <div style={{ padding: '24px 20px' }}>
              {/* Pflicht-Hinweis-Box ganz oben — auffällig orange.
                  v31.9: aus dem handgebauten Kasten wird ein
                  `dex-ui-callout--warn` (Leitfaden 6c: Hinweiskästen nur als
                  Callout, Farbe gleich Bedeutung). Die Aufzählung ist zu
                  einem Satz zusammengezogen, damit der Kasten auf dem Handy
                  nicht die halbe Seite füllt — es geht keine Aussage
                  verloren, und die Zustimmungspflicht bleibt SICHTBAR statt
                  in einem Aufklapper (sie ist der Grund, warum es diesen
                  Kasten gibt). */}
              <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 20 }}>
                <span className="dex-ui-callout-icon"><AlertCircle size={18} /></span>
                <div className="dex-ui-callout-body">
                  <div style={{ fontWeight: 700, marginBottom: 4 }}>
                    {locale === 'de'
                      ? 'Haben alle Teammitglieder zugestimmt?'
                      : 'Has every team member agreed?'}
                  </div>
                  <div>
                    {locale === 'de'
                      ? 'Mit dem Absenden meldest du nicht nur dich selbst an, sondern auch jede unten eingetragene Person. Jede bekommt automatisch eine Anmeldebestätigung per Mail, einen Outlook-Termin im Kalender und sieht das Event in „Meine Events".'
                      : 'By submitting you register yourself AND every person you add below. Each one automatically receives a confirmation email, an Outlook calendar invite, and sees the event in “My Events”.'}
                  </div>
                  <div style={{ marginTop: 4 }}>
                    <strong>
                      {locale === 'de'
                        ? 'Hole die Zustimmung vorher ein — eine Anmeldung ohne Einverständnis ist nicht erlaubt.'
                        : 'Get their consent up front — registering people without their consent is not allowed.'}
                    </strong>
                  </div>
                </div>
              </div>

              {event?.askTeamName && (
                <div className="form-group">
                  {/* v31.9: Beschriftung als Frage statt als Feldname (2c). */}
                  <label className="form-label">
                    <span className="required">*</span> {locale === 'de' ? 'Wie heißt dein Team?' : 'What is your team called?'}
                  </label>
                  <input
                    className="form-input"
                    value={teamName}
                    onChange={e => setTeamName(e.target.value.slice(0, 60))}
                    placeholder={locale === 'de' ? 'z.B. „Die Schnellen"' : 'e.g. “The Quick Ones”'}
                    style={showErrors && isTeamMode && event?.askTeamName && !teamName.trim() ? errorBorder : {}}
                  />
                  <div className="dex-ui-help">
                    {locale === 'de' ? 'Max. 60 Zeichen — der Name steht später in der Teilnehmerliste und auf deiner Event-Karte.' : 'Max 60 characters — the name later appears on the participant list and on your event card.'}
                  </div>
                </div>
              )}

              {/* Member-Slots */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{locale === 'de' ? 'Wer gehört zum Team?' : 'Who is on the team?'}</div>
                {teamMembers.map((mv, idx) => {
                  const slotRequired = !teamPartialAllowed;
                  const parsed = parseTeamMember(mv);
                  const isErr = showErrors && isTeamMode && slotRequired && !parsed;
                  return (
                    <div className="form-group" key={`team-slot-${idx}`}>
                      <label className="form-label">
                        {slotRequired && <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>}
                        {locale === 'de'
                          ? `Mitglied ${idx + 2}${slotRequired ? '' : ' (optional)'}`
                          : `Member ${idx + 2}${slotRequired ? '' : ' (optional)'}`}
                      </label>
                      <UserFieldPicker
                        value={mv}
                        onChange={v => {
                          const next = [...teamMembers];
                          next[idx] = v;
                          setTeamMembers(next);
                        }}
                        searchUsers={async (q, includeIntl) => {
                          const results = await searchUsers(q, includeIntl);
                          return results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location, jobTitle: r.jobTitle }));
                        }}
                        searchUserByEmail={searchUser}
                        placeholder={locale === 'de' ? 'Name oder E-Mail eingeben...' : 'Type a name or email...'}
                        errorStyle={isErr ? errorBorder : {}}
                        forcedIsDe={locale === 'de'}
                      />
                      {/* v18.12: Custom-Fields pro Team-Mitglied — erscheinen,
                          sobald die Person ausgewählt ist (z.B. Essenspräferenz). */}
                      {parsed && teamMemberApplicableFields.length > 0 && (
                        <div style={{ marginTop: 8, marginLeft: 8, paddingLeft: 12, borderLeft: '2px solid var(--dex-gray-200)' }}>
                          {teamMemberApplicableFields
                            .filter(f => {
                              if (!f.showIf) return true;
                              const mv = teamMemberFields[idx] || {};
                              const v = mv[f.showIf.fieldId] || '';
                              const parts = v.split(' | ').map(s => s.trim());
                              return f.showIf.values.some(x => v === x || parts.indexOf(x) >= 0);
                            })
                            .map(f => renderRegField(
                              f,
                              teamMemberFields[idx] || {},
                              next => setTeamMemberFields(prev => ({ ...prev, [idx]: next }))
                            ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* Pflicht-Bestätigungs-Checkbox
                  v31.9: als `dex-ui-toggle-row` (2b: Ja/Nein mit Erklärung).
                  Die rohe Checkbox hatte keinen Hover und keine Fläche — auf
                  dem Handy war nur der kleine Kasten selbst ein Ziel. Der
                  Bestätigungssatz bleibt wortgleich: er ist die
                  Einverständnis-Erklärung, kein Anzeigetext. */}
              <div className="dex-ui-section">
                <label className={cx('dex-ui-toggle-row', teamConsentConfirmed && 'is-active')}>
                  <input
                    type="checkbox"
                    checked={teamConsentConfirmed}
                    onChange={e => setTeamConsentConfirmed(e.target.checked)}
                  />
                  <span className="dex-ui-toggle-row-body">
                    <span className="dex-ui-toggle-row-title">
                      <span className="required" style={{ color: 'var(--dex-red)' }}>*</span>
                      {locale === 'de'
                        ? 'Ich bestätige, dass alle eingetragenen Teammitglieder ihrer Anmeldung zugestimmt haben.'
                        : 'I confirm that every listed team member has consented to this registration.'}
                    </span>
                    <span className="dex-ui-toggle-row-desc">
                      {locale === 'de'
                        ? 'Ohne diese Bestätigung kannst du das Team nicht anmelden.'
                        : 'Without this confirmation you cannot register the team.'}
                    </span>
                  </span>
                </label>
              </div>
            </div>
            </CollapsibleSection>
          </div>
  );
};

/** Liste der Teams mit freien Slots (v18.73). */
export interface OpenTeamsListProps {
  event: DeloitteEvent;
  isMobile: boolean;
  locale: Locale;
  openTeams: { teamId: string; teamName: string; activeCount: number; teamSize: number; leadEmail: string; leadDisplayName: string; }[];
  pendingJoinTeam: { teamId: string; teamName: string; };
  togglePendingJoinTeam: (teamId: string, teamName: string) => void;
}
export const OpenTeamsList: React.FC<OpenTeamsListProps> = (p) => {
  const { event, isMobile, locale, openTeams, pendingJoinTeam, togglePendingJoinTeam } = p;
  return (
          <div className="registration-form" style={{ marginBottom: 16 }}>
            <CollapsibleSection
              isMobile={isMobile}
              icon="People"
              title={locale === 'de' ? 'Offene Teams — einem unvollständigen Team beitreten' : 'Open teams — join an incomplete team'}
            >
            <div style={{ padding: '20px' }}>
              <p className="dex-ui-section-desc" style={{ marginTop: 0 }}>
                {locale === 'de'
                  ? 'Andere Personen haben Teams angemeldet, denen noch Plätze fehlen. Du kannst eines vormerken — fülle dann oben deine persönlichen Daten und unten die event-spezifischen Angaben aus und klicke auf „Anmelden", um beizutreten.'
                  : 'Other people have registered teams with open slots. Pre-select one — then fill in your personal details above and the event-specific information below, and click “Register” to join.'}
                {event.teamJoinRequiresApproval && (
                  <> {locale === 'de'
                    ? <><br /><strong>Hinweis:</strong> der Team-Kapitän muss deinen Beitritt erst bestätigen.</>
                    : <><br /><strong>Note:</strong> the team lead has to approve your join.</>}
                  </>
                )}
              </p>
              {/* v31.9: Aus der Zeile mit angedocktem Knopf wird die Zeile
                  SELBST der Knopf (Leitfaden 2a': eine Zeile, die eine
                  Aktion hat, ist klickbar; das Beispiel des Nutzers sind die
                  Sub-Event-Zeilen in Schritt 1). Der Zustand steht als Text
                  in der Zeile, nicht nur in der Farbe oder im Hover — auf
                  dem Handy gibt es kein Ueberfahren (6b). `btn-primary` ist
                  hier weggefallen: DER Primaerknopf der Anmeldeseite ist
                  „Anmelden" unten, und den gab es damit doppelt
                  (Grundsatz 1.5). Der Aufruf von `togglePendingJoinTeam`
                  ist unveraendert. */}
              <div className="dex-ui-stack" style={{ gap: 6 }}>
                {openTeams.map(t => {
                  const free = t.teamSize - t.activeCount;
                  const isPicked = !!pendingJoinTeam && pendingJoinTeam.teamId === t.teamId;
                  return (
                    <button
                      key={t.teamId}
                      type="button"
                      aria-pressed={isPicked}
                      className={cx('dex-ui-rowbtn', 'dex-ui-row', 'dex-ui-row--framed', isPicked && 'is-active')}
                      onClick={() => togglePendingJoinTeam(t.teamId, t.teamName)}
                      style={{ padding: '10px 12px' }}
                    >
                      <Icon iconName="Group" style={{ fontSize: 16, color: 'var(--dex-green-dark, #4a7c1f)', flexShrink: 0 }} />
                      <span className="dex-ui-row-main" style={{ display: 'block' }}>
                        <span className="dex-ui-row-title dex-ui-row-title--wrap" style={{ display: 'block' }}>
                          {locale === 'de'
                            ? `Team „${t.teamName || 'ohne Namen'}"`
                            : `Team “${t.teamName || 'unnamed'}”`}
                        </span>
                        <span className="dex-ui-row-sub" style={{ display: 'block' }}>
                          {locale === 'de'
                            ? `${t.activeCount}/${t.teamSize} belegt — ${free} Slot${free === 1 ? '' : 's'} frei`
                            : `${t.activeCount}/${t.teamSize} taken — ${free} slot${free === 1 ? '' : 's'} free`}
                        </span>
                        {/* v31.9: Der Ruecknahme-Weg stand vorher nur in der
                            Knopf-Beschriftung („Vorgemerkt ✓ — entfernen").
                            Als Satz in der Zeile bleibt er auch dann lesbar,
                            wenn die Zeile selbst der Knopf ist. */}
                        {isPicked && (
                          <span className="dex-ui-row-sub" style={{ display: 'block' }}>
                            {locale === 'de'
                              ? 'Noch einmal tippen, um die Vormerkung zu entfernen.'
                              : 'Tap again to remove the pre-selection.'}
                          </span>
                        )}
                      </span>
                      <span className="dex-ui-row-actions">
                        <span className={cx('dex-ui-pill', isPicked ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')}>
                          {isPicked
                            ? (locale === 'de' ? 'Vorgemerkt ✓' : 'Pre-selected ✓')
                            : (locale === 'de' ? 'Vormerken' : 'Pre-select')}
                        </span>
                      </span>
                    </button>
                  );
                })}
              </div>
              <p className="dex-ui-help" style={{ marginTop: 10, marginBottom: 0 }}>
                {locale === 'de'
                  ? 'Mitgliedernamen werden aus Privatsphäre-Gründen nicht angezeigt.'
                  : 'Member names are hidden for privacy reasons.'}
              </p>
              {pendingJoinTeam && (
                <div className="dex-ui-callout dex-ui-callout--success" style={{ marginTop: 12 }}>
                  <span className="dex-ui-callout-icon"><Users size={18} /></span>
                  <div className="dex-ui-callout-body">
                    {locale === 'de'
                      ? <>Team <strong>„{pendingJoinTeam.teamName || 'ohne Namen'}“</strong> ist vorgemerkt. Fülle deine Angaben aus und klicke unten auf <strong>„Anmelden“</strong>, um den Beitritt abzuschließen.</>
                      : <>Team <strong>“{pendingJoinTeam.teamName || 'unnamed'}”</strong> is pre-selected. Fill in your details and click <strong>“Register”</strong> below to complete your join.</>}
                  </div>
                </div>
              )}
            </div>
            </CollapsibleSection>
          </div>
  );
};
