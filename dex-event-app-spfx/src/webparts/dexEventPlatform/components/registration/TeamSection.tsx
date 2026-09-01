/* Team-Bloecke der Anmeldeseite — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Die Karte „Ich melde mich + mein Team an" (v11.82) und darunter die Liste der
 * offenen Teams (v18.73). Inhalt zeichengleich uebernommen; die
 * Anzeige-Bedingungen sind beim Aufrufer geblieben. */
import * as React from 'react';
import { CollapsibleSection } from './regHelpers';
import { Icon } from '@fluentui/react/lib/Icon';
import { Locale } from '../../context/LanguageContext';
import { DeloitteEvent, EventSpecificField } from '../../types';
import { UserFieldPicker } from '../UserFieldPicker';

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
              {/* Pflicht-Hinweis-Box ganz oben — auffällig orange. */}
              <div style={{
                marginBottom: 20,
                padding: '14px 16px',
                background: 'rgba(237,139,0,0.10)',
                border: '2px solid var(--dex-orange, #ed8b00)',
                borderRadius: 'var(--dex-radius-md)',
                color: '#7a4a00',
                fontSize: '0.88rem',
                lineHeight: 1.55,
              }}>
                <div style={{ fontWeight: 700, marginBottom: 8, fontSize: '0.95rem' }}>
                  {locale === 'de'
                    ? 'Vorab die Zustimmung jedes Teammitglieds einholen'
                    : 'Get every team member\'s consent up front'}
                </div>
                <div style={{ marginBottom: 8 }}>
                  {locale === 'de'
                    ? 'Mit dem Absenden meldest du nicht nur dich selbst an, sondern auch alle weiter unten eingetragenen Personen. Jedes Teammitglied erhält automatisch:'
                    : 'By submitting you register yourself AND every person you add below. Each team member automatically receives:'}
                </div>
                <ul style={{ margin: '0 0 8px 18px', padding: 0 }}>
                  <li>{locale === 'de' ? 'eine Anmeldebestätigung per Mail' : 'a confirmation email'}</li>
                  <li>{locale === 'de' ? 'einen Outlook-Termin im Kalender' : 'an Outlook calendar invite'}</li>
                  <li>{locale === 'de' ? 'den Event in „Meine Events"' : 'the event in „My Events"'}</li>
                </ul>
                <div style={{ marginTop: 4 }}>
                  {locale === 'de'
                    ? <>Bitte stelle sicher, dass alle Teilnehmer ihrer Anmeldung <strong>VORHER zugestimmt</strong> haben — eine Anmeldung ohne Einverständnis ist nicht erlaubt.</>
                    : <>Please make sure every participant has <strong>consented up front</strong> — registering people without their consent is not allowed.</>}
                </div>
              </div>

              {event?.askTeamName && (
                <div className="form-group">
                  <label className="form-label">
                    <span className="required">*</span> {locale === 'de' ? 'Team-Name' : 'Team name'}
                  </label>
                  <input
                    className="form-input"
                    value={teamName}
                    onChange={e => setTeamName(e.target.value.slice(0, 60))}
                    placeholder={locale === 'de' ? 'z.B. „Die Schnellen"' : 'e.g. „The Quick Ones"'}
                    style={showErrors && isTeamMode && event?.askTeamName && !teamName.trim() ? errorBorder : {}}
                  />
                  <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>
                    {locale === 'de' ? 'Max. 60 Zeichen — wird in der Teilnehmerliste mitgespeichert.' : 'Max 60 characters — stored on the participant list.'}
                  </div>
                </div>
              )}

              {/* Member-Slots */}
              <div style={{ marginTop: 8 }}>
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

              {/* Pflicht-Bestätigungs-Checkbox */}
              <div className="form-group" style={{ marginTop: 18, marginBottom: 0 }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={teamConsentConfirmed}
                    onChange={e => setTeamConsentConfirmed(e.target.checked)}
                    style={{ marginTop: 3 }}
                  />
                  <div style={{ flex: 1, fontSize: '0.88rem', color: 'var(--dex-gray-800)', lineHeight: 1.5 }}>
                    <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                    {locale === 'de'
                      ? 'Ich bestätige, dass alle eingetragenen Teammitglieder ihrer Anmeldung zugestimmt haben.'
                      : 'I confirm that every listed team member has consented to this registration.'}
                  </div>
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
              <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)', marginTop: 0, marginBottom: 12, lineHeight: 1.5 }}>
                {locale === 'de'
                  ? 'Andere Personen haben Teams angemeldet, denen noch Plätze fehlen. Du kannst eines vormerken — fülle dann oben deine persönlichen Daten und unten die event-spezifischen Angaben aus und klicke auf „Anmelden", um beizutreten.'
                  : 'Other people have registered teams with open slots. Pre-select one — then fill in your personal details above and the event-specific information below, and click „Register" to join.'}
                {event.teamJoinRequiresApproval && (
                  <> {locale === 'de'
                    ? <><br /><strong>Hinweis:</strong> der Team-Kapitän muss deinen Beitritt erst bestätigen.</>
                    : <><br /><strong>Note:</strong> the team lead has to approve your join.</>}
                  </>
                )}
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {openTeams.map(t => {
                  const free = t.teamSize - t.activeCount;
                  const isPicked = !!pendingJoinTeam && pendingJoinTeam.teamId === t.teamId;
                  return (
                    <div key={t.teamId} style={{
                      display: 'flex', alignItems: 'center', gap: 12,
                      padding: '10px 12px',
                      background: isPicked ? 'rgba(134,188,37,0.10)' : 'var(--dex-gray-50, #f7f7f7)',
                      borderRadius: 6,
                      border: isPicked ? '2px solid var(--dex-green, #86bc25)' : '1px solid var(--dex-gray-200)',
                    }}>
                      <Icon iconName="Group" style={{ fontSize: 16, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                      <div style={{ flex: 1, fontSize: '0.88rem' }}>
                        <div style={{ fontWeight: 600 }}>
                          {locale === 'de'
                            ? `Team „${t.teamName || 'ohne Namen'}"`
                            : `Team „${t.teamName || 'unnamed'}"`}
                        </div>
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                          {locale === 'de'
                            ? `${t.activeCount}/${t.teamSize} belegt — ${free} Slot${free === 1 ? '' : 's'} frei`
                            : `${t.activeCount}/${t.teamSize} taken — ${free} slot${free === 1 ? '' : 's'} free`}
                        </div>
                      </div>
                      <button
                        type="button"
                        className={isPicked ? 'btn btn-secondary' : 'btn btn-primary'}
                        onClick={() => togglePendingJoinTeam(t.teamId, t.teamName)}
                        style={{ fontSize: '0.82rem', padding: '6px 12px' }}
                      >
                        {isPicked
                          ? (locale === 'de' ? 'Vorgemerkt ✓ — entfernen' : 'Pre-selected ✓ — remove')
                          : (locale === 'de' ? 'Vormerken' : 'Pre-select')}
                      </button>
                    </div>
                  );
                })}
              </div>
              <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 10, marginBottom: 0, lineHeight: 1.4 }}>
                {locale === 'de'
                  ? 'Mitgliedernamen werden aus Privatsphäre-Gründen nicht angezeigt.'
                  : 'Member names are hidden for privacy reasons.'}
              </p>
              {pendingJoinTeam && (
                <div style={{
                  marginTop: 12,
                  padding: '10px 12px',
                  borderRadius: 6,
                  background: 'rgba(134,188,37,0.10)',
                  border: '1px solid var(--dex-green, #86bc25)',
                  color: 'var(--dex-green-dark, #3f5f10)',
                  fontSize: '0.85rem',
                  lineHeight: 1.5,
                }}>
                  {locale === 'de'
                    ? <>Team <strong>„{pendingJoinTeam.teamName || 'ohne Namen'}“</strong> ist vorgemerkt. Fülle deine Angaben aus und klicke unten auf <strong>„Anmelden“</strong>, um den Beitritt abzuschließen.</>
                    : <>Team <strong>“{pendingJoinTeam.teamName || 'unnamed'}”</strong> is pre-selected. Fill in your details and click <strong>“Register”</strong> below to complete your join.</>}
                </div>
              )}
            </div>
            </CollapsibleSection>
          </div>
  );
};
