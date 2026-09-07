/* TeamMailModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 11583-11664 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Mail } from '../../Icons';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

export interface TeamMailModalProps {
  getActiveTeams: () => Array<{    tid: string;    teamName: string;    members: SPRegistration[];}>;
  isDe: boolean;
  selectedEvent: DeloitteEvent;
  sendTeamMails: () => Promise<void>;
  setTeamMailBody: React.Dispatch<React.SetStateAction<string>>;
  setTeamMailInfoByTid: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  setTeamMailOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamMailSubject: React.Dispatch<React.SetStateAction<string>>;
  teamMailBody: string;
  teamMailInfoByTid: Record<string, string>;
  teamMailSending: boolean;
  teamMailSubject: string;
}

export const TeamMailModal: React.FC<TeamMailModalProps> = (p) => {
  const { getActiveTeams, isDe, selectedEvent, sendTeamMails, setTeamMailBody, setTeamMailInfoByTid, setTeamMailOpen, setTeamMailSubject, teamMailBody, teamMailInfoByTid, teamMailSending, teamMailSubject } = p;
  // v31.2: Einmal je Render lesen — Empfänger-Pills, Gruppenliste und der
  // Sende-Knopf zählen so dieselben Teams. Die Zahl im Knopf nennt die Folge
  // („An 47 Personen senden"), bevor jemand klickt.
  const teams = getActiveTeams();
  const recipients = teams.reduce((n, t) => n + t.members.length, 0);
  const termS = selectedEvent.teamTermSingular || 'Team';
  const placeholders = ['{{Vorname}}', '{{Name}}', '{{TeamName}}', '{{EventTitle}}', '{{TeamInfo}}'];
  return (
    <Modal
      open={true}
      onClose={() => { if (!teamMailSending) setTeamMailOpen(false); }}
      dismissable={!teamMailSending}
      maxWidth={720}
      padding={24}
      ariaLabel={isDe ? 'Mail an Teams' : 'Mail to teams'}
      title={isDe ? `Mail an ${selectedEvent.teamTermPlural || 'Teams'}` : `Mail to ${selectedEvent.teamTermPlural || 'teams'}`}
      subtitle={isDe ? 'Jedes aktive Mitglied bekommt eine eigene Mail — mit der Info seiner Gruppe.' : 'Each active member gets an individual mail — with the info of their group.'}
      icon={<Mail size={20} />}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={() => setTeamMailOpen(false)} disabled={teamMailSending}>
          {isDe ? 'Abbrechen' : 'Cancel'}
        </button>
        <button type="button" className="btn btn-primary" onClick={() => { sendTeamMails().catch(() => { /* */ }); }} disabled={teamMailSending || !teamMailSubject.trim() || !teamMailBody.trim()}>
          {teamMailSending ? (isDe ? 'Wird gesendet…' : 'Sending…') : (isDe ? `An ${recipients} Personen senden` : `Send to ${recipients} people`)}
        </button>
      </>}
    >
      {/* v31.2: Reihenfolge = Frage des Organizers: Wer bekommt sie? → Was steht drin? → Was bekommt jede Gruppe extra? */}
      <div>
        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{isDe ? 'Wer bekommt die Mail?' : 'Who gets the mail?'}</div>
          {teams.length === 0 ? (
            <div className="dex-ui-empty">{isDe ? 'Keine aktive Gruppe — es gibt niemanden, der eine Mail bekäme.' : 'No active group — there is nobody to mail.'}</div>
          ) : (
            <div className="dex-ui-inline">
              {teams.map(t => (
                <span key={t.tid} className="dex-ui-pill dex-ui-pill--gray" title={t.members.map(m => m.ParticipantName || m.ParticipantEmail).join(', ')}>
                  {t.teamName || termS} · {t.members.length}
                </span>
              ))}
              <span className="dex-ui-pill dex-ui-pill--green">{recipients} {isDe ? 'Mails' : 'mails'}</span>
            </div>
          )}
        </div>

        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{isDe ? 'Was steht in der Mail?' : 'What does the mail say?'}</div>
          <div className="dex-ui-field">
            <label className="dex-ui-label" htmlFor="dex-teammail-subject">{isDe ? 'Betreff' : 'Subject'}</label>
            <input id="dex-teammail-subject" type="text" className="dex-ui-input" value={teamMailSubject} onChange={e => setTeamMailSubject(e.target.value)} disabled={teamMailSending} />
          </div>
          <div className="dex-ui-field">
            <label className="dex-ui-label" htmlFor="dex-teammail-body">{isDe ? 'Mail-Text' : 'Mail body'}</label>
            <textarea id="dex-teammail-body" className="dex-ui-textarea" value={teamMailBody} onChange={e => setTeamMailBody(e.target.value)} disabled={teamMailSending} rows={8} style={{ fontFamily: 'monospace', fontSize: '0.85rem' }} />
            <div className="dex-ui-help dex-ui-inline" style={{ gap: 6 }}>
              {isDe ? 'HTML erlaubt. Platzhalter:' : 'HTML allowed. Placeholders:'}
              {placeholders.map(ph => <code key={ph} className="dex-ui-pill dex-ui-pill--gray" style={{ fontFamily: 'monospace', fontWeight: 500 }}>{ph}</code>)}
            </div>
          </div>
        </div>

        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{isDe ? 'Was bekommt jede Gruppe extra?' : 'What does each group get extra?'}</div>
          <p className="dex-ui-section-desc">
            {isDe
              ? 'Der Text ersetzt {{TeamInfo}} in der Mail — z.B. der Teams-Einwahllink; Links werden klickbar. Leer heißt: In der Mail steht „(keine zusätzlichen Infos)".'
              : 'The text replaces {{TeamInfo}} in the mail — e.g. the Teams join link; links become clickable. Empty means the mail reads “(no additional info)”.'}
          </p>
          <div className="dex-ui-stack" style={{ maxHeight: 280, overflowY: 'auto' }}>
            {teams.map(t => {
              const tName = t.teamName || termS;
              return (
                <div key={t.tid} className="dex-ui-card dex-ui-card--soft dex-ui-card--hover" style={{ padding: '10px 12px' }}>
                  <label className="dex-ui-label" htmlFor={`dex-teammail-info-${t.tid}`}>
                    {tName} <span className="dex-ui-pill dex-ui-pill--gray">{t.members.length} {isDe ? 'Mitglieder' : 'members'}</span>
                  </label>
                  <textarea
                    id={`dex-teammail-info-${t.tid}`} className="dex-ui-textarea" rows={3} style={{ minHeight: 0, fontSize: '0.85rem' }}
                    value={teamMailInfoByTid[t.tid] || ''}
                    onChange={e => setTeamMailInfoByTid(prev => ({ ...prev, [t.tid]: e.target.value }))}
                    disabled={teamMailSending}
                    placeholder={isDe ? 'z.B. https://teams.microsoft.com/l/meetup-join/…' : 'e.g. https://teams.microsoft.com/l/meetup-join/…'}
                  />
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </Modal>
  );
};

