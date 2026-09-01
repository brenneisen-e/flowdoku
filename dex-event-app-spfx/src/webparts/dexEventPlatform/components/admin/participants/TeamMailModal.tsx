/* TeamMailModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 11583-11664 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
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
  return (
          <Modal
            open={true}
            onClose={() => { if (!teamMailSending) setTeamMailOpen(false); }}
            dismissable={!teamMailSending}
            maxWidth={720}
            padding={24}
            ariaLabel={isDe ? 'Mail an Teams' : 'Mail to teams'}
          >
            <h3 style={{ marginTop: 0, marginBottom: 6, color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {isDe ? `Mail an ${selectedEvent.teamTermPlural || 'Teams'}` : `Mail to ${selectedEvent.teamTermPlural || 'teams'}`}
            </h3>
            <p style={{ marginTop: 0, fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
              {isDe
                ? 'Jedes aktive Mitglied bekommt eine eigene Mail. Pro Gruppe trägst du unten eine eigene Info ein (z.B. einen Teams-Einwahllink) — sie ersetzt den Platzhalter {{TeamInfo}}. Verfügbare Platzhalter: {{Vorname}}, {{Name}}, {{TeamName}}, {{EventTitle}}, {{TeamInfo}}.'
                : 'Each active member receives an individual mail. Per group you enter its own info below (e.g. a Teams join link) — it replaces the {{TeamInfo}} placeholder. Available placeholders: {{Vorname}}, {{Name}}, {{TeamName}}, {{EventTitle}}, {{TeamInfo}}.'}
            </p>

            <label style={{ display: 'block', fontWeight: 600, fontSize: '0.85rem', marginTop: 10 }}>{isDe ? 'Betreff' : 'Subject'}</label>
            <input
              type="text"
              value={teamMailSubject}
              onChange={e => setTeamMailSubject(e.target.value)}
              disabled={teamMailSending}
              style={{ width: '100%', padding: '8px 10px', borderRadius: 8, border: '1px solid var(--dex-gray-300)', fontSize: '0.9rem', boxSizing: 'border-box' }}
            />

            <label style={{ display: 'block', fontWeight: 600, fontSize: '0.85rem', marginTop: 12 }}>{isDe ? 'Mail-Text (HTML erlaubt)' : 'Mail body (HTML allowed)'}</label>
            <textarea
              value={teamMailBody}
              onChange={e => setTeamMailBody(e.target.value)}
              disabled={teamMailSending}
              rows={8}
              style={{ width: '100%', padding: '8px 10px', borderRadius: 8, border: '1px solid var(--dex-gray-300)', fontSize: '0.85rem', fontFamily: 'monospace', boxSizing: 'border-box', resize: 'vertical' }}
            />

            <div style={{ marginTop: 14, fontWeight: 600, fontSize: '0.85rem' }}>
              {isDe ? 'Info pro Gruppe' : 'Info per group'}
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 8, maxHeight: 280, overflowY: 'auto' }}>
              {getActiveTeams().map(t => {
                const tName = t.teamName || (selectedEvent.teamTermSingular || 'Team');
                return (
                  <div key={t.tid} style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: '10px 12px', background: 'var(--dex-gray-50, #f7f7f7)' }}>
                    <div style={{ fontWeight: 600, fontSize: '0.86rem', marginBottom: 6 }}>
                      {tName} <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}>· {t.members.length} {isDe ? 'Mitglieder' : 'members'}</span>
                    </div>
                    <textarea
                      value={teamMailInfoByTid[t.tid] || ''}
                      onChange={e => setTeamMailInfoByTid(prev => ({ ...prev, [t.tid]: e.target.value }))}
                      disabled={teamMailSending}
                      rows={3}
                      placeholder={isDe ? 'z.B. https://teams.microsoft.com/l/meetup-join/…' : 'e.g. https://teams.microsoft.com/l/meetup-join/…'}
                      style={{ width: '100%', padding: '7px 9px', borderRadius: 6, border: '1px solid var(--dex-gray-300)', fontSize: '0.83rem', boxSizing: 'border-box', resize: 'vertical' }}
                    />
                  </div>
                );
              })}
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 18 }}>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => setTeamMailOpen(false)}
                disabled={teamMailSending}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => { sendTeamMails().catch(() => { /* */ }); }}
                disabled={teamMailSending || !teamMailSubject.trim() || !teamMailBody.trim()}
              >
                {teamMailSending
                  ? (isDe ? 'Wird gesendet…' : 'Sending…')
                  : (isDe ? 'Mails senden' : 'Send mails')}
              </button>
            </div>
          </Modal>
  );
};

