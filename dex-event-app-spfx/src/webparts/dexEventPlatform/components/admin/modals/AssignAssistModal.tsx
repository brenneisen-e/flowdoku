/* AssignAssistModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14432-14487 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 * v31.2: Reihenfolge „Wessen Anmeldung? → Wer übernimmt? → Was folgt daraus?";
 * Kopf und Fuß über die Modal-Props, Erklärtext auf eine Zeile Folge plus
 * einen Hinweiskasten verteilt. Handler und Suche unverändert.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Info, Users } from '../../Icons';
import { UserFieldPicker } from '../../UserFieldPicker';
import { ConsolidatedRow } from '../../admin/adminTypes';

export interface AssignAssistModalProps {
  assignAssistBusy: boolean;
  assignAssistRow: ConsolidatedRow;
  assignAssistValue: string;
  isDe: boolean;
  searchUser: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string; company?: string; }>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  setAssignAssistRow: React.Dispatch<React.SetStateAction<ConsolidatedRow>>;
  setAssignAssistValue: React.Dispatch<React.SetStateAction<string>>;
  submitAssignAssistant: () => Promise<void>;
}

export const AssignAssistModal: React.FC<AssignAssistModalProps> = (p) => {
  const { assignAssistBusy, assignAssistRow, assignAssistValue, isDe, searchUser, searchUsers, setAssignAssistRow, setAssignAssistValue, submitAssignAssistant } = p;
  const personName = `${assignAssistRow.vorname || ''} ${assignAssistRow.nachname || ''}`.trim() || assignAssistRow.email;
  const initials = (personName || '?').split(/[\s.@]+/).filter(Boolean).slice(0, 2).map(s => s[0].toUpperCase()).join('');
  const cancel = (): void => { setAssignAssistRow(null); setAssignAssistValue(''); };
  return (
    <Modal
      open={true}
      onClose={() => { if (!assignAssistBusy) cancel(); }}
      maxWidth={540}
      dismissable={!assignAssistBusy}
      ariaLabel={isDe ? 'Assistenz zuordnen' : 'Assign assistant'}
      title={isDe ? 'Assistenz zuordnen' : 'Assign assistant'}
      subtitle={isDe ? 'Wer soll diese Anmeldung künftig verwalten?' : 'Who should manage this registration from now on?'}
      icon={<Users size={20} />}
      footer={<>
        <button type="button" className="btn btn-secondary" disabled={assignAssistBusy} onClick={cancel}>
          {isDe ? 'Abbrechen' : 'Cancel'}
        </button>
        <button type="button" className="btn btn-primary" disabled={assignAssistBusy || !assignAssistValue} onClick={() => { void submitAssignAssistant(); }}>
          {assignAssistBusy ? (isDe ? 'Wird zugeordnet…' : 'Assigning…') : (isDe ? 'Zuordnen' : 'Assign')}
        </button>
      </>}
    >
      <div className="dex-ui-card dex-ui-card--soft" style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '12px 14px' }}>
        <span className="dex-ui-avatar dex-ui-avatar--lg" aria-hidden="true">{initials}</span>
        <div style={{ minWidth: 0 }}>
          <div className="dex-ui-muted" style={{ fontSize: '0.72rem', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.06em' }}>{isDe ? 'Anmeldung von' : 'Registration of'}</div>
          <div className="dex-ui-row-title">{personName}</div>
          <div className="dex-ui-row-sub" style={{ wordBreak: 'break-all' }}>{assignAssistRow.email}</div>
        </div>
      </div>
      <div className="dex-ui-field">
        <label className="dex-ui-label">{isDe ? 'Wer übernimmt die Verwaltung?' : 'Who takes over?'}</label>
        <UserFieldPicker
          value={assignAssistValue}
          onChange={setAssignAssistValue}
          searchUsers={searchUsers}
          searchUserByEmail={searchUser}
          placeholder={isDe ? 'Name oder E-Mail der Assistenz…' : 'Assistant name or email…'}
          errorStyle={{}}
          forcedIsDe={isDe}
        />
        <div className="dex-ui-help">
          {isDe
            ? <>Die Assistenz übernimmt die <strong>Klammer-Anmeldung und alle aktiven Sub-Event-Anmeldungen</strong>. Es gehen <strong>keine Mails</strong> raus.</>
            : <>The assistant takes over the <strong>umbrella registration and all active sub-event registrations</strong>. <strong>No emails</strong> are sent.</>}
        </div>
      </div>
      <div className="dex-ui-callout dex-ui-callout--neutral">
        <span className="dex-ui-callout-icon"><Info size={16} /></span>
        <span>
          {isDe
            ? <>Danach erscheint die Person in der <strong>&bdquo;Assistenz&ldquo;-Kachel</strong> der Assistenz — dort kann sie Angaben anpassen sowie ab- und anmelden.</>
            : <>The person then appears in the assistant&apos;s <strong>&bdquo;Assistant&ldquo; tile</strong> — there they can adjust details, cancel and re-register.</>}
        </span>
      </div>
    </Modal>
  );
};
