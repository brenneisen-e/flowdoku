/* OverbookDecisionModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16437-16560 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.2: Nach docs/ui-leitfaden.md umgebaut — Modal-Kopf/-Fuß, die Situation in
 * einem Satz, Zahlen als KPI (sie folgen den Schaltern live), Ja/Nein als
 * Schalter-Zeilen mit Folge, die zwei „Platz behalten"-Wege als Auswahl-Kacheln
 * mit Konsequenz-Zeile. Die Sprache kommt über useLocaleSafe, weil die Props
 * unverändert bleiben müssen (der Aufrufer wird parallel bearbeitet).
 */
import * as React from 'react';
import Modal from '../../Modal';
import { cx } from '../../dexUi';
import { AlertCircle, Check, ChevronDown } from '../../Icons';
import { useLocaleSafe } from '../../../context/LanguageContext';
import { getCachedLogoBase64, getCachedOrbBase64 } from '../../../services/EmailTemplates';
import { SPRegistration } from '../../../services/EventService';

export interface OverbookDecisionModalProps {
  obBusy: boolean;
  obKeepVariant: "active" | "firstWaitlist";
  obMailBody: string;
  obMailLang: "DE" | "EN";
  obMailSubject: string;
  obRemoveCalendar: boolean;
  obWithMail: boolean;
  overbookModal: { mode: "confirm" | "keep"; targets: SPRegistration[]; };
  runOverbookResolution: () => Promise<void>;
  setObKeepVariant: React.Dispatch<React.SetStateAction<"active" | "firstWaitlist">>;
  setObMailBody: React.Dispatch<React.SetStateAction<string>>;
  setObMailLang: React.Dispatch<React.SetStateAction<"DE" | "EN">>;
  setObMailSubject: React.Dispatch<React.SetStateAction<string>>;
  setObRemoveCalendar: React.Dispatch<React.SetStateAction<boolean>>;
  setObWithMail: React.Dispatch<React.SetStateAction<boolean>>;
  setOverbookModal: React.Dispatch<React.SetStateAction<{ mode: "confirm" | "keep"; targets: SPRegistration[]; }>>;
}

export const OverbookDecisionModal: React.FC<OverbookDecisionModalProps> = (p) => {
  const { obBusy, obKeepVariant, obMailBody, obMailLang, obMailSubject, obRemoveCalendar, obWithMail, overbookModal, runOverbookResolution, setObKeepVariant, setObMailBody, setObMailLang, setObMailSubject, setObRemoveCalendar, setObWithMail, setOverbookModal } = p;
  const isDe = useLocaleSafe() === 'de';
  // v31.2: Die HTML-Vorschau ist standardmäßig zu (Reiter bzw. Aufklapper) —
  // offen füllte sie mit 280 px den halben Dialog, bevor man die Frage gelesen hatte.
  const [showPreview, setShowPreview] = React.useState(false);
  const t = (de: string, en: string): string => (isDe ? de : en);
  const targets = overbookModal.targets;
  const n = targets.length;
  const isConfirm = overbookModal.mode === 'confirm';
  const nameOf = (r: SPRegistration): string => (r.Vorname && r.Nachname) ? `${r.Vorname} ${r.Nachname}` : r.ParticipantName;
  const who = n === 1 ? nameOf(targets[0]) : t(`${n} Personen`, `${n} people`);
  // v31.2: Mail und Kalender-Absage gibt es nur mit Adresse — die KPI zählt genau das.
  const withEmail = targets.filter(r => !!r.ParticipantEmail).length;
  const previewHtml = obMailBody
    .replace(/\{\{LOGO_URL\}\}/g, getCachedLogoBase64() || '')
    .replace(/\{\{ORB_URL\}\}/g, getCachedOrbBase64() || '');
  const preview = (
    <div
      style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 12, maxHeight: 280, overflow: 'auto', background: '#fff' }}
      dangerouslySetInnerHTML={{ __html: previewHtml }}
    />
  );
  return (
        <Modal
          open={true}
          onClose={() => setOverbookModal(null)}
          dismissable={!obBusy}
          maxWidth={600}
          ariaLabel={t('Überbuchung', 'Overbooking')}
          icon={<AlertCircle size={20} />}
          title={isConfirm ? t('Auf die Warteliste setzen', 'Move to the waitlist') : t('Platz behalten', 'Keep the seat')}
          subtitle={isConfirm
            ? (n === 1
              ? t(`${who} wechselt auf die Warteliste der Gruppe. Das Audit-Log vermerkt, dass die Person fälschlich angemeldet war.`,
                  `${who} moves to the group's waitlist. The audit log records that the person was registered by mistake.`)
              : t(`${who} wechseln auf die Warteliste. Das Audit-Log jeder Person vermerkt den Vorgang.`,
                  `${who} move to the waitlist. Each person's audit log records the step.`))
            : t(`${who} verliert den Platz nicht. Wähle, wie er gesichert wird:`, `${who} keeps the seat. Choose how:`)}
          footer={<>
            <button type="button" className="btn btn-secondary" onClick={() => setOverbookModal(null)} disabled={obBusy}>
              {t('Abbrechen', 'Cancel')}
            </button>
            <button type="button" className="btn btn-primary" onClick={() => { runOverbookResolution().catch(() => { /* */ }); }} disabled={obBusy}>
              {obBusy
                ? t('Wird ausgeführt…', 'Running…')
                : (isConfirm ? t('Auf Warteliste setzen & IDs neu vergeben', 'Move to waitlist & reassign IDs') : t('Übernehmen & IDs neu vergeben', 'Apply & reassign IDs'))}
            </button>
          </>}
        >
            {isConfirm ? (
              <>
                <div className="dex-ui-grid-3">
                  <div className="dex-ui-kpi">
                    <div className="dex-ui-kpi-value">{n}</div>
                    <div className="dex-ui-kpi-label">{t(n === 1 ? 'Person' : 'Personen', n === 1 ? 'person' : 'people')}</div>
                  </div>
                  <div className={cx('dex-ui-kpi', obWithMail && withEmail > 0 && 'dex-ui-kpi--green')}>
                    <div className="dex-ui-kpi-value">{obWithMail ? withEmail : 0}</div>
                    <div className="dex-ui-kpi-label">{t('Mails in die Queue', 'Emails queued')}</div>
                  </div>
                  <div className={cx('dex-ui-kpi', obRemoveCalendar && withEmail > 0 && 'dex-ui-kpi--orange')}>
                    <div className="dex-ui-kpi-value">{obRemoveCalendar ? withEmail : 0}</div>
                    <div className="dex-ui-kpi-label">{t('Kalender-Absagen', 'Calendar removals')}</div>
                  </div>
                </div>
                <div className="dex-ui-section">
                  <div className="dex-ui-section-title">{t('Was die Person erfährt', 'What the person hears')}</div>
                  <div className="dex-ui-stack">
                    <label className={cx('dex-ui-toggle-row', obWithMail && 'is-active', obBusy && 'is-disabled')}>
                      <input type="checkbox" checked={obWithMail} onChange={e => setObWithMail(e.target.checked)} disabled={obBusy} />
                      <span className="dex-ui-toggle-row-body">
                        <span className="dex-ui-toggle-row-title">{t('Entschuldigungs-Mail schicken', 'Send an apology email')}</span>
                        <span className="dex-ui-toggle-row-desc">
                          {t('Deloitte-Layout mit der neuen Wartelisten-Position. Geht in die Mail-Queue, nicht direkt raus.',
                             'Deloitte layout with the new waitlist position. Goes into the mail queue, not sent directly.')}
                        </span>
                      </span>
                    </label>
                    {obWithMail && (
                      // v31.2: Abhängiges direkt unter seinem Schalter — Sprache, Text, Vorschau.
                      <div className="dex-ui-card dex-ui-card--soft dex-ui-stack">
                        <div className="dex-ui-field">
                          <div className="dex-ui-label">{t('In welcher Sprache?', 'Which language?')}</div>
                          <div className="dex-ui-inline">
                            {(['DE', 'EN'] as const).map(lng => (
                              <button key={lng} type="button" className={cx('dex-ui-chip', obMailLang === lng && 'is-active')} disabled={obBusy} onClick={() => setObMailLang(lng)}>
                                {lng === 'DE' ? 'Deutsch' : 'English'}
                              </button>
                            ))}
                          </div>
                        </div>
                        {n === 1 ? (
                          <>
                            <div className="dex-ui-tabs" role="tablist">
                              <button type="button" role="tab" aria-selected={!showPreview} className={cx('dex-ui-tab', !showPreview && 'is-active')} onClick={() => setShowPreview(false)}>{t('Text bearbeiten', 'Edit text')}</button>
                              <button type="button" role="tab" aria-selected={showPreview} className={cx('dex-ui-tab', showPreview && 'is-active')} onClick={() => setShowPreview(true)}>{t('Vorschau', 'Preview')}</button>
                            </div>
                            {!showPreview ? (
                              <>
                                <div className="dex-ui-field">
                                  <label className="dex-ui-label" htmlFor="ob-mail-subject">{t('Betreff', 'Subject')}</label>
                                  <input id="ob-mail-subject" className="dex-ui-input" value={obMailSubject} onChange={e => setObMailSubject(e.target.value)} disabled={obBusy} />
                                </div>
                                <div className="dex-ui-field">
                                  <label className="dex-ui-label" htmlFor="ob-mail-body">{t('Text (HTML)', 'Text (HTML)')}</label>
                                  <textarea id="ob-mail-body" className="dex-ui-textarea" value={obMailBody} onChange={e => setObMailBody(e.target.value)} disabled={obBusy} rows={5} style={{ fontFamily: 'monospace', fontSize: '0.72rem' }} />
                                  <div className="dex-ui-help">{t('Vorschlagstext — du kannst ihn ändern. Der Reiter „Vorschau“ zeigt die fertige Deloitte-Mail.', 'Suggested text — you can change it. The “Preview” tab shows the finished Deloitte email.')}</div>
                                </div>
                              </>
                            ) : preview}
                          </>
                        ) : (
                          // v13.0: Preview teilt sich den obMailBody-State mit der
                          // Modal-Open-useEffect — beide rendern den Body der
                          // ersten Person. Vorher wurde buildOverbookApologyEmail
                          // synchron im Render aufgerufen; seit der Template-DB-
                          // Lookup async ist, geht das nicht mehr direkt im JSX.
                          <>
                            <div className="dex-ui-help" style={{ marginTop: 0 }}>
                              {t('Jede Person bekommt den Standardtext personalisiert — mit Namen und eigener Wartelisten-Position.',
                                 'Each person gets the standard text personalised — with their name and their own waitlist position.')}
                            </div>
                            <button type="button" className={cx('dex-ui-disclosure', showPreview && 'is-open')} aria-expanded={showPreview} onClick={() => setShowPreview(v => !v)}>
                              <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                              {t('Vorschau am Beispiel der ersten Person', 'Preview using the first person as example')}
                            </button>
                            {showPreview && <div className="dex-ui-disclosure-body">{preview}</div>}
                          </>
                        )}
                      </div>
                    )}
                    <label className={cx('dex-ui-toggle-row', obRemoveCalendar && 'is-active', obBusy && 'is-disabled')}>
                      <input type="checkbox" checked={obRemoveCalendar} onChange={e => setObRemoveCalendar(e.target.checked)} disabled={obBusy} />
                      <span className="dex-ui-toggle-row-body">
                        <span className="dex-ui-toggle-row-title">{t('Aus dem Kalendereintrag ausladen', 'Remove from the calendar entry')}</span>
                        <span className="dex-ui-toggle-row-desc">{t('Falls die Person eine Outlook-Einladung hat, wird sie zurückgezogen.', 'If the person has an Outlook invitation, it is withdrawn.')}</span>
                      </span>
                    </label>
                  </div>
                </div>
              </>
            ) : (
              <div className="dex-ui-grid-2" role="radiogroup" aria-label={t('Wie bleibt der Platz erhalten?', 'How is the seat kept?')}>
                {([
                  { v: 'firstWaitlist' as const, title: t('Erste(r) auf der Warteliste', 'First on the waitlist'),
                    desc: t('Rückt beim nächsten frei werdenden Platz der Gruppe garantiert als Erste(r) nach. Risikoarm — die Kapazität bleibt eingehalten.',
                            'Guaranteed to move up first when the next seat in the group frees up. Low risk — capacity stays intact.') },
                  { v: 'active' as const, title: t('Bleibt angemeldet (als Letzte(r))', 'Stays registered (as the last one)'),
                    desc: t('Die Gruppe steht +1 über Kapazität; der nächste frei werdende Platz wird einmal nicht nachgerückt, bis die Überzahl absorbiert ist.',
                            'The group is +1 over capacity; the next seat that frees up is not filled once, until the surplus is absorbed.') },
                ]).map(o => (
                  <button key={o.v} type="button" role="radio" aria-checked={obKeepVariant === o.v} disabled={obBusy}
                    className={cx('dex-ui-choice', obKeepVariant === o.v && 'is-active')} onClick={() => setObKeepVariant(o.v)}>
                    <div className="dex-ui-choice-body">
                      <div className="dex-ui-choice-title">{o.title}</div>
                      <div className="dex-ui-choice-desc">{o.desc}</div>
                    </div>
                    <span className="dex-ui-choice-check"><Check size={12} /></span>
                  </button>
                ))}
              </div>
            )}
            <div className="dex-ui-callout dex-ui-callout--neutral">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span>{t('Danach vergibt die App alle TeilnehmerIDs neu (Aktive zuerst, dann Warteliste) und lädt die Liste nach.',
                       'Afterwards the app reassigns all participant IDs (active first, then waitlist) and reloads the list.')}</span>
            </div>
        </Modal>
  );
};
