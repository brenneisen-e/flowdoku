/* CommunicationStep — aus EventCreationPage.tsx ausgelagert (Zeilen 16971-17861 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; einzige
 * Aenderung ist die Anzeige-Bedingung: aus `currentStep === 5` wurde das Prop `visible`.
 * `visible` schaltet display:none statt unmount — Eingaben ueberleben den
 * Schrittwechsel genauso wie vorher. */
import * as React from 'react';
import { InfoTooltip } from '../../InfoTooltip';
import { Icon } from '@fluentui/react/lib/Icon';
import { StepBadge } from '../../wizard/StepBadge';
import WizardHint from '../../WizardHint';
import DatePicker from 'react-datepicker';
import { Check, Plus } from '../../Icons';
import { compressImage } from '../../../utils/imageCompress';
import { RichText } from '@pnp/spfx-controls-react/lib/controls/richText';
import { COMM_TOPICS } from '../logic/commTabs';
import { BundledComm } from '../../../utils/bundledComm';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { EmailOverrideEntry } from '../../wizard/emailOverrideEntry';
export interface CommunicationStepProps {
  visible: boolean;
  activeCommTabIdx: number;
  applyCommToAllSubEvents: () => Promise<void>;
  /** v30.71: ein Thema (s. COMM_TOPICS) vom offenen Reiter auf alle Termine. */
  applyCommTopicToAllSubEvents: (topic: string) => Promise<void>;
  /** v30.71: „Gemeinsam für alle Termine" — Kennzeichen am Haupt-Event. */
  commShared: boolean;
  setCommShared: React.Dispatch<React.SetStateAction<boolean>>;
  flushActiveCommTabToState: () => void;
  resolveTopLevelCommState: () => { emailLanguage: string; emailLogoBase64: string; outlookLogoBase64: string; outlookBody: string; outlookHeading: string; outlookSubheading: string; outlookSubject: string; disableEmails: boolean; disableRegistrationEmail: boolean; disableCancellationEmail: boolean; autoDeregisterOnDecline: boolean; inactiveHandling?: 'notify' | 'autoderegister'; disableOutlook: boolean; emailTemplateOverrides: Record<string, EmailOverrideEntry> };
  applyEventPhotoToLogo: (setter: (b64: string) => void) => Promise<string>;
  autoDeregisterOnDecline: boolean;
  bundledComm: BundledComm;
  childTermPlural: string;
  commToggleRow: (opts: { checked: boolean; onChange: (v: boolean) => void; label: string; short: string; info: React.ReactNode; accent?: string; }) => React.ReactElement;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  disableCancellationEmail: boolean;
  disableEmails: boolean;
  disableOutlook: boolean;
  disableRegistrationEmail: boolean;
  durchstarterCapacity: string;
  effectiveHeaderImage: (kind: 'email' | 'outlook', own: string) => {    src: string;    note: string;};
  emailLanguage: string;
  emailLogoFromPhoto: boolean;
  emailLogoPreview: string;
  emailTemplateOverrides: Record<string, EmailOverrideEntry>;
  emailTemplates: { id: number; templateType: string; language: string; subject: string; heading: string; headingColor: string; bodyHtml: string; }[];
  funstarterCapacity: string;
  imageFile: File;
  imagePreview: string;
  inactiveHandling: "notify" | "autoderegister";
  isDe: boolean;
  mainCommDisabledAck: boolean;
  maxParticipants: string;
  notifyOrgCancelMode: "never" | "always" | "afterDeadline";
  notifyOrgRegisterFromDate: string;
  notifyOrgRegisterMode: "never" | "always" | "fromDate";
  offerLogoToSubEvents: (kind: 'email' | 'outlook', b64: string) => Promise<void>;
  organizer: string;
  outlookBody: string;
  outlookLogoFromPhoto: boolean;
  outlookLogoPreview: string;
  renderHeaderSizeControl: (previewSrc: string, note?: string) => React.ReactElement;
  renderOutlookUpdateButton: () => React.ReactNode;
  renderStepIntro: (_bulletsDe: string[], _bulletsEn: string[]) => React.ReactElement | null;
  setAutoDeregisterOnDecline: React.Dispatch<React.SetStateAction<boolean>>;
  setBundledComm: React.Dispatch<React.SetStateAction<BundledComm>>;
  setDisableCancellationEmail: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableEmails: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableOutlook: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableRegistrationEmail: React.Dispatch<React.SetStateAction<boolean>>;
  setEmailLanguage: React.Dispatch<React.SetStateAction<string>>;
  setEmailLogoFromPhoto: React.Dispatch<React.SetStateAction<boolean>>;
  setEmailLogoPreview: React.Dispatch<React.SetStateAction<string>>;
  setEmailTemplateOverrides: React.Dispatch<React.SetStateAction<Record<string, EmailOverrideEntry>>>;
  setHtmlEditorMode: React.Dispatch<React.SetStateAction<"outlook" | "email" | "description">>;
  setHtmlEditorOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setHtmlEditorTemplateType: React.Dispatch<React.SetStateAction<string>>;
  setInactiveHandling: React.Dispatch<React.SetStateAction<"notify" | "autoderegister">>;
  setLogoCropTarget: React.Dispatch<React.SetStateAction<"outlook" | "email">>;
  setMainCommDisabledAck: React.Dispatch<React.SetStateAction<boolean>>;
  setNotifyOrgCancelMode: React.Dispatch<React.SetStateAction<"never" | "always" | "afterDeadline">>;
  setNotifyOrgRegisterFromDate: React.Dispatch<React.SetStateAction<string>>;
  setNotifyOrgRegisterMode: React.Dispatch<React.SetStateAction<"never" | "always" | "fromDate">>;
  setOutlookLogoFromPhoto: React.Dispatch<React.SetStateAction<boolean>>;
  setOutlookLogoPreview: React.Dispatch<React.SetStateAction<string>>;
  setSubTransfer: React.Dispatch<React.SetStateAction<{ fromIdx: number; groups: string[]; targets: number[]; }>>;
  subEvents: SubEventDraft[];
  subEventsOnlyMode: boolean;
  t: (key: string) => string;
  title: string;
  unlimitedParticipants: boolean;
  useSplitCapacities: boolean;
  waitlistEnabled: boolean;
}
export const CommunicationStep: React.FC<CommunicationStepProps> = (p) => {
  const { visible } = p;
  const { activeCommTabIdx, applyCommTopicToAllSubEvents, commShared, setCommShared, flushActiveCommTabToState, resolveTopLevelCommState, applyEventPhotoToLogo, autoDeregisterOnDecline, bundledComm, childTermPlural, commToggleRow, confirmDialog, disableCancellationEmail, disableEmails, disableOutlook, disableRegistrationEmail, effectiveHeaderImage, emailLanguage, emailLogoFromPhoto, emailLogoPreview, emailTemplateOverrides, emailTemplates, imageFile, imagePreview, inactiveHandling, isDe, mainCommDisabledAck, notifyOrgCancelMode, notifyOrgRegisterFromDate, notifyOrgRegisterMode, offerLogoToSubEvents, organizer, outlookBody, outlookLogoFromPhoto, outlookLogoPreview, renderHeaderSizeControl, renderOutlookUpdateButton, renderStepIntro, setAutoDeregisterOnDecline, setBundledComm, setDisableCancellationEmail, setDisableEmails, setDisableOutlook, setDisableRegistrationEmail, setEmailLanguage, setEmailLogoFromPhoto, setEmailLogoPreview, setEmailTemplateOverrides, setHtmlEditorMode, setHtmlEditorOpen, setHtmlEditorTemplateType, setInactiveHandling, setLogoCropTarget, setMainCommDisabledAck, setNotifyOrgCancelMode, setNotifyOrgRegisterFromDate, setNotifyOrgRegisterMode, setOutlookLogoFromPhoto, setOutlookLogoPreview, subEvents, subEventsOnlyMode, t, title, unlimitedParticipants, waitlistEnabled } = p;

  // v30.89: Ebene 3 („Texte und Bilder anpassen“) — zu, bis jemand sie braucht;
  // die Chip-Zeile öffnet den passenden Reiter. Abmelde-Regel der Organizer-Kopie
  // eingeklappt, solange sie nur gelesen werden muss.
  const [advOpen, setAdvOpen] = React.useState<boolean>(false);
  const [advTab, setAdvTab] = React.useState<'templates' | 'mailLogo' | 'outlook' | 'fine'>('templates');
  const [orgCancelOpen, setOrgCancelOpen] = React.useState<boolean>(false);
  // v30.71: Hilfen für den Schalter "gemeinsam / einzeln" (s. Box oben im Schritt).
  const namedSubCount = subEvents.filter(s => s.title && s.title.trim()).length;
  const mainTabLabel = subEventsOnlyMode ? (isDe ? 'Klammer' : 'Bracket') : (isDe ? 'Haupt-Event' : 'Main event');
  const currentTabLabel = activeCommTabIdx > 0 ? ((subEvents[activeCommTabIdx - 1] && subEvents[activeCommTabIdx - 1].title) || '') : mainTabLabel;
  const childOneDe = 'Jeden Termin';
  // Weicht der offene Termin beim Thema vom Haupt-Event ab? Live-State für die
  // Felder, die der Schritt als Props hat; Betreff/Unterzeile aus dem Slot
  // (die hält der Reiter erst nach dem nächsten Flush aktuell).
  const topicDiffersFromParent = (key: string): boolean => {
    if (activeCommTabIdx <= 0) return false;
    const top = resolveTopLevelCommState();
    const slot: Partial<SubEventDraft> = subEvents[activeCommTabIdx - 1] || {};
    switch (key) {
      case 'language': return !!emailLanguage && emailLanguage !== top.emailLanguage;
      case 'switches': return disableEmails !== top.disableEmails || disableRegistrationEmail !== top.disableRegistrationEmail
        || disableCancellationEmail !== top.disableCancellationEmail || autoDeregisterOnDecline !== top.autoDeregisterOnDecline
        || inactiveHandling !== top.inactiveHandling || disableOutlook !== top.disableOutlook;
      case 'mailLogo': return !!emailLogoPreview && emailLogoPreview !== top.emailLogoBase64;
      case 'outlookLogo': return !!outlookLogoPreview && outlookLogoPreview !== top.outlookLogoBase64;
      case 'outlookText': return (!!outlookBody && outlookBody !== top.outlookBody)
        || !!(slot.outlookSubject || '').trim() || !!(slot.outlookSubheading || '').trim();
      case 'templates': return JSON.stringify(emailTemplateOverrides || {}) !== JSON.stringify(top.emailTemplateOverrides || {});
      default: return false;
    }
  };
  // Umschalten. Auf "gemeinsam": vorher sagen, wie viele Termine bei irgendeinem
  // Thema eigene Werte haben - die bleiben bis zum Speichern erhalten, danach
  // gilt überall der Stand des Haupt-Events.
  const switchCommShared = async (on: boolean): Promise<void> => {
    if (on === commShared) return;
    if (on) {
      flushActiveCommTabToState();
      const top = resolveTopLevelCommState();
      const own = subEvents.filter(x => x.title && x.title.trim()).filter(x =>
        (!!x.emailLanguage && x.emailLanguage !== top.emailLanguage)
        || (!!x.emailLogoBase64 && x.emailLogoBase64 !== top.emailLogoBase64)
        || (!!x.outlookLogoBase64 && x.outlookLogoBase64 !== top.outlookLogoBase64)
        || (!!x.outlookBody && x.outlookBody !== top.outlookBody)
        || !!(x.outlookSubject || '').trim() || !!(x.outlookSubheading || '').trim()
        || (JSON.stringify(x.emailTemplateOverrides || {}) !== '{}' && JSON.stringify(x.emailTemplateOverrides || {}) !== JSON.stringify(top.emailTemplateOverrides || {}))
        || !!x.disableEmails !== !!top.disableEmails || !!x.disableOutlook !== !!top.disableOutlook).length;
      if (own > 0) {
        const term = childTermPlural || (isDe ? 'Termine' : 'dates');
        const ok = await confirmDialog(isDe
          ? `${own} ${own === 1 ? 'Termin hat' : term + ' haben'} eigene Kommunikations-Einstellungen. Im gemeinsamen Modus gelten überall die Werte des Haupt-Events. Die eigenen Werte bleiben erhalten, bis du speicherst — dann werden sie überschrieben. Weiter?`
          : `${own} ${own === 1 ? 'date has' : 'dates have'} their own communication settings. In shared mode the main event's values apply everywhere. The own values stay until you save - then they are overwritten. Continue?`,
          { title: isDe ? 'Gemeinsam für alle Termine' : 'Shared across all dates', confirmLabel: isDe ? 'Gemeinsam' : 'Shared' });
        if (!ok) return;
      }
    }
    setCommShared(on);
  };
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
                <h2 className="dex-step-head-title">
                  {isDe ? 'Schritt 6 — Kommunikation' : 'Step 6 — Communication'}
                </h2>
                <p className="dex-step-head-lead">
                  {isDe
                    ? 'Hier konfigurierst du alle automatischen E-Mails und Outlook-Einladungen — Sprache, Logos, Vorlagen und Versandregeln pro Aktion.'
                    : 'Here you configure all automated emails and Outlook invites — language, logos, templates and per-action send rules.'}
                </p>
                {/* v28.80: Kommunikation eines Sub-Events auf die anderen
                    uebertragen — sonst muss der Organizer Logo, Outlook-Text,
                    Betreff und Mail-Schalter bei jedem Sub-Event einzeln
                    einstellen. Nur sichtbar, wenn ein Sub-Event ausgewaehlt
                    ist; die Klammer-Kommunikation ist eine andere Ebene. */}
                {/* v30.60: Beide Wege stehen jetzt an EINER Stelle und sagen
                    ausdrücklich, was der Normalfall ist. Vorher gab es nur den
                    Weg Sub → Subs, und er erschien ausschließlich auf einem
                    Sub-Reiter — auf der Klammer sah es aus, als gäbe es die
                    Möglichkeit gar nicht, und die Frage „kann man das nicht
                    für alle auf einmal einstellen?" war die logische Folge. */}
                {subEvents.length > 0 && (
                  <div style={{
                    marginBottom: 14, padding: '12px 16px', borderRadius: 10,
                    border: '1px solid var(--dex-gray-200)', background: 'var(--dex-gray-50, #fafafa)',
                  }}>
                    {/* v30.71: Schalter mit Gedächtnis statt Kopier-Knopf.
                        Nutzer-Ansage 02.09.2026: „kein Button, sondern ein
                        Wechselschalter — entweder einzeln oder für alle Termine
                        zusammen." Gemeinsam = die Termin-Reiter sind hier nur
                        Anzeige, gespeichert wird überall der Stand des
                        Haupt-Events (persistSubEvents). Einzeln = Themenliste
                        mit Chip (eigen / wie Haupt-Event) und einem Knopf, der
                        genau EIN Thema auf alle Termine verteilt. */}
                    <div style={{ fontWeight: 600, fontSize: '0.86rem', marginBottom: 4 }}>
                      {isDe ? 'Gelten die Einstellungen für jeden Termin einzeln oder für alle gemeinsam?' : 'Do these settings apply per date or to all dates together?'}
                    </div>
                    <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.55, marginBottom: 10 }}>
                      {isDe
                        ? 'Mails, Kalendereinträge und Texte — du entscheidest einmal, ob jeder Termin seine eigenen bekommt oder alle dieselben.'
                        : 'Emails, calendar entries and texts — decide once whether every date gets its own or all share the same.'}
                    </div>
                    <div role="radiogroup" aria-label={isDe ? 'Modus' : 'Mode'} style={{ display: 'inline-flex', gap: 3, padding: 3, border: '1px solid var(--dex-gray-300)', borderRadius: 999, background: '#fff' }}>
                      {([
                        { on: true, de: `Gemeinsam für alle ${namedSubCount} ${childTermPlural || 'Termine'}`, en: `Shared across all ${namedSubCount} dates`, subDe: 'Einmal einstellen, überall gleich', subEn: 'Set once, same everywhere' },
                        { on: false, de: `${childOneDe} einzeln`, en: 'Each date individually', subDe: 'Jeder Reiter hat eigene Einstellungen', subEn: 'Every tab has its own settings' },
                      ]).map(opt => {
                        const active = commShared === opt.on;
                        return (
                          <button
                            key={String(opt.on)}
                            type="button"
                            role="radio"
                            aria-checked={active}
                            onClick={() => { void switchCommShared(opt.on); }}
                            style={{
                              border: 0, borderRadius: 999, padding: '7px 15px', cursor: 'pointer', textAlign: 'left',
                              background: active ? 'var(--dex-green-dark, #4a7c1f)' : 'transparent',
                              color: active ? '#fff' : 'var(--dex-gray-700)', fontWeight: 600, fontSize: '0.84rem', lineHeight: 1.2,
                            }}
                          >
                            <span style={{ display: 'block' }}>{isDe ? opt.de : opt.en}</span>
                            <span style={{ display: 'block', fontSize: '0.72rem', fontWeight: 400, opacity: active ? 0.85 : 0.7, marginTop: 2 }}>{isDe ? opt.subDe : opt.subEn}</span>
                          </button>
                        );
                      })}
                    </div>
                    {commShared ? (
                      <div style={{ marginTop: 12, padding: '10px 12px', borderRadius: 8, background: 'rgba(134,188,37,0.10)', border: '1px solid var(--dex-green, #86bc25)', fontSize: '0.82rem', lineHeight: 1.5 }}>
                        {activeCommTabIdx === 0
                          ? (isDe
                            ? <><strong>Du stellst alles hier auf dem Reiter „{mainTabLabel}&ldquo; ein.</strong> Was du hier änderst, gilt beim Speichern für alle {namedSubCount} {childTermPlural || 'Termine'}. Die Termin-Reiter zeigen keine eigenen Felder mehr.</>
                            : <><strong>You set everything here on the “{mainTabLabel}” tab.</strong> On save it applies to all {namedSubCount} dates. The date tabs no longer show their own fields.</>)
                          : (isDe
                            ? <><strong>Dieser Termin übernimmt alles vom Reiter „{mainTabLabel}&ldquo;.</strong> Zum Ändern wechsle dorthin — oder stelle oben auf „{childOneDe} einzeln&ldquo;, wenn dieser Termin etwas Eigenes braucht.</>
                            : <><strong>This date takes everything from the “{mainTabLabel}” tab.</strong> Switch there to change it — or choose “Each date individually” above if this date needs something of its own.</>)}
                      </div>
                    ) : (
                      <>
                        <div style={{ marginTop: 12, padding: '10px 12px', borderRadius: 8, background: 'rgba(0,118,168,0.07)', border: '1px solid rgba(0,118,168,0.35)', fontSize: '0.82rem', lineHeight: 1.5 }}>
                          {activeCommTabIdx === 0
                            ? (isDe
                              ? <><strong>Du bearbeitest gerade: {mainTabLabel}.</strong> Jede Zeile unten ist ein Thema. Mit „für alle übernehmen&ldquo; verteilst du genau dieses Thema auf alle {namedSubCount} {childTermPlural || 'Termine'} — mehr nicht.</>
                              : <><strong>You are editing: {mainTabLabel}.</strong> Each row below is one topic. “Apply to all” copies exactly that topic to all {namedSubCount} dates — nothing else.</>)
                            : (isDe
                              ? <><strong>Du bearbeitest gerade: {currentTabLabel}.</strong> Jede Zeile zeigt, ob dieser Termin etwas Eigenes hat oder das Haupt-Event übernimmt. Mit „für alle übernehmen&ldquo; verteilst du genau dieses Thema auf die anderen {childTermPlural || 'Termine'}.</>
                              : <><strong>You are editing: {currentTabLabel}.</strong> Each row shows whether this date has something of its own or inherits from the main event. “Apply to all” copies exactly that topic to the other dates.</>)}
                        </div>
                        <div style={{ marginTop: 10, borderTop: '1px solid var(--dex-gray-200)' }}>
                          {COMM_TOPICS.map(topic => {
                            const own = activeCommTabIdx > 0 && topicDiffersFromParent(topic.key);
                            return (
                              <div key={topic.key} style={{ display: 'grid', gridTemplateColumns: '30px 1fr auto', gap: 10, alignItems: 'center', padding: '8px 0', borderBottom: '1px solid var(--dex-gray-200)' }}>
                                <StepBadge n={topic.step} />
                                <div style={{ minWidth: 0 }}>
                                  <div style={{ fontWeight: 600, fontSize: '0.86rem' }}>{t(topic.labelKey)}</div>
                                </div>
                                <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                                  {activeCommTabIdx > 0 && (
                                    <span style={{
                                      fontSize: '0.72rem', padding: '2px 9px', borderRadius: 999, whiteSpace: 'nowrap',
                                      border: `1px solid ${own ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-300)'}`,
                                      color: own ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-600)',
                                      background: own ? 'rgba(237,139,0,0.10)' : '#fff',
                                    }}>
                                      {own ? (isDe ? 'eigen' : 'own') : (isDe ? 'wie Haupt-Event' : 'as main event')}
                                    </span>
                                  )}
                                  {namedSubCount > (activeCommTabIdx > 0 ? 1 : 0) && (
                                    <button
                                      type="button"
                                      className="btn btn-secondary"
                                      style={{ fontSize: '0.74rem', padding: '3px 10px', borderRadius: 999 }}
                                      onClick={() => { void applyCommTopicToAllSubEvents(topic.key); }}
                                      title={isDe ? `Nur „${t(topic.labelKey)}" von diesem Reiter auf alle ${childTermPlural || 'Termine'} übertragen` : `Copy only this topic from this tab to all dates`}
                                    >
                                      {isDe ? `für alle ${namedSubCount} übernehmen` : `apply to all ${namedSubCount}`}
                                    </button>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </>
                    )}
                    <div style={{ marginTop: 14, paddingTop: 14, borderTop: '1px solid var(--dex-gray-200)' }} />
                    {/* v30.71: Bündel-Block jetzt UNTER dem Schalter - oben geht es um
                        den Inhalt, hier um die Anzahl der Mails (Nutzer-Entscheidung 02.09.). */}
                    {/* v30.61: Der eigentliche Schalter — eine Mail und ein
                        Kalendereintrag fürs ganze Event statt einem je Termin.
                        Das ist etwas anderes als „überall dieselben Texte":
                        Dort verschicken weiterhin zehn Termine zehn Mails, hier
                        verschickt die Klammer EINE. Nur im Klammer-Modus, weil
                        ein buchbares Haupt-Event ohnehin selbst verschickt. */}
                    {subEventsOnlyMode && (
                      <div style={{ marginBottom: 14, paddingBottom: 14, borderBottom: '1px solid var(--dex-gray-200)' }}>
                        <div style={{ fontWeight: 600, fontSize: '0.86rem', marginBottom: 4 }}>
                          {isDe ? 'Wie viele Mails bekommt jemand, der mehrere Termine bucht?' : 'How many emails does someone get who books several dates?'}
                        </div>
                        <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.55, marginBottom: 10 }}>
                          {isDe
                            ? `Wer sich für mehrere ${childTermPlural || 'Termine'} anmeldet, bekommt sonst für jeden eine eigene Bestätigung und einen eigenen Kalendereintrag. Gebündelt kommt stattdessen EINE Bestätigung mit der Liste aller gebuchten Termine — und ein Kalendereintrag über den Gesamtzeitraum.`
                            : `Someone registering for several dates otherwise receives one confirmation and one calendar entry per date. Bundled, they get ONE confirmation listing all booked dates.`}
                        </div>
                        {([
                          { key: 'mail' as const, de: 'Bestätigungs-Mails bündeln', en: 'Bundle confirmation emails', hintDe: 'Eine Mail mit der Liste aller gebuchten Termine.', hintEn: 'One email listing all booked dates.' },
                          { key: 'outlook' as const, de: 'Kalendereintrag bündeln', en: 'Bundle the calendar entry', hintDe: 'Ein Eintrag über den Zeitraum des Gesamt-Events. Wer nur Tag 2 und 4 bucht, bekommt trotzdem einen Eintrag über den ganzen Zeitraum — welche Tage gebucht sind, steht in der Beschreibung.', hintEn: 'One entry spanning the whole event period.' },
                          { key: 'qr' as const, de: 'Einen QR-Code fürs Gesamt-Event', en: 'One QR code for the whole event', hintDe: 'Der Check-in läuft dann über das Haupt-Event statt über die einzelnen Termine.', hintEn: 'Check-in then runs on the main event.' },
                        ]).map(opt => (
                          <label key={opt.key} style={{ display: 'flex', gap: 9, alignItems: 'flex-start', marginBottom: 8, cursor: 'pointer' }}>
                            <input
                              type="checkbox"
                              checked={bundledComm[opt.key]}
                              onChange={e => setBundledComm(prev => ({ ...prev, [opt.key]: e.target.checked }))}
                              style={{ marginTop: 3, flexShrink: 0 }}
                            />
                            <span style={{ minWidth: 0 }}>
                              <span style={{ display: 'block', fontSize: '0.84rem', fontWeight: 600 }}>{isDe ? opt.de : opt.en}</span>
                              <span style={{ display: 'block', fontSize: '0.76rem', color: 'var(--dex-gray-500)', lineHeight: 1.5 }}>{isDe ? opt.hintDe : opt.hintEn}</span>
                            </span>
                          </label>
                        ))}
                        {(bundledComm.mail || bundledComm.outlook) && (
                          <div style={{ marginTop: 8, padding: '8px 12px', borderRadius: 8, background: 'rgba(0,118,168,0.07)', fontSize: '0.77rem', lineHeight: 1.5, color: 'var(--dex-gray-700)' }}>
                            {isDe
                              ? `Die Texte für die gebündelte Mail stehen auf dem Reiter ${subEventsOnlyMode ? 'Klammer' : 'Haupt-Event'}. Was du auf den Termin-Reitern eingestellt hast, bleibt gespeichert, wirkt aber nicht mehr, solange gebündelt wird.`
                              : 'The copy for the bundled email lives on the bracket tab. Per-date settings stay saved but have no effect while bundling is on.'}
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}
                {renderStepIntro(
                  [
                    'Drei Entscheidungen oben: Sprache der Mails, Weg der Kommunikation (Mail und/oder Outlook-Termin) und ob die Organizer in Kopie stehen',
                    'Die Zeile darunter zeigt, was mit diesen Einstellungen tatsächlich rausgeht — Standard heißt: fertig, nichts weiter zu tun',
                    'Nur wer Texte, Bilder oder Feineinstellungen ändern will, öffnet „Texte und Bilder anpassen“ — dort liegen Mail-Vorlagen, Mail-Logo, Outlook-Termin und die seltenen Schalter',
                  ],
                  [
                    'Three decisions at the top: mail language, communication channel (email and/or Outlook invite) and whether organizers are copied',
                    'The line below shows what actually goes out with these settings — default means: done, nothing else to do',
                    'Only if you want to change texts, images or fine-tuning, open “Customise texts and images” — that is where mail templates, mail logo, Outlook invite and the rare switches live',
                  ]
                )}
                <h3 className="mb-16">{t('create.step.communication')}</h3>

                {/* v28.88: Die Reiter-Leiste stand hier ein ZWEITES Mal. Seit
                    v28.78 trägt die Scope-Karte über dem Formular
                    (renderGlobalScopeBar) den Umschalter für alle
                    scope-fähigen Schritte — Kommunikation eingeschlossen, sie
                    hängt über setScope am selben Index. Zwei identische
                    Reiter-Reihen auf einer Seite lesen sich als zwei
                    Navigationen: der Organizer sucht, welche die gültige ist.
                    Die Steps „Ort & Programm", „Kapazität" und „Felder" sind
                    schon in v28.78 entkoppelt worden, Kommunikation blieb
                    übrig. Der Erklär-Tooltip bleibt — er sagt, was pro
                    Sub-Event überhaupt getrennt einstellbar ist. */}
                {subEvents.length > 0 && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 16 }}>
                    <div style={{ flex: 1, minWidth: 0, fontSize: '0.84rem', color: 'var(--dex-gray-600)' }}>
                      {isDe
                        ? 'Die Einstellungen unten gelten für den oben gewählten Reiter.'
                        : 'The settings below apply to the tab selected above.'}
                    </div>
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Was du hier einstellst:</strong> jeder Sub-Event
                        darf seine eigenen Kommunikations-Einstellungen haben —
                        Mail-Sprache, Texte, BCC-Empfänger für Organizer,
                        eigenes Mail- und Outlook-Logo sowie eine eigene
                        Outlook-Termin-Konfiguration (Überschrift, Beschreibung,
                        Ein/Aus-Schalter für den Termin).<br /><br />
                        <strong>Anzeige in der App:</strong> wechselst du auf
                        den Tab eines Sub-Events, werden die Felder unten mit
                        den Werten genau dieses Sub-Events geladen. Speichern
                        am Ende des Wizards persistiert für jeden Sub-Event
                        die zugehörigen Werte separat.<br /><br />
                        <strong>Auswirkung für Teilnehmer:</strong> ein
                        Teilnehmer, der sich für Sub-Event A anmeldet, bekommt
                        die Bestätigungs-Mail in der Sprache und mit dem Text,
                        den du auf dem Tab &bdquo;A&ldquo; eingestellt hast. Anmeldungen
                        zu Sub-Event B verwenden den Tab &bdquo;B&ldquo;. So lassen sich
                        z.B. ein deutsches und ein englisches Sub-Event
                        sauber nebeneinander pflegen.
                      </>
                    ) : (
                      <>
                        <strong>What you set here:</strong> every sub-event may
                        have its own communication settings — email language,
                        copy, BCC recipients for organizers, its own email and
                        Outlook logo, and its own Outlook invite (heading,
                        description, on/off toggle).<br /><br />
                        <strong>Where it shows up:</strong> switching to a
                        sub-event tab loads the fields below with that
                        sub-event&apos;s values. Saving at the end of the
                        wizard persists each sub-event&apos;s values
                        separately.<br /><br />
                        <strong>Effect for attendees:</strong> someone who
                        registers for sub-event A receives the confirmation
                        email in the language and wording you configured on
                        tab &ldquo;A&rdquo;. Registrations for sub-event B
                        use tab &ldquo;B&rdquo;. This lets you cleanly run
                        e.g. a German and an English sub-event side by side.
                      </>
                    )} />
                  </div>
                )}


                {/* v14.8: „Nur Sub-Events"-Modus + auf Haupt-Event-Tab → Banner
                    statt Kommunikations-Settings rendern. Der User soll keine
                    Werte für ein nicht-existentes Hauptevent-Anmelden pflegen. */}
                {subEventsOnlyMode && activeCommTabIdx === 0 && (
                  <div style={{
                    display: 'flex', alignItems: 'flex-start', gap: 10,
                    padding: '14px 16px', marginBottom: 8,
                    background: 'rgba(255,193,7,0.10)',
                    border: '1px solid rgba(255,152,0,0.55)',
                    borderRadius: 'var(--dex-radius, 12px)',
                    fontSize: '0.9rem', color: 'var(--dex-gray-700)',
                    lineHeight: 1.55,
                  }}>
                    <Icon iconName="Info" style={{ fontSize: 20, color: '#e67e22', flexShrink: 0, marginTop: 2 }} />
                    <div>
                      {isDe
                        ? <>
                            <strong>Hauptevent-Kommunikation ist in diesem Modus nicht relevant.</strong><br />
                            Du hast in Schritt 1 (Grundlagen) den Modus <strong>&bdquo;Nur {(childTermPlural || 'Sub-Events').trim() || 'Sub-Events'}&ldquo;</strong> gewählt — Teilnehmer können sich gar nicht fürs Hauptevent anmelden, deshalb gibt es auch keine Bestätigungs-Mails und keinen Outlook-Termin fürs Hauptevent. Wechsle auf den Tab eines Sub-Events, um dort die Kommunikation zu konfigurieren.
                          </>
                        : <>
                            <strong>Main-event communication is not relevant in this mode.</strong><br />
                            You picked the <strong>&bdquo;{(childTermPlural || 'sub-events').trim() || 'sub-events'} only&ldquo;</strong> mode in step 1 (Basics) — attendees cannot register for the main event, so no confirmation emails or Outlook invites are sent for it. Switch to a sub-event tab to configure communication there.
                          </>}
                    </div>
                  </div>
                )}

                {/* v30.71: Im gemeinsamen Modus zeigt ein Termin-Reiter keine
                    eigenen Felder — der Kasten oben sagt, wo man ändert. */}
                {!(subEventsOnlyMode && activeCommTabIdx === 0) && !(commShared && activeCommTabIdx > 0 && subEvents.length > 0) && (
                <>
                {/* v30.89: Ebene 1 — drei Entscheidungen, immer sichtbar (Konzept
                    docs/konzept-kommunikation-schritt.md). Alles Seltene liegt unten
                    hinter „Texte und Bilder anpassen“. */}
                <div className="form-group">
                  <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <StepBadge n={25} />
                    {t('create.emaillanguage')}
                  </label>
                  <div style={{ display: 'flex', gap: 8 }}>
                    {(['DE', 'EN'] as const).map(lang => (
                      <button
                        key={lang}
                        className={`btn ${emailLanguage === lang ? 'btn-primary' : 'btn-secondary'}`}
                        style={{ minWidth: 80 }}
                        onClick={() => setEmailLanguage(lang)}
                      >
                        {lang === 'DE' ? 'DE – Deutsch' : 'EN – English'}
                      </button>
                    ))}
                  </div>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>
                    {t('create.emaillanguage.hint')}
                  </p>
                </div>
                {/* v30.89: Kanal-Karte ersetzt die zwei Master-Haken aus dem alten
                    Kasten 26 — eine Frage, vier Antworten, die Folge steht direkt darunter. */}
                <div className="form-group" style={{ marginTop: 24 }}>
                  <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <StepBadge n={26} />
                    {isDe ? 'Wie erreicht DEX die Teilnehmer?' : 'How does DEX reach attendees?'}
                    <InfoTooltip text={t('create.notifications.hint')} />
                  </label>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(190px, 1fr))', gap: 10 }}>
                    {([
                      { key: 'both', mails: true, outlook: true, de: 'Mail + Outlook-Termin', en: 'Email + Outlook invite', subDe: 'Standard — Bestätigung per Mail und Termin im Kalender', subEn: 'Default — confirmation by email and calendar entry' },
                      { key: 'mail', mails: true, outlook: false, de: 'Nur Mail', en: 'Email only', subDe: 'Kein Kalendereintrag — Teilnehmer planen selbst', subEn: 'No calendar entry — attendees schedule themselves' },
                      { key: 'outlook', mails: false, outlook: true, de: 'Nur Outlook-Termin', en: 'Outlook invite only', subDe: 'Keine Mails — der Termin ist die Bestätigung', subEn: 'No emails — the invite is the confirmation' },
                      { key: 'none', mails: false, outlook: false, de: 'Keine Kommunikation', en: 'No communication', subDe: 'Nichts geht raus — z.B. zum internen Testen', subEn: 'Nothing goes out — e.g. for internal testing' },
                    ]).map(opt => {
                      const selected = !disableEmails === opt.mails && !disableOutlook === opt.outlook;
                      return (
                        <label key={opt.key} style={{
                          display: 'flex', alignItems: 'flex-start', gap: 10, padding: '10px 14px', borderRadius: 10, cursor: 'pointer',
                          border: `1px solid ${selected ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                          background: selected ? 'rgba(134,188,37,0.06)' : '#fff',
                        }}>
                          <input type="radio" name="commChannel" checked={selected} onChange={() => { setDisableEmails(!opt.mails); setDisableOutlook(!opt.outlook); }} style={{ marginTop: 3, flexShrink: 0 }} />
                          <span style={{ minWidth: 0 }}>
                            <span style={{ display: 'block', fontSize: '0.88rem', fontWeight: 700 }}>{isDe ? opt.de : opt.en}</span>
                            <span style={{ display: 'block', fontSize: '0.76rem', color: 'var(--dex-gray-500)', lineHeight: 1.45, marginTop: 2 }}>{isDe ? opt.subDe : opt.subEn}</span>
                          </span>
                        </label>
                      );
                    })}
                  </div>
                  {(disableEmails || disableOutlook) && (
                    <p style={{ margin: '8px 0 0', fontSize: '0.8rem', color: 'var(--dex-orange-dark, #b35a00)', lineHeight: 1.5 }}>
                      {disableEmails && disableOutlook
                        ? (isDe ? 'Achtung: Niemand erfährt von seiner Anmeldung — weder per Mail noch im Kalender. Nur sinnvoll, wenn du selbst einlädst.' : 'Careful: nobody hears about their registration — neither by email nor in the calendar. Only sensible if you invite people yourself.')
                        : disableEmails
                          ? (isDe ? 'Es geht keine einzige Mail an Teilnehmer raus — auch keine Wartelisten- oder Abmelde-Mail.' : 'Not a single email goes out to attendees — no waitlist or cancellation email either.')
                          : (isDe ? 'Kein Termin im Kalender — bei einer Abmeldung gibt es entsprechend auch nichts zu entfernen.' : 'No calendar entry — accordingly nothing is removed on cancellation.')}
                    </p>
                  )}
                  {activeCommTabIdx === 0 && subEvents.length > 0 && (disableEmails || disableOutlook) && (
                    <WizardHint
                      isDe={isDe}
                      title={isDe ? 'Kommunikation für das Hauptevent ist deaktiviert' : 'Communication for the main event is disabled'}
                      style={{ marginTop: 16 }}
                      // Pflicht-Checkbox im Inhalt — muss sichtbar starten,
                      // sonst übersieht der Organizer die Bestätigung und
                      // wundert sich über den blockierten Save.
                      defaultOpen={true}
                    >
                      <div style={{ marginBottom: 10 }}>
                        {isDe
                          ? <>Wer sich <strong>nur für das Hauptevent</strong> anmeldet (und kein Sub-Event auswählt), bekommt damit weder eine Bestätigungs-Mail noch einen Kalender-Termin. Stelle sicher, dass die Teilnehmer im Anmeldeformular <strong>immer mindestens ein Sub-Event</strong> angeben müssen — sonst verlierst du sie kommunikativ.</>
                          : <>Whoever registers <strong>only for the main event</strong> (without picking a sub-event) gets neither a confirmation email nor a calendar invite. Make sure attendees are required to pick <strong>at least one sub-event</strong> in the registration form — otherwise you lose them communication-wise.</>}
                      </div>
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={mainCommDisabledAck}
                          onChange={e => setMainCommDisabledAck(e.target.checked)}
                          style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 2, flexShrink: 0 }}
                        />
                        <span style={{ fontWeight: 600, color: 'var(--dex-gray-800)' }}>
                          {isDe
                            ? 'Ja, mir ist bewusst, dass Teilnehmer sich für mindestens ein Sub-Event anmelden müssen, um Kommunikation zu erhalten.'
                            : 'Yes, I understand attendees need to register for at least one sub-event to receive communication.'}
                        </span>
                      </label>
                    </WizardHint>
                  )}
                </div>
                {/* v30.89: Mitleser-Karte — immer sichtbar (vorher Aufklapper 27). Die
                    Abmelde-Regel bleibt eingeklappt, bis jemand sie abweichend braucht. */}
                <div className="form-group" style={{ marginTop: 24 }}>
                  <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <StepBadge n={27} />
                    {isDe ? 'Sollen die Organizer bei An- und Abmeldungen mitlesen?' : 'Should organizers be looped in on registrations / cancellations?'}
                  </label>
                  <div style={{ padding: '12px 14px', borderRadius: 10, border: '1px solid var(--dex-gray-200)', background: '#fff' }}>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12, lineHeight: 1.5 }}>
                    {/* v28.28: Präzisiert — die Organizer-Kopie ist normalerweise
                        BCC (unsichtbar), bei EXTERNEN Empfängern steht der
                        Organizer aber bewusst sichtbar auf CC (v18.74), damit die
                        externe Person den Ansprechpartner kennt. Die frühere
                        Formulierung „der Teilnehmer sieht nicht, dass jemand
                        mitliest" stimmte für diesen Fall nicht. */}
                    {/* v28.28: Die Organizer-Kopie läuft jetzt durchgehend auf
                        CC (vorher Bcc bei internen Empfängern) — der Organizer
                        steht damit sichtbar im Verteiler, „Allen antworten"
                        landet beim richtigen Ansprechpartner. */}
                    {isDe
                      ? <>Wenn aktiv, steht der Organizer bei der Bestätigungs-Mail an den Teilnehmer sichtbar auf <strong>Kopie (Cc)</strong> — praktisch, um zu wissen, wer sich gerade an- oder abmeldet, und der Teilnehmer sieht direkt, wer sein Ansprechpartner ist. Bei großen Events willst du das vielleicht nicht für jede einzelne Anmeldung — dann kannst du es hier gezielt einschränken (z.B. nur kurz vorm Event, wenn kurzfristige Änderungen wichtig sind).</>
                      : <>When on, the organizer is visibly on <strong>copy (Cc)</strong> of the confirmation email sent to the attendee — handy to know who is signing up or off, and the attendee immediately sees who to contact. For large events you might not want this for every single sign-up — you can narrow it down here (e.g. only close to the event when last-minute changes matter).</>}
                  </p>

                  {/* Anmeldung */}
                  <div style={{ marginBottom: 16 }}>
                    <div style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 6 }}>
                      {isDe ? 'Bei Anmeldungen' : 'On registrations'}
                    </div>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12 }}>
                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', cursor: 'pointer' }}>
                        <input type="radio" name="notifyOrgRegister" checked={notifyOrgRegisterMode === 'never'} onChange={() => setNotifyOrgRegisterMode('never')} />
                        {isDe ? 'Nicht informieren' : 'Don\'t notify'}
                      </label>
                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', cursor: 'pointer' }}>
                        <input type="radio" name="notifyOrgRegister" checked={notifyOrgRegisterMode === 'always'} onChange={() => setNotifyOrgRegisterMode('always')} />
                        {isDe ? 'Bei jeder Anmeldung' : 'On every registration'}
                      </label>
                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', cursor: 'pointer' }}>
                        <input type="radio" name="notifyOrgRegister" checked={notifyOrgRegisterMode === 'fromDate'} onChange={() => setNotifyOrgRegisterMode('fromDate')} />
                        {isDe ? 'Erst ab Datum' : 'Only from date'}
                      </label>
                    </div>
                    {notifyOrgRegisterMode === 'fromDate' && (
                      <div style={{ marginTop: 10, paddingLeft: 24 }}>
                        <DatePicker
                          selected={notifyOrgRegisterFromDate ? new Date(notifyOrgRegisterFromDate) : null}
                          onChange={(date: Date | null) => setNotifyOrgRegisterFromDate(date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}` : '')}
                          showTimeSelect
                          timeFormat="HH:mm"
                          timeIntervals={15}
                          timeCaption={isDe ? 'Uhrzeit' : 'Time'}
                          dateFormat="dd.MM.yyyy, HH:mm"
                          locale="de"
                          placeholderText={isDe ? 'Ab diesem Datum BCC' : 'BCC from this date'}
                          className="form-input"
                          wrapperClassName="dex-datepicker-wrapper"
                          calendarClassName="dex-datepicker-calendar"
                          isClearable
                          autoComplete="off"
                        />
                        <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                          {isDe ? 'Z.B. eine Woche vor dem Event — kurzfristige Anmeldungen werden dann an die Organizer gespiegelt.' : 'E.g. one week before the event — last-minute registrations are mirrored to organizers from then on.'}
                        </p>
                      </div>
                    )}
                  </div>
                  {orgCancelOpen ? (
                    <>
                  {/* Abmeldung */}
                  <div>
                    <div style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 6 }}>
                      {isDe ? 'Bei Abmeldungen' : 'On cancellations'}
                    </div>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12 }}>
                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', cursor: 'pointer' }}>
                        <input type="radio" name="notifyOrgCancel" checked={notifyOrgCancelMode === 'never'} onChange={() => setNotifyOrgCancelMode('never')} />
                        {isDe ? 'Nicht informieren' : 'Don\'t notify'}
                      </label>
                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', cursor: 'pointer' }}>
                        <input type="radio" name="notifyOrgCancel" checked={notifyOrgCancelMode === 'always'} onChange={() => setNotifyOrgCancelMode('always')} />
                        {isDe ? 'Bei jeder Abmeldung' : 'On every cancellation'}
                      </label>
                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', cursor: 'pointer' }}>
                        <input type="radio" name="notifyOrgCancel" checked={notifyOrgCancelMode === 'afterDeadline'} onChange={() => setNotifyOrgCancelMode('afterDeadline')} />
                        {isDe ? 'Erst nach der letzten Abmeldemöglichkeit' : 'Only after the last cancellation date'}
                      </label>
                    </div>
                    <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 6 }}>
                      {isDe
                        ? '„Erst nach der letzten Abmeldemöglichkeit" nutzt das in Schritt 4 (Kapazität & Sichtbarkeit) gesetzte Datum „Letzte Abmeldemöglichkeit". Vor diesem Stichtag gelten Abmeldungen als unproblematisch — danach möchtest du als Organizer aber wissen, wer noch abspringt.'
                        : '„Only after the last cancellation date" uses the date set in step 4 (Capacity & Visibility) under „Last cancellation date". Cancellations before that are considered routine — after that, organizers usually want to know about late drop-outs.'}
                    </p>
                  </div>
                    </>
                  ) : (
                    <button type="button" onClick={() => setOrgCancelOpen(true)} style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', fontSize: '0.8rem', color: 'var(--dex-gray-600)', textDecoration: 'underline' }}>
                      {isDe
                        ? `Bei Abmeldungen: ${notifyOrgCancelMode === 'always' ? 'bei jeder Abmeldung' : notifyOrgCancelMode === 'afterDeadline' ? 'erst nach der letzten Abmeldemöglichkeit' : 'nicht informieren'} — ändern`
                        : `On cancellations: ${notifyOrgCancelMode === 'always' ? 'on every cancellation' : notifyOrgCancelMode === 'afterDeadline' ? 'only after the last cancellation date' : 'do not notify'} — change`}
                    </button>
                  )}
                  </div>
                </div>
                {/* v30.89: Ebene 2 — was mit diesen Einstellungen rausgeht, als Chips.
                    Ersetzt die achtzeilige Übersichtsbox (v19.23) am Kopf des Schritts:
                    Jeder Chip nennt ein Thema und seinen Stand; ein Klick öffnet den
                    passenden Reiter unten. Das ist die Antwort auf „gilt der Standard,
                    wenn ich nichts aufgeklappt habe?“ — ja, und hier steht welcher. */}
                {(() => {
                  const tplCount = Object.keys(emailTemplateOverrides || {}).filter(k => k.charAt(0) !== '_').length;
                  const fineChanged = disableRegistrationEmail || disableCancellationEmail || autoDeregisterOnDecline || inactiveHandling === 'autoderegister';
                  const std = isDe ? 'Standard' : 'default';
                  const chips: Array<{ tab: 'templates' | 'mailLogo' | 'outlook' | 'fine'; label: string; value: string; changed: boolean }> = [
                    { tab: 'templates', label: isDe ? 'Mail-Texte' : 'Mail texts', value: tplCount > 0 ? (isDe ? `${tplCount} angepasst` : `${tplCount} customised`) : std, changed: tplCount > 0 },
                    { tab: 'mailLogo', label: isDe ? 'Mail-Logo' : 'Mail logo', value: emailLogoPreview ? (emailLogoFromPhoto ? (isDe ? 'Event-Foto' : 'event photo') : (isDe ? 'eigenes Bild' : 'own image')) : std, changed: !!emailLogoPreview },
                    { tab: 'outlook', label: isDe ? 'Outlook-Termin' : 'Outlook invite', value: (outlookBody || outlookLogoPreview) ? (isDe ? 'angepasst' : 'customised') : std, changed: !!(outlookBody || outlookLogoPreview) },
                    { tab: 'fine', label: isDe ? 'Feineinstellungen' : 'Fine-tuning', value: fineChanged ? (isDe ? 'angepasst' : 'customised') : std, changed: fineChanged },
                  ];
                  return (
                    <div style={{ marginTop: 28, paddingTop: 16, borderTop: '2px solid var(--dex-gray-200)' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', marginBottom: 10 }}>
                        <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', fontWeight: 600 }}>{isDe ? 'So geht es raus:' : 'What goes out:'}</span>
                        {chips.map(c => (
                          <button key={c.tab} type="button" onClick={() => { setAdvTab(c.tab); setAdvOpen(true); }}
                            title={isDe ? 'Anpassen' : 'Customise'}
                            style={{
                              display: 'inline-flex', alignItems: 'center', gap: 6, padding: '3px 10px', borderRadius: 999, cursor: 'pointer', fontSize: '0.76rem',
                              border: `1px solid ${c.changed ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-300)'}`,
                              background: c.changed ? 'rgba(237,139,0,0.10)' : '#fff', color: c.changed ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-700)',
                            }}>
                            <span style={{ fontWeight: 600 }}>{c.label}:</span> {c.value}
                          </button>
                        ))}
                      </div>
                      <button type="button" className={`btn ${advOpen ? 'btn-primary' : 'btn-secondary'}`} onClick={() => setAdvOpen(v => !v)} style={{ fontSize: '0.85rem', padding: '7px 16px' }}>
                        {advOpen ? (isDe ? 'Anpassen schließen' : 'Close customisation') : (isDe ? 'Texte und Bilder anpassen' : 'Customise texts and images')}
                      </button>
                    </div>
                  );
                })()}
                {/* v30.89: Ebene 3 — vier Reiter statt fünf Aufklapper untereinander.
                    Wenige, lange Inhalte → Reiter (NN/G); die alten Kästen 26 (Rest), 28,
                    29+30 und 31 liegen hier zeichengleich, nur ohne <details>/<summary>. */}
                {advOpen && (
                  <div style={{ marginTop: 14, border: '1px solid var(--dex-gray-200)', borderRadius: 12, background: 'var(--dex-gray-50, #f8f9fa)' }}>
                    <div role="tablist" style={{ display: 'flex', flexWrap: 'wrap', gap: 4, padding: '8px 8px 0', borderBottom: '1px solid var(--dex-gray-200)' }}>
                      {([
                        { key: 'templates' as const, n: 31, label: `${t('create.templates.title')} (${emailLanguage})` },
                        { key: 'mailLogo' as const, n: 28, label: t('create.eventlogo.mail') },
                        { key: 'outlook' as const, n: 29, label: isDe ? 'Outlook-Termin (Bild + Text)' : 'Outlook invite (image + text)' },
                        { key: 'fine' as const, n: 26, label: isDe ? 'Feineinstellungen' : 'Fine-tuning' },
                      ]).map(tab => {
                        const active = advTab === tab.key;
                        return (
                          <button key={tab.key} type="button" role="tab" aria-selected={active} onClick={() => setAdvTab(tab.key)} style={{
                            display: 'inline-flex', alignItems: 'center', gap: 8, padding: '8px 14px', border: 'none', cursor: 'pointer',
                            borderBottom: `3px solid ${active ? 'var(--dex-green, #86bc25)' : 'transparent'}`, background: 'transparent',
                            fontWeight: active ? 700 : 500, fontSize: '0.86rem', color: active ? 'var(--dex-gray-900)' : 'var(--dex-gray-600)',
                          }}>
                            <StepBadge n={tab.n} /> {tab.label}
                          </button>
                        );
                      })}
                    </div>
                    <div style={{ padding: 16 }}>
                    {advTab === 'templates' && (<>
                      <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', margin: '0 0 6px' }}>{isDe ? 'Standard aktiv, solange kein Text angepasst ist. Jede Vorlage lässt sich einzeln zurücksetzen.' : 'Default applies while no text is customised. Each template can be reset individually.'}</p>
                  <div style={{ marginTop: 12 }}>
                <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 12 }}>
                  {t('create.templates.hint')}
                </p>

                {/* TemplateType in DEX_EmailTemplates ist ASCII 'Nachruecken' (Umlaut nicht erlaubt in Choice-Feld).
                    v9.17: Warteliste/Nachrücken-Templates nur anzeigen, wenn das Event eine
                    Warteliste hat — sonst werden sie ohnehin nie genutzt. */}
                {['Anmeldung', 'Warteliste', 'Abmeldung', 'Nachruecken']
                  .filter(tType => {
                    // v9.28: Wartelisten-/Nachrück-Templates nur zeigen, wenn das Event
                    // tatsächlich eine Warteliste haben kann — also Warteliste aktiviert
                    // UND nicht unbegrenzt Teilnehmer (sonst gibt's nie eine volle Kapazität).
                    if (tType === 'Warteliste' || tType === 'Nachruecken') {
                      return waitlistEnabled && !unlimitedParticipants;
                    }
                    return true;
                  })
                  .map(tType => {
                  const defaultTpl = emailTemplates.find(t => t.templateType === tType && t.language === emailLanguage);
                  const override = emailTemplateOverrides[tType];
                  const currentSubject = override?.subject || defaultTpl?.subject || '';
                  const currentBody = override?.bodyHtml || defaultTpl?.bodyHtml || '';
                  const currentHeading = override?.heading || defaultTpl?.heading || '';

                  return (
                    <div key={tType} style={{
                      border: '1px solid var(--dex-gray-200)', borderRadius: 8,
                      padding: 12, marginBottom: 12, background: override ? 'rgba(134,188,37,0.08)' : '#fff',
                    }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <div>
                          <strong style={{ fontSize: '0.85rem' }}>{t(`create.tpl.${tType}`)}</strong>
                          {override && <span style={{ fontSize: '0.7rem', color: 'var(--dex-green)', marginLeft: 8 }}>{t('create.templates.modified')}</span>}
                        </div>
                        <div style={{ display: 'flex', gap: 4 }}>
                          <button
                            className="btn btn-secondary"
                            style={{ fontSize: '0.7rem', padding: '2px 8px' }}
                            onClick={() => {
                              setHtmlEditorMode('email');
                              setHtmlEditorTemplateType(tType);
                              setHtmlEditorOpen(true);
                            }}
                          >
                            {t('create.templates.edit')} & Vorschau
                          </button>
                          {override && (
                            <button
                              className="btn btn-secondary"
                              style={{ fontSize: '0.7rem', padding: '2px 8px', color: 'var(--dex-red, #c00)' }}
                              onClick={() => {
                                const copy = { ...emailTemplateOverrides };
                                delete copy[tType];
                                setEmailTemplateOverrides(copy);
                              }}
                            >
                              {t('create.templates.reset')}
                            </button>
                          )}
                        </div>
                      </div>
                      <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                        {t('create.templates.subject')}: {currentSubject.replace(/\{\{EventTitle\}\}/g, title || '...')}
                      </div>
                      {/* Inline-Editor entfällt — Edit öffnet jetzt das HtmlEditorModal mit Live-Preview */}
                      {false && (
                        <div style={{ marginTop: 8 }}>
                          <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, minHeight: 150, padding: '0 4px' }}>
                            <RichText
                              value={currentBody}
                              onChange={(text: string) => {
                                setEmailTemplateOverrides({
                                  ...emailTemplateOverrides,
                                  [tType]: { subject: currentSubject, heading: currentHeading, bodyHtml: text },
                                });
                                return text;
                              }}
                            />
                          </div>
                          <div style={{ fontSize: '0.65rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>
                            {'{{Name}}'} → {t('create.templates.content') === 'Content' ? 'Participant name' : 'Teilnehmername'} · {'{{EventTitle}}'} → {title || '...'} · {'{{Organizer}}'} → {organizer || '...'} · {'{{WaitlistPosition}}'} → #1, #2, ...
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
                  </div>
                    </>)}
                    {advTab === 'mailLogo' && (<>
                  <div style={{ marginTop: 12 }}>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 8 }}>
                    {t('create.eventlogo.mail.hint')}
                  </p>
                  <label style={{
                    display: 'inline-flex', alignItems: 'center', gap: 8,
                    padding: '8px 16px', borderRadius: 'var(--dex-radius)',
                    border: '2px dashed var(--dex-gray-300)', cursor: 'pointer',
                    fontSize: '0.85rem', color: 'var(--dex-gray-600)',
                    transition: 'border-color 0.2s, background 0.2s',
                  }}>
                    <Plus size={16} />
                    {t('create.eventlogo.select')}
                    <input type="file" accept="image/*" style={{ display: 'none' }} onChange={async (e) => {
                      const file = e.target.files?.[0];
                      if (!file) return;
                      // v9.17: Hinweis vor Upload — Stockfotos / komplexe Bilder
                      // funktionieren nicht zuverlässig in Mails (siehe
                      // EmailImageBase64-Pipeline). Empfehlung sind die
                      // offiziellen Deloitte Circular Motifs.
                      const ok = await confirmDialog(t('create.logoupload.warning'), { confirmLabel: isDe ? 'Trotzdem verwenden' : 'Use anyway' });
                      if (!ok) { e.target.value = ''; return; }
                      const compressed = await compressImage(file, 600, 0.9);
                      const reader = new FileReader();
                      reader.onload = (ev) => { setEmailLogoPreview(ev.target?.result as string || ''); setEmailLogoFromPhoto(false); };
                      reader.readAsDataURL(compressed);
                    }} />
                  </label>
                  {/* v26.95: Event-Foto (falls hinterlegt) mit einem Klick als
                      Mail-Kopfbild übernehmen — kein Extra-Upload nötig. */}
                  {/* v28.29: neutral statt gruen gefüllt. Der gruene Rahmen plus
                      die gruene Füllung lasen sich wie ein AKTIVER Zustand
                      („Event-Foto ist schon übernommen") — war es aber nicht,
                      und genau deshalb blieb der Kopf beim Standardlogo. */}
                  {(imagePreview || imageFile) && (
                    <button
                      type="button"
                      onClick={() => { void (async () => { const b = await applyEventPhotoToLogo(setEmailLogoPreview); if (b) setEmailLogoFromPhoto(true); await offerLogoToSubEvents('email', b); })(); }}
                      style={{
                        marginLeft: 8, display: 'inline-flex', alignItems: 'center', gap: 6, padding: '8px 14px',
                        borderRadius: 'var(--dex-radius)', fontSize: '0.82rem', fontWeight: 600, cursor: 'pointer',
                        border: emailLogoFromPhoto ? '1.5px solid var(--dex-green, #86bc25)' : '1px solid var(--dex-gray-300)',
                        background: emailLogoFromPhoto ? 'rgba(134,188,37,0.12)' : '#fff',
                        color: emailLogoFromPhoto ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-700)',
                      }}
                    >
                      {emailLogoFromPhoto
                        ? <><Check size={14} /> {isDe ? 'Event-Foto \u00fcbernommen' : 'Event photo applied'}</>
                        : <><Icon iconName="Photo2" style={{ fontSize: 14 }} /> {isDe ? 'Event-Foto \u00fcbernehmen' : 'Copy event photo here'}</>}
                    </button>
                  )}
                  {/* v28.30: Zuschneiden/Entfernen sitzen jetzt hier statt in einer
                      eigenen Zeile mit zweitem Vorschaubild — das Bild stand dadurch
                      doppelt auf dem Schirm (Thumbnail oben, Größen-Vorschau unten). */}
                  {emailLogoPreview && (
                    <>
                      <button type="button" className="btn btn-secondary" style={{ marginLeft: 8, fontSize: '0.78rem', padding: '7px 12px' }}
                        onClick={() => setLogoCropTarget('email')}>{isDe ? 'Zuschneiden' : 'Crop'}</button>
                      <button type="button" className="btn btn-secondary" style={{ marginLeft: 8, fontSize: '0.78rem', padding: '7px 12px', color: 'var(--dex-red, #c00)' }}
                        onClick={() => { setEmailLogoPreview(''); setEmailLogoFromPhoto(false); }}>{t('create.eventlogo.remove')}</button>
                    </>
                  )}
                  {/* v28.29: zeigt das TATSÄCHLICH verwendete Kopfbild (eigenes /
                      vom Hauptevent geerbt / Standardlogo) statt blind das Event-Foto. */}
                  {((): React.ReactNode => {
                    const eff = effectiveHeaderImage('email', emailLogoPreview);
                    return renderHeaderSizeControl(eff.src, eff.note);
                  })()}
                  </div>
                    </>)}
                    {advTab === 'outlook' && (<>
                      <div style={{ fontWeight: 600, fontSize: '0.86rem' }}>{t('create.outlooklogo')}</div>
                  <div style={{ marginTop: 12 }}>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 8 }}>
                    {t('create.outlooklogo.hint')}
                  </p>
                  <label style={{
                    display: 'inline-flex', alignItems: 'center', gap: 8,
                    padding: '8px 16px', borderRadius: 'var(--dex-radius)',
                    border: '2px dashed var(--dex-gray-300)', cursor: 'pointer',
                    fontSize: '0.85rem', color: 'var(--dex-gray-600)',
                    transition: 'border-color 0.2s, background 0.2s',
                  }}>
                    <Plus size={16} />
                    {t('create.eventlogo.select')}
                    <input type="file" accept="image/*" style={{ display: 'none' }} onChange={async (e) => {
                      const file = e.target.files?.[0];
                      if (!file) return;
                      const ok = await confirmDialog(t('create.logoupload.warning'), { confirmLabel: isDe ? 'Trotzdem verwenden' : 'Use anyway' });
                      if (!ok) { e.target.value = ''; return; }
                      const compressed = await compressImage(file, 600, 0.9);
                      const reader = new FileReader();
                      reader.onload = (ev) => { setOutlookLogoPreview(ev.target?.result as string || ''); setOutlookLogoFromPhoto(false); };
                      reader.readAsDataURL(compressed);
                    }} />
                  </label>
                  {/* v26.95: Event-Foto mit einem Klick als Outlook-Kopfbild. */}
                  {/* v28.29: neutral statt gruen gefüllt. Der gruene Rahmen plus
                      die gruene Füllung lasen sich wie ein AKTIVER Zustand
                      („Event-Foto ist schon übernommen") — war es aber nicht,
                      und genau deshalb blieb der Kopf beim Standardlogo. */}
                  {(imagePreview || imageFile) && (
                    <button
                      type="button"
                      onClick={() => { void (async () => { const b = await applyEventPhotoToLogo(setOutlookLogoPreview); if (b) setOutlookLogoFromPhoto(true); await offerLogoToSubEvents('outlook', b); })(); }}
                      style={{
                        marginLeft: 8, display: 'inline-flex', alignItems: 'center', gap: 6, padding: '8px 14px',
                        borderRadius: 'var(--dex-radius)', fontSize: '0.82rem', fontWeight: 600, cursor: 'pointer',
                        border: outlookLogoFromPhoto ? '1.5px solid var(--dex-green, #86bc25)' : '1px solid var(--dex-gray-300)',
                        background: outlookLogoFromPhoto ? 'rgba(134,188,37,0.12)' : '#fff',
                        color: outlookLogoFromPhoto ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-700)',
                      }}
                    >
                      {outlookLogoFromPhoto
                        ? <><Check size={14} /> {isDe ? 'Event-Foto \u00fcbernommen' : 'Event photo applied'}</>
                        : <><Icon iconName="Photo2" style={{ fontSize: 14 }} /> {isDe ? 'Event-Foto \u00fcbernehmen' : 'Copy event photo here'}</>}
                    </button>
                  )}
                  {/* v28.30: Zuschneiden/Entfernen sitzen jetzt hier statt in einer
                      eigenen Zeile mit zweitem Vorschaubild — das Bild stand dadurch
                      doppelt auf dem Schirm (Thumbnail oben, Größen-Vorschau unten). */}
                  {outlookLogoPreview && (
                    <>
                      <button type="button" className="btn btn-secondary" style={{ marginLeft: 8, fontSize: '0.78rem', padding: '7px 12px' }}
                        onClick={() => setLogoCropTarget('outlook')}>{isDe ? 'Zuschneiden' : 'Crop'}</button>
                      <button type="button" className="btn btn-secondary" style={{ marginLeft: 8, fontSize: '0.78rem', padding: '7px 12px', color: 'var(--dex-red, #c00)' }}
                        onClick={() => { setOutlookLogoPreview(''); setOutlookLogoFromPhoto(false); }}>{t('create.eventlogo.remove')}</button>
                    </>
                  )}
                  {/* v27.2: Größensteuerung + Vorschau auch hier in Schritt 24. */}
                  {/* v28.29: siehe Schritt 23 — echte statt geratener Vorschau. */}
                  {((): React.ReactNode => {
                    const eff = effectiveHeaderImage('outlook', outlookLogoPreview);
                    return renderHeaderSizeControl(eff.src, eff.note);
                  })()}
                  {renderOutlookUpdateButton()}
                  </div>
                      <div style={{ fontWeight: 600, fontSize: '0.86rem', marginTop: 18, paddingTop: 14, borderTop: '1px solid var(--dex-gray-200)' }}>{t('create.outlookdesc')}</div>
                  <div style={{ marginTop: 12 }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                    <button
                      type="button"
                      className="btn btn-secondary"
                      onClick={() => { setHtmlEditorMode('outlook'); setHtmlEditorOpen(true); }}
                      style={{ fontSize: '0.85rem' }}
                    >
                      {t('create.outlookdesc.edit')}
                    </button>
                    <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)' }}>
                      {outlookBody
                        ? `${outlookBody.replace(/<[^>]+>/g, '').substring(0, 80)}${outlookBody.length > 80 ? '…' : ''}`
                        : t('create.outlookdesc.placeholder')}
                    </span>
                  </div>
                  {renderOutlookUpdateButton()}
                  </div>
                    </>)}
                    {advTab === 'fine' && (<>
                      <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', margin: '0 0 12px' }}>{isDe ? 'Seltene Schalter — die meisten Events lassen sie auf Standard.' : 'Rare switches — most events leave them at default.'}</p>
                      {disableEmails && disableOutlook && (<p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{isDe ? 'Kommunikation ist oben komplett abgeschaltet — hier gibt es nichts feinzustellen.' : 'Communication is switched off above — nothing to fine-tune here.'}</p>)}
                      {!disableEmails && (<div style={{ background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: '10px 14px', marginBottom: 10 }}>
                        <div style={{ fontSize: '0.88rem', fontWeight: 600, marginBottom: 4 }}>{t('create.notifications.email')}</div>
                  {!disableEmails && (
                    <div style={{ marginLeft: 26, paddingLeft: 12, borderLeft: '3px solid var(--dex-green, #86bc25)' }}>
                      {commToggleRow({
                        checked: !disableRegistrationEmail,
                        onChange: v => setDisableRegistrationEmail(!v),
                        label: isDe ? 'Anmelde-Bestätigung' : 'Registration confirmation',
                        short: isDe ? 'Mail bei der Anmeldung (inkl. Wartelisten-Mail).' : 'Email on registration (incl. waitlist email).',
                        info: isDe
                          ? 'Wenn aktiv: Teilnehmer bekommen bei der Anmeldung eine Bestätigungs-Mail (und, falls Warteliste aktiv, die Warteliste-Mail). Haken aus = es geht keine Anmelde-Bestätigung raus — die Abmelde-Mail bleibt davon unberührt.'
                          : 'When active: attendees receive a confirmation email on registration (plus the waitlist email if a waitlist is active). Unchecked = no registration confirmation is sent — the cancellation email is unaffected.',
                      })}
                      {commToggleRow({
                        checked: !disableCancellationEmail,
                        onChange: v => setDisableCancellationEmail(!v),
                        label: isDe ? 'Abmelde-Bestätigung' : 'Cancellation confirmation',
                        short: isDe ? 'Mail bei der Abmeldung.' : 'Email on cancellation.',
                        info: isDe
                          ? 'Wenn aktiv: Teilnehmer bekommen bei einer Abmeldung eine Bestätigungs-Mail. Haken aus = es geht keine Abmelde-Bestätigung raus (z.B. wenn du Teilnehmer still abmeldest) — die Anmelde-Mail bleibt davon unberührt.'
                          : 'When active: attendees receive a confirmation email when cancelled. Unchecked = no cancellation confirmation is sent (e.g. when you remove attendees silently) — the registration email is unaffected.',
                      })}
                    </div>
                  )}
                      </div>)}
                      {!disableOutlook && (<div style={{ background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: '10px 14px', marginBottom: 10 }}>
                        <div style={{ fontSize: '0.88rem', fontWeight: 600, marginBottom: 4 }}>{t('create.notifications.outlook')}</div>
                  {!disableOutlook && (
                    <div style={{ marginLeft: 26, paddingLeft: 12, borderLeft: '3px solid var(--dex-orange, #ed8b00)' }}>
                      {commToggleRow({
                        checked: autoDeregisterOnDecline,
                        onChange: v => setAutoDeregisterOnDecline(v),
                        label: isDe ? 'Outlook-Absage = Abmeldung' : 'Outlook decline = deregistration',
                        short: isDe
                          ? 'Termin abgesagt → Platz wird frei, Warteliste rückt nach.'
                          : 'Invite declined → the spot is freed, the waitlist moves up.',
                        accent: 'var(--dex-orange, #ed8b00)',
                        info: isDe
                          ? 'Wenn aktiv: Sagt ein Teilnehmer den Outlook-Termin ab, wird er automatisch auch vom Event abgemeldet — der Platz wird frei und die Warteliste rückt nach. Ohne diesen Haken bekommt die Person bei einer Outlook-Absage nur eine Erinnerung, sich bei Bedarf selbst abzumelden. Hinweis: Diese Automatik greift erst, sobald die einmalige Anpassung im Outlook-Absage-Verarbeitungsschritt im Tenant eingerichtet ist.'
                          : 'When active: if an attendee declines the Outlook invite, they are automatically deregistered from the event — the spot is freed and the waitlist moves up. Without this, a decline only triggers a reminder asking the person to deregister themselves if needed. Note: this automation only takes effect once the one-time change in the Outlook-decline processing step is set up in the tenant.',
                      })}
                    </div>
                  )}
                      </div>)}
                  {/* inactiveHandling: Verhalten, wenn eine angemeldete Person
                      nicht mehr bei Deloitte arbeitet. 'notify' = Organizer per
                      Mail informieren (Standard), 'autoderegister' = automatisch
                      abmelden (beim Öffnen der App durch einen Organizer). */}
                  <div style={{ background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: '10px 14px' }}>
                    <div style={{ fontSize: '0.88rem', fontWeight: 600, marginBottom: 6 }}>
                      {isDe ? 'Person arbeitet nicht mehr bei Deloitte' : 'Person no longer works at Deloitte'}
                      <InfoTooltip text={isDe
                        ? 'Die App erkennt beim Öffnen durch einen Organizer, wenn das Deloitte-Konto einer angemeldeten Person nicht mehr aktiv ist. „Organizer informieren" schickt dann eine Hinweis-Mail; „Automatisch abmelden" entfernt die Person direkt aus der Teilnehmerliste (Platz wird frei, Warteliste rückt nach).'
                        : 'When an organizer opens the app, it detects registered people whose Deloitte account is no longer active. „Notify organizer" sends an info email; „Auto-deregister" removes the person from the attendee list right away (the spot is freed, the waitlist moves up).'} />
                    </div>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', marginBottom: 6 }}>
                      <input
                        type="radio"
                        name="inactiveHandling"
                        value="notify"
                        checked={inactiveHandling === 'notify'}
                        onChange={() => setInactiveHandling('notify')}
                        style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 1, flexShrink: 0 }}
                      />
                      <span style={{ fontSize: '0.85rem' }}>
                        {isDe ? 'Organizer per E-Mail informieren (Standard)' : 'Notify the organizer by email (default)'}
                      </span>
                    </label>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                      <input
                        type="radio"
                        name="inactiveHandling"
                        value="autoderegister"
                        checked={inactiveHandling === 'autoderegister'}
                        onChange={() => setInactiveHandling('autoderegister')}
                        style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 1, flexShrink: 0 }}
                      />
                      <span style={{ fontSize: '0.85rem' }}>
                        {isDe ? 'Automatisch abmelden (beim Öffnen der App durch einen Organizer)' : 'Auto-deregister (when an organizer opens the app)'}
                      </span>
                    </label>
                  </div>
                    </>)}
                    </div>
                  </div>
                )}
                </>
                )}{/* end !(subEventsOnlyMode && tab===0) wrapper, v14.8 */}

              </div>
  );
};
