/* CommunicationStep — aus EventCreationPage.tsx ausgelagert (Zeilen 16971-17861 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; einzige
 * Aenderung ist die Anzeige-Bedingung: aus `currentStep === 5` wurde das Prop `visible`.
 * `visible` schaltet display:none statt unmount — Eingaben ueberleben den
 * Schrittwechsel genauso wie vorher.
 *
 * v31.2: Optik und Reihenfolge nach docs/ui-leitfaden.md — Kanal vor Sprache,
 * Einführung direkt unter dem Kopf, Geltung/Bündelung als eigene Abschnitte,
 * Vorschau + Chips + Aufklapper in einem Abschnitt, dex-ui-Klassen statt
 * Inline-Kästen (Hover). Bindungen, Setter und Bedingungen sind unverändert;
 * der tote Inline-RichText-Editor (`{false && …}`) ist raus. */
import * as React from 'react';
import { InfoTooltip } from '../../InfoTooltip';
import { Icon } from '@fluentui/react/lib/Icon';
import { StepBadge } from '../../wizard/StepBadge';
import WizardHint from '../../WizardHint';
import DatePicker from 'react-datepicker';
import { AlertCircle, Calendar, Check, ChevronDown, ChevronUp, Info, Mail, Pencil, Plus, Send, X } from '../../Icons';
import { cx } from '../../dexUi';
import { compressImage } from '../../../utils/imageCompress';
import { COMM_TOPICS } from '../logic/commTabs';
import { BundledComm } from '../../../utils/bundledComm';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { AgendaItem } from '../../../types';
import { EmailOverrideEntry } from '../../wizard/emailOverrideEntry';
import { CommPreviewCard } from '../../wizard/CommPreviewCard';
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
  /** v30.95: Programmpunkte des Hauptevents — für {{Programm}} in der Vorschau-Karte. */
  agenda: AgendaItem[];
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
  /** v30.90: für die Vorschau-Karte (Stufe B) und „Testmail an mich" (Stufe C). */
  headerLayoutFor: (logoB64: string) => { imageWidth: number; imagePaddingV: number; imagePaddingH: number };
  location: string;
  startDate: string;
  endDate: string;
  contactEmail: string;
  editEventId: string;
}
export const CommunicationStep: React.FC<CommunicationStepProps> = (p) => {
  const { visible } = p;
  const { activeCommTabIdx, applyCommTopicToAllSubEvents, commShared, setCommShared, flushActiveCommTabToState, resolveTopLevelCommState, applyEventPhotoToLogo, autoDeregisterOnDecline, bundledComm, childTermPlural, confirmDialog, disableCancellationEmail, disableEmails, disableOutlook, disableRegistrationEmail, effectiveHeaderImage, emailLanguage, emailLogoFromPhoto, emailLogoPreview, emailTemplateOverrides, emailTemplates, imageFile, imagePreview, inactiveHandling, isDe, mainCommDisabledAck, notifyOrgCancelMode, notifyOrgRegisterFromDate, notifyOrgRegisterMode, offerLogoToSubEvents, organizer, outlookBody, outlookLogoFromPhoto, outlookLogoPreview, renderHeaderSizeControl, renderOutlookUpdateButton, renderStepIntro, setAutoDeregisterOnDecline, setBundledComm, setDisableCancellationEmail, setDisableEmails, setDisableOutlook, setDisableRegistrationEmail, setEmailLanguage, setEmailLogoFromPhoto, setEmailLogoPreview, setEmailTemplateOverrides, setHtmlEditorMode, setHtmlEditorOpen, setHtmlEditorTemplateType, setInactiveHandling, setLogoCropTarget, setMainCommDisabledAck, setNotifyOrgCancelMode, setNotifyOrgRegisterFromDate, setNotifyOrgRegisterMode, setOutlookLogoFromPhoto, setOutlookLogoPreview, subEvents, subEventsOnlyMode, t, title, unlimitedParticipants, waitlistEnabled } = p;

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
                  <span className="dex-step-eyebrow">{isDe ? 'Schritt 6 von 9' : 'Step 6 of 9'}</span>
                  {t('create.step.communication')}
                </h2>
                <p className="dex-step-head-lead">
                  {isDe
                    ? 'Was bekommen Teilnehmer nach der Anmeldung? Hier legst du Kanal und Sprache fest, siehst die Vorschau — und passt Texte, Bilder und Versandregeln nur an, wenn du willst.'
                    : 'What do attendees get after registering? Set channel and language here, check the preview — and customise texts, images and send rules only if you want to.'}
                </p>
                {/* v31.2: Die Einführung steht jetzt direkt unter dem Kopf — vorher
                    kam sie erst NACH dem Geltungs-Kasten und erklärte damit etwas,
                    das man schon hinter sich hatte. */}
                {renderStepIntro(
                  [
                    'Zwei Entscheidungen oben: Kanal (Mail und/oder Outlook-Termin) und Sprache der Mails',
                    'Darunter siehst du, was mit diesen Einstellungen tatsächlich rausgeht — Standard heißt: fertig, nichts weiter zu tun',
                    'Nur wer Texte, Bilder oder Feineinstellungen ändern will, öffnet „Texte und Bilder anpassen“ — dort liegen Mail-Vorlagen, Mail-Logo, Outlook-Termin, die Organizer-Kopie und die seltenen Schalter',
                  ],
                  [
                    'Two decisions at the top: channel (email and/or Outlook invite) and mail language',
                    'Below you see what actually goes out with these settings — default means: done, nothing else to do',
                    'Only if you want to change texts, images or fine-tuning, open “Customise texts and images” — that is where mail templates, mail logo, Outlook invite, the organizer copy and the rare switches live',
                  ]
                )}
                {/* v28.80: Kommunikation eines Sub-Events auf die anderen
                    uebertragen — sonst muss der Organizer Logo, Outlook-Text,
                    Betreff und Mail-Schalter bei jedem Sub-Event einzeln
                    einstellen. */}
                {/* v30.60: Beide Wege stehen an EINER Stelle und sagen ausdrücklich,
                    was der Normalfall ist — auf der Klammer sah es sonst aus, als
                    gäbe es die Möglichkeit gar nicht. */}
                {/* v30.71: Schalter mit Gedächtnis statt Kopier-Knopf.
                    Nutzer-Ansage 02.09.2026: „kein Button, sondern ein
                    Wechselschalter — entweder einzeln oder für alle Termine
                    zusammen." Gemeinsam = die Termin-Reiter sind hier nur
                    Anzeige, gespeichert wird überall der Stand des
                    Haupt-Events (persistSubEvents). Einzeln = Themenliste
                    mit Pill (eigen / wie Haupt-Event) und einem Knopf, der
                    genau EIN Thema auf alle Termine verteilt. */}
                {subEvents.length > 0 && (
                  <div className="dex-ui-section">
                    <div className="dex-ui-section-title">{isDe ? 'Geltung' : 'Scope'}</div>
                    <div className="dex-ui-label">
                      {isDe ? 'Gelten die Einstellungen für jeden Termin einzeln oder für alle gemeinsam?' : 'Do these settings apply per date or to all dates together?'}
                    </div>
                    <p className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
                      {isDe
                        ? 'Mails, Kalendereinträge und Texte — du entscheidest einmal, ob jeder Termin seine eigenen bekommt oder alle dieselben.'
                        : 'Emails, calendar entries and texts — decide once whether every date gets its own or all share the same.'}
                    </p>
                    <div className="dex-ui-grid-2" role="radiogroup" aria-label={isDe ? 'Modus' : 'Mode'}>
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
                            className={cx('dex-ui-choice', active && 'is-active')}
                            onClick={() => { void switchCommShared(opt.on); }}
                          >
                            <span className="dex-ui-choice-body">
                              <span className="dex-ui-choice-title">{isDe ? opt.de : opt.en}</span>
                              <span className="dex-ui-choice-desc">{isDe ? opt.subDe : opt.subEn}</span>
                            </span>
                            <span className="dex-ui-choice-check">{active && <Check size={12} />}</span>
                          </button>
                        );
                      })}
                    </div>
                    {commShared ? (
                      <div className="dex-ui-callout dex-ui-callout--success" style={{ marginTop: 12 }}>
                        <span className="dex-ui-callout-icon"><Check size={16} /></span>
                        <span>
                          {activeCommTabIdx === 0
                            ? (isDe
                              ? <><strong>Du stellst alles hier auf dem Reiter „{mainTabLabel}&ldquo; ein.</strong> Was du hier änderst, gilt beim Speichern für alle {namedSubCount} {childTermPlural || 'Termine'}. Die Termin-Reiter zeigen keine eigenen Felder mehr.</>
                              : <><strong>You set everything here on the “{mainTabLabel}” tab.</strong> On save it applies to all {namedSubCount} dates. The date tabs no longer show their own fields.</>)
                            : (isDe
                              ? <><strong>Dieser Termin übernimmt alles vom Reiter „{mainTabLabel}&ldquo;.</strong> Zum Ändern wechsle dorthin — oder stelle oben auf „{childOneDe} einzeln&ldquo;, wenn dieser Termin etwas Eigenes braucht.</>
                              : <><strong>This date takes everything from the “{mainTabLabel}” tab.</strong> Switch there to change it — or choose “Each date individually” above if this date needs something of its own.</>)}
                        </span>
                      </div>
                    ) : (
                      <>
                        <div className="dex-ui-callout dex-ui-callout--info" style={{ marginTop: 12 }}>
                          <span className="dex-ui-callout-icon"><Info size={16} /></span>
                          <span>
                            {activeCommTabIdx === 0
                              ? (isDe
                                ? <><strong>Du bearbeitest gerade: {mainTabLabel}.</strong> Jede Zeile unten ist ein Thema. Mit „für alle übernehmen&ldquo; verteilst du genau dieses Thema auf alle {namedSubCount} {childTermPlural || 'Termine'} — mehr nicht.</>
                                : <><strong>You are editing: {mainTabLabel}.</strong> Each row below is one topic. “Apply to all” copies exactly that topic to all {namedSubCount} dates — nothing else.</>)
                              : (isDe
                                ? <><strong>Du bearbeitest gerade: {currentTabLabel}.</strong> Jede Zeile zeigt, ob dieser Termin etwas Eigenes hat oder das Haupt-Event übernimmt. Mit „für alle übernehmen&ldquo; verteilst du genau dieses Thema auf die anderen {childTermPlural || 'Termine'}.</>
                                : <><strong>You are editing: {currentTabLabel}.</strong> Each row shows whether this date has something of its own or inherits from the main event. “Apply to all” copies exactly that topic to the other dates.</>)}
                          </span>
                        </div>
                        <div className="dex-ui-card" style={{ marginTop: 10, padding: '2px 6px' }}>
                          {COMM_TOPICS.map(topic => {
                            const own = activeCommTabIdx > 0 && topicDiffersFromParent(topic.key);
                            return (
                              <div key={topic.key} className="dex-ui-row dex-ui-row--bordered">
                                <StepBadge n={topic.step} />
                                <div className="dex-ui-row-main">
                                  <div className="dex-ui-row-title">{t(topic.labelKey)}</div>
                                </div>
                                {activeCommTabIdx > 0 && (
                                  <span className={cx('dex-ui-pill', own ? 'dex-ui-pill--orange' : 'dex-ui-pill--gray')}>
                                    {own ? (isDe ? 'eigen' : 'own') : (isDe ? 'wie Haupt-Event' : 'as main event')}
                                  </span>
                                )}
                                {namedSubCount > (activeCommTabIdx > 0 ? 1 : 0) && (
                                  <div className="dex-ui-row-actions">
                                    <button
                                      type="button"
                                      className="dex-ui-textbtn"
                                      onClick={() => { void applyCommTopicToAllSubEvents(topic.key); }}
                                      title={isDe ? `Nur „${t(topic.labelKey)}" von diesem Reiter auf alle ${childTermPlural || 'Termine'} übertragen` : `Copy only this topic from this tab to all dates`}
                                    >
                                      {isDe ? `für alle ${namedSubCount} übernehmen` : `apply to all ${namedSubCount}`}
                                    </button>
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      </>
                    )}
                  </div>
                )}
                {/* v30.71: Bündel-Block UNTER dem Schalter — oben geht es um den
                    Inhalt, hier um die Anzahl der Mails (Nutzer-Entscheidung 02.09.). */}
                {/* v30.61: Der eigentliche Schalter — eine Mail und ein
                    Kalendereintrag fürs ganze Event statt einem je Termin.
                    Das ist etwas anderes als „überall dieselben Texte":
                    Dort verschicken weiterhin zehn Termine zehn Mails, hier
                    verschickt die Klammer EINE. Nur im Klammer-Modus, weil
                    ein buchbares Haupt-Event ohnehin selbst verschickt. */}
                {subEvents.length > 0 && subEventsOnlyMode && (
                  <div className="dex-ui-section">
                    <div className="dex-ui-section-title">{isDe ? 'Bündelung' : 'Bundling'}</div>
                    <div className="dex-ui-label">
                      {isDe ? 'Wie viele Mails bekommt jemand, der mehrere Termine bucht?' : 'How many emails does someone get who books several dates?'}
                    </div>
                    <p className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
                      {isDe
                        ? `Wer sich für mehrere ${childTermPlural || 'Termine'} anmeldet, bekommt sonst für jeden eine eigene Bestätigung und einen eigenen Kalendereintrag. Gebündelt kommt stattdessen EINE Bestätigung mit der Liste aller gebuchten Termine — und ein Kalendereintrag über den Gesamtzeitraum.`
                        : `Someone registering for several dates otherwise receives one confirmation and one calendar entry per date. Bundled, they get ONE confirmation listing all booked dates.`}
                    </p>
                    <div className="dex-ui-stack">
                      {([
                        { key: 'mail' as const, de: 'Bestätigungs-Mails bündeln', en: 'Bundle confirmation emails', hintDe: 'Eine Mail mit der Liste aller gebuchten Termine.', hintEn: 'One email listing all booked dates.' },
                        { key: 'outlook' as const, de: 'Kalendereintrag bündeln', en: 'Bundle the calendar entry', hintDe: 'Ein Eintrag über den Zeitraum des Gesamt-Events. Wer nur Tag 2 und 4 bucht, bekommt trotzdem einen Eintrag über den ganzen Zeitraum — welche Tage gebucht sind, steht in der Beschreibung.', hintEn: 'One entry spanning the whole event period.' },
                        { key: 'qr' as const, de: 'Einen QR-Code fürs Gesamt-Event', en: 'One QR code for the whole event', hintDe: 'Der Check-in läuft dann über das Haupt-Event statt über die einzelnen Termine.', hintEn: 'Check-in then runs on the main event.' },
                      ]).map(opt => (
                        <label key={opt.key} className={cx('dex-ui-toggle-row', bundledComm[opt.key] && 'is-active')}>
                          <input
                            type="checkbox"
                            checked={bundledComm[opt.key]}
                            onChange={e => setBundledComm(prev => ({ ...prev, [opt.key]: e.target.checked }))}
                          />
                          <span className="dex-ui-toggle-row-body">
                            <span className="dex-ui-toggle-row-title">{isDe ? opt.de : opt.en}</span>
                            <span className="dex-ui-toggle-row-desc">{isDe ? opt.hintDe : opt.hintEn}</span>
                          </span>
                        </label>
                      ))}
                    </div>
                    {(bundledComm.mail || bundledComm.outlook) && (
                      <div className="dex-ui-callout dex-ui-callout--info" style={{ marginTop: 10 }}>
                        <span className="dex-ui-callout-icon"><Info size={16} /></span>
                        <span>
                          {isDe
                            ? `Die Texte für die gebündelte Mail stehen auf dem Reiter ${subEventsOnlyMode ? 'Klammer' : 'Haupt-Event'}. Was du auf den Termin-Reitern eingestellt hast, bleibt gespeichert, wirkt aber nicht mehr, solange gebündelt wird.`
                            : 'The copy for the bundled email lives on the bracket tab. Per-date settings stay saved but have no effect while bundling is on.'}
                        </span>
                      </div>
                    )}
                  </div>
                )}

                {/* v28.88: Die Reiter-Leiste stand hier ein ZWEITES Mal. Seit
                    v28.78 trägt die Scope-Karte über dem Formular
                    (renderGlobalScopeBar) den Umschalter für alle
                    scope-fähigen Schritte — Kommunikation eingeschlossen, sie
                    hängt über setScope am selben Index. Zwei identische
                    Reiter-Reihen auf einer Seite lesen sich als zwei
                    Navigationen: der Organizer sucht, welche die gültige ist.
                    Der Erklär-Tooltip bleibt — er sagt, was pro Sub-Event
                    überhaupt getrennt einstellbar ist. */}
                {subEvents.length > 0 && (
                  <div className="dex-ui-inline" style={{ marginTop: 22 }}>
                    <span className="dex-ui-muted" style={{ flex: 1, minWidth: 0 }}>
                      {isDe
                        ? 'Die Einstellungen unten gelten für den oben gewählten Reiter.'
                        : 'The settings below apply to the tab selected above.'}
                    </span>
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

                {/* v14.8: „Nur Sub-Events"-Modus + auf Haupt-Event-Tab → Hinweis
                    statt Kommunikations-Settings rendern. Der User soll keine
                    Werte für ein nicht-existentes Hauptevent-Anmelden pflegen. */}
                {subEventsOnlyMode && activeCommTabIdx === 0 && (
                  <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 12 }}>
                    <span className="dex-ui-callout-icon"><AlertCircle size={18} /></span>
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
                {/* v30.89: Kanal-Karte ersetzt die zwei Master-Haken aus dem alten
                    Kasten 26 — eine Frage, vier Antworten, die Folge steht direkt darunter. */}
                {/* v31.2: Kanal VOR Sprache — wer „Keine Kommunikation" wählt, braucht
                    keine Sprache mehr; die Sprach-Karte wird dann gedämpft. Die Badges
                    25/26 bleiben an ihren Fragen. */}
                <div className="dex-ui-section">
                  <div className="dex-ui-section-title">{isDe ? 'Kanal' : 'Channel'}</div>
                  <div className="dex-ui-label">
                    <StepBadge n={26} />
                    {isDe ? 'Wie erreicht DEX die Teilnehmer?' : 'How does DEX reach attendees?'}
                    <InfoTooltip text={t('create.notifications.hint')} />
                  </div>
                  <div className="dex-ui-grid-2" role="radiogroup" aria-label={isDe ? 'Weg der Kommunikation' : 'Communication channel'}>
                    {([
                      { key: 'both', mails: true, outlook: true, icon: <Send size={18} />, de: 'Mail + Outlook-Termin', en: 'Email + Outlook invite', subDe: 'Standard — Bestätigung per Mail und Termin im Kalender', subEn: 'Default — confirmation by email and calendar entry' },
                      { key: 'mail', mails: true, outlook: false, icon: <Mail size={18} />, de: 'Nur Mail', en: 'Email only', subDe: 'Kein Kalendereintrag — Teilnehmer planen selbst', subEn: 'No calendar entry — attendees schedule themselves' },
                      { key: 'outlook', mails: false, outlook: true, icon: <Calendar size={18} strokeWidth={2} />, de: 'Nur Outlook-Termin', en: 'Outlook invite only', subDe: 'Keine Mails — der Termin ist die Bestätigung', subEn: 'No emails — the invite is the confirmation' },
                      { key: 'none', mails: false, outlook: false, icon: <X size={18} />, de: 'Keine Kommunikation', en: 'No communication', subDe: 'Nichts geht raus — z.B. zum internen Testen', subEn: 'Nothing goes out — e.g. for internal testing' },
                    ]).map(opt => {
                      const selected = !disableEmails === opt.mails && !disableOutlook === opt.outlook;
                      return (
                        <button
                          key={opt.key}
                          type="button"
                          role="radio"
                          aria-checked={selected}
                          className={cx('dex-ui-choice', selected && 'is-active')}
                          onClick={() => { setDisableEmails(!opt.mails); setDisableOutlook(!opt.outlook); }}
                        >
                          <span className="dex-ui-choice-icon">{opt.icon}</span>
                          <span className="dex-ui-choice-body">
                            <span className="dex-ui-choice-title">{isDe ? opt.de : opt.en}</span>
                            <span className="dex-ui-choice-desc">{isDe ? opt.subDe : opt.subEn}</span>
                          </span>
                          <span className="dex-ui-choice-check">{selected && <Check size={12} />}</span>
                        </button>
                      );
                    })}
                  </div>
                  {(disableEmails || disableOutlook) && (
                    <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 10 }}>
                      <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                      <span>
                        {disableEmails && disableOutlook
                          ? (isDe ? 'Achtung: Niemand erfährt von seiner Anmeldung — weder per Mail noch im Kalender. Nur sinnvoll, wenn du selbst einlädst.' : 'Careful: nobody hears about their registration — neither by email nor in the calendar. Only sensible if you invite people yourself.')
                          : disableEmails
                            ? (isDe ? 'Es geht keine einzige Mail an Teilnehmer raus — auch keine Wartelisten- oder Abmelde-Mail.' : 'Not a single email goes out to attendees — no waitlist or cancellation email either.')
                            : (isDe ? 'Kein Termin im Kalender — bei einer Abmeldung gibt es entsprechend auch nichts zu entfernen.' : 'No calendar entry — accordingly nothing is removed on cancellation.')}
                      </span>
                    </div>
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
                      <label className={cx('dex-ui-toggle-row', mainCommDisabledAck && 'is-active')}>
                        <input
                          type="checkbox"
                          checked={mainCommDisabledAck}
                          onChange={e => setMainCommDisabledAck(e.target.checked)}
                        />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">
                            {isDe
                              ? 'Ja, mir ist bewusst, dass Teilnehmer sich für mindestens ein Sub-Event anmelden müssen, um Kommunikation zu erhalten.'
                              : 'Yes, I understand attendees need to register for at least one sub-event to receive communication.'}
                          </span>
                        </span>
                      </label>
                    </WizardHint>
                  )}
                </div>
                <div className="dex-ui-section">
                  <div className="dex-ui-section-title">{isDe ? 'Sprache' : 'Language'}</div>
                  <div className={cx('dex-ui-card', disableEmails && disableOutlook && 'dex-ui-card--muted')}>
                    <div className="dex-ui-label">
                      <StepBadge n={25} />
                      {isDe ? 'In welcher Sprache gehen die automatischen Mails raus?' : 'Which language should the automated emails be in?'}
                      <InfoTooltip text={t('create.emaillanguage.hint')} />
                    </div>
                    <div className="dex-ui-inline" role="radiogroup" aria-label={t('create.emaillanguage')}>
                      {(['DE', 'EN'] as const).map(lang => (
                        <button
                          key={lang}
                          type="button"
                          role="radio"
                          aria-checked={emailLanguage === lang}
                          className={cx('dex-ui-chip', emailLanguage === lang && 'is-active')}
                          style={{ padding: '8px 16px', fontSize: '0.84rem' }}
                          onClick={() => setEmailLanguage(lang)}
                        >
                          {lang === 'DE' ? 'DE – Deutsch' : 'EN – English'}
                        </button>
                      ))}
                    </div>
                    <p className="dex-ui-help">
                      {isDe
                        ? 'Gilt für Anmelde-, Abmelde-, Wartelisten- und Nachrück-Mail. Die Texte kannst du unten je Vorlage anpassen.'
                        : 'Applies to the registration, cancellation, waitlist and next-in-line emails. You can adjust each template below.'}
                    </p>
                  </div>
                </div>
                {/* v30.90: Ebene 2 — die gerenderte Vorschau (Stufe B) plus
                    „Testmail an mich" (Stufe C). Überschrift/Unterzeile des
                    Outlook-Termins kommen vom offenen Reiter: Top-Level aus dem
                    aufgelösten State, Termin aus seinem Slot. */}
                {/* v30.89: Ebene 2 — was mit diesen Einstellungen rausgeht, als Chips.
                    Ersetzt die achtzeilige Übersichtsbox (v19.23) am Kopf des Schritts:
                    Jeder Chip nennt ein Thema und seinen Stand; ein Klick öffnet den
                    passenden Reiter unten. Das ist die Antwort auf „gilt der Standard,
                    wenn ich nichts aufgeklappt habe?“ — ja, und hier steht welcher. */}
                {/* v31.2: Vorschau, Chips und Aufklapper in EINEM Abschnitt — sie
                    beantworten dieselbe Frage („was geht raus?“). Der Aufklapper
                    ersetzt den Umschalt-Knopf; seine Zählung sagt, wie viele Themen
                    vom Standard abweichen. */}
                {(() => {
                  const top = resolveTopLevelCommState();
                  const slot: Partial<SubEventDraft> = activeCommTabIdx > 0 ? (subEvents[activeCommTabIdx - 1] || {}) : {};
                  const olHeading = activeCommTabIdx > 0 ? (slot.outlookHeading || '') : (top.outlookHeading || '');
                  const olSub = activeCommTabIdx > 0 ? (slot.outlookSubheading || '') : (top.outlookSubheading || '');
                  const tabTitle = activeCommTabIdx > 0 ? ((slot.title || '').trim() || title) : title;
                  const tplCount = Object.keys(emailTemplateOverrides || {}).filter(k => k.charAt(0) !== '_').length;
                  const fineChanged = disableRegistrationEmail || disableCancellationEmail || autoDeregisterOnDecline || inactiveHandling === 'autoderegister';
                  const std = isDe ? 'Standard' : 'default';
                  const chips: Array<{ tab: 'templates' | 'mailLogo' | 'outlook' | 'fine'; label: string; value: string; changed: boolean }> = [
                    { tab: 'templates', label: isDe ? 'Mail-Texte' : 'Mail texts', value: tplCount > 0 ? (isDe ? `${tplCount} angepasst` : `${tplCount} customised`) : std, changed: tplCount > 0 },
                    { tab: 'mailLogo', label: isDe ? 'Mail-Logo' : 'Mail logo', value: emailLogoPreview ? (emailLogoFromPhoto ? (isDe ? 'Event-Foto' : 'event photo') : (isDe ? 'eigenes Bild' : 'own image')) : std, changed: !!emailLogoPreview },
                    { tab: 'outlook', label: isDe ? 'Outlook-Termin' : 'Outlook invite', value: (outlookBody || outlookLogoPreview) ? (isDe ? 'angepasst' : 'customised') : std, changed: !!(outlookBody || outlookLogoPreview) },
                    { tab: 'fine', label: isDe ? 'Feineinstellungen' : 'Fine-tuning', value: fineChanged ? (isDe ? 'angepasst' : 'customised') : std, changed: fineChanged },
                  ];
                  const changedCount = chips.filter(c => c.changed).length;
                  return (
                    <div className="dex-ui-section">
                      <div className="dex-ui-section-title">{isDe ? 'So geht es raus' : 'What goes out'}</div>
                      <CommPreviewCard
                        isDe={isDe}
                        emailLanguage={emailLanguage}
                        emailTemplates={emailTemplates}
                        emailTemplateOverrides={emailTemplateOverrides}
                        emailLogoPreview={emailLogoPreview}
                        outlookLogoPreview={outlookLogoPreview}
                        effectiveHeaderImage={effectiveHeaderImage}
                        headerLayoutFor={p.headerLayoutFor}
                        title={tabTitle}
                        location={p.location}
                        startDate={p.startDate}
                        endDate={p.endDate}
                        organizer={organizer}
                        contactEmail={p.contactEmail}
                        outlookBody={outlookBody}
                        outlookHeading={olHeading}
                        outlookSubheading={olSub}
                        disableEmails={disableEmails}
                        disableOutlook={disableOutlook}
                        eventId={p.editEventId}
                        agenda={activeCommTabIdx > 0 ? (slot.agenda || []) : p.agenda}
                      />
                      <div className="dex-ui-inline" style={{ marginTop: 14 }}>
                        <span className="dex-ui-muted" style={{ fontWeight: 600 }}>{isDe ? 'Stand:' : 'Status:'}</span>
                        {chips.map(c => (
                          <button key={c.tab} type="button" className="dex-ui-chip" onClick={() => { setAdvTab(c.tab); setAdvOpen(true); }}
                            title={isDe ? 'Anpassen' : 'Customise'}>
                            {c.label}: <span style={{ fontWeight: 500, color: c.changed ? 'var(--dex-orange-dark, #b35a00)' : undefined }}>{c.value}</span>
                          </button>
                        ))}
                      </div>
                      <button type="button" className={cx('dex-ui-disclosure', advOpen && 'is-open')} aria-expanded={advOpen} onClick={() => setAdvOpen(v => !v)} style={{ marginTop: 10 }}>
                        <span className="dex-ui-disclosure-chevron" style={{ transform: 'none' }}>{advOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}</span>
                        {advOpen ? (isDe ? 'Anpassen schließen' : 'Close customisation') : (isDe ? 'Texte und Bilder anpassen' : 'Customise texts and images')}
                        <span className="dex-ui-disclosure-count">{changedCount > 0 ? (isDe ? `${changedCount} angepasst` : `${changedCount} customised`) : (isDe ? 'alles Standard' : 'all default')}</span>
                      </button>
                    </div>
                  );
                })()}
                {/* v30.89: Ebene 3 — vier Reiter statt fünf Aufklapper untereinander.
                    Wenige, lange Inhalte → Reiter (NN/G); die alten Kästen 26 (Rest), 28,
                    29+30 und 31 liegen hier, nur ohne <details>/<summary>. */}
                {/* v31.2: Segment-Reiter (dex-ui-tabs) mit denselben kurzen Namen wie
                    die Chips darüber — Chip und Reiter heißen gleich, sonst sucht man. */}
                {advOpen && (
                  <div className="dex-ui-card dex-ui-card--soft dex-ui-fade-in" style={{ marginTop: 10 }}>
                    <div className="dex-ui-tabs" role="tablist" style={{ marginBottom: 16 }}>
                      {([
                        { key: 'templates' as const, n: 31, label: `${isDe ? 'Mail-Texte' : 'Mail texts'} (${emailLanguage})` },
                        { key: 'mailLogo' as const, n: 28, label: isDe ? 'Mail-Logo' : 'Mail logo' },
                        { key: 'outlook' as const, n: 29, label: isDe ? 'Outlook-Termin' : 'Outlook invite' },
                        { key: 'fine' as const, n: 26, label: isDe ? 'Feineinstellungen' : 'Fine-tuning' },
                      ]).map(tab => {
                        const active = advTab === tab.key;
                        return (
                          <button key={tab.key} type="button" role="tab" aria-selected={active} className={cx('dex-ui-tab', active && 'is-active')} style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={() => setAdvTab(tab.key)}>
                            <StepBadge n={tab.n} /> {tab.label}
                          </button>
                        );
                      })}
                    </div>
                    <div>
                    {advTab === 'templates' && (<>
                      <div className="dex-ui-label">
                        <StepBadge n={31} />
                        {t('create.templates.title')}
                        <InfoTooltip text={t('create.templates.hint')} />
                      </div>
                      <p className="dex-ui-help" style={{ margin: '0 0 12px' }}>{isDe ? 'Standard aktiv, solange kein Text angepasst ist. Jede Vorlage lässt sich einzeln zurücksetzen; Änderungen gelten nur für dieses Event.' : 'Default applies while no text is customised. Each template can be reset individually; changes apply to this event only.'}</p>
                      {/* TemplateType in DEX_EmailTemplates ist ASCII 'Nachruecken' (Umlaut nicht erlaubt in Choice-Feld).
                          v9.17: Warteliste/Nachrücken-Templates nur anzeigen, wenn das Event eine
                          Warteliste hat — sonst werden sie ohnehin nie genutzt. */}
                      <div className="dex-ui-card" style={{ padding: '2px 6px' }}>
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
                            // Bearbeiten öffnet das HtmlEditorModal mit Live-Preview (kein Inline-Editor).
                            return (
                              <div key={tType} className="dex-ui-row dex-ui-row--bordered">
                                <div className="dex-ui-row-main">
                                  <div className="dex-ui-row-title" style={{ display: 'flex', alignItems: 'center', gap: 8, whiteSpace: 'normal' }}>
                                    {t(`create.tpl.${tType}`)}
                                    {override && <span className="dex-ui-pill dex-ui-pill--green">{t('create.templates.modified')}</span>}
                                  </div>
                                  <div className="dex-ui-row-sub">
                                    {t('create.templates.subject')}: {currentSubject.replace(/\{\{EventTitle\}\}/g, title || '...')}
                                  </div>
                                </div>
                                <div className="dex-ui-row-actions">
                                  {override && (
                                    <button
                                      type="button"
                                      className="dex-ui-textbtn dex-ui-textbtn--danger"
                                      onClick={() => {
                                        const copy = { ...emailTemplateOverrides };
                                        delete copy[tType];
                                        setEmailTemplateOverrides(copy);
                                      }}
                                    >
                                      {t('create.templates.reset')}
                                    </button>
                                  )}
                                  <button
                                    type="button"
                                    className="btn btn-secondary dex-ui-btn-sm"
                                    onClick={() => {
                                      setHtmlEditorMode('email');
                                      setHtmlEditorTemplateType(tType);
                                      setHtmlEditorOpen(true);
                                    }}
                                  >
                                    <Pencil size={14} /> {isDe ? `${t('create.templates.edit')} & Vorschau` : `${t('create.templates.edit')} & preview`}
                                  </button>
                                </div>
                              </div>
                            );
                          })}
                      </div>
                    </>)}
                    {advTab === 'mailLogo' && (<>
                      <div className="dex-ui-field">
                        <div className="dex-ui-label">
                          <StepBadge n={28} />
                          {t('create.eventlogo.mail')}
                          <InfoTooltip text={t('create.eventlogo.mail.hint')} />
                        </div>
                        <p className="dex-ui-help" style={{ margin: '0 0 10px' }}>
                          {isDe ? 'Erscheint oben in jeder automatischen Mail. Leer = neutrales Standard-Logo; einfache Grafiken wirken in Mails am saubersten.' : 'Appears at the top of every automated email. Empty = neutral default logo; simple graphics look cleanest in emails.'}
                        </p>
                        <div className="dex-ui-inline">
                          <label className="btn btn-secondary dex-ui-btn-sm" style={{ cursor: 'pointer' }}>
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
                          {/* v28.29: neutral, solange NICHT übernommen — der grüne Zustand
                              hieß sonst „schon drin", obwohl der Kopf beim Standardlogo blieb. */}
                          {(imagePreview || imageFile) && (
                            <button
                              type="button"
                              className={cx('dex-ui-chip', emailLogoFromPhoto && 'is-active')}
                              onClick={() => { void (async () => { const b = await applyEventPhotoToLogo(setEmailLogoPreview); if (b) setEmailLogoFromPhoto(true); await offerLogoToSubEvents('email', b); })(); }}
                            >
                              {emailLogoFromPhoto
                                ? <><Check size={14} /> {isDe ? 'Event-Foto übernommen' : 'Event photo applied'}</>
                                : <><Icon iconName="Photo2" style={{ fontSize: 14 }} /> {isDe ? 'Event-Foto übernehmen' : 'Copy event photo here'}</>}
                            </button>
                          )}
                          {/* v28.30: Zuschneiden/Entfernen direkt neben dem Upload — sonst
                              stand das Bild doppelt auf dem Schirm. */}
                          {emailLogoPreview && (
                            <>
                              <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => setLogoCropTarget('email')}>{isDe ? 'Zuschneiden' : 'Crop'}</button>
                              <button type="button" className="dex-ui-textbtn dex-ui-textbtn--danger" onClick={() => { setEmailLogoPreview(''); setEmailLogoFromPhoto(false); }}>{t('create.eventlogo.remove')}</button>
                            </>
                          )}
                        </div>
                        {/* v28.29: zeigt das TATSÄCHLICH verwendete Kopfbild (eigenes /
                            vom Hauptevent geerbt / Standardlogo) statt blind das Event-Foto. */}
                        {((): React.ReactNode => {
                          const eff = effectiveHeaderImage('email', emailLogoPreview);
                          return renderHeaderSizeControl(eff.src, eff.note);
                        })()}
                      </div>
                    </>)}
                    {advTab === 'outlook' && (<>
                      {/* v31.2: Bild und Text des Termins als zwei Felder in EINEM Reiter,
                          Badges 29/30 wie in COMM_TOPICS; der Outlook-Update-Knopf steht
                          einmal am Ende statt zweimal (er ist für beides derselbe). */}
                      <div className="dex-ui-field">
                        <div className="dex-ui-label">
                          <StepBadge n={29} />
                          {t('create.outlooklogo')}
                          <InfoTooltip text={t('create.outlooklogo.hint')} />
                        </div>
                        <p className="dex-ui-help" style={{ margin: '0 0 10px' }}>
                          {isDe ? 'Steht oben im Kalendereintrag der Teilnehmer — z.B. ein Foto vom Ort. Leer = Standard-Logo; das Deloitte-Branding bleibt immer.' : 'Shown at the top of the attendees’ calendar entry — e.g. a venue photo. Empty = default logo; the Deloitte branding always stays.'}
                        </p>
                        <div className="dex-ui-inline">
                          <label className="btn btn-secondary dex-ui-btn-sm" style={{ cursor: 'pointer' }}>
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
                          {/* v28.29: neutral, solange NICHT übernommen — der grüne Zustand
                              hieß sonst „schon drin", obwohl der Kopf beim Standardlogo blieb. */}
                          {(imagePreview || imageFile) && (
                            <button
                              type="button"
                              className={cx('dex-ui-chip', outlookLogoFromPhoto && 'is-active')}
                              onClick={() => { void (async () => { const b = await applyEventPhotoToLogo(setOutlookLogoPreview); if (b) setOutlookLogoFromPhoto(true); await offerLogoToSubEvents('outlook', b); })(); }}
                            >
                              {outlookLogoFromPhoto
                                ? <><Check size={14} /> {isDe ? 'Event-Foto übernommen' : 'Event photo applied'}</>
                                : <><Icon iconName="Photo2" style={{ fontSize: 14 }} /> {isDe ? 'Event-Foto übernehmen' : 'Copy event photo here'}</>}
                            </button>
                          )}
                          {/* v28.30: Zuschneiden/Entfernen direkt neben dem Upload — sonst
                              stand das Bild doppelt auf dem Schirm. */}
                          {outlookLogoPreview && (
                            <>
                              <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => setLogoCropTarget('outlook')}>{isDe ? 'Zuschneiden' : 'Crop'}</button>
                              <button type="button" className="dex-ui-textbtn dex-ui-textbtn--danger" onClick={() => { setOutlookLogoPreview(''); setOutlookLogoFromPhoto(false); }}>{t('create.eventlogo.remove')}</button>
                            </>
                          )}
                        </div>
                        {/* v27.2 / v28.29: Größensteuerung mit dem TATSÄCHLICH verwendeten Kopfbild. */}
                        {((): React.ReactNode => {
                          const eff = effectiveHeaderImage('outlook', outlookLogoPreview);
                          return renderHeaderSizeControl(eff.src, eff.note);
                        })()}
                      </div>
                      <hr className="dex-ui-divider" />
                      <div className="dex-ui-field">
                        <div className="dex-ui-label">
                          <StepBadge n={30} />
                          {t('create.outlookdesc')}
                        </div>
                        <div className="dex-ui-inline">
                          <button
                            type="button"
                            className="btn btn-secondary dex-ui-btn-sm"
                            onClick={() => { setHtmlEditorMode('outlook'); setHtmlEditorOpen(true); }}
                          >
                            <Pencil size={14} /> {t('create.outlookdesc.edit')}
                          </button>
                          <span className="dex-ui-muted" style={{ flex: 1, minWidth: 200 }}>
                            {outlookBody
                              ? `${outlookBody.replace(/<[^>]+>/g, '').substring(0, 80)}${outlookBody.length > 80 ? '…' : ''}`
                              : t('create.outlookdesc.placeholder')}
                          </span>
                        </div>
                      </div>
                      {renderOutlookUpdateButton()}
                    </>)}
                    {advTab === 'fine' && (<>
                      <p className="dex-ui-section-desc" style={{ margin: '0 0 6px' }}>{isDe ? 'Seltene Schalter — die meisten Events lassen sie auf Standard.' : 'Rare switches — most events leave them at default.'}</p>
                      {disableEmails && disableOutlook && (<p className="dex-ui-muted">{isDe ? 'Kommunikation ist oben komplett abgeschaltet — hier gibt es nichts feinzustellen.' : 'Communication is switched off above — nothing to fine-tune here.'}</p>)}
                      {/* v31.2: Die drei Schalter als dex-ui-toggle-row statt über
                          commToggleRow — dieselben Werte an denselben Settern, aber mit
                          Hover und der Folge direkt unter dem Titel. */}
                      {!disableEmails && (<div className="dex-ui-section">
                        <div className="dex-ui-section-title">{isDe ? 'Mails an Teilnehmer' : 'Emails to attendees'}</div>
                        <div className="dex-ui-stack">
                          <label className={cx('dex-ui-toggle-row', !disableRegistrationEmail && 'is-active')}>
                            <input type="checkbox" checked={!disableRegistrationEmail} onChange={e => setDisableRegistrationEmail(!e.target.checked)} />
                            <span className="dex-ui-toggle-row-body">
                              <span className="dex-ui-toggle-row-title">
                                {isDe ? 'Teilnehmer bekommen eine Anmelde-Bestätigung' : 'Attendees get a registration confirmation'}
                                <InfoTooltip text={isDe
                                  ? 'Wenn aktiv: Teilnehmer bekommen bei der Anmeldung eine Bestätigungs-Mail (und, falls Warteliste aktiv, die Warteliste-Mail). Haken aus = es geht keine Anmelde-Bestätigung raus — die Abmelde-Mail bleibt davon unberührt.'
                                  : 'When active: attendees receive a confirmation email on registration (plus the waitlist email if a waitlist is active). Unchecked = no registration confirmation is sent — the cancellation email is unaffected.'} />
                              </span>
                              <span className="dex-ui-toggle-row-desc">{isDe ? 'Mail bei der Anmeldung (inkl. Wartelisten-Mail). Aus = keine Anmelde-Mail, die Abmelde-Mail bleibt.' : 'Email on registration (incl. waitlist email). Off = no registration email; the cancellation email stays.'}</span>
                            </span>
                          </label>
                          <label className={cx('dex-ui-toggle-row', !disableCancellationEmail && 'is-active')}>
                            <input type="checkbox" checked={!disableCancellationEmail} onChange={e => setDisableCancellationEmail(!e.target.checked)} />
                            <span className="dex-ui-toggle-row-body">
                              <span className="dex-ui-toggle-row-title">
                                {isDe ? 'Teilnehmer bekommen eine Abmelde-Bestätigung' : 'Attendees get a cancellation confirmation'}
                                <InfoTooltip text={isDe
                                  ? 'Wenn aktiv: Teilnehmer bekommen bei einer Abmeldung eine Bestätigungs-Mail. Haken aus = es geht keine Abmelde-Bestätigung raus (z.B. wenn du Teilnehmer still abmeldest) — die Anmelde-Mail bleibt davon unberührt.'
                                  : 'When active: attendees receive a confirmation email when cancelled. Unchecked = no cancellation confirmation is sent (e.g. when you remove attendees silently) — the registration email is unaffected.'} />
                              </span>
                              <span className="dex-ui-toggle-row-desc">{isDe ? 'Mail bei der Abmeldung. Aus = still abmelden, die Anmelde-Mail bleibt.' : 'Email on cancellation. Off = cancel silently; the registration email stays.'}</span>
                            </span>
                          </label>
                        </div>
                      </div>)}
                      {!disableOutlook && (<div className="dex-ui-section">
                        <div className="dex-ui-section-title">{isDe ? 'Outlook-Termin' : 'Outlook invite'}</div>
                        {/* v30.94: ausformuliert — „Termin abgesagt → Platz wird frei"
                            las sich als Pfeil-Rätsel (Nutzer: „das versteht man
                            textuell nicht"). Jetzt steht da, WER was tut und was
                            ohne den Haken passiert. */}
                        <label className={cx('dex-ui-toggle-row', autoDeregisterOnDecline && 'is-active')}>
                          <input type="checkbox" checked={autoDeregisterOnDecline} onChange={e => setAutoDeregisterOnDecline(e.target.checked)} />
                          <span className="dex-ui-toggle-row-body">
                            <span className="dex-ui-toggle-row-title">
                              {isDe ? 'Outlook-Absage gilt als Abmeldung' : 'Declining in Outlook counts as cancelling'}
                              <InfoTooltip text={isDe
                                ? 'Wenn aktiv: Sagt ein Teilnehmer den Outlook-Termin ab, wird er automatisch auch vom Event abgemeldet — der Platz wird frei und die Warteliste rückt nach. Ohne diesen Haken bekommt die Person bei einer Outlook-Absage nur eine Erinnerung, sich bei Bedarf selbst abzumelden. Hinweis: Diese Automatik greift erst, sobald die einmalige Anpassung im Outlook-Absage-Verarbeitungsschritt im Tenant eingerichtet ist.'
                                : 'When active: if an attendee declines the Outlook invite, they are automatically deregistered from the event — the spot is freed and the waitlist moves up. Without this, a decline only triggers a reminder asking the person to deregister themselves if needed. Note: this automation only takes effect once the one-time change in the Outlook-decline processing step is set up in the tenant.'} />
                            </span>
                            <span className="dex-ui-toggle-row-desc">{isDe
                              ? 'Sagt ein Teilnehmer den Kalendertermin in Outlook ab, meldet DEX ihn automatisch vom Event ab: Sein Platz wird frei, die Warteliste rückt nach. Ohne diesen Haken bekommt er nur eine Erinnerung, sich in der App selbst abzumelden.'
                              : 'If an attendee declines the calendar invite in Outlook, DEX automatically cancels their registration: the seat is freed and the waitlist moves up. Without this, they only get a reminder to cancel in the app themselves.'}</span>
                          </span>
                        </label>
                      </div>)}
                {!disableEmails && (<>
                {/* v30.89: Mitleser-Karte. v30.94: aus Ebene 1 in die Feineinstellungen
                    verlegt (Nutzer-Ansage 07.09.2026: das ist Feineinstellung). Die
                    Abmelde-Regel bleibt eingeklappt, bis jemand sie abweichend braucht. */}
                <div className="dex-ui-section">
                  <div className="dex-ui-section-title">{isDe ? 'Organizer in Kopie' : 'Organizers in copy'}</div>
                  <div className="dex-ui-card">
                    <div className="dex-ui-label">
                      <StepBadge n={27} />
                      {isDe ? 'Sollen die Organizer bei An- und Abmeldungen mitlesen?' : 'Should organizers be looped in on registrations / cancellations?'}
                      {/* v28.28: Präzisiert — die Organizer-Kopie läuft durchgehend auf
                          CC (vorher Bcc bei internen Empfängern) — der Organizer steht
                          damit sichtbar im Verteiler, „Allen antworten" landet beim
                          richtigen Ansprechpartner. */}
                      <InfoTooltip text={isDe
                        ? <>Wenn aktiv, steht der Organizer bei der Bestätigungs-Mail an den Teilnehmer sichtbar auf <strong>Kopie (Cc)</strong> — praktisch, um zu wissen, wer sich gerade an- oder abmeldet, und der Teilnehmer sieht direkt, wer sein Ansprechpartner ist. Bei großen Events willst du das vielleicht nicht für jede einzelne Anmeldung — dann kannst du es hier gezielt einschränken (z.B. nur kurz vorm Event, wenn kurzfristige Änderungen wichtig sind).</>
                        : <>When on, the organizer is visibly on <strong>copy (Cc)</strong> of the confirmation email sent to the attendee — handy to know who is signing up or off, and the attendee immediately sees who to contact. For large events you might not want this for every single sign-up — you can narrow it down here (e.g. only close to the event when last-minute changes matter).</>} />
                    </div>
                    <p className="dex-ui-help" style={{ margin: '0 0 12px' }}>
                      {isDe
                        ? 'Wenn ja, steht der Organizer sichtbar auf Cc der Bestätigungs-Mail — der Teilnehmer sieht so direkt seinen Ansprechpartner. Bei großen Events lässt sich das hier eingrenzen.'
                        : 'If yes, the organizer is visibly on Cc of the confirmation email — so the attendee sees who to contact. For large events you can narrow it down here.'}
                    </p>
                    {/* Anmeldung */}
                    <div className="dex-ui-field">
                      <div className="dex-ui-label" style={{ fontSize: '0.82rem' }}>{isDe ? 'Bei Anmeldungen' : 'On registrations'}</div>
                      <div className="dex-ui-inline" role="radiogroup" aria-label={isDe ? 'Bei Anmeldungen' : 'On registrations'}>
                        {([
                          { v: 'never' as const, de: 'Nicht informieren', en: 'Don\'t notify' },
                          { v: 'always' as const, de: 'Bei jeder Anmeldung', en: 'On every registration' },
                          { v: 'fromDate' as const, de: 'Erst ab Datum', en: 'Only from date' },
                        ]).map(o => (
                          <button key={o.v} type="button" role="radio" aria-checked={notifyOrgRegisterMode === o.v} className={cx('dex-ui-chip', notifyOrgRegisterMode === o.v && 'is-active')} onClick={() => setNotifyOrgRegisterMode(o.v)}>
                            {isDe ? o.de : o.en}
                          </button>
                        ))}
                      </div>
                      {notifyOrgRegisterMode === 'fromDate' && (
                        <div style={{ marginTop: 10 }}>
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
                          <p className="dex-ui-help">
                            {isDe ? 'Z.B. eine Woche vor dem Event — kurzfristige Anmeldungen werden dann an die Organizer gespiegelt.' : 'E.g. one week before the event — last-minute registrations are mirrored to organizers from then on.'}
                          </p>
                        </div>
                      )}
                    </div>
                    {orgCancelOpen ? (
                      /* Abmeldung */
                      <div className="dex-ui-field">
                        <div className="dex-ui-label" style={{ fontSize: '0.82rem' }}>{isDe ? 'Bei Abmeldungen' : 'On cancellations'}</div>
                        <div className="dex-ui-inline" role="radiogroup" aria-label={isDe ? 'Bei Abmeldungen' : 'On cancellations'}>
                          {([
                            { v: 'never' as const, de: 'Nicht informieren', en: 'Don\'t notify' },
                            { v: 'always' as const, de: 'Bei jeder Abmeldung', en: 'On every cancellation' },
                            { v: 'afterDeadline' as const, de: 'Erst nach der letzten Abmeldemöglichkeit', en: 'Only after the last cancellation date' },
                          ]).map(o => (
                            <button key={o.v} type="button" role="radio" aria-checked={notifyOrgCancelMode === o.v} className={cx('dex-ui-chip', notifyOrgCancelMode === o.v && 'is-active')} onClick={() => setNotifyOrgCancelMode(o.v)}>
                              {isDe ? o.de : o.en}
                            </button>
                          ))}
                        </div>
                        <p className="dex-ui-help">
                          {isDe
                            ? '„Erst nach der letzten Abmeldemöglichkeit" nutzt das in Schritt 4 (Kapazität & Sichtbarkeit) gesetzte Datum „Letzte Abmeldemöglichkeit". Vor diesem Stichtag gelten Abmeldungen als unproblematisch — danach möchtest du als Organizer aber wissen, wer noch abspringt.'
                            : '„Only after the last cancellation date" uses the date set in step 4 (Capacity & Visibility) under „Last cancellation date". Cancellations before that are considered routine — after that, organizers usually want to know about late drop-outs.'}
                        </p>
                      </div>
                    ) : (
                      <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" style={{ marginLeft: -8 }} onClick={() => setOrgCancelOpen(true)}>
                        <Pencil size={14} />
                        {isDe
                          ? `Bei Abmeldungen: ${notifyOrgCancelMode === 'always' ? 'bei jeder Abmeldung' : notifyOrgCancelMode === 'afterDeadline' ? 'erst nach der letzten Abmeldemöglichkeit' : 'nicht informieren'} — ändern`
                          : `On cancellations: ${notifyOrgCancelMode === 'always' ? 'on every cancellation' : notifyOrgCancelMode === 'afterDeadline' ? 'only after the last cancellation date' : 'do not notify'} — change`}
                      </button>
                    )}
                  </div>
                </div>
                </>)}
                  {/* inactiveHandling: Verhalten, wenn eine angemeldete Person
                      nicht mehr bei Deloitte arbeitet. 'notify' = Organizer per
                      Mail informieren (Standard), 'autoderegister' = automatisch
                      abmelden (beim Öffnen der App durch einen Organizer). */}
                  <div className="dex-ui-section">
                    <div className="dex-ui-section-title">{isDe ? 'Konto nicht mehr aktiv' : 'Account no longer active'}</div>
                    <div className="dex-ui-label">
                      {isDe ? 'Was passiert, wenn eine angemeldete Person nicht mehr bei Deloitte arbeitet?' : 'What happens when a registered person no longer works at Deloitte?'}
                      <InfoTooltip text={isDe
                        ? 'Die App erkennt beim Öffnen durch einen Organizer, wenn das Deloitte-Konto einer angemeldeten Person nicht mehr aktiv ist. „Organizer informieren" schickt dann eine Hinweis-Mail; „Automatisch abmelden" entfernt die Person direkt aus der Teilnehmerliste (Platz wird frei, Warteliste rückt nach).'
                        : 'When an organizer opens the app, it detects registered people whose Deloitte account is no longer active. „Notify organizer" sends an info email; „Auto-deregister" removes the person from the attendee list right away (the spot is freed, the waitlist moves up).'} />
                    </div>
                    <div className="dex-ui-grid-2">
                      <label className={cx('dex-ui-toggle-row', inactiveHandling === 'notify' && 'is-active')}>
                        <input type="radio" name="inactiveHandling" value="notify" checked={inactiveHandling === 'notify'} onChange={() => setInactiveHandling('notify')} />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">{isDe ? 'Organizer per E-Mail informieren' : 'Notify the organizer by email'}</span>
                          <span className="dex-ui-toggle-row-desc">{isDe ? 'Standard — die Anmeldung bleibt, du bekommst eine Hinweis-Mail.' : 'Default — the registration stays, you get an info email.'}</span>
                        </span>
                      </label>
                      <label className={cx('dex-ui-toggle-row', inactiveHandling === 'autoderegister' && 'is-active')}>
                        <input type="radio" name="inactiveHandling" value="autoderegister" checked={inactiveHandling === 'autoderegister'} onChange={() => setInactiveHandling('autoderegister')} />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">{isDe ? 'Automatisch abmelden' : 'Auto-deregister'}</span>
                          <span className="dex-ui-toggle-row-desc">{isDe ? 'Sobald ein Organizer die App öffnet — der Platz wird frei, die Warteliste rückt nach.' : 'As soon as an organizer opens the app — the spot is freed, the waitlist moves up.'}</span>
                        </span>
                      </label>
                    </div>
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
