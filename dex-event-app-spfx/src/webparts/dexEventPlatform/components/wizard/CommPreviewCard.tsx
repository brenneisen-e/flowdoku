/* CommPreviewCard — v30.90 (Kommunikations-Schritt, Stufe B + C des Konzepts
 * docs/konzept-kommunikation-schritt.md).
 *
 * Ebene 2 des Schritts: EINE gerenderte Vorschau der Anmeldebestätigung und
 * des Outlook-Termins, so wie sie mit den aktuellen Einstellungen rausgehen —
 * statt acht Zeilen Erklärtext. Eine Vorschau beantwortet „was kommt raus?"
 * schneller als jede Beschreibung und zeigt sofort, ob Logo, Sprache und
 * Text zusammenpassen.
 *
 * Gebaut wird EXAKT mit den Funktionen des echten Versands
 * (`buildEmailFromTemplate`, `buildOutlookBody`, `wrapTemplate` darunter) —
 * eine zweite Vorschau-Logik würde irgendwann anders aussehen als die Mail.
 * Dieselbe Regel gilt im HtmlEditorModal; hier fehlt nur der Editor.
 *
 * „Testmail an mich" schickt genau dieses HTML an die angemeldete Person über
 * die normale Mail-Queue. {{ORB_URL}}/{{LOGO_URL}} werden dafür fest
 * eingebettet: Der Flow löst sie sonst über die Event-Zeile auf, und die gibt
 * es beim Anlegen noch nicht.
 *
 * v31.2: Optik auf die dex-ui-Klassen umgestellt (Reiter mit Hover, Warnungen
 * als Callouts, Karte als dex-ui-card). Aufbau, Memos, Testmail-Pfad unverändert. */
import * as React from 'react';
import { buildEmailFromTemplate, buildOutlookBody, replacePlaceholders, getCachedLogoBase64, getCachedOrbBase64 } from '../../services/EmailTemplates';
import { formatOrganizerList } from '../../context/eventTextHelpers';
import { outlookDefaultBodyTemplate } from '../../utils/outlookDefaultBody';
import { buildProgramHtml, applyProgramPlaceholder } from '../../utils/programPlaceholder';
import { AgendaItem } from '../../types';
import { useCurrentUser } from '../../context/UserContext';
import { EventService } from '../../services/EventService';
import { EmailOverrideEntry } from './emailOverrideEntry';
// v31.2: Gemeinsame Klassen (Reiter, Callouts, Karte) statt Inline-Styles —
// Inline kann kein :hover, und die Reiter hatten bis dahin keinen.
import { cx } from '../dexUi';
import { Mail, Calendar, Send, Check, AlertCircle } from '../Icons';

export interface CommPreviewCardProps {
  isDe: boolean;
  emailLanguage: string;
  emailTemplates: { id: number; templateType: string; language: string; subject: string; heading: string; headingColor: string; bodyHtml: string; }[];
  emailTemplateOverrides: Record<string, EmailOverrideEntry>;
  emailLogoPreview: string;
  outlookLogoPreview: string;
  effectiveHeaderImage: (kind: 'email' | 'outlook', own: string) => { src: string; note: string };
  headerLayoutFor: (logoB64: string) => { imageWidth: number; imagePaddingV: number; imagePaddingH: number };
  title: string;
  location: string;
  startDate: string;
  endDate: string;
  organizer: string;
  contactEmail: string;
  outlookBody: string;
  outlookHeading: string;
  outlookSubheading: string;
  disableEmails: boolean;
  disableOutlook: boolean;
  /** Id des bearbeiteten Events; leer beim Anlegen. */
  eventId: string;
  /** v30.95: Programmpunkte des offenen Reiters für {{Programm}}. */
  agenda: AgendaItem[];
}

export const CommPreviewCard: React.FC<CommPreviewCardProps> = (p) => {
  const { isDe } = p;
  const { currentUser } = useCurrentUser();
  const [pane, setPane] = React.useState<'mail' | 'outlook'>(p.disableEmails && !p.disableOutlook ? 'outlook' : 'mail');
  const [testState, setTestState] = React.useState<'idle' | 'busy' | 'sent' | 'error'>('idle');

  const fmt = (iso: string, withTime: boolean): string => {
    if (!iso) return '';
    const d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    return withTime
      ? d.toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })
      : d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  };
  // Dieselben Variablen wie die Live-Vorschau des HtmlEditorModal (WizardModals).
  const meName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || (isDe ? 'Max Mustermann' : 'Jane Doe');
  const vars: Record<string, string> = {
    EventTitle: p.title || (isDe ? 'Event-Titel' : 'Event title'),
    Name: meName,
    Organizer: formatOrganizerList([p.organizer], p.emailLanguage) || p.organizer || 'Organizer',
    ContactEmail: (p.contactEmail || '').trim() || 'kontakt@deloitte.de',
    AppUrl: 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView',
    WaitlistPosition: '1',
    Location: p.location || (isDe ? 'Veranstaltungsort' : 'Venue'),
    StartDate: fmt(p.startDate, true),
    EndDate: fmt(p.endDate, true),
    EventDate: fmt(p.startDate, true),
    // v30.95: Programm-Tabelle (buildEmailFromTemplate setzt sie roh ein).
    Programm: buildProgramHtml(p.agenda, p.emailLanguage),
  };

  // v30.98: Layout und Bildquelle VOR den Memos auflösen und als Abhängigkeit
  // führen. Vorher hingen beide Memos nur am Logo-String — ein Klick auf
  // „Volle Breite" änderte den Layout-State, die Karte rechnete nicht neu und
  // zeigte das Bild weiter klein (Nutzer-Screenshot 07.09.2026).
  const mailLayout = p.headerLayoutFor(p.emailLogoPreview);
  const mailSrc = p.effectiveHeaderImage('email', p.emailLogoPreview).src || '';
  const olLayout = p.headerLayoutFor(p.outlookLogoPreview);
  const olSrc = p.effectiveHeaderImage('outlook', p.outlookLogoPreview).src || '';

  const mail = React.useMemo((): { subject: string; html: string } => {
    const tpl = p.emailTemplates.find(x => x.templateType === 'Anmeldung' && x.language === p.emailLanguage);
    const ov = p.emailTemplateOverrides['Anmeldung'];
    const fallbackBody = isDe
      ? '<p>Hallo {{Name}},</p><p>deine Anmeldung zu <strong>{{EventTitle}}</strong> ist bestätigt.</p><p><strong>Wann:</strong> {{StartDate}}<br><strong>Wo:</strong> {{Location}}</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>'
      : '<p>Hi {{Name}},</p><p>your registration for <strong>{{EventTitle}}</strong> is confirmed.</p><p><strong>When:</strong> {{StartDate}}<br><strong>Where:</strong> {{Location}}</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>';
    const merged = {
      subject: (ov && ov.subject) || (tpl && tpl.subject) || (isDe ? 'Anmeldebestätigung: {{EventTitle}}' : 'Registration confirmed: {{EventTitle}}'),
      headingColor: (ov && ov.headingColor) || (tpl && tpl.headingColor) || '#86bc25',
      heading: (ov && ov.heading) || (tpl && tpl.heading) || (isDe ? 'Anmeldung erfolgreich' : 'Registration successful'),
      subheading: ov && ov.subheading !== undefined ? ov.subheading : undefined,
      bodyHtml: (ov && ov.bodyHtml) || (tpl && tpl.bodyHtml) || fallbackBody,
      headingFontSize: ov && ov.headingFontSize, headingBold: ov && ov.headingBold, headingItalic: ov && ov.headingItalic,
      subheadingColor: ov && ov.subheadingColor, subheadingFontSize: ov && ov.subheadingFontSize,
      subheadingBold: ov && ov.subheadingBold, subheadingItalic: ov && ov.subheadingItalic,
      ...mailLayout,
    };
    const built = buildEmailFromTemplate(merged, vars);
    const html = built.body
      .replace(/\{\{LOGO_URL\}\}/g, getCachedLogoBase64() || '')
      .replace(/\{\{ORB_URL\}\}/g, mailSrc || getCachedOrbBase64() || '');
    return { subject: built.subject, html };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [p.emailTemplates, p.emailTemplateOverrides, p.emailLanguage, p.emailLogoPreview, p.title, p.location, p.startDate, p.endDate, p.organizer, p.contactEmail, isDe, meName, mailSrc, mailLayout.imageWidth, mailLayout.imagePaddingV, mailLayout.imagePaddingH, p.agenda]);

  const outlookHtml = React.useMemo((): string => {
    // v30.94: Leerer Body = der Standard-Text, der auch gespeichert wird
    // (utils/outlookDefaultBody) — nicht ein erfundener Vorschau-Satz.
    const bodyRaw = p.outlookBody || outlookDefaultBodyTemplate(p.emailLanguage);
    const heading = replacePlaceholders(p.outlookHeading || '', vars) || vars.EventTitle;
    const sub = replacePlaceholders(p.outlookSubheading || '', vars) || vars.Location;
    return buildOutlookBody(heading, applyProgramPlaceholder(replacePlaceholders(bodyRaw, vars), vars.Programm), sub, olLayout, undefined, isDe)
      .replace(/\{\{LOGO_URL\}\}/g, getCachedLogoBase64() || '')
      .replace(/\{\{ORB_URL\}\}/g, olSrc || getCachedOrbBase64() || '');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [p.outlookBody, p.outlookHeading, p.outlookSubheading, p.outlookLogoPreview, p.title, p.location, p.startDate, p.endDate, isDe, meName, olSrc, olLayout.imageWidth, olLayout.imagePaddingV, olLayout.imagePaddingH, p.agenda, p.emailLanguage]);

  const sendTest = async (): Promise<void> => {
    if (testState === 'busy') return;
    const to = (currentUser.email || '').trim();
    if (!to) { setTestState('error'); return; }
    setTestState('busy');
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (!ctx) throw new Error('Kein SPFx-Kontext');
      const svc = new EventService(ctx);
      const subject = `${isDe ? '[Test] ' : '[Test] '}${mail.subject}`;
      const ok = await svc.queueEmail(subject, to, meName, mail.html, 'Info', p.title || '', p.eventId || '');
      setTestState(ok ? 'sent' : 'error');
    } catch (err) {
      console.warn('[DEX] Testmail fehlgeschlagen:', err);
      setTestState('error');
    }
  };

  const nothing = p.disableEmails && p.disableOutlook;
  const frameStyle: React.CSSProperties = { width: '100%', height: 420, border: '1px solid var(--dex-gray-200, #e8e8e8)', borderRadius: 12, background: '#fff', display: 'block' };
  // v31.2: Reiter als dex-ui-tabs (Hover, aktiver Zustand kommen aus der
  // Klasse). Ein abgeschalteter Kanal wird nicht mehr durchgestrichen —
  // das las sich wie „gelöscht" — sondern trägt den Zusatz „aus".
  const tabBtn = (key: 'mail' | 'outlook', icon: React.ReactNode, label: string, off: boolean): React.ReactElement => (
    <button key={key} type="button" className={cx('dex-ui-tab', pane === key && 'is-active', off && 'is-off')} onClick={() => setPane(key)}
      aria-pressed={pane === key} style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
      {icon}{label}{off && <span style={{ fontWeight: 500, fontSize: '0.72rem' }}>· {isDe ? 'aus' : 'off'}</span>}
    </button>
  );
  const warn = (text: string): React.ReactElement => (
    <div className="dex-ui-callout dex-ui-callout--warn" role="status">
      <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span><span>{text}</span>
    </div>
  );

  return (
    <div className="dex-ui-card dex-ui-card--soft">
      {/* v31.2: Kein eigener Titel mehr — der Schritt setzt darüber bereits die
          Abschnitts-Überschrift „So geht es raus"; zwei Überschriften für eine
          Karte lasen sich wie zwei Kästen. Reiter links, Hinweis rechts. */}
      <div className="dex-ui-inline" style={{ justifyContent: 'space-between', marginBottom: 12 }}>
        <div className="dex-ui-tabs" role="tablist" aria-label={isDe ? 'Vorschau wählen' : 'Choose preview'}>
          {tabBtn('mail', <Mail size={14} />, isDe ? 'Anmeldebestätigung' : 'Registration email', p.disableEmails)}
          {tabBtn('outlook', <Calendar size={14} strokeWidth={2} />, isDe ? 'Outlook-Termin' : 'Outlook invite', p.disableOutlook)}
        </div>
        <span className="dex-ui-muted" style={{ fontSize: '0.76rem' }}>
          {isDe ? 'Vorschau mit deinen aktuellen Einstellungen · Empfänger: du' : 'Preview with your current settings · recipient: you'}
        </span>
      </div>
      {nothing ? warn(isDe ? 'Kommunikation ist abgeschaltet — es gibt nichts zu zeigen. Wähle oben „Mail + Outlook-Termin", „Nur Mail" oder „Nur Outlook-Termin".' : 'Communication is switched off — nothing to show. Pick “Email + Outlook invite”, “Email only” or “Outlook invite only” above.') : (
        <>
          {pane === 'mail' && (
            p.disableEmails
              ? warn(isDe ? 'Mails sind für dieses Event abgeschaltet — diese Mail geht nicht raus.' : 'Emails are switched off for this event — this mail is not sent.')
              : (
                <>
                  <div className="dex-ui-muted" style={{ marginBottom: 8, color: 'var(--dex-gray-600, #666)' }}>
                    <strong>{isDe ? 'Betreff:' : 'Subject:'}</strong> {mail.subject}
                  </div>
                  <iframe title={isDe ? 'Vorschau Anmeldebestätigung' : 'Registration email preview'} sandbox="" srcDoc={mail.html} style={frameStyle} />
                </>
              )
          )}
          {pane === 'outlook' && (
            p.disableOutlook
              ? warn(isDe ? 'Der Outlook-Termin ist für dieses Event abgeschaltet.' : 'The Outlook invite is switched off for this event.')
              : <iframe title={isDe ? 'Vorschau Outlook-Termin' : 'Outlook invite preview'} sandbox="" srcDoc={outlookHtml} style={frameStyle} />
          )}
          {/* v31.2: Knopf und Folge in EINER Zeile; Erfolg und Fehler als
              farbige Callouts statt eines grauen Satzes, den man übersieht. */}
          <div className="dex-ui-inline" style={{ marginTop: 12, gap: 10 }}>
            <button type="button" className="btn btn-secondary dex-ui-btn-sm" disabled={p.disableEmails || testState === 'busy'} onClick={() => { void sendTest(); }}
              title={p.disableEmails ? (isDe ? 'Mails sind für dieses Event abgeschaltet.' : 'Emails are switched off for this event.') : undefined}>
              <Send size={14} />{testState === 'busy' ? (isDe ? 'Wird verschickt…' : 'Sending…') : (isDe ? 'Testmail an mich' : 'Send me a test email')}
            </button>
            {testState === 'sent' ? (
              <span className="dex-ui-callout dex-ui-callout--success dex-ui-callout--sm" role="status">
                <span className="dex-ui-callout-icon"><Check size={14} /></span>
                <span>{isDe ? `Unterwegs an ${currentUser.email} — kommt in wenigen Minuten, Betreff beginnt mit [Test].` : `On its way to ${currentUser.email} — arrives within minutes, subject starts with [Test].`}</span>
              </span>
            ) : testState === 'error' ? (
              <span className="dex-ui-callout dex-ui-callout--danger dex-ui-callout--sm" role="alert">
                <span className="dex-ui-callout-icon"><AlertCircle size={14} /></span>
                <span>{isDe ? 'Konnte nicht in die Mail-Warteschlange geschrieben werden — bitte später erneut.' : 'Could not be queued — please try again later.'}</span>
              </span>
            ) : (
              <span className="dex-ui-muted" style={{ fontSize: '0.76rem' }}>
                {isDe ? 'Schickt genau diese Anmeldebestätigung an deine Adresse, ohne etwas am Event zu ändern.' : 'Sends exactly this registration email to your address without changing the event.'}
              </span>
            )}
          </div>
        </>
      )}
    </div>
  );
};

export default CommPreviewCard;
