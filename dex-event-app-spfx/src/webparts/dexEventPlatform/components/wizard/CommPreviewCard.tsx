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
 * es beim Anlegen noch nicht. */
import * as React from 'react';
import { buildEmailFromTemplate, buildOutlookBody, replacePlaceholders, getCachedLogoBase64, getCachedOrbBase64 } from '../../services/EmailTemplates';
import { formatOrganizerList } from '../../context/eventTextHelpers';
import { outlookDefaultBodyTemplate } from '../../utils/outlookDefaultBody';
import { buildProgramHtml, applyProgramPlaceholder } from '../../utils/programPlaceholder';
import { AgendaItem } from '../../types';
import { useCurrentUser } from '../../context/UserContext';
import { EventService } from '../../services/EventService';
import { EmailOverrideEntry } from './emailOverrideEntry';

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
      ...p.headerLayoutFor(p.emailLogoPreview),
    };
    const built = buildEmailFromTemplate(merged, vars);
    const eff = p.effectiveHeaderImage('email', p.emailLogoPreview);
    const html = built.body
      .replace(/\{\{LOGO_URL\}\}/g, getCachedLogoBase64() || '')
      .replace(/\{\{ORB_URL\}\}/g, eff.src || getCachedOrbBase64() || '');
    return { subject: built.subject, html };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [p.emailTemplates, p.emailTemplateOverrides, p.emailLanguage, p.emailLogoPreview, p.title, p.location, p.startDate, p.endDate, p.organizer, p.contactEmail, isDe, meName]);

  const outlookHtml = React.useMemo((): string => {
    // v30.94: Leerer Body = der Standard-Text, der auch gespeichert wird
    // (utils/outlookDefaultBody) — nicht ein erfundener Vorschau-Satz.
    const bodyRaw = p.outlookBody || outlookDefaultBodyTemplate(p.emailLanguage);
    const heading = replacePlaceholders(p.outlookHeading || '', vars) || vars.EventTitle;
    const sub = replacePlaceholders(p.outlookSubheading || '', vars) || vars.Location;
    const eff = p.effectiveHeaderImage('outlook', p.outlookLogoPreview);
    return buildOutlookBody(heading, applyProgramPlaceholder(replacePlaceholders(bodyRaw, vars), vars.Programm), sub, p.headerLayoutFor(p.outlookLogoPreview), undefined, isDe)
      .replace(/\{\{LOGO_URL\}\}/g, getCachedLogoBase64() || '')
      .replace(/\{\{ORB_URL\}\}/g, eff.src || getCachedOrbBase64() || '');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [p.outlookBody, p.outlookHeading, p.outlookSubheading, p.outlookLogoPreview, p.title, p.location, p.startDate, p.endDate, isDe, meName]);

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
  const frameStyle: React.CSSProperties = { width: '100%', height: 420, border: '1px solid var(--dex-gray-200)', borderRadius: 8, background: '#fff' };
  const tabBtn = (key: 'mail' | 'outlook', label: string, off: boolean): React.ReactElement => (
    <button key={key} type="button" onClick={() => setPane(key)} style={{
      padding: '6px 14px', borderRadius: 999, cursor: 'pointer', fontSize: '0.8rem', fontWeight: 600,
      border: `1px solid ${pane === key ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
      background: pane === key ? 'rgba(134,188,37,0.12)' : '#fff', color: off ? 'var(--dex-gray-400)' : 'var(--dex-gray-800)',
      textDecoration: off ? 'line-through' : 'none',
    }}>{label}</button>
  );

  return (
    <div style={{ marginTop: 24, padding: 14, borderRadius: 12, border: '1px solid var(--dex-gray-200)', background: 'var(--dex-gray-50, #f8f9fa)' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', marginBottom: 10 }}>
        <strong style={{ fontSize: '0.9rem' }}>{isDe ? 'So sieht es aus' : 'This is what goes out'}</strong>
        <span style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)' }}>
          {isDe ? '— Vorschau mit deinen aktuellen Einstellungen, Empfänger: du' : '— preview with your current settings, recipient: you'}
        </span>
        <span style={{ flex: 1 }} />
        {tabBtn('mail', isDe ? 'Anmeldebestätigung' : 'Registration email', p.disableEmails)}
        {tabBtn('outlook', isDe ? 'Outlook-Termin' : 'Outlook invite', p.disableOutlook)}
      </div>
      {nothing ? (
        <div style={{ padding: '14px 16px', borderRadius: 8, background: 'rgba(237,139,0,0.09)', fontSize: '0.85rem', lineHeight: 1.5 }}>
          {isDe ? 'Kommunikation ist abgeschaltet — es gibt nichts zu zeigen. Wähle oben „Mail + Outlook-Termin", „Nur Mail" oder „Nur Outlook-Termin".' : 'Communication is switched off — nothing to show. Pick “Email + Outlook invite”, “Email only” or “Outlook invite only” above.'}
        </div>
      ) : (
        <>
          {pane === 'mail' && (
            p.disableEmails
              ? <div style={{ padding: '12px 14px', borderRadius: 8, background: 'rgba(237,139,0,0.09)', fontSize: '0.82rem' }}>{isDe ? 'Mails sind für dieses Event abgeschaltet — diese Mail geht nicht raus.' : 'Emails are switched off for this event — this mail is not sent.'}</div>
              : (
                <>
                  <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginBottom: 6 }}>
                    <strong>{isDe ? 'Betreff:' : 'Subject:'}</strong> {mail.subject}
                  </div>
                  <iframe title={isDe ? 'Vorschau Anmeldebestätigung' : 'Registration email preview'} sandbox="" srcDoc={mail.html} style={frameStyle} />
                </>
              )
          )}
          {pane === 'outlook' && (
            p.disableOutlook
              ? <div style={{ padding: '12px 14px', borderRadius: 8, background: 'rgba(237,139,0,0.09)', fontSize: '0.82rem' }}>{isDe ? 'Der Outlook-Termin ist für dieses Event abgeschaltet.' : 'The Outlook invite is switched off for this event.'}</div>
              : <iframe title={isDe ? 'Vorschau Outlook-Termin' : 'Outlook invite preview'} sandbox="" srcDoc={outlookHtml} style={frameStyle} />
          )}
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginTop: 10 }}>
            <button type="button" className="btn btn-secondary" disabled={p.disableEmails || testState === 'busy'} onClick={() => { void sendTest(); }} style={{ fontSize: '0.82rem', padding: '6px 14px' }}>
              {testState === 'busy' ? (isDe ? 'Wird verschickt…' : 'Sending…') : (isDe ? 'Testmail an mich' : 'Send me a test email')}
            </button>
            <span style={{ fontSize: '0.76rem', color: testState === 'error' ? 'var(--dex-red, #c00)' : 'var(--dex-gray-500)' }}>
              {testState === 'sent'
                ? (isDe ? `Unterwegs an ${currentUser.email} — kommt in wenigen Minuten, Betreff beginnt mit [Test].` : `On its way to ${currentUser.email} — arrives within minutes, subject starts with [Test].`)
                : testState === 'error'
                  ? (isDe ? 'Konnte nicht in die Mail-Warteschlange geschrieben werden — bitte später erneut.' : 'Could not be queued — please try again later.')
                  : (isDe ? 'Schickt genau diese Anmeldebestätigung an deine Adresse, ohne etwas am Event zu ändern.' : 'Sends exactly this registration email to your address without changing the event.')}
            </span>
          </div>
        </>
      )}
    </div>
  );
};

export default CommPreviewCard;
