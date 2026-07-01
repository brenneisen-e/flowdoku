/**
 * v24.98: Globaler Mail-Vorlagen-Editor (Admin-only).
 *
 * Zeigt alle Standard-Mail-Vorlagen aus DEX_EmailTemplates und lässt Admins
 * Betreff, Überschrift, Unter-Überschrift, Überschrift-Farbe und den HTML-Body
 * pro Vorlage (DE/EN) bearbeiten — mit Live-Vorschau im Deloitte-Layout.
 *
 * Erreichbar über den Admin-Hub (Kachel „Einstellungen & Templates").
 */
import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { useLanguage } from '../context/LanguageContext';
import { useDialog } from '../context/DialogContext';
import { useIsMobile } from '../utils/useIsMobile';
import { Mail, Check, X } from './Icons';
import { wrapTemplate } from '../services/EmailTemplates';

interface Tpl { id: number; templateType: string; language: string; subject: string; headingColor: string; heading: string; subheading: string; bodyHtml: string }
interface Draft { subject: string; heading: string; subheading: string; headingColor: string; bodyHtml: string }

// Klartext-Bezeichnungen der wichtigsten Vorlagen-Typen (Fallback: Roh-Typ).
const TYPE_LABELS: Record<string, string> = {
  Anmeldung: 'Anmeldebestätigung',
  Warteliste: 'Warteliste — Bestätigung',
  Abmeldung: 'Abmeldebestätigung',
  AbmeldungAuto: 'Abmeldung (automatisch nach Outlook-Absage)',
  Nachruecken: 'Nachrücken von der Warteliste',
  EventErstellt: 'Event angelegt (an Organisatoren)',
  OrgNachruecker: 'Organisator-Info: Nachrücker',
  OutlookDeclineReminder: 'Outlook-Absage — Erinnerung',
  OutlookDeclineReminder_OnBehalfOf: 'Outlook-Absage — Erinnerung (stellvertretend)',
  OutlookForwardNotification: 'Outlook-Weiterleitung — Hinweis',
  OutlookDeclineDigest: 'Outlook-Absagen — Sammelmeldung',
  TeamMemberJoined: 'Team: Mitglied hinzugefügt',
  TeamJoinRequest: 'Team: Beitritts-Anfrage',
  TeamJoinRejected: 'Team: Beitritt abgelehnt',
  TeamLeadTransferred: 'Team: Lead-Rolle übergeben',
  TeamMemberCancelled: 'Team: Mitglied abgemeldet',
  RoommateRequest: 'Zimmerpartner-Anfrage',
  GroupSwitchConfirmed: 'Gruppenwechsel bestätigt',
  GroupSwitchWaitlist: 'Gruppenwechsel — Warteliste',
  OverbookingApology: 'Überbuchung — Entschuldigung',
  WeeklyReport: 'Wöchentlicher Admin-Bericht',
};

const PLACEHOLDERS = ['{{Name}}', '{{Vorname}}', '{{EventTitle}}', '{{AppUrl}}', '{{OrganizerHtml}}', '{{WaitlistPosition}}'];

export default function EmailTemplatesPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { isAdmin, originalIsAdmin } = useRoles();
  const { getAllEmailTemplates, updateEmailTemplate } = useEvents();
  const { locale } = useLanguage();
  const { showAlert } = useDialog();
  const isDe = locale === 'de';
  const isMobile = useIsMobile();
  const adminLike = isAdmin || originalIsAdmin;

  const [templates, setTemplates] = React.useState<Tpl[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [search, setSearch] = React.useState('');
  const [expandedId, setExpandedId] = React.useState<number | null>(null);
  const [drafts, setDrafts] = React.useState<Record<number, Draft>>({});
  const [savingId, setSavingId] = React.useState<number | null>(null);

  React.useEffect(() => {
    if (!adminLike) { navigate('start'); return; }
    let cancelled = false;
    setLoading(true);
    getAllEmailTemplates().then(all => {
      if (cancelled) return;
      // Pseudo-Einträge wie „_Config" (Logo/KPI-Zähler) ausblenden.
      const real = all.filter(t => t.templateType && t.templateType.indexOf('_') !== 0)
        .sort((a, b) => a.templateType.localeCompare(b.templateType) || a.language.localeCompare(b.language));
      setTemplates(real);
      setLoading(false);
    }).catch(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [adminLike]);

  if (!adminLike) return <div className="page-container" />;

  const typeLabel = (t: string): string => TYPE_LABELS[t] || t;

  const toggle = (t: Tpl): void => {
    setDrafts(prev => prev[t.id] ? prev : { ...prev, [t.id]: { subject: t.subject, heading: t.heading, subheading: t.subheading, headingColor: t.headingColor || '#86bc25', bodyHtml: t.bodyHtml } });
    setExpandedId(expandedId === t.id ? null : t.id);
  };
  const setField = (id: number, field: keyof Draft, value: string): void => {
    setDrafts(prev => ({ ...prev, [id]: { ...prev[id], [field]: value } }));
  };
  const isDirty = (t: Tpl): boolean => {
    const d = drafts[t.id];
    return !!d && (d.subject !== t.subject || d.heading !== t.heading || d.subheading !== t.subheading || (d.headingColor || '') !== (t.headingColor || '') || d.bodyHtml !== t.bodyHtml);
  };
  const reset = (t: Tpl): void => {
    setDrafts(prev => ({ ...prev, [t.id]: { subject: t.subject, heading: t.heading, subheading: t.subheading, headingColor: t.headingColor || '#86bc25', bodyHtml: t.bodyHtml } }));
  };
  const save = async (t: Tpl): Promise<void> => {
    const d = drafts[t.id]; if (!d) return;
    setSavingId(t.id);
    const ok = await updateEmailTemplate(t.id, d);
    setSavingId(null);
    if (ok) {
      setTemplates(prev => prev.map(x => x.id === t.id ? { ...x, ...d } : x));
      showAlert(isDe ? 'Vorlage gespeichert.' : 'Template saved.', { variant: 'success' });
    } else {
      showAlert(isDe ? 'Speichern fehlgeschlagen.' : 'Saving failed.', { variant: 'error' });
    }
  };

  const q = search.trim().toLowerCase();
  const filtered = templates.filter(t => !q
    || typeLabel(t.templateType).toLowerCase().indexOf(q) >= 0
    || t.templateType.toLowerCase().indexOf(q) >= 0
    || t.subject.toLowerCase().indexOf(q) >= 0);

  const labelStyle: React.CSSProperties = { display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 3 };
  const inputStyle: React.CSSProperties = { width: '100%', padding: '7px 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.85rem', boxSizing: 'border-box' };
  const cardStyle: React.CSSProperties = { background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 12, marginBottom: 10, overflow: 'hidden' };

  return (
    <div className="page-container">
      <h1 style={{ marginTop: 0, display: 'flex', alignItems: 'center', gap: 10 }}>
        <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><Mail size={26} /></span>
        {isDe ? 'Mail-Vorlagen' : 'Mail templates'}
      </h1>
      <p style={{ color: 'var(--dex-gray-600)', marginTop: 0 }}>
        {isDe
          ? 'Die globalen Standard-Mails der App (Anmeldung, Warteliste, Abmeldung …) — hier kannst du Betreff, Überschriften und Text pro Sprache anpassen. Änderungen gelten für ALLE Events, sofern ein Event keine eigene Anpassung hat.'
          : 'The global default emails of the app — edit subject, headings and body per language here. Changes apply to ALL events unless an event has its own override.'}
      </p>

      {/* Platzhalter-Hinweis */}
      <div style={{ background: 'rgba(134,188,37,0.08)', border: '1px solid var(--dex-green, #86bc25)', borderRadius: 10, padding: '10px 14px', margin: '8px 0 16px', fontSize: '0.82rem', color: 'var(--dex-gray-700)' }}>
        {isDe ? 'Verfügbare Platzhalter (werden beim Versand ersetzt): ' : 'Available placeholders (replaced when sending): '}
        {PLACEHOLDERS.map(p => (
          <code key={p} style={{ background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 5, padding: '1px 6px', margin: '0 4px 4px 0', display: 'inline-block', fontFamily: 'Consolas, monospace', fontSize: '0.78rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>{p}</code>
        ))}
        <div style={{ marginTop: 4, color: 'var(--dex-gray-500)' }}>
          {isDe ? 'Nicht jeder Platzhalter passt zu jeder Vorlage — nutze nur die, die zum Mail-Typ passen.' : 'Not every placeholder fits every template — use only those that match the mail type.'}
        </div>
      </div>

      {/* Suche */}
      <input
        type="text"
        value={search}
        onChange={e => setSearch(e.target.value)}
        placeholder={isDe ? 'Vorlage suchen (Name, Typ, Betreff) …' : 'Search template (name, type, subject) …'}
        style={{ ...inputStyle, maxWidth: 380, marginBottom: 16 }}
      />

      {loading ? (
        <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{isDe ? 'Vorlagen werden geladen …' : 'Loading templates …'}</p>
      ) : filtered.length === 0 ? (
        <p style={{ color: 'var(--dex-gray-400)' }}>{isDe ? 'Keine Vorlage gefunden.' : 'No template found.'}</p>
      ) : filtered.map(t => {
        const open = expandedId === t.id;
        const d = drafts[t.id] || { subject: t.subject, heading: t.heading, subheading: t.subheading, headingColor: t.headingColor, bodyHtml: t.bodyHtml };
        const dirty = isDirty(t);
        const previewHtml = wrapTemplate(d.headingColor || '#86bc25', d.heading || '', d.subheading || '', d.bodyHtml || '');
        return (
          <div key={t.id} style={cardStyle}>
            <button
              type="button"
              onClick={() => toggle(t)}
              style={{ width: '100%', textAlign: 'left', display: 'flex', alignItems: 'center', gap: 12, padding: '12px 16px', background: open ? 'var(--dex-gray-50, #fafafa)' : '#fff', border: 'none', cursor: 'pointer' }}
            >
              <span style={{ color: 'var(--dex-gray-400)', fontSize: '0.9rem' }}>{open ? '▾' : '▸'}</span>
              <span style={{ flex: 1, minWidth: 0 }}>
                <span style={{ display: 'block', fontWeight: 700, color: 'var(--dex-gray-800)' }}>{typeLabel(t.templateType)}</span>
                <span style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{t.subject || (isDe ? '(kein Betreff)' : '(no subject)')}</span>
              </span>
              <span style={{ flexShrink: 0, fontSize: '0.7rem', fontWeight: 700, padding: '2px 9px', borderRadius: 999, background: t.language === 'DE' ? 'rgba(0,118,168,0.10)' : 'rgba(134,188,37,0.16)', color: t.language === 'DE' ? 'var(--dex-blue, #0076a8)' : 'var(--dex-green-dark, #4a7c1f)' }}>{t.language}</span>
              {dirty && <span style={{ flexShrink: 0, fontSize: '0.68rem', fontWeight: 700, color: 'var(--dex-orange-dark, #b35a00)' }}>{isDe ? '• geändert' : '• edited'}</span>}
            </button>

            {open && (
              <div style={{ padding: '4px 16px 16px', borderTop: '1px solid var(--dex-gray-100)' }}>
                <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : 'minmax(280px, 1fr) minmax(280px, 1fr)', gap: 18, alignItems: 'start' }}>
                  {/* Editor */}
                  <div>
                    <div style={{ marginBottom: 10 }}>
                      <label style={labelStyle}>{isDe ? 'Betreff' : 'Subject'}</label>
                      <input style={inputStyle} value={d.subject} onChange={e => setField(t.id, 'subject', e.target.value)} />
                    </div>
                    <div style={{ display: 'flex', gap: 10, marginBottom: 10 }}>
                      <div style={{ flex: 1 }}>
                        <label style={labelStyle}>{isDe ? 'Überschrift' : 'Heading'}</label>
                        <input style={inputStyle} value={d.heading} onChange={e => setField(t.id, 'heading', e.target.value)} />
                      </div>
                      <div>
                        <label style={labelStyle}>{isDe ? 'Farbe' : 'Color'}</label>
                        <input type="color" value={/^#[0-9a-fA-F]{6}$/.test(d.headingColor) ? d.headingColor : '#86bc25'} onChange={e => setField(t.id, 'headingColor', e.target.value)} style={{ width: 44, height: 38, padding: 2, border: '1px solid var(--dex-gray-300)', borderRadius: 8, cursor: 'pointer', background: '#fff' }} />
                      </div>
                    </div>
                    <div style={{ marginBottom: 10 }}>
                      <label style={labelStyle}>{isDe ? 'Unter-Überschrift' : 'Subheading'}</label>
                      <input style={inputStyle} value={d.subheading} onChange={e => setField(t.id, 'subheading', e.target.value)} placeholder={isDe ? '(leer = Event-Titel)' : '(empty = event title)'} />
                    </div>
                    <div style={{ marginBottom: 12 }}>
                      <label style={labelStyle}>{isDe ? 'Text (HTML)' : 'Body (HTML)'}</label>
                      <textarea
                        value={d.bodyHtml}
                        onChange={e => setField(t.id, 'bodyHtml', e.target.value)}
                        spellCheck={false}
                        style={{ ...inputStyle, minHeight: 200, fontFamily: 'Consolas, monospace', fontSize: '0.8rem', lineHeight: 1.5, resize: 'vertical' }}
                      />
                    </div>
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button className="btn btn-primary" style={{ fontSize: '0.85rem', padding: '8px 18px' }} disabled={!dirty || savingId === t.id} onClick={() => { void save(t); }}>
                        {savingId === t.id ? (isDe ? 'Wird gespeichert …' : 'Saving …') : <><Check size={14} /> {isDe ? 'Speichern' : 'Save'}</>}
                      </button>
                      <button className="btn btn-secondary" style={{ fontSize: '0.85rem', padding: '8px 14px' }} disabled={!dirty || savingId === t.id} onClick={() => reset(t)}>
                        <X size={14} /> {isDe ? 'Verwerfen' : 'Discard'}
                      </button>
                    </div>
                  </div>

                  {/* Live-Vorschau */}
                  <div>
                    <label style={labelStyle}>{isDe ? 'Vorschau' : 'Preview'}</label>
                    <iframe
                      title={`preview-${t.id}`}
                      srcDoc={previewHtml}
                      style={{ width: '100%', height: 460, border: '1px solid var(--dex-gray-200)', borderRadius: 8, background: '#fff' }}
                    />
                  </div>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
