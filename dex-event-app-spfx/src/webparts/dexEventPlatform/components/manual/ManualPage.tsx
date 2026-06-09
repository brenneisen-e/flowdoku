/**
 * Handbuch-Seite (In-App-Manual)
 *
 * Layout: Sidebar mit Sektionsliste (gefiltert nach Rolle) + Content-Bereich
 * mit ausgewaehlter Sektion. Pro Sektion werden die Perspektiven (User /
 * Organizer / Admin) als Tabs dargestellt, innerhalb einer Perspektive die
 * nummerierten Schritte mit Text + Visual.
 */

import * as React from 'react';
import { useRoles } from '../../context/RoleContext';
import { useEvents } from '../../context/EventContext';
import { useCurrentUser } from '../../context/UserContext';
import { useLanguage } from '../../context/LanguageContext';
// v20.4: App-Modal statt window.prompt als Copy-Fallback.
import { useDialog } from '../../context/DialogContext';
import { ManualSection, ManualPerspectiveBlock, ManualStep, ManualPerspective } from './types';
import { PlaceholderShot } from './ManualMockups';
import { getManualSections } from './handbookContent';
import { APP_VERSION } from '../../version';

const PERSPECTIVE_LABEL: Record<ManualPerspective, { de: string; en: string }> = {
  user: { de: 'Als User:in', en: 'As User' },
  organizer: { de: 'Als Organizer:in', en: 'As Organizer' },
  admin: { de: 'Als Admin:in', en: 'As Admin' },
};

const CATEGORY_LABEL: Record<string, { de: string; en: string }> = {
  general: { de: 'Grundlagen', en: 'Basics' },
  organizer: { de: 'Für Organizer', en: 'For Organizers' },
  admin: { de: 'Administration', en: 'Administration' },
  architecture: { de: 'Architektur & Technik', en: 'Architecture & Tech' },
};

export default function ManualPage(): React.ReactElement {
  const { currentUserRole } = useRoles();
  const { events } = useEvents();
  const { currentUser } = useCurrentUser();
  const { locale, setLocale } = useLanguage();
  const isDe = locale === 'de';
  // v20.4: App-Modal statt window.prompt als Copy-Fallback.
  const { showAlert } = useDialog();

  // v13.9: User mit Check-In-Rechten pro Event (qrScannerEmails / organizerEmails
  // / coOrganizerEmails) bekommen im Handbuch die Organizer-Sicht freigeschaltet,
  // damit sie die Anleitungen für Check-In, Teilnehmer-Verwaltung etc. sehen —
  // auch ohne globale Organizer-Rolle in DEX_Roles.
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const hasPerEventCheckInRights = !!currentEmailLc && (events || []).some(e => {
    const qr = (e.qrScannerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
    if (qr) return true;
    const org = (e.organizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
    if (org) return true;
    return (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
  });
  const effectiveRole: 'User' | 'Organizer' | 'Admin' =
    currentUserRole === 'User' && hasPerEventCheckInRights ? 'Organizer' : (currentUserRole as 'User' | 'Organizer' | 'Admin');

  const allSections = React.useMemo(() => getManualSections(locale), [locale]);
  const visibleSections = React.useMemo(() => {
    return allSections.filter(s => s.visibleFor.indexOf(effectiveRole) >= 0);
  }, [allSections, effectiveRole]);

  // v6.23: Deep-Link-Support. Wenn die URL `?section=<id>` enthält (z.B.
  // ?action=manual&section=check-in), starten wir direkt in dieser Sektion —
  // sofern der User sie laut Rolle sehen darf. Andernfalls Default-Section.
  const [activeId, setActiveId] = React.useState<string>(() => {
    try {
      const params = new URLSearchParams(window.location.search);
      const sec = params.get('section');
      if (sec && visibleSections.some(s => s.id === sec)) return sec;
    } catch { /* URL-Parsing-Fehler ignorieren */ }
    return visibleSections[0]?.id || '';
  });
  const [search, setSearch] = React.useState('');
  const [copiedLink, setCopiedLink] = React.useState(false);
  const activeSection = visibleSections.find(s => s.id === activeId) || visibleSections[0];

  // Deep-Link zur aktuell offenen Sektion in die Zwischenablage kopieren.
  // Entfernt bestehende action/section/event-Parameter, fügt action=manual
  // + section=<activeId> sauber hinzu, damit die Ziel-URL immer auf diese
  // Sektion zeigt (unabhängig davon, mit welchem Deep-Link der User gerade
  // selbst hier ist).
  const copyDeepLink = React.useCallback((): void => {
    if (!activeSection) return;
    try {
      const url = new URL(window.location.href);
      url.searchParams.delete('event');
      url.searchParams.set('action', 'manual');
      url.searchParams.set('section', activeSection.id);
      const link = url.toString();
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(link).then(() => {
          setCopiedLink(true);
          setTimeout(() => setCopiedLink(false), 2000);
        }).catch(() => {
          showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{link}</span>, { title: isDe ? 'Link manuell kopieren' : 'Copy link manually' });
        });
      } else {
        showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{link}</span>, { title: isDe ? 'Link manuell kopieren' : 'Copy link manually' });
      }
    } catch { /* URL-Manipulation fehlgeschlagen — ignorieren */ }
  }, [activeSection?.id, isDe]);

  const [activePerspective, setActivePerspective] = React.useState<ManualPerspective | null>(null);
  React.useEffect(() => {
    if (activeSection) setActivePerspective(activeSection.perspectives[0]?.perspective || null);
  }, [activeSection?.id]);

  const grouped = React.useMemo(() => {
    const out: Record<string, ManualSection[]> = {};
    const q = search.trim().toLowerCase();
    for (const s of visibleSections) {
      if (q && s.title.toLowerCase().indexOf(q) < 0 && s.description.toLowerCase().indexOf(q) < 0) continue;
      (out[s.category] = out[s.category] || []).push(s);
    }
    return out;
  }, [visibleSections, search]);

  const label = (pair: { de: string; en: string }): string => isDe ? pair.de : pair.en;

  if (!activeSection) {
    return (
      <div style={{ padding: 40, textAlign: 'center', color: 'var(--dex-gray-600)' }}>
        {isDe ? 'Keine Handbuch-Sektionen verfügbar.' : 'No manual sections available.'}
      </div>
    );
  }

  return (
    <div style={{ display: 'grid', gridTemplateColumns: 'minmax(240px, 300px) 1fr', gap: 24, padding: '20px 24px', maxWidth: 1400, margin: '0 auto' }}>
      {/* Sidebar */}
      <aside style={{ position: 'sticky', top: 72, alignSelf: 'start', maxHeight: 'calc(100vh - 100px)', overflowY: 'auto' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 4 }}>
          <h2 style={{ fontSize: '1.2rem', fontWeight: 700, margin: 0 }}>
            {isDe ? 'Handbuch' : 'Handbook'}
          </h2>
          <div style={{ display: 'inline-flex', border: '1px solid var(--dex-gray-300)', borderRadius: 999, overflow: 'hidden', fontSize: '0.72rem' }}>
            <button
              onClick={() => setLocale('de')}
              style={{
                padding: '3px 10px', border: 'none',
                background: isDe ? 'var(--dex-green)' : '#fff',
                color: isDe ? '#fff' : 'var(--dex-gray-600)',
                fontWeight: isDe ? 600 : 400, cursor: 'pointer',
              }}
            >DE</button>
            <button
              onClick={() => setLocale('en')}
              style={{
                padding: '3px 10px', border: 'none',
                background: !isDe ? 'var(--dex-green)' : '#fff',
                color: !isDe ? '#fff' : 'var(--dex-gray-600)',
                fontWeight: !isDe ? 600 : 400, cursor: 'pointer',
              }}
            >EN</button>
          </div>
        </div>
        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>
          {isDe
            ? `Rolle: ${effectiveRole}${effectiveRole !== currentUserRole ? ' (Check-In-Helfer)' : ''} · ${visibleSections.length} Kapitel`
            : `Role: ${effectiveRole}${effectiveRole !== currentUserRole ? ' (check-in helper)' : ''} · ${visibleSections.length} chapters`}
        </div>
        <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', marginBottom: 14 }}>
          {isDe ? 'App-Version' : 'App version'} <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>{APP_VERSION}</strong>
        </div>
        <input
          className="form-input"
          value={search}
          onChange={e => setSearch(e.target.value)}
          placeholder={isDe ? 'Kapitel durchsuchen…' : 'Search chapters…'}
          style={{ marginBottom: 14 }}
        />
        {Object.keys(grouped).map(cat => (
          <div key={cat} style={{ marginBottom: 18 }}>
            <div style={{ fontSize: '0.72rem', fontWeight: 700, textTransform: 'uppercase', letterSpacing: 0.5, color: 'var(--dex-gray-500)', marginBottom: 6 }}>
              {label(CATEGORY_LABEL[cat] || { de: cat, en: cat })}
            </div>
            {grouped[cat].map(s => (
              <button
                key={s.id}
                onClick={() => setActiveId(s.id)}
                style={{
                  display: 'block', width: '100%', textAlign: 'left',
                  padding: '8px 12px', marginBottom: 3,
                  background: s.id === activeSection.id ? 'var(--dex-green)' : 'transparent',
                  color: s.id === activeSection.id ? '#fff' : 'var(--dex-gray-800)',
                  border: 'none', borderRadius: 6, cursor: 'pointer',
                  fontSize: '0.86rem', fontWeight: s.id === activeSection.id ? 600 : 400,
                  lineHeight: 1.35,
                }}
              >
                {s.title}
              </button>
            ))}
          </div>
        ))}
      </aside>

      {/* Content */}
      <section style={{ maxWidth: 920, minWidth: 0 }}>
        <header style={{ marginBottom: 18 }}>
          <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4 }}>
            {label(CATEGORY_LABEL[activeSection.category] || { de: activeSection.category, en: activeSection.category })}
          </div>
          <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 12 }}>
            <h1 style={{ fontSize: '1.8rem', fontWeight: 700, margin: 0 }}>{activeSection.title}</h1>
            {/* v6.23: Deep-Link zur aktuellen Sektion in die Zwischenablage */}
            <button
              type="button"
              onClick={copyDeepLink}
              title={isDe ? 'Direkt-Link zu dieser Sektion kopieren' : 'Copy direct link to this section'}
              style={{
                flexShrink: 0,
                background: copiedLink ? 'var(--dex-green)' : 'var(--dex-gray-100)',
                color: copiedLink ? '#fff' : 'var(--dex-gray-600)',
                border: '1px solid var(--dex-gray-200)',
                borderRadius: 8, padding: '6px 10px', cursor: 'pointer',
                fontSize: '0.78rem', fontWeight: 500,
                display: 'inline-flex', alignItems: 'center', gap: 6,
                whiteSpace: 'nowrap',
                transition: 'all 0.15s ease',
              }}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
              {copiedLink
                ? (isDe ? 'Kopiert!' : 'Copied!')
                : (isDe ? 'Link kopieren' : 'Copy link')}
            </button>
          </div>
          <p style={{ fontSize: '0.95rem', color: 'var(--dex-gray-600)', marginTop: 6, lineHeight: 1.5 }}>{activeSection.description}</p>
        </header>

        {activeSection.perspectives.length > 1 && (
          <div style={{ display: 'flex', gap: 4, borderBottom: '1px solid var(--dex-gray-200)', marginBottom: 20 }}>
            {activeSection.perspectives.map(p => (
              <button
                key={p.perspective}
                onClick={() => setActivePerspective(p.perspective)}
                style={{
                  padding: '8px 14px',
                  background: 'transparent', border: 'none',
                  borderBottom: activePerspective === p.perspective ? '2px solid var(--dex-green)' : '2px solid transparent',
                  fontSize: '0.88rem', fontWeight: activePerspective === p.perspective ? 600 : 400,
                  color: activePerspective === p.perspective ? 'var(--dex-green-dark)' : 'var(--dex-gray-600)',
                  cursor: 'pointer', marginBottom: -1,
                }}
              >
                {p.title || label(PERSPECTIVE_LABEL[p.perspective])}
              </button>
            ))}
          </div>
        )}

        {activeSection.perspectives
          .filter(p => !activePerspective || p.perspective === activePerspective)
          .map(p => (
            <PerspectiveView key={p.perspective} block={p} />
          ))}

        <footer style={{ marginTop: 40, padding: '14px 0', borderTop: '1px solid var(--dex-gray-200)', color: 'var(--dex-gray-400)', fontSize: '0.78rem', textAlign: 'center' }}>
          {isDe
            ? 'Handbuch v5.0 · Fragen zur Plattform? Rollen-Info im Admin Center.'
            : 'Handbook v5.0 · Questions? See the Admin Center for support.'}
        </footer>
      </section>
    </div>
  );
}

function PerspectiveView({ block }: { block: ManualPerspectiveBlock }): React.ReactElement {
  return (
    <div>
      {block.intro && (
        <div style={{ background: 'var(--dex-gray-50)', borderRadius: 8, padding: '14px 16px', marginBottom: 20, fontSize: '0.9rem', lineHeight: 1.6, color: 'var(--dex-gray-700)' }}>
          {block.intro}
        </div>
      )}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
        {block.steps.map(step => (
          <StepView key={step.number} step={step} />
        ))}
      </div>
    </div>
  );
}

function StepView({ step }: { step: ManualStep }): React.ReactElement {
  const visual = step.liveDemo || step.mockup
    || (step.screenshotPath ? <img src={step.screenshotPath} alt="" style={{ width: '100%', borderRadius: 8 }} /> : null)
    || (step.visualHint ? <PlaceholderShot caption={step.visualHint} /> : null);

  return (
    <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)', gap: 20, alignItems: 'start', padding: '18px 0', borderBottom: '1px solid var(--dex-gray-100)' }}>
      <div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 10 }}>
          <span style={{
            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
            width: 36, height: 36, borderRadius: '50%', background: 'var(--dex-green)',
            color: '#fff', fontWeight: 700, fontSize: '1rem', flexShrink: 0,
          }}>{step.number}</span>
          <h3 style={{ fontSize: '1.05rem', fontWeight: 600, margin: 0 }}>{step.title}</h3>
        </div>
        <div style={{ fontSize: '0.92rem', lineHeight: 1.6, color: 'var(--dex-gray-700)' }}>{step.description}</div>
        {step.tip && (
          <div style={{ marginTop: 12, padding: '8px 12px', background: 'rgba(255,193,7,0.08)', borderLeft: '3px solid #f0ad2e', borderRadius: 4, fontSize: '0.82rem', color: '#8a5d00' }}>
            💡 <strong>Tipp:</strong> {step.tip}
          </div>
        )}
        {step.warning && (
          <div style={{ marginTop: 12, padding: '8px 12px', background: 'rgba(200,30,30,0.06)', borderLeft: '3px solid #c9302c', borderRadius: 4, fontSize: '0.82rem', color: '#7a1b1b' }}>
            ⚠ <strong>Achtung:</strong> {step.warning}
          </div>
        )}
      </div>
      <div style={{ minWidth: 0 }}>
        {visual}
      </div>
    </div>
  );
}
