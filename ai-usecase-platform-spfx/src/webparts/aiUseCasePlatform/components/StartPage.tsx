/**
 * Die Start-Übersicht — was kommt nach „Start" auf der Landing Page.
 *
 * Aufbau wie `StartPage` in DEX: beschriftete Abschnitte mit großen Kacheln.
 * Aus „Aktuelle Events" wird „Use Cases", aus „Organizer" das „Use Case
 * Studio" (Nutzer-Ansage 29.09.2026). Die Kachelwand mit den Demos selbst
 * liegt eine Ebene tiefer (`UseCasesPage`).
 *
 * Wie in DEX sind die Menüpunkte DATEN, und Kachel (Desktop) und Zeile
 * (Handy) sind nur zwei Darstellungen derselben Liste. In DEX standen sie bis
 * v31.8 zweimal im Code und liefen auseinander — andere Untertitel, ein
 * fehlender Hinweis auf dem Handy. Zwei Quellen für dieselbe Aussage laufen
 * immer auseinander, nicht vielleicht.
 *
 * Der Punkt „Use Case Studio" steht für JEDEN da. Wer keine Organizer-Rechte
 * hat, sieht ihn ausgegraut mit „Organizer werden?" — der Weg zur Rolle ist
 * damit sichtbar, statt dass die Funktion für Nutzer schlicht nicht existiert
 * (DEX v12.5).
 */

import * as React from 'react';
import { ChevronRight, LayoutGrid, History, MessageSquare, Settings, Users } from './Icons';
import { cx, ensureDexUiStyles } from './dexUi';
import { useNavigation, Page } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { useHilfe } from '../context/HilfeContext';
import { useIsMobile } from '../utils/useIsMobile';

interface MenuItem {
  key: string;
  icon: (size: number, strokeWidth: number) => React.ReactNode;
  title: string;
  desc: string;
  ziel?: Page;
  onClick?: () => void;
  /** Ohne Rechte: ausgegraut, mit dem Weg zur Rolle statt einer Sackgasse. */
  anfrage?: boolean;
  cta?: string;
  /** Sichtbarer Grund, warum der Punkt gesperrt ist — auf BEIDEN Wegen. */
  note?: string | null;
  /** Zusatzklasse der Kachel (Farbakzent bzw. Symbol-Animation aus dem SCSS-Modul). */
  cardClass?: string;
}

export default function StartPage(): React.ReactElement {
  // Die Zeilen-Ansicht unten nutzt `dex-ui-`-Klassen; die Seite hängt weder an
  // einem Modal noch am Formular, die das Stylesheet sonst einziehen.
  ensureDexUiStyles();

  const isMobile = useIsMobile();
  const { navigate } = useNavigation();
  const { isAdmin, isOrganizer, isRolesLoading, rolesReadStatus, erstinstallation } = useRoles();
  const { t } = useLanguage();
  const { openKontakt } = useHilfe();

  // Ein 403 auf der Rollenliste heißt „nicht lesbar", nicht „kein Organizer".
  // Ohne den Satz sieht jemand, der Organizer IST, nur „Organizer werden?" —
  // und niemand weiß, warum die Kachel grau ist (DEX v30.81).
  //
  // Scheitert das Speichern des ersten Admin-Eintrags (frische Plattform),
  // wird niemand still Admin — und ohne diesen Satz sähe die Person nur eine
  // normale Oberfläche mit „Organizer werden?" und wüsste nicht, warum
  // (Review 29.09.2026). Sie ist ja gerade NICHT Admin und erreicht deshalb
  // keine Seite, die es sagen könnte.
  const rechteHinweis = erstinstallation === 'nicht-gespeichert'
    ? t('Die Erstinstallation konnte nicht gespeichert werden — du bist deshalb noch nicht als Admin eingetragen. Lade die Seite neu, um es erneut zu versuchen.',
      'The initial setup could not be saved — so you are not registered as admin yet. Reload the page to try again.')
    : rolesReadStatus === 'forbidden'
      ? t('Deine Rolle konnte nicht geprüft werden (fehlendes Leserecht). Bist du bereits Organizer? Dann bitte einen Admin, in der Rollenverwaltung „Rechte prüfen" auszuführen.',
        'Your role could not be checked (missing read access). Already an organizer? Ask an admin to run "Check rights" in role management.')
      : null;

  const itemUseCases: MenuItem = {
    key: 'usecases',
    icon: (s, w) => <LayoutGrid size={s} strokeWidth={w} />,
    title: 'Use Cases',
    desc: t('Demos ansehen und starten', 'View and start demos'),
    ziel: 'usecases',
  };

  const itemStudio: MenuItem = isOrganizer ? {
    key: 'studio', cardClass: 'start-card--admin',
    icon: (s, w) => <Settings size={s} strokeWidth={w} />,
    title: 'Use Case Studio',
    desc: t('Use Cases anlegen und pflegen', 'Create and maintain use cases'),
    ziel: 'studio',
  } : {
    key: 'studio', cardClass: 'start-card--admin',
    icon: (s, w) => <Settings size={s} strokeWidth={w} />,
    title: 'Use Case Studio',
    desc: t('Use Cases anlegen und pflegen', 'Create and maintain use cases'),
    // Während die Rollen noch laden, weiß niemand, ob die Person Organizer
    // ist: ausgegraut, aber ohne Angebot — sonst blitzt „Organizer werden?"
    // bei jedem Organizer kurz auf.
    anfrage: !isRolesLoading,
    cta: t('Organizer werden?', 'Become an organizer?'),
    note: rechteHinweis,
  };

  const itemProtokoll: MenuItem = {
    key: 'protokoll',
    icon: (s, w) => <History size={s} strokeWidth={w} />,
    title: t('Protokoll', 'Log'),
    desc: t('Wer hat was geändert', 'Who changed what'),
    ziel: 'protokoll',
  };

  const itemFragen: MenuItem = {
    key: 'fragen',
    icon: (s, w) => <MessageSquare size={s} strokeWidth={w} />,
    title: t('Fragen & Feedback', 'Questions & feedback'),
    desc: t('Frage stellen oder Use Case vorschlagen', 'Ask a question or suggest a use case'),
    onClick: () => openKontakt('frage'),
  };

  const itemRollen: MenuItem = {
    key: 'rollen',
    icon: (s, w) => <Users size={s} strokeWidth={w} />,
    title: t('Rollenverwaltung', 'Role management'),
    desc: t('Rollen vergeben und Rechte prüfen', 'Assign roles and check rights'),
    ziel: 'rollen',
  };

  const clusters: Array<{ key: string; title: string; items: MenuItem[] }> = [
    { key: 'entdecken', title: t('Entdecken', 'Discover'), items: [itemUseCases] },
    { key: 'organisation', title: t('Organisation', 'Organization'), items: [itemStudio, ...(isOrganizer ? [itemProtokoll] : [])] },
    { key: 'support', title: t('Support', 'Support'), items: [itemFragen] },
    { key: 'verwaltung', title: t('Verwaltung', 'Administration'), items: isAdmin ? [itemRollen] : [] },
  ].filter(c => c.items.length > 0);

  const oeffne = (it: MenuItem): void => {
    if (it.anfrage) { openKontakt('organizer'); return; }
    if (it.onClick) { it.onClick(); return; }
    if (it.ziel) navigate(it.ziel);
  };

  if (isMobile) {
    return (
      <div className="dex-start-rows" style={{ maxWidth: 520, margin: '0 auto', display: 'flex', flexDirection: 'column', gap: 20 }}>
        {clusters.map(c => (
          <div key={c.key}>
            <div className="dex-ui-section-title" style={{ margin: '0 4px 8px' }}>{c.title}</div>
            <div className="dex-ui-card dex-ui-card--list">
              {c.items.map((it, i) => (
                <button
                  key={it.key}
                  type="button"
                  className={cx('dex-ui-rowbtn', 'dex-ui-row')}
                  // Innenabstand und Trennlinie stehen bewusst hier: Der
                  // Knopf-Reset `dex-ui-rowbtn` steht im Stylesheet HINTER
                  // `dex-ui-row` und setzt `padding: 0; border: none` — die
                  // Klassen allein ergäben eine randlose, gequetschte Zeile.
                  style={{ minHeight: 58, padding: '10px 12px', borderTop: i === 0 ? undefined : '1px solid var(--dex-gray-100)' }}
                  onClick={() => oeffne(it)}
                >
                  <span style={{
                    flex: '0 0 auto', width: 42, height: 42, borderRadius: '50%',
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    background: 'rgba(134,188,37,0.12)', color: 'var(--dex-green-dark, #4a7c1f)',
                  }}>{it.icon(24, 1.6)}</span>
                  <span className="dex-ui-row-main">
                    {/* Titel und Untertitel brechen um, statt zu kürzen — was
                        hier abgeschnitten wird, steht auf dem Handy nirgends
                        sonst. */}
                    <span className="dex-ui-row-title dex-ui-row-title--wrap" style={{ display: 'block', fontSize: '0.98rem' }}>{it.title}</span>
                    <span className="dex-ui-row-sub" style={{ display: 'block', fontSize: '0.8rem' }}>{it.desc}</span>
                    {it.anfrage && it.cta && (
                      <span className="dex-ui-row-link" style={{ display: 'block', marginTop: 4, fontWeight: 600, fontSize: '0.8rem' }}>{it.cta}</span>
                    )}
                    {it.note && (
                      <span className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" style={{ display: 'flex', marginTop: 6 }}>{it.note}</span>
                    )}
                  </span>
                  <span style={{ flex: '0 0 auto', color: 'var(--dex-gray-400)', display: 'inline-flex', marginLeft: 2 }} aria-hidden="true">
                    <ChevronRight size={18} />
                  </span>
                </button>
              ))}
            </div>
          </div>
        ))}
      </div>
    );
  }

  return (
    <div>
      {/* Keyframes inline: SPFx hasht die Namen im .module.scss, und die
          Animation im :global-Block fände sie sonst nicht (DEX v9.36).
          Das Layout ist DEX v26: jeder Abschnitt in einer EIGENEN Zeile,
          darin gleich große, quadratische Kacheln — passen zwei nebeneinander,
          stehen sie in einer Zeile, sonst brechen sie um, statt
          unterschiedlich zu schrumpfen. */}
      <style>{`
        @keyframes dexStartIconBounce {
          0%   { transform: translateY(0) scale(1); }
          30%  { transform: translateY(-8px) scale(1.1); }
          60%  { transform: translateY(0) scale(1); }
          80%  { transform: translateY(-3px) scale(1.04); }
          100% { transform: translateY(0) scale(1); }
        }
        @keyframes dexStartIconWiggle {
          0%   { transform: rotate(0deg) scale(1); }
          20%  { transform: rotate(-10deg) scale(1.08); }
          40%  { transform: rotate(8deg) scale(1.08); }
          60%  { transform: rotate(-6deg) scale(1.05); }
          80%  { transform: rotate(4deg) scale(1.03); }
          100% { transform: rotate(0deg) scale(1); }
        }
        @keyframes dexStartIconSpin {
          0%   { transform: rotate(0deg) scale(1); }
          100% { transform: rotate(360deg) scale(1); }
        }
        .dex-cluster-grid { display: grid; grid-template-columns: 1fr; gap: 28px; max-width: 800px; margin: 0 auto; }
        .dex-cluster-title {
          font-size: 0.9rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px;
          color: var(--dex-gray-500); margin: 0 0 14px; padding-bottom: 8px;
          border-bottom: 1px solid var(--dex-gray-200);
        }
        .dex-cluster-tiles { display: flex; flex-wrap: wrap; gap: 22px; align-items: stretch; }
        .dex-cluster .start-card {
          flex: 0 0 auto !important; width: 380px !important; max-width: 100% !important;
          padding: 32px 22px !important; min-height: 0 !important; aspect-ratio: 1 / 1; gap: 10px !important;
        }
        /* Echter KREIS hinter dem Symbol: flex 0 0 auto + aspect-ratio verhindern,
           dass die Flex-Stauchung den Kreis zur Ellipse macht. */
        .dex-cluster .start-card__icon { width: 104px !important; height: 104px !important; flex: 0 0 auto !important; aspect-ratio: 1 / 1 !important; border-radius: 50% !important; margin-bottom: 6px !important; }
        .dex-cluster .start-card__icon svg { width: 54px !important; height: 54px !important; }
        .dex-cluster .start-card h2 { font-size: 1.32rem !important; }
        .dex-cluster .start-card p { font-size: 0.98rem !important; white-space: normal !important; }
        /* Laptop-Stufe (DEX v31.32): Windows skaliert Laptop-Displays meist auf
           125–150 %, es bleiben rund 1280 CSS-Pixel — zwei Kacheln à 380 px
           füllten dann fast die halbe Höhe. Bei 300 px passen drei nebeneinander.
           Steht VOR den 820/480-Stufen: Bei gleicher Spezifität gewinnt die
           letzte passende Regel. */
        @media (max-width: 1440px) {
          .dex-cluster-tiles { gap: 16px; }
          .dex-cluster .start-card { width: 300px !important; padding: 24px 18px !important; gap: 8px !important; }
          .dex-cluster .start-card__icon { width: 84px !important; height: 84px !important; margin-bottom: 2px !important; }
          .dex-cluster .start-card__icon svg { width: 44px !important; height: 44px !important; }
          .dex-cluster .start-card h2 { font-size: 1.14rem !important; }
          .dex-cluster .start-card p { font-size: 0.9rem !important; }
          .dex-cluster-title { margin-bottom: 10px; padding-bottom: 6px; }
        }
        @media (max-width: 820px) { .dex-cluster-grid { max-width: 420px; } }
      `}</style>
      <div className="dex-cluster-grid">
        {clusters.map(c => (
          <div key={c.key} className="dex-cluster">
            <div className="dex-cluster-title">{c.title}</div>
            <div className="dex-cluster-tiles">
              {c.items.map(it => (
                <div
                  key={it.key}
                  className={cx('card', !it.anfrage && 'card-clickable', 'start-card', it.cardClass)}
                  style={it.anfrage ? { position: 'relative', cursor: 'default', opacity: 0.55 } : (isRolesLoading && it.key === 'studio' && !isOrganizer ? { opacity: 0.55, cursor: 'default' } : undefined)}
                  onClick={it.anfrage ? undefined : () => oeffne(it)}
                  role={it.anfrage ? undefined : 'button'}
                  tabIndex={it.anfrage ? undefined : 0}
                  onKeyDown={it.anfrage ? undefined : e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); oeffne(it); } }}
                >
                  <div className="start-card__icon">{it.icon(64, 1)}</div>
                  <h2>{it.title}</h2>
                  <p>{it.desc}</p>
                  {it.anfrage && (
                    <div style={{ position: 'absolute', inset: 0, display: 'flex', flexDirection: 'column', gap: 8, alignItems: 'center', justifyContent: 'center', padding: 10 }}>
                      <button
                        type="button"
                        onClick={e => { e.stopPropagation(); oeffne(it); }}
                        style={{
                          background: 'var(--dex-green, #86bc25)', color: '#fff', border: 'none', borderRadius: 14,
                          padding: '8px 12px', fontWeight: 700, cursor: 'pointer', boxShadow: '0 4px 14px rgba(0,0,0,0.18)',
                          fontSize: '0.74rem', lineHeight: 1.25, textAlign: 'center', fontFamily: 'inherit',
                        }}
                      >{it.cta}</button>
                      {it.note && (
                        <div style={{ fontSize: '0.7rem', lineHeight: 1.3, textAlign: 'center', color: 'var(--dex-orange-dark, #b35a00)', background: 'rgba(255,255,255,0.92)', borderRadius: 8, padding: '6px 8px', maxWidth: 240 }}>
                          {it.note}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
