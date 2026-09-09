/**
 * Modul-Ebene der Anmeldeseite — Formatierer, Sanitizer und die
 * einklappbare Formular-Sektion.
 *
 * v30.66: aus `RegistrationPage.tsx` herausgezogen (Zeilen 48-251 des
 * urspruenglichen Stands). Der Inhalt ist zeichengleich uebernommen; einzige
 * Aenderung ist das vorangestellte `export` je Deklaration. Der Grund fuer die
 * eigene Datei ist der Modul-Zyklus: die ausgelagerten Teilbaeume der
 * Anmeldeseite brauchen diese Helfer und duerfen sie nicht aus der Seite
 * zurueckimportieren.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { isExternalEmail } from '../../utils/deloitteDomain';
import { ChevronDown } from '../Icons';
import { cx } from '../dexUi';

export function formatDate(iso: string): string {
  const d = new Date(iso);
  return (
    // v26.82: Wochentag (z.B. „Mi,") mit anzeigen.
    d.toLocaleDateString('de-DE', { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric' }) +
    ' ' +
    d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })
  );
}

// v11.94: Kompakt-Format für die Datum-Badge in der Event-Card —
// wenn Start- und End-Tag identisch sind, nur einmal das Datum +
// "HH:MM - HH:MM". Sonst beide voll mit "-" dazwischen.
export function formatDateRange(startIso: string, endIso: string): string {
  if (!startIso) return '';
  const start = new Date(startIso);
  const end = endIso ? new Date(endIso) : null;
  // v26.82: Wochentag (z.B. „Mi,") mit anzeigen.
  const dayFmt = { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric' } as const;
  const timeFmt = { hour: '2-digit', minute: '2-digit' } as const;
  if (!end || isNaN(end.getTime())) {
    return `${start.toLocaleDateString('de-DE', dayFmt)} ${start.toLocaleTimeString('de-DE', timeFmt)}`;
  }
  const sameDay = start.toDateString() === end.toDateString();
  if (sameDay) {
    return `${start.toLocaleDateString('de-DE', dayFmt)} ${start.toLocaleTimeString('de-DE', timeFmt)} – ${end.toLocaleTimeString('de-DE', timeFmt)}`;
  }
  return `${start.toLocaleDateString('de-DE', dayFmt)} ${start.toLocaleTimeString('de-DE', timeFmt)} – ${end.toLocaleDateString('de-DE', dayFmt)} ${end.toLocaleTimeString('de-DE', timeFmt)}`;
}

// v18.74/v27.11: Externe Adresse = kein Deloitte-Postfach. Seit v27.11 zählt
// JEDE Member-Firm-Domain als intern (@deloitte.at, @deloitte.com, …) — die
// International-Suche (v26.57) findet diese Kolleg:innen, also darf die
// Anmeldung sie nicht als extern behandeln.
export const isExternalEmailAddr = (e: string): boolean => isExternalEmail(e);

// v26.75: Die Vorfilter-Kategorie-Auswahl liegt transient unter dem Schlüssel
// '<fieldId>__cat' im Antwort-Store — sie ist reine UI-Hilfe zum Filtern der
// Optionsliste und wird NICHT als Antwort gespeichert.
export function stripPrefilterKeys(o: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const k of Object.keys(o || {})) { if (!k.endsWith('__cat')) out[k] = o[k]; }
  return out;
}

// v18.74: Strengere Plausibilitätsprüfung gegen Tippfehler bei externen
// Adressen — fängt fehlende/zu kurze TLD, doppelte Punkte, mehrere @, führende/
// abschließende Punkte und Whitespace/Kommas ab. Verifiziert NICHT die Existenz
// des Postfachs (das geht clientseitig nicht), aber blockt offensichtliche
// Vertipper.
export const isPlausibleEmail = (e: string): boolean => {
  const v = (e || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[a-zA-Z]{2,}$/.test(v)) return false;
  if ((v.match(/@/g) || []).length !== 1) return false;
  if (v.indexOf('..') >= 0) return false;
  if (/[\s,;]/.test(v)) return false;
  const [local, domain] = v.split('@');
  if (!local || !domain) return false;
  if (local.startsWith('.') || local.endsWith('.')) return false;
  if (domain.startsWith('.') || domain.endsWith('.') || domain.startsWith('-')) return false;
  return true;
};

// v26.91: Feld-Beschreibungen dürfen ein kleines Markdown-Subset tragen:
//   **fett**            → <strong>
//   [Text](https://…)   → Link (nur http/https/mailto)
//   nackte http(s)-URL  → automatisch verlinkt
// v26.96: Neu erfasste Beschreibungen kommen aus einem echten Rich-Text-Editor
// und sind bereits HTML. In diesem Fall wird das HTML (organizer-authored,
// sicherer Origin) nur von gefährlichen Teilen befreit und direkt ausgegeben.
// Alt-Beschreibungen (reiner Text / Markdown) gehen weiter durch das Subset.
export function renderFieldDescHtml(raw: string): string {
  if (!raw) return '';
  // Enthält der Text echte HTML-Tags (Rich-Text-Editor), NICHT escapen —
  // nur script/style/Event-Handler/javascript: entfernen.
  if (/<[a-z][\s\S]*>/i.test(raw)) {
    return raw
      .replace(/<\s*script[\s\S]*?<\s*\/\s*script\s*>/gi, '')
      .replace(/<\s*style[\s\S]*?<\s*\/\s*style\s*>/gi, '')
      .replace(/\son\w+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
      .replace(/javascript:/gi, '');
  }
  const linkStyle = 'color:var(--dex-green,#86bc25);font-weight:600;text-decoration:underline;';
  let html = raw
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  // [Text](url) — nur sichere Schemata
  html = html.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+|mailto:[^\s)]+)\)/g,
    (_m, label, url) => `<a href="${url}" target="_blank" rel="noopener noreferrer" style="${linkStyle}">${label}</a>`);
  // nackte http(s)-URLs (nicht die, die schon in einem href="…" stehen)
  html = html.replace(/(^|[\s(])(https?:\/\/[^\s<)]+)/g,
    (_m, pre, url) => `${pre}<a href="${url}" target="_blank" rel="noopener noreferrer" style="${linkStyle}">${url}</a>`);
  // **fett**
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  // Zeilenumbrüche erhalten
  html = html.replace(/\n/g, '<br />');
  return html;
}

// v29.27: Sub-Event-Beschreibungen kommen seit v28.89 aus demselben
// Rich-Text-Editor wie die Hauptevent-Beschreibung — sie können HTML und
// HTML-Entities tragen. Die Sub-Event-Karten renderten sie aber als ROHEN
// Text: ein „&nbsp;" aus dem Editor stand wörtlich auf der Anmeldeseite.
// Plain-Texte mit Entities werden zuerst dekodiert (BEWUSST ohne &lt;/&gt; —
// sonst könnte aus escaptem Text nachträglich Markup werden), dann geht
// alles durch renderFieldDescHtml: echtes HTML wird sanitisiert, reiner
// Text escaped + Markdown-Subset.
export function subEventDescHtml(raw: string): string {
  if (!raw) return '';
  const cleaned = /<[a-z][\s\S]*>/i.test(raw)
    ? raw
    : raw
      .replace(/&nbsp;|&#160;/gi, ' ')
      .replace(/&quot;/gi, '"')
      .replace(/&#39;|&apos;/gi, "'")
      .replace(/&amp;/gi, '&');
  return renderFieldDescHtml(cleaned);
}

/**
 * Einklappbare Formular-Sektion für die Handy-Ansicht.
 *
 * Auf dem Desktop (isMobile=false) wird EXAKT wie bisher gerendert: der
 * Header (mit optionalen Action-Buttons in `headerExtra`) plus der Body sind
 * immer sichtbar, ohne Chevron und ohne Toggle — das Verhalten bleibt
 * unverändert.
 *
 * Auf dem Handy (isMobile=true) wird der Header zu einer antippbaren Zeile mit
 * Chevron (▸/▾). Der Body ist per Default eingeklappt (`defaultOpen=false`),
 * damit die Anmeldemaske kompakt bleibt und man nicht ewig scrollen muss.
 * Etwaige Action-Buttons aus `headerExtra` bleiben auch eingeklappt bedienbar
 * (sie stehen weiter im Header, nur der reine Titel-Bereich toggelt).
 *
 * WICHTIG: Es werden keine Feldnamen, kein State und keine Validierung
 * verändert — nur die Sichtbarkeit des bereits gerenderten Bodys.
 */
export function CollapsibleSection(props: {
  isMobile: boolean;
  icon: string;
  title: React.ReactNode;
  /** Zusätzlicher Header-Inhalt rechts (z.B. Toggle-Buttons). */
  headerExtra?: React.ReactNode;
  defaultOpen?: boolean;
  /** v26.37: false = auf dem Handy NICHT einklappbar (immer sichtbar), z.B. für
   *  die eventspezifischen Felder, die der Teilnehmer aktiv ausfüllen muss. */
  collapsible?: boolean;
  children: React.ReactNode;
}): React.ReactElement {
  const { isMobile, icon, title, headerExtra, defaultOpen, collapsible, children } = props;
  const [open, setOpen] = React.useState<boolean>(defaultOpen ?? !isMobile);

  // v31.9: Symbol und Titel standen zweimal wortgleich im Markup — einmal je
  // Zweig. Der Handy-Zweig ist eine eigene Ansicht und driftet genau so ab
  // (Leitfaden 6b). Ein gemeinsames Fragment hält beide Zweige zusammen; das
  // gerenderte DOM ist unverändert.
  const headLabel = (
    <>
      <Icon iconName={icon} style={{ fontSize: 16 }} />
      {title}
    </>
  );
  // v31.9: Die Kopfzeile hielt ihre Zusatz-Aktionen mit `space-between` am
  // rechten Rand — genau das Muster, das der Leitfaden (2a′) verbietet: ein
  // Knopf, der zum Inhalt gehört, steht links beim Inhalt. Beide Zweige
  // nutzen dieselbe Zeile, damit sie nicht wieder auseinanderlaufen.
  const headRowStyle: React.CSSProperties = { display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' };

  // Desktop ODER explizit nicht-einklappbar: unverändertes Markup (Header + Body
  // immer sichtbar, kein Chevron).
  if (!isMobile || collapsible === false) {
    return (
      <>
        {headerExtra ? (
          <div style={headRowStyle}>
            <div className="section-header" style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              {headLabel}
            </div>
            {headerExtra}
          </div>
        ) : (
          <div className="section-header" style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
            {headLabel}
          </div>
        )}
        {children}
      </>
    );
  }

  // Handy: antippbarer Header + einklappbarer Body.
  return (
    <>
      <div style={headRowStyle}>
        <div
          className="section-header"
          role="button"
          tabIndex={0}
          aria-expanded={open}
          onClick={() => setOpen(o => !o)}
          onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setOpen(o => !o); } }}
          style={{ display: 'inline-flex', alignItems: 'center', gap: 8, cursor: 'pointer', flex: headerExtra ? '0 1 auto' : '1 1 auto', userSelect: 'none' }}
        >
          {/* v31.9: Aus „▸/▾" wird der Aufklapper-Chevron der App. Ein
              Text-Dreieck ist kein Bedienhinweis (Leitfaden 6c) und war auf
              dem Handy der einzige Hinweis darauf, dass sich hier etwas
              öffnet. `color: inherit`, weil die Kopfzeile weiß auf Grün ist —
              die Klasse färbt sonst grau. */}
          <span
            className={cx('dex-ui-disclosure-chevron', open && 'is-open')}
            aria-hidden="true"
            style={{ color: 'inherit' }}
          >
            <ChevronDown size={16} />
          </span>
          {headLabel}
        </div>
        {headerExtra}
      </div>
      {open && children}
    </>
  );
}

// v28.19: Ergebnis der Bildform-Analyse (Content-Ratio, v28.9) pro Bild-URL
// modulweit merken — beim erneuten Öffnen derselben Anmeldeseite steht das
// Kreis-Layout dann schon im ersten Render fest (kein Umspringen mehr).
export const IMG_ASPECT_CACHE: Record<string, number> = {};
