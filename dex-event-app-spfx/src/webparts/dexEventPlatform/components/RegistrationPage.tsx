/**
 * Registrierungsseite für ein einzelnes Event
 *
 * Drei-Spalten-Layout: Event-Info | persönliche Daten | eventspezifische Felder
 * Speichert die Registrierung in der SharePoint-Teilnehmerliste des Events.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents, collectCcEmailsFromFields } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
// v22.10: Sub-Sections nach ihrer EIGENEN Sichtbarkeit filtern (gleiche Logik
// wie die Event-Liste) — sonst sieht jeder Hauptevent-Teilnehmer alle Sub-Events.
import { isEventVisibleForUser } from './EventListPage';
import { useCachedImage, useCachedImageWithFallback } from '../utils/imageCache';
import { useIsMobile } from '../utils/useIsMobile';
import { isRegistrationFullyClosed } from '../utils/eventFormat';
import { selfCancelLocked } from '../utils/cancelPolicy';
import { isDeloitteInternalEmail, isExternalEmail } from '../utils/deloitteDomain';
import { useLanguage, translations as appTranslations, Locale } from '../context/LanguageContext';
// v20.4: modernes Alert-Modal statt window.alert.
import { useDialog } from '../context/DialogContext';
import { Salutation, EventSpecificField, DeloitteEvent } from '../types';
import { Icon } from '@fluentui/react/lib/Icon';
import { Send, X, Mail } from './Icons';
import { InfoTooltip } from './InfoTooltip';
import { MultiSelectDropdown } from './MultiSelectDropdown';
import OrganizerList from './OrganizerList';
import Modal from './Modal';
import InternationalSearchToggle from './InternationalSearchToggle';
import { UserFieldPicker } from './UserFieldPicker';
import StayRangePicker from './StayRangePicker';
// v28.95: Platzhalter fuer Events ohne eigenes Foto. Zuerst das im Admin
// Center unter „Logo & Branding" hinterlegte DEX-Orb (DefaultImageBase64 im
// _Config-Eintrag von DEX_EmailTemplates) — das ist die Stelle, an der es
// ausgetauscht wird, und dann soll der Tausch ueberall greifen. Das
// gebuendelte PNG ist nur der Rueckfall, solange der Cache noch nicht
// geladen ist (frischer Tab, erster Render).
import { DEX_ORB_PNG } from '../data/brandLogos';
import { getCachedOrbBase64 } from '../services/EmailTemplates';

function formatDate(iso: string): string {
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
function formatDateRange(startIso: string, endIso: string): string {
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
const isExternalEmailAddr = (e: string): boolean => isExternalEmail(e);

// v26.75: Die Vorfilter-Kategorie-Auswahl liegt transient unter dem Schlüssel
// '<fieldId>__cat' im Antwort-Store — sie ist reine UI-Hilfe zum Filtern der
// Optionsliste und wird NICHT als Antwort gespeichert.
function stripPrefilterKeys(o: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const k of Object.keys(o || {})) { if (!k.endsWith('__cat')) out[k] = o[k]; }
  return out;
}

// v18.74: Strengere Plausibilitätsprüfung gegen Tippfehler bei externen
// Adressen — fängt fehlende/zu kurze TLD, doppelte Punkte, mehrere @, führende/
// abschließende Punkte und Whitespace/Kommas ab. Verifiziert NICHT die Existenz
// des Postfachs (das geht clientseitig nicht), aber blockt offensichtliche
// Vertipper.
const isPlausibleEmail = (e: string): boolean => {
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
function renderFieldDescHtml(raw: string): string {
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
function subEventDescHtml(raw: string): string {
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
function CollapsibleSection(props: {
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

  // Desktop ODER explizit nicht-einklappbar: unverändertes Markup (Header + Body
  // immer sichtbar, kein Chevron).
  if (!isMobile || collapsible === false) {
    return (
      <>
        {headerExtra ? (
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
            <div className="section-header" style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Icon iconName={icon} style={{ fontSize: 16 }} />
              {title}
            </div>
            {headerExtra}
          </div>
        ) : (
          <div className="section-header" style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
            <Icon iconName={icon} style={{ fontSize: 16 }} />
            {title}
          </div>
        )}
        {children}
      </>
    );
  }

  // Handy: antippbarer Header + einklappbarer Body.
  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8, flexWrap: 'wrap' }}>
        <div
          className="section-header"
          role="button"
          tabIndex={0}
          aria-expanded={open}
          onClick={() => setOpen(o => !o)}
          onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setOpen(o => !o); } }}
          style={{ display: 'inline-flex', alignItems: 'center', gap: 8, cursor: 'pointer', flex: headerExtra ? '0 1 auto' : '1 1 auto', userSelect: 'none' }}
        >
          <span aria-hidden="true" style={{ fontSize: 12, width: 12, display: 'inline-block' }}>{open ? '▾' : '▸'}</span>
          <Icon iconName={icon} style={{ fontSize: 16 }} />
          {title}
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
const IMG_ASPECT_CACHE: Record<string, number> = {};

export default function RegistrationPage(): React.ReactElement {
  // v11.98: Beim Mount nach oben scrollen. Sonst behält der scrollende
  // .main-content-Container die Position aus der vorherigen Seite (z.B.
  // wenn man weit unten in der Events-Kachel war und dann auf Register
  // klickt — die Register-Page erscheint dann mittendrin statt am Anfang).
  React.useEffect(() => {
    const main = document.querySelector('.main-content');
    if (main && typeof (main as HTMLElement).scrollTo === 'function') {
      (main as HTMLElement).scrollTo({ top: 0, behavior: 'auto' });
    } else if (main) {
      (main as HTMLElement).scrollTop = 0;
    }
    if (typeof window !== 'undefined') {
      window.scrollTo({ top: 0, behavior: 'auto' });
    }
  }, []);

  const { selectedEventId, navigate, navIntent, clearIntent } = useNavigation();
  const { events, isEventsLoading, registerForEvent, registerTeam, cancelRegistration, declineEvent, checkRegistrationByEmail, getMyRegistration, getAllRegistrations, childEventsOf, listOpenTeamsForEvent, joinTeam, createTeamJoinRequest, updateMyRegistration, uploadFieldDocument, delegateRegistrationToAssistant, recordProxyDelegation, getLiveCounterStats, subscribeEventRealtime, getEventNumbersForEmail } = useEvents();
  const { currentUser, groupEmails } = useCurrentUser();
  const { searchUsers, searchUser, isAdmin } = useRoles();
  const { locale: appLocale } = useLanguage();
  // v20.4: App-Modal statt nativem Browser-Alert.
  const { showAlert, confirmDialog } = useDialog();
  // Mobile-Breakpoint für kompaktere Handy-Darstellung (einklappbare Sektionen etc.).
  const isMobile = useIsMobile();
  const event = events.find(e => e.id === selectedEventId);

  // v18.35: Erzwungene Anmeldesprache. Hat der Organizer für dieses Event eine
  // feste Anmeldesprache gesetzt ('de'/'en'), wird die GESAMTE Anmeldeseite
  // (App-Chrome, Form-Chrome, Inline-Texte UND Disclaimer) in dieser Sprache
  // angezeigt — unabhängig von der App-Sprache des Teilnehmers. Wir
  // überschreiben dazu `locale` und `t` lokal; alle bestehenden Verwendungen
  // im Rest der Datei greifen damit automatisch auf die erzwungene Sprache zu.
  const forcedRegLang: Locale | undefined =
    (event?.registrationLanguage === 'de' || event?.registrationLanguage === 'en') ? event.registrationLanguage : undefined;
  const locale: Locale = forcedRegLang || appLocale;
  const t = React.useCallback(
    (key: string): string => appTranslations[locale][key] || appTranslations['en'][key] || appTranslations['de'][key] || key,
    [locale]
  );

  // v11.56/v17.20/v19.18: tEvent() liefert die Form-Chrome-Strings (Placeholder,
  // Hints, Sub-Event-/Auswahl-Sektion).
  // v19.18 FIX: Form-Chrome folgt jetzt IMMER der App-/erzwungenen Anmeldesprache
  // (`locale`) — NICHT mehr der Event-MAIL-Sprache (`emailLanguage`). Vorher
  // entstand eine verwirrende Misch-Anzeige: der Großteil der Anmeldeseite in der
  // App-Sprache (z.B. DE), aber die Sub-Event-Auswahl + Platzhalter in der
  // Event-Mail-Sprache (z.B. EN bei einem B2Run-Event mit emailLanguage='EN').
  // Die Anmeldeseite zeigt der Person jetzt durchgängig EINE Sprache: die
  // erzwungene Anmeldesprache (falls per Event gesetzt), sonst die App-Sprache
  // des Teilnehmers. Die Mail-Sprache (`emailLanguage`) steuert weiterhin NUR die
  // tatsächlichen E-Mails — nicht die Formular-Anzeige. Der Bilingual-Toggle
  // steuert davon unberührt weiter die EN-Varianten der Custom-Field-Labels
  // (siehe `useEnVariants` unten).
  const eventLocale: Locale = locale;
  // v29.13: Die Anmeldeseite zeigt das Event-Bild aus Schritt 1. Es ist NICHT
  // dasselbe Bild wie das Mail-Logo aus dem Kommunikations-Schritt — Mails und
  // Outlook-Termin nehmen das, die Seite hier nicht. Wer nur eines von beiden
  // pflegt, pflegt meist das Mail-Logo (man sieht es sofort im Postfach) und
  // wundert sich, warum die Anmeldeseite den generischen DEX-Kreis zeigt.
  // Deshalb: kein Event-Bild → das Mail-Logo des Events übernehmen. Es bleibt
  // ein Rückfall; ist ein Event-Bild da, hat es immer Vorrang.
  const heroImgUrl = (event?.imageUrl || '') || (event?.mailImageBase64 || '');
  const usesMailImage = !event?.imageUrl && !!event?.mailImageBase64;
  // v19.22: Event-Bild über den IndexedDB-Cache (sofort beim zweiten Aufruf).
  // Base64 gehört nicht in den Cache — es liegt bereits vollständig vor.
  const cachedImage0 = useCachedImage(event?.imageUrl);
  const cachedImage = usesMailImage ? heroImgUrl : cachedImage0;
  // v28.11: Vergrößerte Hover-Ansicht des Event-Bilds — zeigt bevorzugt das
  // unbeschnittene Querformat-Original (falls vorhanden), sonst das Event-Bild.
  // v29.34: mit Rückfall auf das Event-Bild — die Original-URL kann ins Leere
  // zeigen (siehe useCachedImageWithFallback), die Lupe blieb dann leer.
  const cachedZoomImage0 = useCachedImageWithFallback(event?.imageOrigUrl, event?.imageUrl);
  const cachedZoomImage = usesMailImage ? heroImgUrl : cachedZoomImage0;
  // v28.12: Kein Auto-Zoom mehr beim Hover — der Hover zeigt nur ein
  // Lupen-Icon, erst der KLICK darauf öffnet die Großansicht (Lightbox).
  const [imgHovered, setImgHovered] = React.useState(false);
  const [imgZoomed, setImgZoomed] = React.useState(false);
  const tEvent = React.useCallback((key: string): string => {
    return appTranslations[eventLocale][key] || appTranslations['en'][key] || appTranslations['de'][key] || t(key) || key;
  }, [eventLocale, t]);
  // v17.20: Lookup-Helfer für die EN-Varianten eines Custom-Fields. Greift
  // nur, wenn der Bilingual-Toggle des Events an ist UND die App-Locale des
  // Teilnehmers `en` ist. Sonst still Fallback auf den DE-Wert. Index-Mapping
  // der Optionen ist positional — DE-Option i ↔ EN-Option i.
  // v17.22: `useEnVariants` steuert NUR die Anzeige-Labels. Die gespeicherten
  // Werte bleiben in JEDEM Fall die kanonischen DE-Originale: Single-Select
  // rendert `<option value={DE-Wert}>{EN-Anzeige}</option>`, Multi-Select gibt
  // `options={field.options}` (DE) als Wert weiter und nutzt `optionLabels`
  // nur für die Darstellung. Deshalb ist auch der „Register-for-Other"-Pfad
  // unkritisch: meldet ein EN-Organizer eine DE-Person an, sieht der Organizer
  // die EN-Labels (er füllt das Formular), gespeichert wird aber der
  // DE-Wert — die Zielperson und die Bestätigungs-Mail (event.emailLanguage)
  // bekommen also keine sprachlich falschen Daten.
  const useEnVariants = !!event?.bilingualFields && locale === 'en';
  const pickFieldLabel = React.useCallback((f: EventSpecificField): string =>
    (useEnVariants && f.labelEn && f.labelEn.trim()) ? f.labelEn : f.label,
  [useEnVariants]);
  const pickFieldHelp = React.useCallback((f: EventSpecificField): string | undefined =>
    (useEnVariants && f.helpTextEn && f.helpTextEn.trim()) ? f.helpTextEn : f.helpText,
  [useEnVariants]);
  const pickFieldConfirmLabel = React.useCallback((f: EventSpecificField): string | undefined =>
    (useEnVariants && f.confirmLabelEn && f.confirmLabelEn.trim()) ? f.confirmLabelEn : f.confirmLabel,
  [useEnVariants]);
  const pickOptionLabel = React.useCallback((f: EventSpecificField, optIdx: number, fallback: string): string => {
    if (useEnVariants && f.optionsEn && f.optionsEn[optIdx] && f.optionsEn[optIdx].trim()) {
      return f.optionsEn[optIdx];
    }
    return fallback;
  }, [useEnVariants]);

  // Per-Event-Organizer-Check: ist der eingeloggte User Haupt- ODER Co-Organizer
  // dieses Events? Nur dann darf er a) nach Deadline registrieren und b)
  // "Register for another person" nutzen. Ein Organizer von EVENT A darf NICHT
  // für EVENT B solche Admin-Aktionen ausführen. Admin darf global alles.
  // v19.6: Co-Organizer (event.coOrganizerEmails) zählen hier ausdrücklich
  // mit — vorher sah ein Co-Organizer den „Für andere registrieren"-Button
  // nicht, obwohl er das Event mitorganisiert. Serverseitig wird derselbe
  // Personenkreis in canRegisterForOthers() akzeptiert.
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isEventOrganizer = !!event && (
    event.organizerEmails.some(e => (e || '').toLowerCase() === currentEmailLc) ||
    (event.coOrganizerEmails || []).some(e => (e || '').toLowerCase() === currentEmailLc)
  );
  const isOrganizer = isEventOrganizer; // alten Namen behalten für Referenzen unten
  const canCreateEvents = isEventOrganizer || isAdmin; // statt tenant-weitem Organizer

  // Assistant-Ausnahme: User mit JobTitle "Assistant" / "Senior Assistant" dürfen
  // "Register for another person" nutzen, allerdings NUR für Director/Partner und
  // NUR für Events für die sie sich eh selber anmelden könnten (also nicht nach
  // Deadline). Der Deadline-Schutz greift automatisch, weil RegistrationPage für
  // normale User nach Deadline komplett die "closed"-Seite zeigt und gar nicht
  // zum Button-Rendering kommt.
  const currentJobTitleLc = (currentUser.jobTitle || '').toLowerCase();
  const isAssistant = currentJobTitleLc.includes('assistant');
  const ALLOWED_TARGET_TITLES = ['partner', 'director'];
  const isAllowedTargetForAssistant = (jt: string): boolean => {
    const lc = (jt || '').toLowerCase();
    return ALLOWED_TARGET_TITLES.some(t => lc === t || lc.indexOf(t) >= 0);
  };
  const canRegisterForOther = canCreateEvents || isAssistant;

  // Sichtbarkeits-Check: Würde dieses Event dem User als normaler User angezeigt werden?
  const showLocationBanner = canCreateEvents && event && (() => {
    const locFilters = event.locationAudience;
    // Audience-Filter normalisieren: 'All'/'DEALL' = "kein Audience-Filter"
    const audFilters = (event.audienceFilter || [])
      .map(s => s.trim())
      .filter(s => s && s.toLowerCase() !== 'all' && s.toLowerCase() !== 'deall');
    const hasLoc = locFilters.length > 0;
    const hasAud = audFilters.length > 0;
    if (!hasLoc && !hasAud) return false; // kein Filter = alle sehen es

    const loc = (currentUser.location || '').toLowerCase();
    const email = currentUser.email.toLowerCase();

    const locMatch = !hasLoc || locFilters.some(f => {
      const fl = f.trim().toLowerCase();
      if (fl === 'all') return true;
      const norm = fl.replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ä/g, 'ae');
      return loc.indexOf(fl) >= 0 || loc.indexOf(norm) >= 0;
    });

    const audMatch = !hasAud || audFilters.some(f => {
      const fl = f.trim().toLowerCase();
      if (fl.indexOf('@') >= 0) return email === fl;
      if (fl.startsWith('de')) {
        const city = fl.substring(2);
        const norm = city.replace(/oe/g, 'ö').replace(/ue/g, 'ü').replace(/ae/g, 'ä');
        return loc.indexOf(city) >= 0 || loc.indexOf(norm) >= 0;
      }
      return false;
    });

    // Default: AND (Schnittmenge). Nur OR wenn explizit gesetzt.
    // Wichtig: wenn nur EIN Filter gesetzt ist, zählt nur dieser - egal ob AND/OR.
    const mode = event.filterMode || 'AND';
    let visible: boolean;
    if (mode === 'OR') {
      if (hasLoc && hasAud) visible = locMatch || audMatch;
      else if (hasLoc) visible = locMatch;
      else visible = audMatch;
    } else {
      // AND
      if (hasLoc && hasAud) visible = locMatch && audMatch;
      else if (hasLoc) visible = locMatch;
      else visible = audMatch;
    }
    return !visible;
  })();

  const [salutation, setSalutation] = React.useState<Salutation | ''>('');
  const [firstName, setFirstName] = React.useState(currentUser.firstName);
  const [surname, setSurname] = React.useState(currentUser.surname);
  const [email, setEmail] = React.useState(currentUser.email);
  const [registerForOther, setRegisterForOther] = React.useState(false);
  // v24.41: „Meine Assistenz beauftragen" — Admins/Directoren können bei der
  // eigenen Anmeldung eine Assistenz angeben, die danach die Anmeldung in ihrer
  // „Assistenz"-Kachel verwaltet (Delegation) und auf CC der Bestätigung kommt.
  const [delegateAssistEnabled, setDelegateAssistEnabled] = React.useState(false);
  const [delegateAssistValue, setDelegateAssistValue] = React.useState('');
  // v18.74: „Person außerhalb Deloitte" — explizit eine externe Person
  // stellvertretend anmelden. Blendet den Deloitte-People-Picker aus und macht
  // Vorname/Nachname/E-Mail frei eintragbar. Die Zustimmung ist hier SCHRIFTLICH
  // einzuholen; es wird kein Outlook-Termin versendet (Bestätigungs-Mail mit
  // Organizer auf CC).
  const [externalPerson, setExternalPerson] = React.useState(false);

  // Wenn die Seite mit Intent 'register-other' geöffnet wird (z.B. via "Register another person"
  // Button auf einer Karte, für die der Organizer/Admin schon selbst registriert ist),
  // direkt in den "Für andere registrieren"-Modus springen und Felder leeren.
  React.useEffect(() => {
    if (navIntent === 'register-other' && (canCreateEvents)) {
      setRegisterForOther(true);
      setFirstName(''); setSurname(''); setEmail('');
      clearIntent();
    }
  }, [navIntent, canCreateEvents]);
  const [eventSpecific, setEventSpecific] = React.useState<Record<string, string>>({});
  // v19.0: Pro Dokument-Custom-Feld die ausgewählte Datei (vor dem Absenden).
  // Wird NICHT in customData geschrieben — nach erfolgreicher Anmeldung als
  // Attachment an die Teilnehmer-Zeile gehängt.
  const [pendingDocFiles, setPendingDocFiles] = React.useState<Record<string, File | null>>({});
  const [preferredStarterType, setPreferredStarterType] = React.useState<string>('');
  // Seit v6.5: Fallback-Dialog wenn B2Run-Wunschtyp voll, aber Alternative frei.
  const [fallbackDialog, setFallbackDialog] = React.useState<{ wunsch: string; alt: string; altFree: number } | null>(null);
  // v19.19: zusätzlich zu den aktiven Belegungen pro Gruppe (durch/fun) auch
  // die Wartelisten-Zahlen pro Gruppe (durchWait/funWait) — für die
  // Kapazitäts-/Warteliste-Anzeige in der Gruppen-Auswahl.
  const [starterCounts, setStarterCounts] = React.useState<{ durch: number; fun: number; durchWait: number; funWait: number } | null>(null);
  const [submitted, setSubmitted] = React.useState(false);
  // v18.11: „Ich nehme nicht teil"-Absage.
  const [declined, setDeclined] = React.useState(false);
  const [isDeclining, setIsDeclining] = React.useState(false);
  // v18.13/v18.14: Massenimport von Drittpersonen (nur Organizer/Admin im
  // „Für andere registrieren"-Modus). Zwei Schritte: (1) Liste einfügen →
  // gegen das Verzeichnis auflösen, (2) Vorschau-Tabelle prüfen → anmelden.
  const [massImportOpen, setMassImportOpen] = React.useState(false);
  const [massImportText, setMassImportText] = React.useState('');
  const [massImportMode, setMassImportMode] = React.useState<'mail' | 'nomail' | 'silent'>('mail');
  const [massImportStep, setMassImportStep] = React.useState<'input' | 'preview'>('input');
  const [massImportResolving, setMassImportResolving] = React.useState(false);
  const [massImportRows, setMassImportRows] = React.useState<Array<{
    email: string; firstName: string; lastName: string; jobTitle: string; location: string;
    status: 'ok' | 'duplicate' | 'notfound'; raw: string;
  }>>([]);
  const [massImportBusy, setMassImportBusy] = React.useState(false);
  const [massImportProgress, setMassImportProgress] = React.useState('');
  const [massImportResult, setMassImportResult] = React.useState<{ ok: number; failed: string[] } | null>(null);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  // v11.33: Submit-Overlay mit Fortschrittsanzeige (Prozent + Label).
  // Bei vielen Sub-Events / vielen Custom-Fields kann der Submit
  // mehrere Sekunden dauern — vorher hat der User nur einen disabled
  // Button gesehen ohne Feedback was gerade passiert.
  const [submitProgress, setSubmitProgress] = React.useState(0);
  const [submitProgressLabel, setSubmitProgressLabel] = React.useState('');
  // v29.29: Angezeigter Fortschritt. `submitProgress` ist der ZIELWERT, den die
  // Anmelde-Schritte setzen — er springt naturgemäß (5 → 30 → 50 → 100), weil
  // dazwischen einzelne, unterschiedlich lange SharePoint-Aufrufe liegen. Der
  // Balken lief dadurch in drei Sätzen statt gleichmäßig. Der Anzeigewert
  // nähert sich dem Ziel jetzt in kleinen Schritten und kriecht, solange ein
  // Schritt dauert, langsam weiter (gedeckelt vor der nächsten Stufe) — so
  // steht der Balken nie still und überholt den echten Stand auch nicht.
  const [displayProgress, setDisplayProgress] = React.useState(0);
  // Start und Ende eines Laufs setzen die Anzeige zurück. Das gehört in einen
  // EIGENEN Effekt: Der Animations-Effekt unten läuft bei jeder Ziel-Änderung
  // neu an und würde die Anzeige sonst mitten im Lauf auf 0 zurückwerfen.
  React.useEffect(() => { setDisplayProgress(0); }, [isSubmitting]);
  React.useEffect(() => {
    if (!isSubmitting) return undefined;
    let tick = 0;
    const id = window.setInterval(() => {
      tick++;
      setDisplayProgress(prev => {
        const target = Math.min(100, Math.max(0, submitProgress));
        // Abschluss: zügig auf 100 aufziehen.
        if (target >= 100) return Math.min(100, prev + Math.max(1, Math.ceil((100 - prev) / 4)));
        // Annäherung ans Ziel — je größer der Rückstand, desto größer der Schritt.
        if (prev < target) return Math.min(target, prev + Math.max(1, Math.ceil((target - prev) / 6)));
        // Ziel erreicht, der Schritt läuft noch: langsam weiterkriechen.
        //
        // Die Stufen liegen weit auseinander (30 → 50 → 95): Die Anmeldung des
        // Haupt-Events ist EIN langer Aufruf ohne Zwischenstand. Ein Deckel von
        // +8 (erste Fassung) ließ den Balken deshalb bei 38 stehen, bis am Ende
        // alles auf einmal kam — „38 % und dann direkt auf 100 %". Der Kriech-
        // Bereich deckt jetzt bis zu 25 Punkte ab und wird dabei immer
        // langsamer (erst alle ~0,25 s ein Prozent, später über eine Sekunde),
        // damit er die Lücke füllt, ohne dem echten Stand davonzulaufen.
        //
        // WICHTIG: Die Anzeige darf hier NICHT wieder auf den Zielwert
        // zurückgezogen werden. Genau das tat eine frühere Fassung („liegt die
        // Anzeige über dem Ziel, übernimm das Ziel") — der Kriech-Schritt ging
        // auf 31, die nächste Runde sprang zurück auf 30, und der Balken
        // flackerte 30/31/30/31. Ein Lauf startet ohnehin bei 0 (Effekt oben).
        const creepCap = Math.min(95, target + 25);
        if (prev >= creepCap) return prev;
        const every = 4 + (prev - target); // je weiter vorgekrochen, desto träger
        return (tick % Math.max(4, every) === 0) ? prev + 1 : prev;
      });
    }, 60);
    return () => window.clearInterval(id);
  }, [isSubmitting, submitProgress]);
  const [error, setError] = React.useState('');
  const [showErrors, setShowErrors] = React.useState(false);
  // v11.91: showDescription wurde entfernt — Beschreibung ist immer offen.
  const [thirdPartyCheck, setThirdPartyCheck] = React.useState<{ alreadyRegistered: boolean; notInAudience: boolean; registeredName?: string; registeredDate?: string } | null>(null);
  // v27.13: Profil-Karte („Persönliche Informationen") — Toggle für die
  // vollständige Liste der automatisch übernommenen Profildaten.
  // v28.1: standardmäßig AUSGEKLAPPT (Wunsch E.B.) — volle Transparenz ohne
  // Klick; über den Minus-Button weiterhin einklappbar.
  const [profileCardExpanded, setProfileCardExpanded] = React.useState(true);

  // Seit v6.14: integrierte Session-Auswahl direkt auf der Registrierungsseite.
  // Der User kann auf EINER Seite wählen, ob er sich für das Haupt-Event und/oder
  // einzelne Sub-Events anmelden möchte. Bei B2Run-Parents zusätzlich pro Session
  // eine Durchstarter/Funstarter-Auswahl.
  // v22.10 (Bugfix): Sub-Sections werden jetzt nach ihrer EIGENEN Sichtbarkeit
  // gefiltert. Vorher sah jeder, der das Hauptevent sehen konnte, ALLE
  // Sub-Events — auch wenn das Sub-Event einen eigenen Empfängerkreis hatte.
  // Organizer/Admins (und der „Für andere registrieren"-Modus) sehen weiterhin
  // alle Sub-Sections, damit sie stellvertretend buchen können.
  const childEvents = React.useMemo(() => {
    if (!event) return [];
    // v28.2: Sub-Events SOFT-deaktiviert (_subEventsDisabled) — für ALLE
    // ausblenden (auch Organizer/Stellvertreter: es soll niemand mehr auf
    // deaktivierte Sub-Events gebucht werden). Bestehende Anmeldungen
    // bleiben unberührt (MyEvents/Admin lesen die Kinder direkt).
    if (event.subEventsDisabled) return [];
    const all = childEventsOf(event.id);
    if (canCreateEvents || registerForOther) return all;
    // v22.68: Sub-Events im Entwurf (isFictive) sind für reguläre Teilnehmer
    // NICHT buchbar — vorher wurden sie nicht gefiltert und waren trotz
    // „Entwurf" buchbar, solange die Klammer sichtbar war. Organizer/
    // Stellvertreter (oben) sehen Entwürfe weiterhin.
    return all.filter(ce => !ce.isFictive && isEventVisibleForUser(ce, currentUser.email, currentUser.location, groupEmails, currentUser.jobTitle));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [event?.id, canCreateEvents, registerForOther, currentUser.email, currentUser.location, groupEmails]);
  /**
   * v29.9: Wie viele buchbare Sub-Events hat das Event, die dieser Person der
   * Zielgruppen-Filter oben WEGGENOMMEN hat?
   *
   * Das ist der Unterschied zwischen „es gibt keine" und „für dich ist keines
   * freigegeben" — und der ist wichtig: Ist die Klammer weiter gefasst als ihre
   * Sub-Events, kann jemand die Anmeldeseite öffnen und findet nichts zum
   * Anklicken. Die Meldung dazu behauptete bis v29.8, es sei „aktuell keines
   * angelegt". Das ist für die betroffene Person nachweislich falsch und
   * schickt sie mit der falschen Frage zu den Organizern.
   */
  const hiddenChildCount = React.useMemo(() => {
    if (!event || event.subEventsDisabled) return 0;
    if (canCreateEvents || registerForOther) return 0;
    const bookable = childEventsOf(event.id).filter(ce => !ce.isFictive);
    return Math.max(0, bookable.length - childEvents.length);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [event?.id, event?.subEventsDisabled, canCreateEvents, registerForOther, childEvents.length]);
  // v15.10: vom Organizer konfigurierbare Bezeichnung (z.B. „Event-Sections",
  // „Workshops"). Wenn gesetzt überschreibt das die Default-Übersetzung
  // („Sessions" / „Sub-Events") überall im RegistrationPage-UI.
  // v29.13: Besteht das Event ausschließlich aus Sub-Events, ist „Sub-Event"
  // die falsche Bezeichnung — es gibt kein Haupt-Event, unter dem sie hingen.
  // Für den Teilnehmer sind das schlicht DIE Events. Ohne eigenen Begriff des
  // Organizers wird deshalb hier der Default umgestellt; ein gesetzter eigener
  // Begriff hat weiterhin Vorrang. Weil damit alle Texte über diese beiden
  // Konstanten laufen, verschwindet „Sub-Event" in diesem Fall überall auf der
  // Anmeldeseite auf einmal — statt an einem Dutzend Einzelstellen.
  const subOnlyTerms = !!(event && event.subEventsOnlyMode);
  const childTermSingular = (event && event.childEventTermSingular)
    || (subOnlyTerms ? (locale === 'de' ? 'Event' : 'event') : '');
  const childTermPlural = (event && event.childEventTermPlural)
    || (subOnlyTerms ? (locale === 'de' ? 'Events' : 'events') : '');
  /**
   * v29.13: „ein Event" — aber „eine Session". Der unbestimmte Artikel hing
   * bisher fest an „eine …", was schon mit dem Default „Sub-Event" falsch war
   * („eine Sub-Event") und mit dem neuen Default „Event" auffällt. Wir raten
   * das Geschlecht nicht, sondern führen die wenigen femininen Begriffe, die
   * als Bezeichnung realistisch vorkommen; alles andere ist maskulin/neutral.
   */
  const childOneDe = React.useMemo(() => {
    const term = childTermSingular || 'Sub-Event';
    return /(session|veranstaltung|einheit|runde|reihe|tour|führung|schicht|woche|gruppe|stunde|session)$/i.test(term)
      ? `eine ${term}`
      : `ein ${term}`;
  }, [childTermSingular]);
  // v24.58: Anzeige-Präfix des Haupt-Events in der Sub-Event-Auswahl.
  // 'none' → kein Präfix (null), 'custom' → freier Text, sonst der mitgegebene
  // Default („Haupt-Event"/„Main event").
  const resolveMainEventLabel = React.useCallback((defaultLabel: string): string | null => {
    const mode = event && event.mainEventLabelMode;
    if (mode === 'none') return null;
    if (mode === 'custom' && event && event.mainEventLabel && event.mainEventLabel.trim()) return event.mainEventLabel.trim();
    return defaultLabel;
  }, [event]);
  // v24.73: Live-Plätze aus dem (für alle lesbaren) Sitzplatz-Counter. Die
  // Teilnehmerliste selbst ist item-level-gesichert — ein normaler Teilnehmer
  // sieht darüber NICHT die echte Gesamtzahl. Der Counter (aktiv = SeatsTaken,
  // Warteliste = WaitlistTaken) ist für alle lesbar und liefert die korrekten
  // Werte. Wird beim Öffnen + bei Fenster-Fokus leise nachgeladen (kein
  // sichtbares Nachladen — nur die Zahl ändert sich). Der Live-Push folgt in v24.74.
  const [liveStats, setLiveStats] = React.useState<{ active: number; waitlist: number } | null>(null);
  React.useEffect(() => {
    if (!event || !event.id || !(event.maxParticipants > 0)) { setLiveStats(null); return undefined; }
    let cancelled = false;
    const load = (): void => {
      getLiveCounterStats(event.id).then(s => { if (!cancelled && s) setLiveStats(s); }).catch(() => { /* best-effort */ });
    };
    load();
    const onFocus = (): void => load();
    window.addEventListener('focus', onFocus);
    // v24.75: Echtzeit-Push auf den (für alle lesbaren) Counter — bei jeder
    // An-/Abmeldung am Event meldet SharePoint die Counter-Änderung → der Wert
    // aktualisiert sich live, ohne Polling. Best-effort: klappt der Socket nicht,
    // bleibt der Lade-/Fokus-Refresh.
    let cleanupSocket: (() => void) | null = null;
    subscribeEventRealtime(event.id, 'counter', load)
      .then(c => { if (cancelled) c(); else cleanupSocket = c; })
      .catch(() => { /* best-effort */ });
    return () => {
      cancelled = true;
      window.removeEventListener('focus', onFocus);
      if (cleanupSocket) cleanupSocket();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [event?.id, event?.maxParticipants]);
  const [registerForParent, setRegisterForParent] = React.useState(true);
  const [selectedSessions, setSelectedSessions] = React.useState<Set<string>>(new Set());
  const [sessionStarterType, setSessionStarterType] = React.useState<Record<string, string>>({});
  // v10.12+: pro Sub-Event eigene Custom-Field-Werte. Wird beim Check eines
  // Sub-Events in dem Pop-Up-Modal abgefragt (siehe pendingSubEventModal weiter
  // unten) und beim Submit pro Sub-Event-Registrierung an registerForEvent
  // weitergereicht.
  const [sessionFieldValues, setSessionFieldValues] = React.useState<Record<string, Record<string, string>>>({});
  // Modal-State: wenn ein Sub-Event angecheckt wird das Custom-Fields hat,
  // wird hier die ID gemerkt + ein Draft der Field-Werte. Beim „Bestätigen"
  // wandern die Werte in sessionFieldValues + die Session in selectedSessions.
  // Beim „Abbrechen" wird der Modal geschlossen und die Session NICHT angecheckt.
  const [pendingSubEventModal, setPendingSubEventModal] = React.useState<{
    subEventId: string;
    draftValues: Record<string, string>;
  } | null>(null);
  const [sessionMeta, setSessionMeta] = React.useState<Record<string, { count: number; wasRegistered: boolean }>>({});
  const [myParentReg, setMyParentReg] = React.useState<{ Status?: string } | null>(null);
  const [sessionsOnlySubmitted, setSessionsOnlySubmitted] = React.useState(false);
  // v18.67: echtes Anmelde-Ergebnis (Angemeldet/Warteliste) aus der
  // Haupt-Registrierung — das Ergebnis-Modal nutzt das statt der gecachten
  // isFull-Schätzung, die nach Cancel/Re-Register veraltet sein konnte und
  // fälschlich „Warteliste" zeigte, obwohl der User angemeldet wurde.
  const [submittedAsWaitlist, setSubmittedAsWaitlist] = React.useState(false);

  // v11.82: Team-Anmeldung — UI-State.
  // - isTeamMode: User hat den Toggle „Ich melde mich + mein Team an" angehakt.
  // - teamName: optionaler Team-Name (nur sichtbar wenn event.askTeamName).
  // - teamMembers: N-1 People-Picker-Slots, jeder „<DisplayName> <email>".
  // - teamConsentConfirmed: Pflicht-Checkbox „alle Mitglieder haben zugestimmt".
  // v14.5: Wenn der Organizer `requireSubEventSelection` aktiviert hat, ist
  // Team-Anmeldung nicht kombinierbar — der Team-Flow registriert nur fürs
  // Hauptevent (keine Sub-Event-Auswahl möglich), würde also entweder am
  // Submit-Gate scheitern (verwirrend) oder das Event-Setup unterlaufen.
  // Deshalb hier den Toggle hart ausblenden, damit die Inkonsistenz gar
  // nicht erst entsteht.
  const isTeamCapable = !!event?.teamRegistrationEnabled
    && (event?.teamSize || 0) >= 2
    // v22.78: Wenn Teilnehmer keine neuen Teams erstellen dürfen, wird der
    // „Ich melde mich + mein Team an"-Toggle ausgeblendet (Organizer ordnet zu).
    && !event?.teamMembersCannotCreate
    && !(event?.requireSubEventSelection && childEvents.length > 0);
  const teamSize = event?.teamSize || 0;
  const teamPartialAllowed = !!event?.teamPartialAllowed;
  const [isTeamMode, setIsTeamMode] = React.useState(false);
  const [teamName, setTeamName] = React.useState('');
  const [teamMembers, setTeamMembers] = React.useState<string[]>([]);
  // v18.12: Custom-Field-Antworten pro Team-Mitglied (Slot-Index → {fieldId: value}).
  // So kann der Lead z.B. die Essenspräferenz auch für jedes Teammitglied angeben.
  const [teamMemberFields, setTeamMemberFields] = React.useState<Record<number, Record<string, string>>>({});
  const [teamConsentConfirmed, setTeamConsentConfirmed] = React.useState(false);
  // v15.16: Bei „Für andere registrieren" (registerForOther) braucht es
  // ebenfalls eine explizite Bestätigung, dass die Person der Anmeldung
  // zugestimmt hat — analog zur Team-Anmelde-Pflicht.
  const [otherConsentConfirmed, setOtherConsentConfirmed] = React.useState(false);
  // v26.76: Geführter Wizard für die stellvertretende Anmeldung (interner Fall):
  // 0 = geschlossen, 1 = Person suchen, 2 = Zustimmung. Nach „OK" ist die Person
  // übernommen und die persönlichen Felder vorbefüllt.
  const [proxyStep, setProxyStep] = React.useState<0 | 1 | 2>(0);
  // v27.6: Der frühere INLINE-Ablauf für „Für andere Person anmelden"
  // (Personensuche + Extern-Umschalter + große Zustimmungs-Box direkt im
  // Anmeldeformular) ist vollständig in den geführten Wizard-Modal gewandert.
  // Auf dem Hauptformular steht danach nur eine kompakte Zusammenfassung +
  // die grauen (read-only) Felder. Dieses Flag lässt die Alt-UI dauerhaft aus,
  // bewahrt sie aber als Referenz (kein Setter → immer false).
  const [showInlineProxyPicker] = React.useState(false);
  // v11.83: Offene Teams (Slots-frei) + Beitritts-Flow.
  const [openTeams, setOpenTeams] = React.useState<Array<{ teamId: string; teamName: string; activeCount: number; teamSize: number; leadEmail: string; leadDisplayName: string }>>([]);
  const [openTeamsLoaded, setOpenTeamsLoaded] = React.useState(false);
  // v18.73: Beitritt zu einem offenen Team wird nur VORGEMERKT — die eigentliche
  // Anmeldung (inkl. der ausgefüllten persönlichen + event-spezifischen Felder)
  // passiert erst beim Klick auf „Anmelden" unten (performJoinSelectedTeam).
  // Vorher wurde der Beitritt sofort beim Klick committet, ohne dass der User
  // seine event-spezifischen Infos angeben konnte.
  const [pendingJoinTeam, setPendingJoinTeam] = React.useState<{ teamId: string; teamName: string } | null>(null);
  // v18.73: Erfolgsscreen-Variante bei Team-Beitritt ('joined' = direkt
  // angemeldet, 'requested' = Anfrage an den Team-Lead gesendet).
  const [submittedJoinKind, setSubmittedJoinKind] = React.useState<null | 'joined' | 'requested'>(null);
  // Beim Aktivieren des Team-Modus: Member-Slots initialisieren (teamSize-1 Slots).
  React.useEffect(() => {
    if (isTeamMode && teamMembers.length !== Math.max(0, teamSize - 1)) {
      setTeamMembers(Array.from({ length: Math.max(0, teamSize - 1) }, () => ''));
    }
    if (!isTeamMode) {
      setTeamMembers([]);
      setTeamName('');
      setTeamConsentConfirmed(false);
      setTeamMemberFields({});
    }
  }, [isTeamMode, teamSize]);
  // Parser für People-Picker-Values im Format „DisplayName <email>".
  const parseTeamMember = (v: string): { displayName: string; email: string } | null => {
    const m = (v || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
    if (!m) return null;
    return { displayName: m[1].trim(), email: m[2].trim().toLowerCase() };
  };
  const teamMembersParsed = teamMembers.map(parseTeamMember);
  // v18.12: Custom-Fields, die pro Team-Mitglied abgefragt werden — alle
  // event-spezifischen Felder AUSSER Personen-Pickern (user/roommate) und
  // B2Run-Spezialfeldern; gruppen-spezifische Felder nur „für alle".
  const teamMemberApplicableFields = (event?.eventSpecificFields || []).filter(f =>
    f.type !== 'user' && f.type !== 'roommate' &&
    f.id !== 'b2run_startblock' && f.id !== 'b2run_mobilnummer' &&
    (!f.onlyForGroup || f.onlyForGroup === 'all')
  );
  // Validation des Team-Submits — Lead-Email darf nicht in der Member-Liste
  // sein, Member-Emails müssen untereinander disjunkt sein, im Pflicht-Modus
  // müssen alle Slots gefüllt sein.
  const teamValidation = ((): { ok: boolean; reason?: string } => {
    if (!isTeamMode) return { ok: true };
    const leadEmail = (email || '').trim().toLowerCase();
    const filled = teamMembersParsed.filter(m => !!m);
    if (!teamPartialAllowed && filled.length < (teamSize - 1)) {
      return { ok: false, reason: locale === 'de' ? `Bitte alle ${teamSize - 1} Team-Mitglieder auswählen.` : `Please pick all ${teamSize - 1} team members.` };
    }
    const seen = new Set<string>();
    if (leadEmail) seen.add(leadEmail);
    for (const m of filled) {
      if (!m) continue;
      if (seen.has(m.email)) {
        return { ok: false, reason: locale === 'de' ? `„${m.displayName}" ist doppelt im Team.` : `„${m.displayName}" appears twice in the team.` };
      }
      seen.add(m.email);
    }
    if (event?.askTeamName && !teamName.trim()) {
      return { ok: false, reason: locale === 'de' ? 'Bitte Team-Name angeben.' : 'Please enter a team name.' };
    }
    if (teamName.trim().length > 60) {
      return { ok: false, reason: locale === 'de' ? 'Team-Name max. 60 Zeichen.' : 'Team name must be 60 characters or fewer.' };
    }
    if (!teamConsentConfirmed) {
      return { ok: false, reason: locale === 'de' ? 'Bitte bestätige, dass alle Mitglieder zugestimmt haben.' : 'Please confirm that all members have consented.' };
    }
    return { ok: true };
  })();

  // v11.83: Offene Teams nachladen — sobald wir wissen, dass das Event
  // Team-Anmeldung erlaubt UND der Organizer „Offene Slots oeffentlich
  // sichtbar" aktiviert hat UND der User selbst noch nicht angemeldet
  // ist. Lazy: einmal pro Event-Wechsel.
  React.useEffect(() => {
    setOpenTeamsLoaded(false);
    setOpenTeams([]);
    setPendingJoinTeam(null); // v18.73: Vormerkung beim Event-/Modus-Wechsel zurücksetzen
    if (!event) return;
    if (!event.teamRegistrationEnabled || !event.teamOpenSlotsVisible) return;
    if (registerForOther) return; // Stellvertreter-Modus nicht unterstützt für Beitritt
    (async () => {
      try {
        const list = await listOpenTeamsForEvent(event.id);
        setOpenTeams(list);
      } catch {
        setOpenTeams([]);
      } finally {
        setOpenTeamsLoaded(true);
      }
    })().catch(() => setOpenTeamsLoaded(true));
  }, [event?.id, event?.teamRegistrationEnabled, event?.teamOpenSlotsVisible, registerForOther, listOpenTeamsForEvent]);

  // v18.73: Team-Beitritt nur VORMERKEN (Toggle). Erneuter Klick auf dasselbe
  // Team hebt die Vormerkung wieder auf. Gegenseitig exklusiv zum „Ich melde
  // mich + mein Team an"-Modus (man kann nicht gleichzeitig ein neues Team
  // anlegen und einem bestehenden beitreten). Die eigentliche Anmeldung läuft
  // erst über den „Anmelden"-Button (performJoinSelectedTeam).
  const togglePendingJoinTeam = (teamId: string, teamName: string): void => {
    setError('');
    setPendingJoinTeam(prev => (prev && prev.teamId === teamId) ? null : { teamId, teamName });
    setIsTeamMode(false);
  };

  // Vorbelegen: Parent-Reg prüfen + Sessions-Meta laden (bereits-registrierte
  // Sessions werden als angehakt voreingestellt).
  React.useEffect(() => {
    if (!event) return;
    // Parent-Reg-Vorbelegung nur im Selbst-Modus — sie beruht auf
    // getMyRegistration des eingeloggten Users und ist im Stellvertreter-
    // Modus bedeutungslos.
    if (!registerForOther) {
      (async () => {
        try {
          const r = await getMyRegistration(event.id) as { Status?: string; StarterType?: string; PreferredStarterType?: string } | null;
          setMyParentReg(r);
          if (r && r.Status !== 'Abgemeldet') {
            setRegisterForParent(false);
            // v11.10: Bei bereits angemeldetem Parent den existierenden
            // Starter-Typ in die Group-Selection vorladen, damit auch im
            // Sessions-Only-Modus eine Gruppe sichtbar gewählt ist und
            // Sub-Events sauber davon erben.
            const existing = r.StarterType || r.PreferredStarterType;
            if (existing && (existing === 'Durchstarter' || existing === 'Funstarter')) {
              setPreferredStarterType(existing);
            }
          }
        } catch { /* */ }
      })();
    }
    if (childEvents.length > 0) {
      (async () => {
        try {
          const meta: Record<string, { count: number; wasRegistered: boolean }> = {};
          const preselect = new Set<string>();
          const starterPre: Record<string, string> = {};
          for (const ce of childEvents) {
            if (registerForOther) {
              // v18.37: Stellvertreter-Modus — nur die Belegungszahl laden
              // (für die „X/Y belegt"-/Voll-Anzeige). KEINE Self-Vorbelegung,
              // weil getMyRegistration die Daten des eingeloggten Users liefert,
              // nicht die der angemeldeten Person. Der Assistent wählt die
              // Sub-Events frisch aus.
              const allRegs = await getAllRegistrations(ce.id);
              const count = (allRegs || []).filter(r => {
                const s = r.Status || '';
                return s === 'Angemeldet' || s === 'QR versendet' || s === 'Eingecheckt';
              }).length;
              meta[ce.id] = { count, wasRegistered: false };
            } else {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const [myReg, allRegs] = await Promise.all([
                getMyRegistration(ce.id) as Promise<{ Status?: string; StarterType?: string; PreferredStarterType?: string } | null>,
                getAllRegistrations(ce.id),
              ]);
              const wasRegistered = !!myReg && myReg.Status !== 'Abgemeldet';
              const count = (allRegs || []).filter(r => {
                const s = r.Status || '';
                return s === 'Angemeldet' || s === 'QR versendet' || s === 'Eingecheckt';
              }).length;
              meta[ce.id] = { count, wasRegistered };
              if (wasRegistered) {
                preselect.add(ce.id);
                const existingType = myReg?.StarterType || myReg?.PreferredStarterType;
                if (existingType) starterPre[ce.id] = existingType;
              }
            }
          }
          setSessionMeta(meta);
          // Im Stellvertreter-Modus startet die Auswahl leer (frische Anmeldung).
          setSelectedSessions(registerForOther ? new Set<string>() : preselect);
          setSessionStarterType(prev => ({ ...starterPre, ...prev }));
        } catch { /* */ }
      })();
    }
  }, [event?.id, registerForOther]);

  // v28.4: Seitenverhältnis des Event-Bildes erkennen — Querformat-Fotos
  // bekommen im „Geführte Schritte"-Layout einen BREITEREN Bild-Slot (420px
  // statt 300px), damit sie nicht winzig in der Ecke hängen; Hochkant/
  // Quadrat bleibt beim kompakten 300er-Slot. Das Bild sitzt vertikal
  // mittig neben den Infos (kein toter Leerraum mehr unter dem Foto).
  // v28.19: Kein Layout-Umspringen mehr — die Analyse läuft asynchron, daher:
  // (a) Ergebnis pro URL im Modul-Cache (IMG_ASPECT_CACHE), damit der zweite
  //     Besuch synchron im ersten Render die richtige Form kennt, und
  // (b) `imgAspectReady` als Gate: Der Bild-Slot wird erst gerendert, wenn
  //     die Form feststeht (oder die Analyse fehlschlug) — das Bild erscheint
  //     dann direkt an der richtigen Stelle statt kurz rechts zu starten.
  const [imgProbe, setImgProbe] = React.useState<{ url: string; ratio: number | null } | null>(null);
  const imgAspectCached = heroImgUrl ? IMG_ASPECT_CACHE[heroImgUrl] : undefined;
  const imgAspect: number | null = imgAspectCached !== undefined
    ? imgAspectCached
    : (imgProbe && imgProbe.url === heroImgUrl ? imgProbe.ratio : null);
  const imgAspectReady = imgAspectCached !== undefined
    || (!!imgProbe && imgProbe.url === heroImgUrl);
  React.useEffect(() => {
    if (!heroImgUrl) return undefined;
    if (IMG_ASPECT_CACHE[heroImgUrl] !== undefined) return undefined;
    let cancelled = false;
    const probeUrl = heroImgUrl;
    const img = new Image();
    img.onload = () => {
      if (cancelled || img.naturalHeight <= 0) return;
      const fileRatio = img.naturalWidth / img.naturalHeight;
      let ratio = fileRatio;
      // v28.9: CONTENT-Ratio statt reiner Datei-Ratio. Logos/Kreis-Grafiken
      // liegen oft mit transparentem oder einfarbigem Rand in einer breiten
      // Datei — die Datei-Ratio sortierte sie als „Querformat" ein und das
      // Kreis-Layout (v28.7) griff nie. Wir rastern das Bild klein, prüfen
      // ob die vier Ecken einen einheitlichen Rand bilden (transparent oder
      // eine Farbe), trimmen diesen Rand und nehmen das Seitenverhältnis
      // des sichtbaren Inhalts. Randlose Fotos (uneinige Ecken) und
      // Canvas-Fehler behalten die Datei-Ratio.
      try {
        const W = 96;
        const H = Math.max(1, Math.round(W / fileRatio));
        const canvas = document.createElement('canvas');
        canvas.width = W; canvas.height = H;
        const ctx = canvas.getContext('2d');
        if (ctx) {
          ctx.drawImage(img, 0, 0, W, H);
          const data = ctx.getImageData(0, 0, W, H).data;
          const px = (x: number, y: number): number[] => {
            const i = (y * W + x) * 4;
            return [data[i], data[i + 1], data[i + 2], data[i + 3]];
          };
          const corners = [px(0, 0), px(W - 1, 0), px(0, H - 1), px(W - 1, H - 1)];
          const dist = (a: number[], b: number[]): number =>
            Math.max(Math.abs(a[0] - b[0]), Math.abs(a[1] - b[1]), Math.abs(a[2] - b[2]));
          const allTransparent = corners.every(c => c[3] <= 24);
          const uniform = allTransparent || corners.every(c => c[3] > 24 && dist(c, corners[0]) <= 20);
          if (uniform) {
            const bg = corners[0];
            let minX = W, minY = H, maxX = -1, maxY = -1;
            for (let y = 0; y < H; y++) {
              for (let x = 0; x < W; x++) {
                const p = px(x, y);
                const isContent = p[3] > 24 && (allTransparent || dist(p, bg) > 28);
                if (isContent) {
                  if (x < minX) minX = x;
                  if (x > maxX) maxX = x;
                  if (y < minY) minY = y;
                  if (y > maxY) maxY = y;
                }
              }
            }
            if (maxX >= minX && maxY >= minY) {
              const bw = maxX - minX + 1;
              const bh = maxY - minY + 1;
              // Nur übernehmen, wenn wirklich Rand weggefallen ist — sonst
              // ist die Datei-Ratio genauer (voller Inhalt bis zur Kante).
              if (bh > 0 && (bw < W || bh < H)) ratio = bw / bh;
            }
          }
        }
      } catch { /* tainted canvas o.ä. → Datei-Ratio behalten */ }
      IMG_ASPECT_CACHE[probeUrl] = ratio;
      setImgProbe({ url: probeUrl, ratio });
    };
    // Ladefehler: Form bleibt unbekannt (ratio null → Standard-Slot), aber
    // das Gate öffnet, damit das Bild/Fallback nicht dauerhaft versteckt ist.
    img.onerror = () => { if (!cancelled) setImgProbe({ url: probeUrl, ratio: null }); };
    img.src = probeUrl;
    return () => { cancelled = true; };
  }, [heroImgUrl]);
  // v28.6: Slot-Größe hängt von der BILDFORM ab — Kreis-/Quadrat-Bilder
  // (Ratio ~1, z.B. aus dem Zuschnitt-Tool) brauchen keinen 300er-Block,
  // Querformat bekommt Breite, Hochkant Höhe.
  const imgSlotW = imgAspect == null ? 280 : (imgAspect >= 1.2 ? 420 : (imgAspect >= 0.8 ? 210 : 240));
  const imgSlotH = imgAspect == null ? 260 : (imgAspect >= 1.2 ? 260 : (imgAspect >= 0.8 ? 210 : 300));
  // v28.7: Kreis-/Quadrat-Bilder (Ratio ~1, typisch der Kreis-Zuschnitt aus
  // dem Wizard mit transparenten Ecken) sitzen NICHT mehr seitlich neben den
  // Infos, sondern als eigener Kreis OBEN MITTIG, der die Oberkante der
  // Event-Karte überlappt („eingebautes" Profilbild-Muster). Banner-Modus
  // hat weiter Vorrang; Quer-/Hochformat behält den Seiten-Slot.
  const imgCircleNotch = !!heroImgUrl && !event?.imageBanner && imgAspect != null && imgAspect >= 0.8 && imgAspect < 1.2;
  const circleSize = isMobile ? 140 : 170;
  // v28.91: Kein Event-Foto → das DEX-Bild steht als KREIS oben mittig,
  // genau dort, wo auch ein rundes Event-Logo sitzt (imgCircleNotch). Im
  // Seiten-Slot rechts wirkte es wie ein Foto des Events, das es nicht ist.
  const showOrbPlaceholder = !heroImgUrl;

  // B2Run Split-Capacity: aktuelle Auslastung pro Typ laden
  // Split-UI nur wenn BEIDE Starter-Typen verfügbar sind (>0). Wenn der Admin eine
  // Kapazität auf 0 gesetzt hat, gibt es faktisch nur einen Typ — dann keine Auswahl
  // anzeigen und den einzig verfügbaren Typ automatisch setzen (siehe useEffect unten).
  const durchCap = (event && typeof event.durchstarterCapacity === 'number') ? event.durchstarterCapacity : 0;
  const funCap = (event && typeof event.funstarterCapacity === 'number') ? event.funstarterCapacity : 0;
  const isSplitGroup = !!event && durchCap > 0 && funCap > 0;
  // v10.20: frei wählbare Bezeichnungen aus dem Event laden, mit Fallback auf
  // die historischen B2Run-Defaults 'Durchstarter' / 'Funstarter'. Die internen
  // Werte für SP-Persistenz (StarterType-Spalte) bleiben unverändert — das
  // Label ist reines UI.
  const splitLabelA = (event?.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
  const splitLabelB = (event?.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
  const singleStarterType: string = (!event || (durchCap <= 0 && funCap <= 0))
    ? '' // kein B2Run-Event überhaupt
    : (durchCap > 0 && funCap <= 0) ? 'Durchstarter'
    : (funCap > 0 && durchCap <= 0) ? 'Funstarter'
    : ''; // beide > 0 -> User muss wählen (Split-UI)

  // Auto-Set: wenn nur ein Starter-Typ verfügbar ist, direkt diesen Typ als
  // preferredStarterType speichern — damit registerForEvent ihn trotzdem auf den
  // Teilnehmer-Eintrag schreiben kann, obwohl das Split-UI nicht angezeigt wird.
  React.useEffect(() => {
    if (singleStarterType && preferredStarterType !== singleStarterType) {
      setPreferredStarterType(singleStarterType);
    }
  }, [singleStarterType]);

  // v6.15: Starter-Typ → Startblock-Auto-Mapping. Wenn der Admin für dieses Event
  // einen Block an den Starter-Typ gebunden hat, wird das zugehörige
  // b2run_startblock-Custom-Field automatisch gesetzt — der User muss den Block
  // nicht extra auswählen (das Custom-Field wird dann im UI ausgeblendet).
  const durchstarterBlock = event?.durchstarterStartblock || '';
  const funstarterBlock = event?.funstarterStartblock || '';
  const hasStarterBlockMapping = !!(durchstarterBlock || funstarterBlock);
  React.useEffect(() => {
    if (!hasStarterBlockMapping || !preferredStarterType) return;
    const mappedBlock = preferredStarterType === 'Durchstarter' ? durchstarterBlock : funstarterBlock;
    if (!mappedBlock) return;
    if (eventSpecific.b2run_startblock === mappedBlock) return;
    setEventSpecific(prev => ({ ...prev, b2run_startblock: mappedBlock }));
  }, [preferredStarterType, durchstarterBlock, funstarterBlock, hasStarterBlockMapping]);
  // v26.74: Vorauswahl bei Single-Select-Feldern — den vom Organizer gesetzten
  // Default einmal vorbelegen (nur wenn der Teilnehmer das Feld noch nicht
  // berührt hat; ein bewusstes Leeren bleibt erhalten). Läuft, sobald die
  // Event-Felder verfügbar sind bzw. sich ihre Defaults ändern.
  const selectDefaultsSig = (event?.eventSpecificFields || [])
    .filter(f => f.type === 'select' && !f.multi && f.defaultValue)
    .map(f => `${f.id}=${f.defaultValue}`).join('|');
  React.useEffect(() => {
    const fields = event?.eventSpecificFields || [];
    const defaults: Record<string, string> = {};
    for (const f of fields) {
      if (f.type === 'select' && !f.multi && f.defaultValue && (f.options || []).indexOf(f.defaultValue) >= 0) {
        defaults[f.id] = f.defaultValue;
      }
    }
    if (Object.keys(defaults).length === 0) return;
    setEventSpecific(prev => {
      let changed = false;
      const next = { ...prev };
      for (const k of Object.keys(defaults)) {
        if (next[k] === undefined) { next[k] = defaults[k]; changed = true; }
      }
      return changed ? next : prev;
    });
  }, [selectDefaultsSig]);
  React.useEffect(() => {
    if (!isSplitGroup || !event?.subsiteUrl) return;
    (async () => {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ctx = (window as any).__dexSpfxContext;
        if (!ctx) return;
        const { EventService } = await import('../services/EventService');
        const svc = new EventService(ctx);
        const allRegs = await svc.getAllRegistrations(event.subsiteUrl!);
        const active = allRegs.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
        // v19.19: Warteliste pro Gruppe über die effektive Gruppe
        // (StarterType || PreferredStarterType) — Wartelisten-Einträge haben
        // i.d.R. noch keinen StarterType, ihr Gruppen-Wunsch steht in
        // PreferredStarterType.
        const waiting = allRegs.filter(r => r.Status === 'Warteliste');
        const effGroup = (r: { StarterType?: string; PreferredStarterType?: string }): string => r.StarterType || r.PreferredStarterType || '';
        setStarterCounts({
          durch: active.filter(r => r.StarterType === 'Durchstarter').length,
          fun: active.filter(r => r.StarterType === 'Funstarter').length,
          durchWait: waiting.filter(r => effGroup(r) === 'Durchstarter').length,
          funWait: waiting.filter(r => effGroup(r) === 'Funstarter').length,
        });
      } catch { /* ignore */ }
    })();
  }, [isSplitGroup, event?.subsiteUrl]);
  // v19.17: Der frühere 5-Sekunden-Poll auf dem Anmelde-Screen wurde wieder
  // entfernt — er verursachte einen sichtbaren Re-Render. Die Belegungszahl
  // kommt jetzt aus dem Context-Stand beim Öffnen (die Übersicht lädt sie beim
  // Navigieren frisch nach). Kein Polling, kein Refresh, Formular bleibt stabil.
  // Deloitte-Mitarbeitersuche
  const [userSearch, setUserSearch] = React.useState('');
  const [userResults, setUserResults] = React.useState<Array<{ email: string; displayName: string; location: string; jobTitle: string }>>([]);
  // v11.97: nach Picker-Auswahl im "Für andere Person registrieren"-Modus
  // halten wir das volle Profil (Department + Mobile zusätzlich), damit
  // die Personal-Info-Card die gleichen Read-only-Felder zeigt wie beim
  // Self-Register-Modus.
  const [pickedUserProfile, setPickedUserProfile] = React.useState<{
    jobTitle?: string;
    department?: string;
    location?: string;
    mobilePhone?: string;
    // v28.11: Unternehmenszugehörigkeit der ausgewählten Person — vorher
    // fehlte das Feld und die Profil-Karte zeigte „— nicht hinterlegt".
    company?: string;
  } | null>(null);
  const [isSearchingUser, setIsSearchingUser] = React.useState(false);
  const [userSearchIncludeIntl, setUserSearchIncludeIntl] = React.useState(false);
  const searchTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  if (!event) {
    // v28.7: Beim Browser-Refresh restauriert der NavigationContext die
    // Anmeldeseite SOFORT, während die Events noch aus SharePoint laden —
    // vorher stand dann fälschlich „Event nicht gefunden". Solange die
    // Events laden, zeigen wir den Spinner (gleiches Muster wie die
    // Event-Liste); „nicht gefunden" kommt erst, wenn das Event nach dem
    // Laden wirklich fehlt.
    if (isEventsLoading) {
      return (
        <div className="page-container text-center">
          <div style={{ padding: 48 }}>
            <svg width={48} height={48} viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" style={{ display: 'block', margin: '0 auto 16px' }}>
              <circle cx={24} cy={24} r={20} fill="none" stroke="rgba(134,188,37,0.20)" strokeWidth={4} />
              <path d="M 24 4 A 20 20 0 0 1 44 24" fill="none" stroke="#86bc25" strokeWidth={4} strokeLinecap="round">
                <animateTransform attributeName="transform" type="rotate" from="0 24 24" to="360 24 24" dur="1s" repeatCount="indefinite" />
              </path>
            </svg>
            <p style={{ color: 'var(--dex-gray-400)' }}>{locale === 'de' ? 'Event wird geladen …' : 'Loading event …'}</p>
          </div>
        </div>
      );
    }
    return (
      <div className="page-container text-center">
        <h2>{t('reg.eventnotfound')}</h2>
        <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>
          {t('reg.backtoevents')}
        </button>
      </div>
    );
  }

  // Registrierungs-Deadline prüfen.
  // v22.54: „Anmeldung geschlossen" greift nur, wenn das Hauptevent UND alle
  // Sub-Events zu sind. Solange mindestens ein Sub-Event noch offen ist, kommt
  // der Teilnehmer rein und kann sich für die offenen Sub-Events anmelden —
  // auch wenn die (Klammer-/Hauptevent-)Frist abgelaufen ist.
  // v28.20: Auch die explizite Klammer-Frist zählt (Organizer/Admin-Banner +
  // Parent-Reg-Block; für reguläre User greift ohnehin die Fully-Closed-Seite).
  const isDeadlinePassed = (!!event.registrationDeadline && new Date(event.registrationDeadline) < new Date())
    || (!!event.klammerDeadline && new Date(event.klammerDeadline) < new Date());
  const isFullyClosed = isRegistrationFullyClosed(event, childEvents);

  // v23.14: Vorschau vor Aktivierung — reguläre User dürfen die Anmeldeseite
  // erst ab dem „Aktiv ab"-Zeitpunkt öffnen (Deep-Link-Schutz; die Karte
  // blockiert den Klick ohnehin). Organizer/Admin dürfen vorbereiten.
  const notYetActive = !!event.activeFrom && new Date(event.activeFrom) > new Date();
  if (notYetActive && !isOrganizer && !isAdmin) {
    const activeFromStr = new Date(event.activeFrom as string).toLocaleString(locale === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    return (
      <div className="page-container">
        <div className="card" style={{ position: 'relative', overflow: 'hidden' }}>
          <div style={{
            height: 200,
            background: event.imageUrl ? `url(${cachedImage}) center/cover no-repeat` : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
            borderRadius: '16px 16px 0 0',
          }} />
          <div style={{ padding: 32, textAlign: 'center' }}>
            <Icon iconName="Clock" style={{ fontSize: 48, color: 'var(--dex-orange)', marginBottom: 16 }} />
            <h2 style={{ marginBottom: 8 }}>{event.title}</h2>
            <p style={{ color: 'var(--dex-gray-700)', marginBottom: 8, fontWeight: 600 }}>
              {locale === 'de' ? 'Die Anmeldung ist noch nicht geöffnet.' : 'Registration is not open yet.'}
            </p>
            <p style={{ color: 'var(--dex-gray-500)', fontSize: '0.9rem' }}>
              {locale === 'de' ? 'Anmeldung ab' : 'Registration opens'}: <strong>{activeFromStr}</strong>
            </p>
            <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>
              {t('reg.backtoevents')}
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (isFullyClosed && !isOrganizer && !isAdmin) {
    return (
      <div className="page-container">
        <div className="card" style={{ position: 'relative', overflow: 'hidden' }}>
          <div style={{
            height: 200,
            background: event.imageUrl
              ? `url(${cachedImage}) center/cover no-repeat`
              : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
            borderRadius: '16px 16px 0 0',
          }} />
          <div style={{ padding: 32, textAlign: 'center' }}>
            <Icon iconName="Clock" style={{ fontSize: 48, color: 'var(--dex-orange)', marginBottom: 16 }} />
            <h2 style={{ marginBottom: 8 }}>{t('reg.deadlinepassed.title')}</h2>
            <p style={{ color: 'var(--dex-gray-600)', marginBottom: 8 }}>
              {t('reg.deadlinepassed.text')}
            </p>
            {/* v28.20: Bei Klammern mit expliziter Frist DIE anzeigen — die
                Spalten-Frist ist dort ein wirkungsloser Alt-Wert (und kann
                leer sein → Invalid Date). */}
            {(() => {
              const d = event.klammerDeadline || event.registrationDeadline;
              return d ? (
                <p style={{ color: 'var(--dex-gray-400)', fontSize: '0.85rem' }}>
                  {t('reg.deadlinepassed.date')}: {new Date(d).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}
                </p>
              ) : null;
            })()}
            <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>
              {t('reg.backtoevents')}
            </button>
          </div>
        </div>
      </div>
    );
  }

  const errorBorder = { border: '2px solid var(--dex-red)' };

  const parentAlreadyRegistered = !!(myParentReg && myParentReg.Status && myParentReg.Status !== 'Abgemeldet');
  // "Hauptevent wird jetzt angemeldet" gilt nur, wenn der Parent-Checkbox an ist
  // UND der User nicht bereits angemeldet ist. Bei bereits angemeldetem Parent
  // wird die Parent-Registrierung nicht nochmal ausgelöst.
  // v22.54: Ist die Hauptevent-Frist abgelaufen, kann ein normaler Teilnehmer
  // das Hauptevent nicht mehr buchen — die offenen Sub-Events bleiben aber
  // wählbar. Organizer/Admins dürfen weiterhin (manuelle Anmeldung).
  // v24.41: „Meine Assistenz beauftragen"-Prompt — nur bei der EIGENEN
  // Anmeldung (kein Stellvertreter-/Team-Modus), nur für Admins ODER Directoren,
  // und NICHT, wenn das Event bereits ein Assistenz-CC-Feld hat (dann hätte der
  // Organizer den CC schon eingebaut → kein doppeltes Abfragen).
  // v24.45: Partner ODER Director (P/D) — vorher wurde nur „director" geprüft,
  // Partner gingen leer aus. „Senior Director"/„Associate Partner" matchen mit.
  const isPartnerOrDirector = /(partner|director)/i.test(currentUser.jobTitle || '');
  // v24.48: Unterdrücken, wenn das Event bereits ein eigenes Assistenz-Feld hat
  // — entweder ein People-Picker mit „auf CC setzen" ODER ein People-Picker,
  // dessen Bezeichnung auf „Assistenz/Assistant" hindeutet (z.B. „Your assistant").
  const hasAssistantCcField = (event?.eventSpecificFields || []).some(f => {
    if (f.type !== 'user' && f.type !== 'roommate') return false;
    if (f.ccOnEmails) return true;
    return /assist|assistenz|assistenten?/i.test(`${f.label || ''} ${f.labelEn || ''}`);
  });
  // v27.8: NUR noch Partner/Director — die Abfrage „Für Partner & Directoren"
  // erschien vorher auch jedem Admin bei der eigenen Anmeldung, obwohl ein
  // Admin nicht zwangsläufig P/D ist. Admins, die zugleich P/D sind, matchen
  // weiter über isPartnerOrDirector.
  const canDelegateAssistant = isPartnerOrDirector && !registerForOther && !isTeamMode && !pendingJoinTeam && !hasAssistantCcField;
  const parsedDelegateAssist = (() => {
    const m = (delegateAssistValue || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
    return m ? { name: m[1].trim(), email: m[2].trim() } : null;
  })();
  // v27.11: Voll & Warteliste vom Organizer deaktiviert → Hauptevent nicht
  // mehr buchbar. Vorher lief die Anmeldung still auf die (abgeschaltete)
  // Warteliste — der WaitlistEnabled-Toggle war wirkungslos. Sub-Events
  // bleiben über die parentRegBlocked-Mechanik weiterhin einzeln buchbar.
  const pfActive = liveStats ? liveStats.active : (event ? (event.currentParticipants || 0) : 0);
  const pfWaitlist = (liveStats && liveStats.waitlist >= 0) ? liveStats.waitlist : (event ? (event.waitlistCount || 0) : 0);
  const parentFullNoWaitlist = !!event && event.maxParticipants > 0 && event.waitlistEnabled === false
    && Math.max(0, event.maxParticipants - pfActive - pfWaitlist) <= 0;
  const parentRegBlocked = ((isDeadlinePassed && !isOrganizer && !isAdmin) || parentFullNoWaitlist) && !parentAlreadyRegistered;
  const willRegisterParent = registerForParent && !parentAlreadyRegistered && !registerForOther && !(event && event.subEventsOnlyMode) && !parentRegBlocked;
  // Fürs Registrieren für andere bleibt der alte Flow: Parent wird immer registriert,
  // keine Session-Auswahl (siehe Render).
  const isSessionsOnlyMode = !willRegisterParent && !registerForOther && !parentAlreadyRegistered;

  // v28.88: Gibt es überhaupt etwas abzuschicken?
  //
  // Bisher hing die Sperre allein an `selectedSessions.size === 0`. Das ist zu
  // grob: Wer bereits gebuchte Sub-Events ALLE abwählt, will sie abmelden —
  // die Auswahl ist dann leer, es gibt aber sehr wohl etwas zu tun (der
  // Abmelde-Pfad im Sub-Event-Loop weiter unten). Deshalb zählt hier jede
  // Abweichung zwischen Vorbelegung (sessionMeta.wasRegistered) und aktueller
  // Auswahl als Änderung.
  const sessionsChanged = childEvents.some(ce => {
    const wasReg = !!sessionMeta[ce.id]?.wasRegistered;
    const isSel = selectedSessions.has(ce.id);
    return (isSel && !wasReg) || (!isSel && wasReg && !registerForOther);
  });
  // Nichts anzumelden, nichts zu ändern, kein Team-Vorgang → der
  // „Registrieren"-Klick hätte keine Wirkung.
  const nothingToSubmit = !willRegisterParent && !registerForOther && !isTeamMode
    && !pendingJoinTeam && selectedSessions.size === 0 && !sessionsChanged;

  // v9.22: Warning-Modal für externe Email-Anmeldung (durch Organizer für
  // Drittpersonen die noch kein Deloitte-Postfach haben). Default: nicht
  // erlaubt; Organizer kann nach Bestätigung trotzdem fortfahren — die
  // Bestätigungsmail geht dann nicht an die externe Adresse, sondern an
  // den Organizer mit Datenschutz-Hinweis-Header.
  const [externalEmailWarning, setExternalEmailWarning] = React.useState(false);
  // v18.75: Sicherheitshinweis vor dem Absenden (pro Event konfiguriert). Der
  // Dialog erscheint nach dem „Anmelden"-Klick und vor der eigentlichen
  // (Normal-)Anmeldung. confirmDraft* halten die — in der Auswahl-Übersicht
  // editierbare — Auswahl, bis der User bestätigt.
  const [confirmDialogOpen, setConfirmDialogOpen] = React.useState(false);
  const [confirmDialogAck, setConfirmDialogAck] = React.useState(false);
  const [confirmDraftParent, setConfirmDraftParent] = React.useState(true);
  const [confirmDraftSessions, setConfirmDraftSessions] = React.useState<Set<string>>(new Set());
  const confirmDialogConfirmedRef = React.useRef(false);
  const externalEmailConfirmedRef = React.useRef(false);
  // v19.6: Bei stellvertretender Anmeldung einer INTERNEN Person (Deloitte)
  // fragt nach dem „Anmelden"-Klick ein Modal, ob der/die Anmeldende (Organizer,
  // Co-Organizer oder Assistenz) selbst auf CC der Bestätigungs-Mail gesetzt
  // werden soll. ccSelfDecidedRef merkt sich, dass die Frage in diesem
  // Submit-Durchlauf bereits beantwortet wurde (analog confirmDialogConfirmedRef);
  // ccSelfRef hält die Entscheidung (true = auf CC).
  const [ccSelfModalOpen, setCcSelfModalOpen] = React.useState(false);
  const ccSelfDecidedRef = React.useRef(false);
  const ccSelfRef = React.useRef(false);
  // v24.48: Assistenz-Abfrage als Modal beim Register-Klick (Partner/Director).
  const [assistantModalOpen, setAssistantModalOpen] = React.useState(false);
  const assistantModalDecidedRef = React.useRef(false);
  // v24.49: Auswahl SYNCHRON im Ref festhalten — der Re-Submit aus dem Modal
  // läuft sonst mit dem alten State-Wert (setState ist async) und die CC würde
  // verloren gehen. { enabled, value } wird beim Klick im Modal gesetzt.
  const delegateChoiceRef = React.useRef<{ enabled: boolean; value: string } | null>(null);

  // v29.27: Sub-Event-Fragen INLINE in der Sub-Event-Karte — nicht mehr im
  // Bestätigen-Modal. Der Teilnehmer sieht damit direkt an der Kachel, welche
  // Frage zu welchem Termin gehört (die Hauptevent-Fragen stehen darunter mit
  // eigener Überschrift). Die Werte hängen live an sessionFieldValues[ce.id];
  // die Pflicht-Prüfung, die vorher das Modal erzwang, sitzt jetzt im Submit.
  // Kalender-Modus und der Team-Beitritts-Dialog nutzen weiter das Modal —
  // dort gibt es keine Karte, die die Felder tragen könnte.
  const renderSubEventInlineFields = (ce: DeloitteEvent): React.ReactElement | null => {
    const values = sessionFieldValues[ce.id] || {};
    const useEnHere = locale === 'en' && !!ce.bilingualFields;
    const fLabel = (f: EventSpecificField): string =>
      (useEnHere && f.labelEn && f.labelEn.trim()) ? f.labelEn : f.label;
    const fHelp = (f: EventSpecificField): string | undefined =>
      (useEnHere && f.helpTextEn && f.helpTextEn.trim()) ? f.helpTextEn : f.helpText;
    const fOpt = (f: EventSpecificField, opt: string, idx: number): string =>
      (useEnHere && f.optionsEn && f.optionsEn[idx] && f.optionsEn[idx].trim()) ? f.optionsEn[idx] : opt;
    const fields = (ce.eventSpecificFields || [])
      .filter(f => f && f.label)
      .filter(f => {
        if (!f.showIf || !f.showIf.fieldId) return true;
        const raw = (values[f.showIf.fieldId] || '').trim();
        if (!raw) return false;
        const answers = raw.indexOf(' | ') >= 0
          ? raw.split(' | ').map(s => s.trim()).filter(Boolean)
          : [raw];
        return answers.some(a => f.showIf!.values.indexOf(a) >= 0);
      });
    if (fields.length === 0) return null;
    const setValue = (fieldId: string, value: string): void => {
      setSessionFieldValues(prev => ({ ...prev, [ce.id]: { ...(prev[ce.id] || {}), [fieldId]: value } }));
    };
    return (
      <div style={{ marginTop: 10, padding: '10px 12px', borderRadius: 8, background: '#fff', border: '1px solid var(--dex-gray-200)' }}>
        <div style={{ fontSize: '0.78rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', marginBottom: 8 }}>
          {locale === 'de'
            ? `Fragen zu diesem ${childTermSingular || 'Sub-Event'}`
            : `Questions for this ${childTermSingular || 'sub-event'}`}
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          {fields.map(f => {
            const val = values[f.id] || '';
            const missing = showErrors && f.required && (f.type === 'checkbox' ? val !== 'true' : !val.trim());
            return (
              <div key={f.id}>
                <label className="form-label" style={{ display: 'block', fontSize: '0.82rem', marginBottom: 4, ...(missing ? { color: 'var(--dex-red, #c00)' } : {}) }}>
                  {fLabel(f)}
                  {f.required && <span style={{ color: 'var(--dex-red, #c00)', marginLeft: 4 }}>*</span>}
                  {fHelp(f) && <InfoTooltip text={fHelp(f)} />}
                </label>
                {f.type === 'select' && f.multi ? (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    {(f.options || []).map((opt, optIdx) => {
                      const current = val.split(' | ').map(s => s.trim()).filter(Boolean);
                      const checked = current.indexOf(opt) >= 0;
                      return (
                        <label key={opt} style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.84rem', cursor: 'pointer' }}>
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={e => {
                              const next = e.target.checked ? [...current, opt] : current.filter(x => x !== opt);
                              setValue(f.id, next.join(' | '));
                            }}
                          />
                          {fOpt(f, opt, optIdx)}
                        </label>
                      );
                    })}
                  </div>
                ) : f.type === 'select' ? (
                  <select className="form-input" value={val} onChange={e => setValue(f.id, e.target.value)} style={{ width: '100%', fontSize: '0.88rem' }}>
                    <option value="">{locale === 'de' ? '— bitte wählen —' : '— please select —'}</option>
                    {(f.options || []).map((opt, optIdx) => <option key={opt} value={opt}>{fOpt(f, opt, optIdx)}</option>)}
                  </select>
                ) : f.type === 'checkbox' ? (
                  <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.84rem' }}>
                    <input type="checkbox" checked={val === 'true'} onChange={e => setValue(f.id, e.target.checked ? 'true' : 'false')} />
                    {locale === 'de' ? 'Ja' : 'Yes'}
                  </label>
                ) : f.type === 'number' ? (
                  <input type="number" className="form-input" value={val} onChange={e => setValue(f.id, e.target.value)} style={{ width: '100%', fontSize: '0.88rem' }} />
                ) : (
                  <input type="text" className="form-input" value={val} onChange={e => setValue(f.id, e.target.value)} style={{ width: '100%', fontSize: '0.88rem' }} />
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  // v29.28: Der Hauptevent-Felder-Block als Render-Funktion — er wird an
  // ZWEI möglichen Orten gebraucht: bei Events MIT Sub-Events direkt unter
  // der Haupt-Event-Kachel in der Auswahl-Box (dort, wo die Fragen
  // hingehören — die Sub-Event-Fragen stecken seit v29.27 in deren Karten),
  // sonst an der bisherigen Stelle unter der Auswahl. Inhalt 1:1 der
  // bisherige Block (v11.2/v26.91) — als Ganzes gehoben, nicht geschnitten.
  const renderMainFieldsSection = (): React.ReactElement => (
    event.eventSpecificFields.length === 0 && !isSplitGroup ? (
      <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{t('reg.noadditional')}</p>
    ) : (
      // v11.2 / v11.5: Custom-Fields ohne Pro-Gruppe-Constraint im
      // 2-Spalten-Grid. Group-spezifische Felder werden bereits
      // oben innerhalb der Gruppen-Auswahl-Box gerendert und hier
      // ausgefiltert.
      <div className="dex-reg-fields-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
      {/* v29.27: Zuordnung klarmachen — die Sub-Event-Fragen stehen
          in den Karten, diese hier gehören zum Haupt-Event (bzw. bei
          einer Klammer zur Anmeldung insgesamt). */}
      {childEvents.length > 0 && event.eventSpecificFields.length > 0 && (
        <div style={{ gridColumn: '1 / -1', fontSize: '0.8rem', fontWeight: 700, color: 'var(--dex-gray-600)', marginBottom: -6 }}>
          {event.subEventsOnlyMode
            ? (locale === 'de' ? 'Allgemeine Fragen zur Anmeldung' : 'General questions for your registration')
            : (locale === 'de' ? 'Fragen zum Haupt-Event' : 'Questions for the main event')}
        </div>
      )}
      {(() => {
        // v26.91: Zuerst die WIRKLICH sichtbaren Felder ermitteln, dann mit
        // Index rendern — so kann renderRegField pro 2-Spalten-Zeile
        // entscheiden, ob es leeren Beschreibungs-Platz reservieren muss.
        const visibleSpecificFields = event.eventSpecificFields
          .filter(f => f.id !== 'b2run_mobilnummer' || eventSpecific['b2run_infoservice'] === 'true')
          .filter(f => !(f.id === 'b2run_startblock' && hasStarterBlockMapping))
          .filter(f => {
            if (!f.showIf || !f.showIf.fieldId) return true;
            const raw = (eventSpecific[f.showIf.fieldId] || '').trim();
            if (!raw) return false;
            const answers = raw.indexOf(' | ') >= 0
              ? raw.split(' | ').map(s => s.trim()).filter(Boolean)
              : [raw];
            return answers.some(a => f.showIf!.values.indexOf(a) >= 0);
          })
          // v11.5: Group-Spec NICHT hier rendern — die kommen oben
          // in der Gruppen-Auswahl-Box. Hier nur 'all' / undefined
          // (oder Events ohne Split-Capacity).
          .filter(f => {
            const grp = f.onlyForGroup;
            if (!grp || grp === 'all') return true;
            if (!isSplitGroup) return true;
            return false;
          });
        return visibleSpecificFields.map((f, i) => renderRegField(f, undefined, undefined, i, visibleSpecificFields));
      })()}
      </div>
    )
  );

  const handleSubmit = async (): Promise<void> => {
    // v17.25: Demo-Showcase-Event — keine echte Anmeldung. Freundlicher
    // Hinweis statt SP-Roundtrip; der Context-Guard würde ohnehin no-oppen.
    if (event?.isDemoShowcase) {
      setError(locale === 'de'
        ? 'Dies ist ein Demo-Event — es wird keine echte Anmeldung gespeichert. Du kannst die Bereiche oben frei ausprobieren.'
        : 'This is a demo event — no real registration is stored. Feel free to explore the sections above.');
      return;
    }
    // Validierung Pflichtfelder
    setShowErrors(true);

    // Wenn der Haupt-Event-Checkbox aus ist und keine Session ausgewählt ist,
    // gibt es nichts zu tun.
    //
    // v28.88: …aber der GRUND muss zum Fall passen. „Bitte wähle mindestens das
    // Haupt-Event oder ein Sub-Event aus" setzt voraus, dass es etwas zu wählen
    // gibt. Bei einer bestehenden Anmeldung ist `willRegisterParent` immer
    // false (parentAlreadyRegistered, s.o.), und hat das Event keine
    // Sub-Events, steht auf der Seite überhaupt keine Auswahl. Wer bereits
    // angemeldet war und auf „Registrieren" klickte, bekam deshalb eine
    // Aufforderung, die ins Leere zeigt. Dasselbe im „Nur Sub-Events"-Modus und
    // bei gesperrtem Hauptevent (Frist abgelaufen / voll ohne Warteliste): Das
    // Haupt-Event ist dort gar nicht wählbar, die Meldung nannte es trotzdem.
    if (nothingToSubmit) {
      const oneSub = childTermSingular || (locale === 'de' ? 'Sub-Event' : 'sub-event');
      const hasSubs = childEvents.length > 0;
      if (parentAlreadyRegistered) {
        setError(locale === 'de'
          ? (hasSubs
            ? `Du bist für dieses Event bereits angemeldet. Möchtest du zusätzlich ${childOneDe} buchen, wähle es oben aus — abmelden kannst du dich über „Meine Events".`
            : 'Du bist für dieses Event bereits angemeldet — es gibt nichts weiter abzuschicken. Abmelden kannst du dich über „Meine Events".')
          : (hasSubs
            ? `You are already registered for this event. To additionally book a ${oneSub}, pick it above — you can cancel via „My events".`
            : 'You are already registered for this event — there is nothing further to submit. You can cancel via „My events".'));
        return;
      }
      if ((event && event.subEventsOnlyMode) || parentRegBlocked) {
        if (hasSubs) {
          // v29.13: Der Zusatz „das Haupt-Event ist hier nicht buchbar" hilft
          // nur, wenn es für den Teilnehmer sichtbar EIN Haupt-Event gibt. Im
          // reinen Sub-Event-Modus gibt es das nicht — dort ist die Auswahl
          // oben schlicht die Liste der Events, und der Nachsatz erfände eine
          // zweite Ebene, die nirgends auftaucht.
          setError(locale === 'de'
            ? (subOnlyTerms
              ? `Bitte wähle mindestens ${childOneDe} aus, um dich anzumelden.`
              : `Bitte wähle mindestens ${childOneDe} aus, um dich anzumelden — das Haupt-Event ist hier nicht buchbar.`)
            : (subOnlyTerms
              ? `Please select at least one ${oneSub} to register.`
              : `Please select at least one ${oneSub} to register — the main event cannot be booked here.`));
          return;
        }
        // v29.9: „keines angelegt" nur sagen, wenn wirklich keines existiert.
        // Sind welche da, aber keines für diese Person freigegeben, ist das der
        // Grund — und die Person soll wissen, dass sie nach einer Freigabe
        // fragen muss und nicht nach einem fehlenden Termin.
        setError(locale === 'de'
          ? (parentFullNoWaitlist
            ? 'Alle Plätze sind belegt und die Warteliste ist für dieses Event deaktiviert — eine Anmeldung ist nicht mehr möglich.'
            : (isDeadlinePassed
              ? 'Die Anmeldefrist dieses Events ist abgelaufen — eine Anmeldung ist nicht mehr möglich.'
              : (hiddenChildCount > 0
                ? 'Die Anmeldung läuft hier ausschließlich über die einzelnen Programmpunkte — für dich ist aktuell keiner davon freigegeben. Wenn du teilnehmen möchtest, wende dich bitte an die Organizer.'
                : 'Für dieses Event läuft die Anmeldung ausschließlich über Sub-Events — aktuell ist keines angelegt. Bitte wende dich an die Organizer.')))
          : (parentFullNoWaitlist
            ? 'All seats are taken and the waitlist is disabled for this event — registration is no longer possible.'
            : (isDeadlinePassed
              ? 'The registration deadline for this event has passed — registration is no longer possible.'
              : (hiddenChildCount > 0
                ? 'Registration here runs exclusively via the individual programme items — none of them is currently released for you. If you would like to attend, please contact the organizers.'
                : 'Registration for this event runs exclusively via sub-events — none exists yet. Please contact the organizers.'))));
        return;
      }
      setError(t('reg.nothing.selected') || 'Bitte wähle mindestens Haupt-Event oder eine Session aus.');
      return;
    }

    // v24.64: Pflicht-Sub-Events. Pro Sub-Event kann der Organizer im Wizard
    // („Sub-Events"-Schritt) „Pflichtanmeldung" setzen — ein so markiertes
    // Sub-Event MUSS ausgewählt sein, sonst ist die Anmeldung nicht möglich.
    // (Löst das alte, in der UI nicht mehr einstellbare
    // requireSubEventSelection ab — das wird hier bewusst NICHT mehr geprüft.)
    const subShortName = (c: { title?: string }): string => {
      const tt = (c.title || '').trim();
      const parts = tt.split('|');
      return (parts.length > 1 ? parts[parts.length - 1] : tt).trim();
    };
    const mandatoryMissing = childEvents.filter(c => c.mandatoryRegistration && !selectedSessions.has(c.id));
    if (mandatoryMissing.length > 0) {
      const names = mandatoryMissing.map(subShortName).filter(Boolean).join(', ');
      setError(locale === 'de'
        ? `Bitte wähle die Pflicht-${mandatoryMissing.length === 1 ? (childTermSingular || 'Sub-Event') : (childTermPlural || 'Sub-Events')} aus, um dich anzumelden: ${names}.`
        : `Please select the mandatory ${mandatoryMissing.length === 1 ? (childTermSingular || 'sub-event') : (childTermPlural || 'sub-events')} to register: ${names}.`);
      return;
    }
    // v24.64: Im „Nur Sub-Events"-Modus ist das Haupt-Event nicht buchbar —
    // dann muss mindestens ein Sub-Event gewählt sein.
    if (event && event.subEventsOnlyMode && childEvents.length > 0 && selectedSessions.size === 0) {
      setError(locale === 'de'
        ? `Bitte wähle mindestens ${childOneDe} aus, um dich anzumelden.`
        : `Please select at least one ${childTermSingular || 'sub-event'} to register.`);
      return;
    }

    // v15.16: Pflicht-Bestätigung bei „Für andere registrieren" —
    // analog zur Team-Anmeldung muss die Zustimmung der Person
    // explizit bestätigt werden.
    if (registerForOther && !otherConsentConfirmed) {
      setError(locale === 'de'
        ? 'Bitte bestätige, dass die Person ihrer stellvertretenden Anmeldung zugestimmt hat.'
        : 'Please confirm that the person has consented to this registration.');
      return;
    }

    // Basis-Felder sind immer Pflicht (Name + Email), auch im Sessions-Only-Modus.
    if (!firstName.trim() || !surname.trim() || !email.trim()) {
      setError(t('reg.requiredfields'));
      return;
    }
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) {
      setError(t('reg.invalidemail') || 'Ungültige E-Mail-Adresse');
      return;
    }

    // v19.6: Stellvertretende Anmeldung — verhindern, dass der Organizer
    // versehentlich SICH SELBST als „andere Person" einträgt. Im
    // „Für andere registrieren"-Modus MUSS eine andere Person ausgewählt sein.
    // Ist die Teilnehmer-E-Mail leer ODER identisch zur eigenen, läuft sonst die
    // Doppel-Anmelde-Prüfung gegen den eingeloggten User und meldet
    // irreführend „bereits angemeldet" (Beobachtung des Users: „bei einer
    // Anmeldung einer dritten Person wird weiterhin geprüft, ob ich selber
    // angemeldet werde"). Hier klar abbrechen statt still die Selbst-Prüfung
    // auszulösen.
    if (registerForOther && email.trim().toLowerCase() === (currentUser.email || '').toLowerCase()) {
      setError(t('reg.error.selfasother'));
      return;
    }

    // v19.8: Bereits angemeldete Zielperson hart blocken — kein CC-Modal, keine
    // Anmeldung. Greift auch, falls der Button-Disable wegen einer noch
    // laufenden Vorab-Prüfung kurz nicht aktiv war.
    if (registerForOther && thirdPartyCheck && thirdPartyCheck.alreadyRegistered) {
      setError(t('reg.error.other'));
      return;
    }

    // v18.74: Bei stellvertretender Anmeldung einer EXTERNEN Adresse zwei
    // Stufen: (1) strengere Plausibilitätsprüfung gegen Tippfehler (fehlende
    // TLD, doppelte Punkte, mehrere @, …) → harter Fehler; (2) ein
    // Bestätigungs-Dialog, der die Adresse groß anzeigt und zum Gegenlesen
    // auffordert (man kann an externe Adressen keinen Tippfehler korrigieren).
    if (registerForOther && email.trim() && isExternalEmailAddr(email)) {
      if (!isPlausibleEmail(email)) {
        setError(locale === 'de'
          ? 'Die externe E-Mail-Adresse sieht ungültig aus — bitte auf Tippfehler prüfen (z.B. fehlende Domain-Endung wie „.de" oder „.com").'
          : 'The external email address looks invalid — please check for typos (e.g. a missing domain ending like „.de" or „.com").');
        return;
      }
      if (!externalEmailConfirmedRef.current) {
        setExternalEmailWarning(true);
        return; // Bestätigungs-Dialog (Tippfehler-Gegenlesen) zeigen
      }
    }

    // Nur wenn der User das Haupt-Event (neu) anmelden möchte, gelten
    // Anrede + Custom-Fields + B2Run-Starter-Typ als Pflicht.
    // v18.53 BUG-FIX: Im subEventsOnlyMode wird das Hauptevent zwar nicht
    // direkt angemeldet, aber als „Schatten-Registrierung" mitgeschrieben
    // (s.u., shouldShadowRegisterParent) — inkl. der Hauptevent-Custom-Fields
    // (z.B. „Selection as above confirmed", Food Preferences, Hotel). Diese
    // werden in dem Modus trotzdem angezeigt und persistiert, also MÜSSEN sie
    // auch validiert werden. Vorher konnte man im Nur-Sub-Events-Modus mit
    // leerem Pflichtfeld absenden, weil `willRegisterParent` hier immer false
    // ist. Dieselbe Bedingung wie shouldShadowRegisterParent unten verwenden.
    const isSubOnlyModeValidate = !!(event && event.subEventsOnlyMode);
    // v18.57 BUG-FIX: Bedingung robust gemacht. Vorher hing sie an
    // `sessionsBeingAddedValidate` (= sessionMeta[...]?.wasRegistered), das in
    // manchen Re-Submit-/Reload-Fällen unzuverlässig war → Validierungsblock
    // wurde übersprungen → Pflichtfeld-Bypass. Jetzt reicht: subEventsOnlyMode
    // aktiv UND mindestens eine Section ausgewählt. Die übergreifenden
    // Hauptevent-Pflichtfelder (insb. die pro Absenden neu leere „Bestätigung"-
    // Checkbox) werden in dem Modus IMMER angezeigt und müssen IMMER validiert
    // werden. KEIN `!myParentReg`, KEIN sessionMeta-Abhängigkeit mehr.
    // v29.27: Pflichtfelder der ausgewählten Sub-Events prüfen — die Fragen
    // stehen jetzt inline in der Karte, also muss der Submit erzwingen, was
    // vorher das Bestätigen-Modal erzwungen hat. Bereits bestehende
    // Anmeldungen (wasRegistered) sind ausgenommen: ihre Antworten liegen in
    // der Teilnehmer-Zeile und werden hier nicht neu erfasst.
    {
      const subMissing: string[] = [];
      childEvents.forEach(ce => {
        if (!selectedSessions.has(ce.id)) return;
        if (sessionMeta[ce.id]?.wasRegistered) return;
        const values = sessionFieldValues[ce.id] || {};
        (ce.eventSpecificFields || []).filter(f => f && f.label && f.required && f.type !== 'document').forEach(f => {
          if (f.showIf && f.showIf.fieldId) {
            const raw = (values[f.showIf.fieldId] || '').trim();
            const answers = !raw ? [] : (raw.indexOf(' | ') >= 0 ? raw.split(' | ').map(s => s.trim()).filter(Boolean) : [raw]);
            if (!answers.some(a => f.showIf!.values.indexOf(a) >= 0)) return;
          }
          const filled = f.type === 'checkbox' ? values[f.id] === 'true' : !!(values[f.id] || '').trim();
          if (!filled) subMissing.push(`${ce.title || (locale === 'de' ? 'Sub-Event' : 'sub-event')}: ${f.label}`);
        });
      });
      if (subMissing.length > 0) {
        setError(`${t('reg.requiredcustom')}: ${subMissing.join(', ')}`);
        return;
      }
    }
    const willCollectMainFields = willRegisterParent || registerForOther
      || (isSubOnlyModeValidate && selectedSessions.size > 0 && !registerForOther)
      // v18.73: Beim vorgemerkten Team-Beitritt gelten dieselben Pflichtfelder
      // (Anrede + event-spezifische Felder) wie bei einer normalen Anmeldung —
      // der ganze Sinn der Änderung ist, dass diese Infos nicht übersprungen
      // werden können.
      || !!pendingJoinTeam;
    if (willCollectMainFields) {
      // v11.80: Anrede ist nur dann Pflichtfeld, wenn das Event das
      // Anrede-Dropdown auch tatsächlich abfragt (event.askSalutation === true).
      // Sonst wird die Anrede gar nicht gerendert und bleibt leer.
      if (event.askSalutation && !salutation) {
        setError(t('reg.requiredfields'));
        return;
      }

      // Pflicht-Custom-Fields validieren. Checkbox-Pflichtfelder müssen 'true' sein,
      // alle anderen dürfen nach trim nicht leer sein.
      // B2Run: Mobilnummer ist nur Pflicht wenn Infoservice aktiviert; ansonsten
      // gilt das Feld als versteckt und wird übersprungen.
      const missingRequired = event.eventSpecificFields
        .filter(f => {
          if (f.id === 'b2run_mobilnummer') {
            if (eventSpecific['b2run_infoservice'] !== 'true') return false;
            return !eventSpecific[f.id]?.trim();
          }
          // v7.21: Felder mit nicht erfüllter Sichtbarkeitsbedingung sind
          // ausgeblendet und dürfen die Validation nicht blockieren.
          if (f.showIf && f.showIf.fieldId) {
            const raw = (eventSpecific[f.showIf.fieldId] || '').trim();
            const answers = !raw
              ? []
              : raw.indexOf(' | ') >= 0
                ? raw.split(' | ').map(s => s.trim()).filter(Boolean)
                : [raw];
            const conditionMet = answers.some(a => f.showIf!.values.indexOf(a) >= 0);
            if (!conditionMet) return false;
          }
          // v10.24: Pro-Gruppe-Constraint — wenn das Feld auf eine andere
          // Gruppe als die aktuell vom User gewählte beschränkt ist, ist
          // es ausgeblendet und blockt die Validation nicht.
          if (f.onlyForGroup && f.onlyForGroup !== 'all' && isSplitGroup) {
            const want = f.onlyForGroup === 'A' ? 'Durchstarter' : 'Funstarter';
            if (preferredStarterType !== want) return false;
          }
          if (!f.required) return false;
          // v19.0: Pflicht-Dokument — bei NEUER Anmeldung muss eine Datei gewählt
          // sein. Bei bereits angemeldeter Person läuft die Verwaltung über
          // „Meine Events", daher dort nicht blockieren.
          if (f.type === 'document') return !parentAlreadyRegistered && !pendingDocFiles[f.id];
          return f.type === 'checkbox'
            ? eventSpecific[f.id] !== 'true'
            : !eventSpecific[f.id]?.trim();
        });
      if (missingRequired.length > 0) {
        setError(`${t('reg.requiredcustom')}: ${missingRequired.map(f => f.label).join(', ')}`);
        return;
      }

      // B2Run: Starter-Typ Pflichtfeld. v18.73: Beim Team-Beitritt NICHT
      // erzwingen — der Beitretende erbt die Gruppe des Teams (siehe
      // addTeamMember), er wählt sie nicht selbst.
      if (isSplitGroup && !preferredStarterType && !pendingJoinTeam) {
        // v11.7: generische Fehlermeldung — vorher hatte der Translation-Key
        // 'B2Run Starter-Typ' als Fallback. Bei generischen Split-Capacity-
        // Events mit eigenen Labels (z.B. 'Vormittag' / 'Nachmittag') passt
        // das nicht. Inline-Text mit den frei wählbaren Labels.
        const lblA = (event?.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
        const lblB = (event?.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
        setError(locale === 'de'
          ? `Bitte wähle eine der zwei Gruppen aus (${lblA} oder ${lblB}).`
          : `Please pick one of the two groups (${lblA} or ${lblB}).`);
        return;
      }

      // v6.15: Leistungsnachweis-Pflicht bei Durchstarter (Admin-Option)
      if (event.durchstarterRequiresProof && preferredStarterType === 'Durchstarter' && eventSpecific['b2run_leistungsnachweis'] !== 'true') {
        setError(t('reg.starter.proof.required') || 'Bitte Leistungsnachweis bestätigen.');
        return;
      }
    }

    // v11.10: Sub-Events erben preferredStarterType vom oberen
    // Group-Selection-Block. Wenn das Event Split-Capacity hat und Sessions
    // ausgewählt sind, muss eine Gruppe gewählt sein — egal ob Parent dabei
    // ist oder nur Sessions registriert werden.
    if (isSplitGroup && selectedSessions.size > 0 && !preferredStarterType && !pendingJoinTeam) {
      setError(locale === 'de'
        ? 'Bitte wähle eine Gruppe.'
        : 'Please pick a group.');
      return;
    }

    // Assistant-Ausnahme: defense-in-depth check beim Submit — Target muss
    // Partner oder Director sein. Der Fall tritt nur ein wenn der User weder
    // Organizer des Events noch Admin ist, aber via Assistant-Ausnahme für
    // eine andere Person registrieren will. JobTitle entweder aus dem zuletzt
    // geladenen Search-Result oder per Live-Lookup.
    if (registerForOther && isAssistant && !canCreateEvents) {
      try {
        const emailLc = email.trim().toLowerCase();
        let targetJobTitle = userResults.find(u => u.email.toLowerCase() === emailLc)?.jobTitle || '';
        if (!targetJobTitle) {
          const fresh = await searchUsers(email.trim(), userSearchIncludeIntl);
          targetJobTitle = fresh.find(u => u.email.toLowerCase() === emailLc)?.jobTitle || '';
        }
        if (!isAllowedTargetForAssistant(targetJobTitle)) {
          setError('As an Assistant you can only register Partners or Directors for events.');
          return;
        }
      } catch {
        setError('Unable to verify job title of selected person. Please select again from the dropdown.');
        return;
      }
    }

    // Seit v6.5: B2Run-Split-Kapazitäten. Wenn der gewählte Starter-Typ voll ist,
    // aber der andere Typ noch freie Plätze hat, zeigen wir einen Dialog —
    // der User entscheidet explizit:
    //   (a) auf den anderen Typ umsteigen, oder
    //   (b) auf die Warteliste für den gewünschten Typ.
    // Kein stiller Auto-Fallback mehr. Beide Typen voll → direkt auf Warteliste
    // (kein Dialog, Logik in EventContext setzt Status=Warteliste).
    if ((willRegisterParent || registerForOther) && isSplitGroup && preferredStarterType && !pendingJoinTeam) {
      const durchFree = Math.max(0, durchCap - starterCounts.durch);
      const funFree = Math.max(0, funCap - starterCounts.fun);
      const wunschFree = preferredStarterType === 'Durchstarter' ? durchFree : funFree;
      const altType = preferredStarterType === 'Durchstarter' ? 'Funstarter' : 'Durchstarter';
      const altFree = altType === 'Durchstarter' ? durchFree : funFree;
      if (wunschFree === 0 && altFree > 0) {
        // Dialog zeigen, Submit warten.
        setFallbackDialog({ wunsch: preferredStarterType, alt: altType, altFree });
        return;
      }
    }

    // v18.73: Vorgemerkter Team-Beitritt — committet hier (mit den oben
    // ausgefüllten persönlichen + event-spezifischen Feldern), statt einer
    // normalen Einzel-Anmeldung. Steht vor dem Team-Anmelde-Pfad, weil beide
    // sich gegenseitig ausschließen.
    if (pendingJoinTeam) {
      await performJoinSelectedTeam();
      return;
    }

    // v11.82: Team-Anmeldung — separater Submit-Pfad.
    if (isTeamMode) {
      if (!teamValidation.ok) {
        setError(teamValidation.reason || (locale === 'de' ? 'Team-Anmeldung unvollständig.' : 'Team registration incomplete.'));
        return;
      }
      await performTeamRegistration(preferredStarterType);
      return;
    }

    // v18.75: Sicherheitshinweis vor dem Absenden (nur normale Anmeldung —
    // Team-/Beitritts-Pfade sind oben bereits abgehandelt). Beim ersten
    // Submit öffnet sich der Dialog; nach Bestätigung läuft handleSubmit erneut
    // (Ref gesetzt) und überspringt den Dialog.
    if (event && event.confirmDialogEnabled && !confirmDialogConfirmedRef.current) {
      setConfirmDraftParent(willRegisterParent || registerForOther);
      setConfirmDraftSessions(new Set(selectedSessions));
      setConfirmDialogAck(false);
      setConfirmDialogOpen(true);
      return;
    }

    // v19.6: Stellvertretende Anmeldung einer INTERNEN Person — fragen, ob
    // der/die Anmeldende selbst auf CC der Bestätigungs-Mail soll. Nur einmal
    // pro Submit-Durchlauf (ccSelfDecidedRef); externe Personen sind ausgenommen
    // (dort ist der Organizer-Kreis ohnehin schon auf CC). Nach der Entscheidung
    // läuft handleSubmit erneut und überspringt das Modal.
    if (registerForOther && !externalPerson && !ccSelfDecidedRef.current) {
      setCcSelfModalOpen(true);
      return;
    }

    // v24.48: Assistenz-Abfrage als Modal beim Register-Klick — nur für
    // Partner/Director (canDelegateAssistant), einmal pro Submit-Durchlauf.
    // Nach der Entscheidung läuft handleSubmit erneut und überspringt das Modal.
    if (canDelegateAssistant && !assistantModalDecidedRef.current) {
      setAssistantModalOpen(true);
      return;
    }

    await performRegistration(preferredStarterType);
  };

  // v11.82: Team-Anmeldung absenden — Lead + N-1 Mitglieder per registerTeam.
  // Sub-Events werden im Team-Modus NICHT mitangemeldet — das ist Phase 2
  // (siehe Manual). Lead und Member bekommen jeweils nur den Hauptevent.
  const performTeamRegistration = async (starterTypeToUse: string): Promise<void> => {
    setError('');
    // v23.2: Harter Riegel gegen Doppel-Anmeldung über den Team-Pfad. Wer als
    // eingeloggte Person bereits aktiv beim Event angemeldet ist (solo oder in
    // einem anderen Team), darf NICHT erneut ein Team anlegen — der Solo-Pfad
    // ist seit jeher so abgesichert, der Team-Pfad war es nicht (Ursache der
    // Doppel-Anmeldung bei Team-Events). Die Team-Karte ist in dem Fall bereits
    // ausgeblendet; das hier ist das Sicherheitsnetz, falls der Status erst
    // nach dem Aufklappen geladen wurde.
    if (parentAlreadyRegistered) {
      setError(locale === 'de'
        ? 'Du bist bereits für dieses Event angemeldet — eine zusätzliche Team-Anmeldung ist nicht möglich. Bitte zuerst über „Meine Events" abmelden, falls du in ein anderes Team wechseln möchtest.'
        : 'You are already registered for this event — an additional team registration is not possible. Please cancel via „My Events" first if you want to switch to another team.');
      return;
    }
    setIsSubmitting(true);
    setSubmitProgress(5);
    setSubmitProgressLabel(locale === 'de' ? 'Team-Anmeldung wird vorbereitet…' : 'Preparing team registration…');
    try {
      const customData: Record<string, string> = { salutation, ...stripPrefilterKeys(eventSpecific) };
      const leadEmail = email.trim();
      const leadFirstName = firstName.trim();
      const leadLastName = surname.trim();
      // v18.12: Custom-Field-Antworten pro Mitglied (nach Slot-Index) mitgeben.
      const members = teamMembersParsed
        .map((m, idx) => m ? { email: m.email, displayName: m.displayName, customData: stripPrefilterKeys(teamMemberFields[idx] || {}) } : null)
        .filter((m): m is { displayName: string; email: string; customData: Record<string, string> } => !!m);
      setSubmitProgress(30);
      setSubmitProgressLabel(locale === 'de'
        ? `Team wird angemeldet (${1 + members.length} Personen)…`
        : `Registering team (${1 + members.length} people)…`);
      const result = await registerTeam(
        selectedEventId!,
        {
          firstName: leadFirstName,
          lastName: leadLastName,
          email: leadEmail,
          salutation,
          customData,
          preferredStarterType: starterTypeToUse || undefined,
        },
        members,
        event?.askTeamName ? (teamName.trim() || undefined) : undefined
      );
      setSubmitProgress(90);
      if (!result.ok) {
        if (result.reason && result.reason.indexOf('already-registered:') === 0) {
          const dupEmail = result.reason.substring('already-registered:'.length);
          setError(locale === 'de'
            ? `Person bereits angemeldet: ${dupEmail}. Bitte aus dem Team entfernen und erneut versuchen.`
            : `Person already registered: ${dupEmail}. Please remove from the team and try again.`);
        } else {
          setError(t('reg.error') || (locale === 'de' ? 'Fehler bei der Team-Anmeldung.' : 'Team registration failed.'));
        }
        return;
      }
      setSubmitted(true);
    } catch {
      setError(t('reg.genericerror') || (locale === 'de' ? 'Unerwarteter Fehler.' : 'Unexpected error.'));
    } finally {
      setSubmitProgress(100);
      setTimeout(() => {
        setIsSubmitting(false);
        setSubmitProgress(0);
        setSubmitProgressLabel('');
      }, 250);
    }
  };

  // v18.73: Vorgemerkten Team-Beitritt absenden — committet erst hier (auf
  // „Anmelden"), inkl. der ausgefüllten event-spezifischen Felder. Direkter
  // Beitritt (joinTeam) bzw. Anfrage an den Lead (createTeamJoinRequest), je
  // nach event.teamJoinRequiresApproval.
  const performJoinSelectedTeam = async (): Promise<void> => {
    if (!pendingJoinTeam || !event) return;
    setError('');
    setIsSubmitting(true);
    setSubmitProgress(10);
    setSubmitProgressLabel(locale === 'de' ? 'Beitritt wird verarbeitet…' : 'Processing your join…');
    try {
      // Event-spezifische Antworten des Beitretenden (wie bei der normalen
      // Anmeldung) — werden an den Team-Beitritt durchgereicht.
      const customData: Record<string, string> = { salutation, ...stripPrefilterKeys(eventSpecific) };
      setSubmitProgress(50);
      if (event.teamJoinRequiresApproval) {
        const r = await createTeamJoinRequest(event.id, pendingJoinTeam.teamId, customData);
        if (!r.ok) {
          setError(r.reason === 'already-registered'
            ? (locale === 'de' ? 'Du bist bereits beim Event angemeldet.' : 'You are already registered for this event.')
            : (locale === 'de' ? 'Beitritts-Anfrage fehlgeschlagen.' : 'Join request failed.'));
          return;
        }
        setSubmittedJoinKind('requested');
        setSubmitted(true);
      } else {
        const r = await joinTeam(event.id, pendingJoinTeam.teamId, pendingJoinTeam.teamName, customData);
        if (!r.ok) {
          setError(r.reason && r.reason.startsWith('already-registered')
            ? (locale === 'de' ? 'Du bist bereits beim Event angemeldet.' : 'You are already registered for this event.')
            : r.reason === 'team-full'
              ? (locale === 'de' ? 'Das Team ist inzwischen voll.' : 'The team has filled up in the meantime.')
              : (locale === 'de' ? 'Beitritt fehlgeschlagen.' : 'Joining failed.'));
          return;
        }
        setSubmittedJoinKind('joined');
        setSubmittedAsWaitlist(r.status === 'Warteliste');
        setSubmitted(true);
      }
    } catch {
      setError(locale === 'de' ? 'Unerwarteter Fehler beim Beitritt.' : 'Unexpected error while joining.');
    } finally {
      setSubmitProgress(100);
      setTimeout(() => { setIsSubmitting(false); setSubmitProgress(0); setSubmitProgressLabel(''); }, 250);
    }
  };

  // Eigentliche Registrierung — entkoppelt vom Validation/Submit-Trigger,
  // damit sie auch vom Fallback-Dialog aufgerufen werden kann (mit ggf. geändertem Starter-Typ).
  const performRegistration = async (starterTypeToUse: string): Promise<void> => {
    setError('');
    setIsSubmitting(true);
    setSubmitProgress(5);
    setSubmitProgressLabel(locale === 'de' ? 'Anmeldung wird vorbereitet…' : 'Preparing registration…');
    try {
      const customData: Record<string, string> = {
        salutation,
        ...stripPrefilterKeys(eventSpecific),
      };
      const participantEmail = email.trim();
      const firstTrim = firstName.trim();
      const surnameTrim = surname.trim();

      // v19.6: CC-Wunsch bei stellvertretender INTERNER Anmeldung. Wenn der/die
      // Anmeldende im Modal „Ja, auf CC" gewählt hat, landet die eigene E-Mail
      // als CC auf der Bestätigungs-Mail (NICHT im Outlook-Termin — wie bei den
      // Feld-CCs). Externe Anmeldungen sind ausgenommen.
      const ccSelfEmail = (registerForOther && !externalPerson && ccSelfRef.current)
        ? (currentUser.email || '').trim()
        : '';

      // v24.41: Delegation an die eigene Assistenz (nur Selbst-Anmeldung von
      // Admin/Director, Assistenz im Picker gewählt). Die Assistenz kommt auf CC
      // der Bestätigung; die eigentliche Zugriffs-Übergabe (Zeilen-Autor) läuft
      // nach erfolgreicher Anmeldung über delegateRegistrationToAssistant.
      // v24.49: Auswahl aus dem Ref (synchron gesetzt) lesen, Fallback auf State.
      const choice = delegateChoiceRef.current;
      const chosenVal = choice ? choice.value : (delegateAssistEnabled ? delegateAssistValue : '');
      const chosenEnabled = choice ? choice.enabled : delegateAssistEnabled;
      const chosenParsed = (() => {
        const m = (chosenVal || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
        return m ? { name: m[1].trim(), email: m[2].trim() } : null;
      })();
      const delegateAssist = (canDelegateAssistant && chosenEnabled && chosenParsed
        && chosenParsed.email.toLowerCase() !== participantEmail.toLowerCase())
        ? chosenParsed : null;
      const delegateCc = delegateAssist ? delegateAssist.email : '';

      let anySuccess = false;
      let parentOk = true;
      let lastSubReason: string | undefined;
      // v23.10: Assistenz-Proxy-Anmeldung — der Client hat (Picker-Greyout +
      // Submit-Validierung oben) bereits sichergestellt, dass eine Assistenz nur
      // Partner/Director anmeldet. Dieses Flag wird der Registrierung als
      // vertrauenswürdig durchgereicht, damit die fragile serverseitige
      // Job-Title-Ableitung legitime Assistenzen nicht mehr fälschlich ablehnt.
      const actorAllowedAsAssistant = registerForOther && isAssistant && !canCreateEvents;
      // v23.9: Übersetzt den konkreten Misserfolgs-Grund aus registerForEvent in
      // eine verständliche Meldung — statt pauschal „bereits registriert".
      const regFailMessage = (reason?: string): string => {
        if (reason === 'not-allowed') {
          return locale === 'de'
            ? 'Du bist nicht berechtigt, diese Person für dieses Event anzumelden. Bitte wende dich an die Organizer des Events.'
            : 'You are not allowed to register this person for this event. Please contact the event organizers.';
        }
        if (reason === 'deadline') {
          return locale === 'de'
            ? 'Die Anmeldefrist dieses Events ist abgelaufen — eine Anmeldung ist nicht mehr möglich. Organizer und Admins können nach Ablauf weiterhin anmelden.'
            : 'The registration deadline for this event has passed — registration is no longer possible. Organizers and admins can still register after the deadline.';
        }
        if (reason === 'insert-failed') {
          return locale === 'de'
            ? 'Die Anmeldung konnte nicht gespeichert werden (technischer Fehler an der Teilnehmerliste). Bitte erneut versuchen; hält es an, im Organizer Center „Spalten fixen" ausführen.'
            : 'The registration could not be saved (technical error on the participant list). Please try again; if it persists, run „Fix columns" in the organizer center.';
        }
        // v27.11: aktiver Duplikat-Treffer (deckt jetzt auch Externe ab).
        if (reason === 'already-registered') {
          return registerForOther
            ? t('reg.thirdparty.alreadyregistered')
            : (locale === 'de'
              ? 'Du bist für dieses Event bereits angemeldet.'
              : 'You are already registered for this event.');
        }
        // v27.11: Event voll und Warteliste vom Organizer abgeschaltet.
        if (reason === 'full') {
          return locale === 'de'
            ? 'Alle Plätze sind belegt und die Warteliste ist für dieses Event deaktiviert — eine Anmeldung ist nicht mehr möglich.'
            : 'All seats are taken and the waitlist is disabled for this event — registration is no longer possible.';
        }
        // Fallback (unbekannt / kein Grund) — bisherige generische Meldung.
        return t(registerForOther ? 'reg.error.other' : 'reg.error');
      };

      // Verfeinerte Progress-Stufen je nach Anzahl Sub-Events:
      // - parent: 5 → 30 → 50 (wenn Parent-Anmeldung lief)
      // - sub-events: 50 → 90 (gleichmäßig verteilt)
      // - finalize: 90 → 100
      const subOps = childEvents.filter(ce => {
        const wasReg = sessionMeta[ce.id]?.wasRegistered;
        const isSel = selectedSessions.has(ce.id);
        return (isSel && !wasReg) || (!isSel && wasReg && !registerForOther);
      }).length;

      // 1) Haupt-Event anmelden (nur wenn Checkbox an und noch nicht angemeldet).
      // v15.25: Im subEventsOnlyMode wird die Parent-Anmeldung trotzdem
      // durchgeführt — als „Schatten-Registrierung" rein zur Daten-
      // Vollständigkeit. Damit hat jeder Teilnehmer auch im Parent-
      // Teilnehmer-Schema eine Zeile mit den Antworten auf die Hauptevent-
      // Custom-Fields (Food Preferences, Hotel, Travel etc.). Mails +
      // Outlook werden für diese Schatten-Anmeldung in EventContext
      // unterdrückt — der User soll keine Bestätigung fürs Hauptevent
      // bekommen, da er da gar nicht „teilnimmt", sondern nur für Sub-
      // Events.
      const isSubOnlyMode = !!(event && event.subEventsOnlyMode);
      const sessionsBeingAdded = childEvents.some(ce => selectedSessions.has(ce.id) && !sessionMeta[ce.id]?.wasRegistered);
      const parentAlreadyHasRow = !!myParentReg;
      // v26.67 (B): deckt jetzt Selbst- UND Fremd-Anmeldung ab (das frühere
      // `!registerForOther` entfällt — die Klammer läuft im subEventsOnly-Modus
      // in beiden Fällen ZUM SCHLUSS über den Schritt-3-Block unten).
      const shouldShadowRegisterParent = isSubOnlyMode && sessionsBeingAdded && !parentAlreadyHasRow;

      // v28.22: UNSICHTBARE Doppel-Anmeldung abfangen.
      //
      // Die Teilnehmerlisten laufen mit Item-Level-Security („nur eigene
      // Elemente", geprüft am Zeilen-AUTOR). Meldet eine Assistenz jemanden an,
      // bleibt sie Autor der Zeile, solange der nachträgliche Autor-Wechsel
      // mangels Rechten scheitert (Contribute reicht dafür nicht; der
      // DEX_AccessFix-Flow bzw. der Admin-Auto-Fix zieht ihn erst später nach).
      // Bis dahin ist die Zeile für die betroffene Person UNSICHTBAR — weder in
      // „Meine Events" noch für den Vorab-Check beim Anmelden. Der Check läuft
      // fail-open (lieber eine Zeile zu viel als eine blockierte Anmeldung) und
      // legte deshalb eine ZWEITE Anmeldung an.
      //
      // Gegenmittel: zusätzlich DEX_Participants fragen. Die Liste liegt auf der
      // Haupt-Site, kennt keine Item-Level-Security und wird bei JEDER An-/
      // Abmeldung mitgeschrieben — sie sieht also auch fremd angelegte Zeilen.
      // Bewusst nur eine RÜCKFRAGE, keine harte Sperre: Sollte der Eintrag mal
      // veraltet sein (Abmeldung ohne erfolgreiches Nachziehen), bleibt eine
      // legitime Anmeldung möglich.
      const hiddenDupTitles: string[] = [];
      try {
        const nums = await getEventNumbersForEmail(participantEmail);
        const knownNumbers = new Set<number>([...nums.registered, ...nums.waitlisted]);
        const isKnown = (n?: number): boolean => typeof n === 'number' && n > 0 && knownNumbers.has(n);
        // Hauptevent/Klammer: nur prüfen, wenn wir jetzt wirklich eine Zeile
        // anlegen würden und uns keine sichtbare bekannt ist.
        const willTouchParent = (registerForParent && !parentAlreadyRegistered) || shouldShadowRegisterParent;
        if (willTouchParent && !parentAlreadyHasRow && isKnown(event.eventNumber)) {
          hiddenDupTitles.push(event.title);
        }
        for (const ce of childEvents) {
          if (!selectedSessions.has(ce.id)) continue;
          if (sessionMeta[ce.id]?.wasRegistered) continue;
          if (isKnown(ce.eventNumber)) hiddenDupTitles.push(ce.title);
        }
      } catch { /* best-effort — im Zweifel wie bisher weiter */ }
      if (hiddenDupTitles.length > 0) {
        const list = hiddenDupTitles.map(x => `• ${x}`).join('\n');
        const who = registerForOther ? (`${firstTrim} ${surnameTrim}`.trim() || participantEmail) : '';
        const proceed = await confirmDialog(
          locale === 'de'
            ? (registerForOther
              ? `${who} ist laut unseren Daten hier bereits angemeldet:\n\n${list}\n\nMöglicherweise hat sich die Person selbst angemeldet oder eine andere Assistenz hat das übernommen — dann siehst du die Zeile wegen der Zugriffsrechte auf der Teilnehmerliste nicht. Eine erneute Anmeldung würde einen ZWEITEN Platz belegen.\n\nTrotzdem anmelden?`
              : `Du bist laut unseren Daten hier bereits angemeldet:\n\n${list}\n\nMöglicherweise hat dich jemand angemeldet (z.B. deine Assistenz) — dann siehst du die Anmeldung wegen der Zugriffsrechte auf der Teilnehmerliste nicht in „Meine Events". Eine erneute Anmeldung würde einen ZWEITEN Platz belegen.\n\nTrotzdem anmelden?`)
            : (registerForOther
              ? `According to our records ${who} is already registered for:\n\n${list}\n\nThe person may have registered themselves, or another assistant did it — in that case the row is hidden from you by the attendee list's permissions. Registering again would take a SECOND seat.\n\nRegister anyway?`
              : `According to our records you are already registered for:\n\n${list}\n\nSomeone may have registered you (e.g. your assistant) — in that case the attendee list's permissions hide it from „My events". Registering again would take a SECOND seat.\n\nRegister anyway?`),
          { danger: true, confirmLabel: locale === 'de' ? 'Trotzdem anmelden' : 'Register anyway' },
        );
        if (!proceed) {
          setIsSubmitting(false);
          setSubmitProgress(0);
          setSubmitProgressLabel('');
          return;
        }
      }
      // v26.67 (B): Gemeinsame Klammer-/Parent-Anmelde-Routine. `bestEffort` =
      // true bei der subEventsOnly-Schatten-Zeile, die JETZT NACH den Sub-Events
      // angelegt wird — ein Fehlschlag darf die (gültigen) Sub-Event-Anmeldungen
      // nicht als Fehler markieren.
      const doParentRegistration = async (bestEffort: boolean): Promise<void> => {
        const parentResult = await registerForEvent(
          selectedEventId!,
          customData,
          firstTrim,
          surnameTrim,
          participantEmail,
          starterTypeToUse || undefined,
          // v18.74: Bei stellvertretender Anmeldung den Zustimmungs-Nachweis
          // mitschreiben (Pflicht-Checkbox wurde oben validiert).
          // v19.6: ccSelfEmail (Anmeldende:r auf CC der Bestätigungs-Mail).
          registerForOther
            ? { proxyConsentConfirmed: true, actorAllowedAsAssistant, ...(ccSelfEmail ? { extraCc: ccSelfEmail } : {}) }
            : (delegateCc ? { extraCc: delegateCc } : undefined)
        );
        if (bestEffort) {
          // Schatten-Klammer: Erfolg zählt mit, aber kein Fehler-Durchschlag.
          if (parentResult.ok) anySuccess = true;
          return;
        }
        parentOk = parentResult.ok;
        if (parentOk) {
          anySuccess = true;
          // v18.67: echten Status fürs Ergebnis-Modal merken (nicht isFull).
          setSubmittedAsWaitlist(parentResult.status === 'Warteliste');
        }
        // v23.9: KONKRETE Fehlermeldung statt pauschal „bereits registriert" —
        // der echte Grund (Berechtigung / Deadline / technischer Fehler) wird
        // jetzt aus registerForEvent durchgereicht.
        else setError(regFailMessage(parentResult.reason));
      };
      // v26.67 (B) BUG-FIX: Im subEventsOnly-Modus ist die „Parent"-Anmeldung nur
      // eine Schatten-/Klammer-Zeile (Daten-Vollständigkeit) — sie wird jetzt ZUM
      // SCHLUSS angelegt (siehe Schritt 3), nachdem mind. ein Sub-Event steht.
      // Bricht der Vorgang vorher ab, ist die Person sichtbar in ihren Sub-Events
      // angemeldet statt als unsichtbarer, blockierender „Geist" zurückzubleiben.
      // Bei NORMALEN Events (nicht subEventsOnly) bleibt der Parent die eigentliche
      // Anmeldung und läuft weiterhin ZUERST. (willRegisterParent ist im
      // subEventsOnly-Modus immer false.)
      if (!isSubOnlyMode && (willRegisterParent || registerForOther)) {
        setSubmitProgress(30);
        setSubmitProgressLabel(locale === 'de' ? 'Haupt-Event wird angemeldet…' : 'Registering for main event…');
        await doParentRegistration(false);
        setSubmitProgress(50);
      } else if (isSubOnlyMode && parentAlreadyHasRow && sessionsBeingAdded && !registerForOther) {
        // v18.59: Die Schatten-Parent-Zeile existiert bereits (frühere
        // Section-Anmeldung) → sie wird NICHT neu registriert. Trotzdem die
        // übergreifenden Hauptevent-Antworten (Food Preferences, Hotel etc.)
        // mit den aktuellen Formular-Werten aktualisieren, damit Änderungen
        // beim Nach-Anmelden einer weiteren Section persistiert werden. Vorher
        // gingen sie verloren (Audit-Befund #2).
        setSubmitProgress(40);
        try { await updateMyRegistration(selectedEventId!, customData); } catch { /* best-effort */ }
        setSubmitProgress(50);
      } else {
        setSubmitProgress(50);
      }

      // 2) Ausgewählte Sessions an-/abmelden (unabhängig vom Parent).
      //    - Session ausgewählt + nicht angemeldet → anmelden
      //    - Session nicht ausgewählt + angemeldet → abmelden
      //    - Starter-Typ: wenn der User sich gleichzeitig fürs Haupt-Event anmeldet,
      //      wird dessen Starter-Typ auch auf die Session-Teilnehmerliste geschrieben
      //      (shared) — sonst die pro-Session-Auswahl. So steht in der TN-Liste jeder
      //      Session korrekt, ob der Teilnehmer Durchstarter oder Funstarter ist.
      // v11.10: Sub-Events erben grundsätzlich preferredStarterType
      // (bzw. starterTypeToUse vom Fallback-Dialog). Vorher hingen sie
      // an sessionStarterType pro-Session, was zu redundanten UI-Radios
      // pro Sub-Event geführt hat.
      const inheritedStarterType = starterTypeToUse || preferredStarterType || '';
      // v18.53: Im subEventsOnlyMode sind die Hauptevent-CC-Felder (z.B.
      // „Your assistant") übergreifend — sie gelten für die Sub-Events. Daher
      // die CC einmal aus dem Hauptformular ziehen und an jede Sub-Event-
      // Anmeldung mitgeben, damit deren Bestätigungsmails ebenfalls an die
      // Assistenz auf CC gehen (das „Hauptevent" ist nicht anmeldbar und seine
      // Schatten-Registrierung verschickt keine Mail).
      const crossCutCc = (event && event.subEventsOnlyMode)
        ? collectCcEmailsFromFields(event.eventSpecificFields, customData, participantEmail)
        : '';
      // v10.15+: Sub-Event-Anmeldungen laufen auch beim Stellvertreter-Modus
      // (registerForOther) durch. registerForEvent akzeptiert ja participantFirstName/
      // -LastName/-Email als Argumente, daher kann der Assistent jede beliebige
      // Person sowohl auf das Hauptevent als auch auf alle gewählten Sub-Events
      // anmelden. Vorher war der Sub-Event-Loop hinter !registerForOther
      // versteckt — Beobachtung des Users: 'beim register for someone else kann
      // man nur fürs Main Event anmelden, nicht für die Sub-Events'. Fix.
      let subOpsDone = 0;
      // v26.67 (B): mind. eine NEUE Sub-Event-Anmeldung erfolgreich? Gate für
      // die nachgelagerte Schatten-Klammer.
      let anySubRegSuccess = false;
      // v29.25: Abwahlen, die wegen der Abmelde-Sperre NICHT abgemeldet wurden.
      const lockedCancelTitles: string[] = [];
      for (const ce of childEvents) {
        const wasReg = sessionMeta[ce.id]?.wasRegistered;
        const isSel = selectedSessions.has(ce.id);
        if (isSel && !wasReg) {
          setSubmitProgressLabel(locale === 'de'
            ? `${childTermSingular || 'Sub-Event'} „${ce.title || '?'}" wird angemeldet…`
            : `Registering for ${childTermSingular || 'sub-event'} „${ce.title || '?'}"…`);
          const sType = isSplitGroup ? (inheritedStarterType || undefined) : undefined;
          // Pro-Sub-Event Custom-Field-Werte aus dem Modal-Flow (sessionFieldValues
          // wird beim Bestätigen des Sub-Event-Modals befüllt). Default: {}.
          // v11.34: Anrede (salutation) zusätzlich mitgeben — vorher fehlte sie
          // bei Sub-Event-Anmeldungen, die Teilnehmerliste hatte dann „-" in
          // der Anrede-Spalte. Salutation kommt aus dem Hauptformular und ist
          // pro User identisch für alle Sub-Event-Anmeldungen.
          // v15.25: Im subEventsOnlyMode landen die Hauptevent-CF-Antworten
          // jetzt in der Schatten-Parent-Registrierung (s.o.) — die Sub-
          // Events bekommen nur ihre eigenen CFs aus dem Modal-Flow.
          const seFieldValues = { salutation, ...(sessionFieldValues[ce.id] || {}) };
          // v18.74: extraCc (übergreifende CC) + proxyConsentConfirmed (Nachweis
          // bei stellvertretender Anmeldung) zusammen in die Opts.
          // v19.6: ccSelfEmail zusätzlich in die CC der Sub-Event-Bestätigung
          // mergen (deduppt serverseitig).
          const seExtraCc = [crossCutCc, ccSelfEmail, delegateCc].filter(Boolean).join(';');
          const seOpts = (seExtraCc || registerForOther)
            ? { ...(seExtraCc ? { extraCc: seExtraCc } : {}), ...(registerForOther ? { proxyConsentConfirmed: true, actorAllowedAsAssistant } : {}) }
            : undefined;
          const subRes = await registerForEvent(ce.id, seFieldValues, firstTrim, surnameTrim, participantEmail, sType, seOpts);
          if (subRes.ok) { anySuccess = true; anySubRegSuccess = true; }
          else lastSubReason = subRes.reason;
          subOpsDone++;
          setSubmitProgress(50 + Math.floor((subOpsDone / Math.max(subOps, 1)) * 40));
        } else if (!isSel && wasReg && !registerForOther) {
          // v29.25: Selbst-Abmeldung nach der Frist gesperrt (Organizer-
          // Option) — das Abwählen darf hier nicht still abmelden. Der
          // Haken bleibt technisch abgewählt, die Anmeldung besteht weiter;
          // die betroffenen Termine werden nach dem Absenden benannt.
          if (selfCancelLocked(ce, event)) {
            lockedCancelTitles.push(ce.title || (locale === 'de' ? 'Sub-Event' : 'sub-event'));
            subOpsDone++;
            setSubmitProgress(50 + Math.floor((subOpsDone / Math.max(subOps, 1)) * 40));
            continue;
          }
          setSubmitProgressLabel(locale === 'de'
            ? `${childTermSingular || 'Sub-Event'} „${ce.title || '?'}" wird abgemeldet…`
            : `Cancelling ${childTermSingular || 'sub-event'} „${ce.title || '?'}"…`);
          // Cancel-Pfad bleibt aufs Selbst-Anmelden begrenzt: ein Stellvertreter
          // soll nicht aus Versehen einen Sub-Event-Slot des Anderen freigeben
          // weil er den Haken nicht gesetzt hat. Wer einen TN abmelden will,
          // macht das aktiv im Admin Center.
          await cancelRegistration(ce.id);
          anySuccess = true;
          subOpsDone++;
          setSubmitProgress(50 + Math.floor((subOpsDone / Math.max(subOps, 1)) * 40));
        }
      }
      // v29.25: Gesperrte Abmeldungen benennen — sonst sähe das Absenden wie
      // eine erfolgreiche Abmeldung aus, obwohl die Anmeldung weiter besteht.
      if (lockedCancelTitles.length > 0) {
        showAlert((event && event.noSelfCancel)
          ? (locale === 'de'
            ? `Nicht abgemeldet: ${lockedCancelTitles.join(', ')}. Die Organizer haben die Selbst-Abmeldung für dieses Event deaktiviert — bitte wende dich zum Abmelden an die Organizer. Deine Anmeldung bleibt bestehen.`
            : `Not cancelled: ${lockedCancelTitles.join(', ')}. The organizers have disabled self-cancellation for this event — please contact the organizers to cancel. Your registration remains in place.`)
          : (locale === 'de'
            ? `Nicht abgemeldet: ${lockedCancelTitles.join(', ')}. Die Abmeldefrist ist abgelaufen und die Organizer haben die Selbst-Abmeldung danach deaktiviert — bitte wende dich zum Abmelden an die Organizer. Deine Anmeldung bleibt bestehen.`
            : `Not cancelled: ${lockedCancelTitles.join(', ')}. The cancellation deadline has passed and the organizers have disabled self-cancellation after it — please contact the organizers to cancel. Your registration remains in place.`),
          { variant: 'error' });
      }
      // v26.67 (B) Schritt 3: Schatten-/Klammer-Zeile im subEventsOnly-Modus
      // JETZT anlegen — erst nachdem mind. ein Sub-Event erfolgreich angemeldet
      // wurde. So kann kein unsichtbarer „Geist" entstehen (Klammer ohne
      // Sub-Event). Deckt Selbst- (shouldShadowRegisterParent) UND
      // Fremd-Anmeldung (registerForOther) ab; ohne neue Sub-Event-Anmeldung
      // (z. B. nur Abmeldungen) wird keine leere Klammer erzeugt.
      if (shouldShadowRegisterParent && anySubRegSuccess) {
        setSubmitProgress(92);
        setSubmitProgressLabel(locale === 'de' ? 'Hauptevent-Daten werden gespeichert…' : 'Saving main-event data…');
        await doParentRegistration(true);
      }
      setSubmitProgress(95);
      setSubmitProgressLabel(locale === 'de' ? 'Bestätigungen werden versandt…' : 'Confirmations are being queued…');

      if (anySuccess) {
        // v19.0: ausgewählte Dokument-Dateien als Attachment an die Teilnehmer-
        // Zeile hängen — das Item existiert jetzt. Bei stellvertretender
        // Anmeldung an die Teilnehmer-E-Mail, sonst an den eingeloggten User.
        const docFields = (event?.eventSpecificFields || []).filter(f => f.type === 'document');
        const anyDoc = docFields.some(df => !!pendingDocFiles[df.id]);
        if (anyDoc) {
          setSubmitProgressLabel(locale === 'de' ? 'Dokumente werden hochgeladen…' : 'Uploading documents…');
          for (const df of docFields) {
            const file = pendingDocFiles[df.id];
            if (!file) continue;
            try { await uploadFieldDocument(selectedEventId!, df.id, file, registerForOther ? participantEmail : undefined); }
            catch { /* best-effort — Anmeldung bleibt gültig, Upload kann später über „Meine Events" nachgeholt werden */ }
          }
        }
        // Flag: wenn ausschließlich Sessions angemeldet/geändert wurden (kein
        // Parent diesmal oder schon vorher angemeldet), zeigen wir auf der
        // Success-Seite den Sessions-Only-Hinweis.
        setSessionsOnlySubmitted(!willRegisterParent && !registerForOther);
        // v24.41 Szenario A: Assistenz verknüpfen (Info + Anforderung). Der
        // Owner bleibt der/die Anmeldende; die Assistenz sieht es als Info.
        if (delegateAssist) {
          try { await delegateRegistrationToAssistant(selectedEventId!, delegateAssist); }
          catch { /* best-effort — Anmeldung bleibt gültig */ }
        }
        // v24.46: Hat das Event ein Assistenz-CC-Feld (Organizer hat es selbst
        // eingebaut → KEIN Modal) und der User dort eine Person angegeben, läuft
        // dieselbe Info-Freischaltung automatisch über dieses Feld — für JEDEN
        // Anmelder, nicht nur Partner/Director. Greift nur bei Selbst-Anmeldung
        // (für andere: die andere Person ist die angemeldete, nicht der CC).
        if (!registerForOther) {
          const ccFields = (event?.eventSpecificFields || []).filter(f => (f.type === 'user' || f.type === 'roommate') && !!f.ccOnEmails);
          const seenAssist = new Set<string>();
          for (const f of ccFields) {
            const raw = (customData[f.id] || '').trim();
            for (const part of raw.split(';').map(s => s.trim()).filter(Boolean)) {
              const m = part.match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
              const aEmail = m ? m[2].trim() : '';
              const aName = m ? m[1].trim() : '';
              if (!aEmail) continue;
              const key = aEmail.toLowerCase();
              if (key === participantEmail.toLowerCase() || seenAssist.has(key)) continue;
              seenAssist.add(key);
              try { await delegateRegistrationToAssistant(selectedEventId!, { email: aEmail, name: aName }); }
              catch { /* best-effort */ }
            }
          }
        }
        // v24.41 Szenario B: Bei stellvertretender Anmeldung (für andere) einen
        // Info-Link anlegen — der/die Anmeldende ist Owner, die angemeldete
        // Person sieht die Anmeldung als Info unter „Meine Events".
        if (registerForOther && !externalPerson && participantEmail
          && participantEmail.toLowerCase() !== (currentUser.email || '').toLowerCase()) {
          try { await recordProxyDelegation(selectedEventId!, { email: participantEmail, name: `${firstTrim} ${surnameTrim}`.trim() || participantEmail }); }
          catch { /* best-effort */ }
        }
        setSubmitted(true);
      } else if (!parentOk) {
        // Parent-Fehler wurde schon in setError oben gesetzt.
      } else {
        setError(regFailMessage(lastSubReason));
      }
    } catch {
      setError(t('reg.genericerror'));
    } finally {
      setSubmitProgress(100);
      setSubmitProgressLabel(locale === 'de' ? 'Fertig!' : 'Done!');
      // v19.6: CC-Frage-Entscheidung zurücksetzen, damit der nächste
      // Submit-Durchlauf (z.B. nächste stellvertretende Anmeldung) wieder fragt.
      ccSelfDecidedRef.current = false;
      ccSelfRef.current = false;
      // v24.48: Assistenz-Entscheidung zurücksetzen (nächster Submit fragt neu).
      assistantModalDecidedRef.current = false;
      delegateChoiceRef.current = null;
      // Kleine Verzögerung damit der User die 100%-Anzeige kurz sieht
      // bevor das Overlay wieder verschwindet.
      setTimeout(() => {
        setIsSubmitting(false);
        setSubmitProgress(0);
        setSubmitProgressLabel('');
      }, 250);
    }
  };

  // v26.37: „Zurücksetzen"-Button (und handleClear) entfernt — auf Wunsch.

  // v18.11: Proaktive Absage („Ich nehme nicht teil"). Keine Pflichtfelder
  // nötig — der User signalisiert nur, dass er nicht kommt.
  const handleDecline = async (): Promise<void> => {
    if (!event || isDeclining) return;
    if (event.isDemoShowcase) {
      setError(locale === 'de'
        ? 'Dies ist ein Demo-Event — es wird nichts gespeichert.'
        : 'This is a demo event — nothing is stored.');
      return;
    }
    setIsDeclining(true);
    setError('');
    try {
      // v29.32: Die Absage gilt für das GANZE Event — Klammer/Haupt-Event UND
      // alle sichtbaren Sub-Events. Vorher landete sie nur in der
      // Hauptevent-Liste: Im „Nur Sub-Events"-Modus ist das eine Schattenzeile,
      // und in den Sub-Event-Listen (die der Organizer tatsächlich auswertet)
      // stand die Person weiter als „hat nicht geantwortet". Eine Auswahl ist
      // dafür bewusst NICHT nötig — wer absagt, sagt für alles ab.
      // declineEvent je Ziel macht das Richtige: bestehende Anmeldung →
      // regulärer Abmelde-Pfad (Platz frei, Mail, Nachrücken), sonst eine
      // Absage-Zeile.
      const ok = await declineEvent(event.id);
      let subFailed = 0;
      for (const ce of childEvents) {
        try { if (!(await declineEvent(ce.id))) subFailed++; }
        catch { subFailed++; }
      }
      if (ok && subFailed === 0) setDeclined(true);
      else if (ok) {
        // Klammer steht, einzelne Sub-Events nicht — den Teilablauf benennen,
        // statt eine vollständige Absage zu behaupten.
        setDeclined(true);
        setError(locale === 'de'
          ? `Deine Absage ist erfasst — bei ${subFailed} ${subFailed === 1 ? (childTermSingular || 'Sub-Event') : (childTermPlural || 'Sub-Events')} hat es nicht geklappt. Bitte melde dich bei den Organizern.`
          : `Your decline was recorded — it failed for ${subFailed} sub-event(s). Please contact the organizers.`);
      }
      else setError(t('reg.genericerror') || 'Ein Fehler ist aufgetreten. Bitte versuche es erneut.');
    } catch {
      setError(t('reg.genericerror') || 'Ein Fehler ist aufgetreten. Bitte versuche es erneut.');
    } finally {
      setIsDeclining(false);
    }
  };

  // v18.14: Schritt 1 — eingefügte Liste gegen das Verzeichnis auflösen und
  // eine Vorschau-Tabelle (Vorname/Nachname/Position/Standort/E-Mail) bauen.
  // Pro Zeile: E-Mail erkennen (dann Profil-Lookup per E-Mail) ODER nur ein
  // Name (dann Personensuche → bester Treffer). Duplikate + nicht auflösbare
  // Zeilen werden markiert.
  const splitName = (raw: string): { firstName: string; lastName: string } => {
    const dn = (raw || '').trim();
    if (!dn) return { firstName: '', lastName: '' };
    if (dn.indexOf(',') >= 0) { const p = dn.split(',').map(s => s.trim()); return { firstName: p[1] || '', lastName: p[0] || '' }; }
    const p = dn.split(/\s+/).filter(Boolean);
    if (p.length <= 1) return { firstName: p[0] || '', lastName: '' };
    return { firstName: p[0], lastName: p.slice(1).join(' ') };
  };
  const resolveMassImport = async (): Promise<void> => {
    if (massImportResolving) return;
    const lines = massImportText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    const EMAIL_RE = /([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})/;
    setMassImportResolving(true);
    const rows: typeof massImportRows = [];
    const seen = new Set<string>();
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      setMassImportProgress(`${i + 1} / ${lines.length}`);
      const m = line.match(EMAIL_RE);
      let email = m ? m[1].toLowerCase() : '';
      const nameRaw = line.replace(m ? m[1] : '', '').replace(/[<>;,\t]/g, ' ').trim();
      let jobTitle = ''; let location = ''; let displayName = nameRaw;
      if (email) {
        // Profil per E-Mail nachschlagen (für Position + Standort + Name).
        try { const p = await searchUser(email); if (p) { displayName = p.displayName || nameRaw; jobTitle = p.jobTitle || ''; location = p.location || ''; } } catch { /* */ }
      } else if (nameRaw) {
        // Kein E-Mail in der Zeile → Personensuche, besten Treffer nehmen.
        try {
          const results = await searchUsers(nameRaw, false);
          if (results && results.length > 0) {
            const best = results[0];
            email = (best.email || '').toLowerCase();
            displayName = best.displayName || nameRaw;
            jobTitle = best.jobTitle || '';
            location = best.location || '';
          }
        } catch { /* */ }
      }
      const { firstName, lastName } = splitName(displayName);
      let status: 'ok' | 'duplicate' | 'notfound';
      if (!email) status = 'notfound';
      else if (seen.has(email)) status = 'duplicate';
      else { seen.add(email); status = 'ok'; }
      rows.push({ email, firstName, lastName, jobTitle, location, status, raw: line });
    }
    setMassImportRows(rows);
    setMassImportStep('preview');
    setMassImportResolving(false);
    setMassImportProgress('');
  };

  // v18.14: Schritt 2 — die aufgelösten, gültigen Zeilen anmelden.
  const runMassImport = async (): Promise<void> => {
    if (!event || massImportBusy) return;
    const toRegister = massImportRows.filter(r => r.status === 'ok');
    if (toRegister.length === 0) return;
    const suppressMail = massImportMode === 'nomail' || massImportMode === 'silent';
    const suppressOutlook = massImportMode === 'silent';
    setMassImportBusy(true);
    setMassImportResult(null);
    let ok = 0;
    const failed: string[] = [];
    for (let i = 0; i < toRegister.length; i++) {
      const r = toRegister[i];
      setMassImportProgress(`${i + 1} / ${toRegister.length} — ${r.email}`);
      try {
        const success = (await registerForEvent(event.id, {}, r.firstName, r.lastName, r.email, undefined, { suppressMail, suppressOutlook })).ok;
        if (success) ok++; else failed.push(r.email);
      } catch { failed.push(r.email); }
    }
    setMassImportBusy(false);
    setMassImportProgress('');
    setMassImportResult({ ok, failed });
  };

  if (declined) {
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '48px 32px', maxWidth: 640, margin: '0 auto' }}>
          <h2 style={{ marginTop: 0 }}>
            {locale === 'de' ? 'Absage erfasst' : 'Decline recorded'}
          </h2>
          <p className="mt-8" style={{ color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
            {locale === 'de'
              ? <>Danke für die Rückmeldung — wir haben vermerkt, dass du <strong>nicht</strong> an &bdquo;{event?.title}&ldquo; teilnimmst. Falls sich das ändert, kannst du dich jederzeit über diese Seite anmelden.</>
              : <>Thanks for letting us know — we noted that you will <strong>not</strong> attend &bdquo;{event?.title}&ldquo;. If that changes, you can register any time via this page.</>}
          </p>
          <div style={{ marginTop: 28 }}>
            <button className="btn btn-primary" onClick={() => navigate('register')}>
              {t('reg.backtoevents') || (locale === 'de' ? 'Zurück zu Events' : 'Back to events')}
            </button>
          </div>
        </div>
      </div>
    );
  }

  // v18.73: Eigener Erfolgsscreen für den Team-Beitritt (direkt angemeldet
  // bzw. Beitritts-Anfrage gesendet) — vor der generischen Anmelde-Logik.
  if (submitted && submittedJoinKind) {
    const isReq = submittedJoinKind === 'requested';
    const headline = isReq
      ? (locale === 'de' ? 'Beitritts-Anfrage gesendet' : 'Join request sent')
      : (submittedAsWaitlist
          ? (locale === 'de' ? 'Auf der Warteliste' : 'On the waitlist')
          : (locale === 'de' ? 'Team-Beitritt erfolgreich' : 'Joined the team'));
    const body = isReq
      ? (locale === 'de'
          ? `Deine Anfrage zum Beitritt wurde an den Team-Kapitän gesendet. Sobald er entscheidet, bekommst du eine E-Mail mit dem Ergebnis. Deine Angaben werden bei der Bestätigung automatisch übernommen.`
          : `Your join request has been sent to the team lead. Once they decide, you will receive an email with the result. Your details will be applied automatically upon approval.`)
      : (submittedAsWaitlist
          ? (locale === 'de'
              ? `Das Team war voll — du stehst jetzt auf der Warteliste für „${event.title}". Sobald ein Platz frei wird, rückst du automatisch nach und bekommst eine Bestätigung. Details findest du unter „Meine Events".`
              : `The team was full — you are now on the waitlist for „${event.title}". You will be moved up automatically when a spot opens and receive a confirmation. See „My Events" for details.`)
          : (locale === 'de'
              ? `Du bist dem Team beigetreten und für „${event.title}" angemeldet. Du bekommst eine Bestätigungs-E-Mail und einen Outlook-Termin. Details findest du unter „Meine Events".`
              : `You joined the team and are registered for „${event.title}". You will receive a confirmation email and an Outlook invite. See „My Events" for details.`));
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '48px 32px', maxWidth: 720, margin: '0 auto' }}>
          {event.imageUrl && (
            <div style={{
              width: '100%', maxWidth: 480, height: 200, margin: '0 auto 24px',
              borderRadius: 'var(--dex-radius-lg)',
              background: `url(${cachedImage}) center/cover no-repeat`,
            }} />
          )}
          <h2 style={{ marginTop: 0 }}>{headline}</h2>
          <p style={{ fontSize: '0.95rem', color: 'var(--dex-gray-700)', lineHeight: 1.6, maxWidth: 560, margin: '0 auto' }}>{body}</p>
        </div>
      </div>
    );
  }

  if (submitted) {
    const sessionsOnlyHint = sessionsOnlySubmitted;
    // v26.71: Externe stellvertretende Anmeldung — die Person ist NICHT final
    // angemeldet, sondern nur in der Teilnehmerliste hinterlegt (Datenschutz-
    // rückmeldung offen). Es geht KEINE Mail an die externe Adresse und KEIN
    // Kalendereintrag; der Anmelder verschickt die Einladung selbst und
    // bestätigt danach. Deshalb hier NICHT „erfolgreich registriert" texten.
    const isExternalProxy = registerForOther && isExternalEmailAddr((email || '').trim());
    const proxyName = `${firstName} ${surname}`.trim() || email;
    const successHeadline = isExternalProxy
      ? (locale === 'de' ? 'In der Teilnehmerliste hinterlegt' : 'Added to the participant list')
      : sessionsOnlyHint
      ? (childTermPlural
          ? (locale === 'de' ? `Für ${childTermPlural} angemeldet` : `Registered for ${childTermPlural}`)
          : (t('reg.success.sessionsonly.title') || 'Für Sessions angemeldet'))
      : (submittedAsWaitlist ? t('reg.waitlisttitle') : t('reg.success'));
    const successBody = isExternalProxy
      ? (locale === 'de'
          ? `${proxyName} wurde in der Teilnehmerliste hinterlegt — mit dem Status „Angemeldet (Datenschutzrückmeldung offen)". Da „${email}" eine externe Adresse ist, versendet die App KEINE Mail dorthin und KEINEN Kalendereintrag. Du bekommst eine E-Mail (Organisator:innen in Kopie) mit einem Button, über den du den fertigen Einladungs-Entwurf direkt herunterlädst — leite ihn aus deinem eigenen Postfach an ${proxyName} weiter und bestätige nach ihrer Rückmeldung die Datenschutz-Zustimmung in der Teilnehmerliste.`
          : `${proxyName} has been added to the participant list — with status „Registered (privacy confirmation pending)". Since „${email}" is an external address, the app sends NO email there and NO calendar entry. You'll receive an email (organizers in copy) with a button to download the ready-made invitation draft directly — forward it from your own mailbox to ${proxyName}, and confirm the privacy consent in the participant list once they reply.`)
      : sessionsOnlyHint
      ? (event.subEventsOnlyMode
          ? (childTermSingular
              ? (locale === 'de'
                  ? `Du hast dich für die ausgewählten ${childTermPlural || `${childTermSingular}s`} im Rahmen von „${event.title}" angemeldet. Du bekommst pro ${childTermSingular} eine separate Bestätigungsmail und einen eigenen Outlook-Kalendereintrag.`
                  : `You registered for the selected items within "${event.title}". You will receive a separate confirmation email and Outlook calendar entry per ${childTermSingular}.`)
              : (locale === 'de'
                  ? `Du hast dich für die ausgewählten Sub-Events im Rahmen von „${event.title}" angemeldet. Du bekommst pro Sub-Event eine separate Bestätigungsmail und einen eigenen Outlook-Kalendereintrag.`
                  : `You registered for the selected sub-events within "${event.title}". You will receive a separate confirmation email and Outlook calendar entry per sub-event.`))
          : (childTermPlural && childTermSingular
              ? (locale === 'de'
                  ? `Du hast dich ausschließlich für die ausgewählten ${childTermPlural} angemeldet — NICHT für das Haupt-Event „${event.title}". Du bekommst pro ${childTermSingular} eine separate Bestätigungsmail und einen eigenen Outlook-Kalendereintrag.`
                  : `You registered exclusively for the selected ${childTermPlural} — NOT for the main event "${event.title}". You will receive a separate confirmation email and Outlook calendar entry per ${childTermSingular}.`)
              : (t('reg.success.sessionsonly.msg') || 'Du hast dich ausschließlich für die ausgewählten Sessions angemeldet — NICHT für das Haupt-Event "{title}". Du bekommst pro Session eine separate Bestätigungsmail und einen eigenen Outlook-Kalendereintrag.').replace('{title}', event.title)))
      : (submittedAsWaitlist
          ? (registerForOther
              ? t('reg.waitlistmsg.other').replace('{name}', `${firstName} ${surname}`.trim()).replace('{title}', event.title).replace('{email}', email)
              : t('reg.waitlistmsg').replace('{title}', event.title))
          : (() => {
              // v28.16: Statt der nackten E-Mail-Adresse konkret sagen, was
              // automatisch verschickt wird — Mail-Bestätigung und Outlook-
              // Termin jeweils nur, wenn sie für das Event aktiv sind.
              const mailActive = !event.disableEmails && !event.disableRegistrationEmail;
              const outlookActive = !event.disableOutlook && !!event.startDate;
              let confirmTail = '';
              if (registerForOther) {
                const who = `${firstName} ${surname}`.trim() || email;
                if (mailActive && outlookActive) confirmTail = locale === 'de'
                  ? ` ${who} erhält automatisch eine Bestätigung per E-Mail (an ${email}) sowie einen Outlook-Kalendereintrag.`
                  : ` ${who} will automatically receive a confirmation email (to ${email}) and an Outlook calendar invitation.`;
                else if (mailActive) confirmTail = locale === 'de'
                  ? ` ${who} erhält automatisch eine Bestätigung per E-Mail an ${email}.`
                  : ` ${who} will automatically receive a confirmation email to ${email}.`;
                else if (outlookActive) confirmTail = locale === 'de'
                  ? ` ${who} erhält automatisch einen Outlook-Kalendereintrag.`
                  : ` ${who} will automatically receive an Outlook calendar invitation.`;
              } else {
                if (mailActive && outlookActive) confirmTail = locale === 'de'
                  ? ' Du erhältst automatisch eine Bestätigung per E-Mail sowie einen Outlook-Kalendereintrag.'
                  : ' You will automatically receive a confirmation email and an Outlook calendar invitation.';
                else if (mailActive) confirmTail = locale === 'de'
                  ? ' Du erhältst automatisch eine Bestätigung per E-Mail.'
                  : ' You will automatically receive a confirmation email.';
                else if (outlookActive) confirmTail = locale === 'de'
                  ? ' Du erhältst automatisch einen Outlook-Kalendereintrag.'
                  : ' You will automatically receive an Outlook calendar invitation.';
              }
              const base = registerForOther
                ? t('reg.successmsg.other').replace('{name}', `${firstName} ${surname}`.trim()).replace('{title}', event.title)
                : t('reg.successmsg').replace('{title}', event.title);
              return base + confirmTail;
            })());
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '48px 32px', maxWidth: 720, margin: '0 auto' }}>
          {/* v15.20: Event-Foto + Organizer-Info auch auf der
              Erfolgs-Seite anzeigen — analog zur Event-Karte. */}
          {event.imageUrl && (
            <div style={{
              width: '100%', maxWidth: 480, height: 200, margin: '0 auto 24px',
              borderRadius: 'var(--dex-radius-lg)',
              background: `url(${cachedImage}) center/cover no-repeat`,
            }} />
          )}
          <h2 style={{ marginTop: 0 }}>{successHeadline}</h2>
          {/* v18.54: Im subEventsOnlyMode strukturierter Bestätigungstext —
              Begrüßung, Verweis auf das (nicht anwählbare) Hauptevent, Bullet-
              Liste der gewählten Sections (dynamische Organizer-Bezeichnung) und
              der Mail/Outlook-Satz NUR wenn für die gewählten Sections wirklich
              Mail bzw. Outlook aktiv ist. */}
          {!isExternalProxy && sessionsOnlyHint && event.subEventsOnlyMode ? (() => {
            const selectedChildren = childEvents.filter(ce => selectedSessions.has(ce.id));
            const anyEmail = selectedChildren.some(ce => !ce.disableEmails);
            const anyOutlook = selectedChildren.some(ce => !ce.disableOutlook);
            const sectionPlural = childTermPlural || (locale === 'de' ? 'Event-Sections' : 'event-sections');
            const sectionSingular = childTermSingular || (locale === 'de' ? 'Event-Section' : 'event-section');
            const greetingName = (firstName || '').trim();
            let confirmLine = '';
            if (anyEmail && anyOutlook) confirmLine = locale === 'de'
              ? `Du erhältst pro ${sectionSingular} eine E-Mail-Bestätigung und einen Outlook-Termin.`
              : `You will receive a confirmation email and an Outlook invitation per ${sectionSingular}.`;
            else if (anyEmail) confirmLine = locale === 'de'
              ? `Du erhältst pro ${sectionSingular} eine E-Mail-Bestätigung.`
              : `You will receive a confirmation email per ${sectionSingular}.`;
            else if (anyOutlook) confirmLine = locale === 'de'
              ? `Du erhältst pro ${sectionSingular} einen Outlook-Termin.`
              : `You will receive an Outlook invitation per ${sectionSingular}.`;
            return (
              <div className="mt-8" style={{ color: 'var(--dex-gray-700)', textAlign: 'left', maxWidth: 520, margin: '8px auto 0', lineHeight: 1.6 }}>
                <p style={{ margin: '0 0 10px' }}>
                  {locale === 'de'
                    ? <>Hallo{greetingName ? <> <strong>{greetingName}</strong></> : ''},</>
                    : <>Hi{greetingName ? <> <strong>{greetingName}</strong></> : ''},</>}
                </p>
                <p style={{ margin: '0 0 10px' }}>
                  {locale === 'de'
                    ? <>du hast dich erfolgreich für das <strong>{event.title}</strong> angemeldet. Wir haben deine Anmeldung für die folgenden {sectionPlural} erhalten:</>
                    : <>you have successfully registered for <strong>{event.title}</strong>. We received your registration for the following {sectionPlural}:</>}
                </p>
                <ul style={{ margin: '0 0 10px', paddingLeft: 22 }}>
                  {selectedChildren.map(ce => {
                    // v19.33: nur den reinen Section-Namen zeigen (Parent-Präfix
                    // „<Hauptevent> | …" strippen) + dahinter „ | <Datum>".
                    const full = (ce.title || '').trim();
                    const pipe = full.lastIndexOf('|');
                    const name = (pipe >= 0 ? full.substring(pipe + 1).trim() : full) || (locale === 'de' ? 'ohne Titel' : 'untitled');
                    const date = ce.startDate ? formatDate(ce.startDate) : '';
                    return (
                      <li key={ce.id} style={{ marginBottom: 3 }}>{date ? `${name} | ${date}` : name}</li>
                    );
                  })}
                </ul>
                {confirmLine && <p style={{ margin: 0 }}>{confirmLine}</p>}
              </div>
            );
          })() : (
            <p className="mt-8" style={{ color: 'var(--dex-gray-600)' }}>{successBody}</p>
          )}
          {(() => {
            // v24.15: „Organizer ausblenden" ohne Einzel-Modus = ALLE aus.
            if (event.hideOrganizer && !event.hideOrganizerIndividualOnly) return null;
            // Organizer als Chips mit Foto (gleicher Stil wie auf der
            // Anmelde-Seite). „Nachname, Vorname" → „Vorname Nachname".
            const orgs = event.organizers.reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map(o => {
              const trimmed = o.trim();
              const parts = trimmed.split(',').map(s => s.trim());
              return parts.length === 2 ? `${parts[1]} ${parts[0]}` : trimmed;
            }).filter(Boolean);
            if (orgs.length === 0) return null;
            return (
              <div style={{ marginTop: 24, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8 }}>
                <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, fontWeight: 600 }}>Organizer</div>
                <OrganizerList names={orgs} emails={event.organizerEmails} hiddenEmails={(event.hideOrganizer && event.hideOrganizerIndividualOnly) ? event.hiddenOrganizerEmails : []} forceIsDe={locale === 'de'} size="md" display={event.organizerDisplayLarge ? 'card' : 'chip'} nameFontSize="1.05rem" hideContactPrompt={!!(event.contactName || event.contactEmail || event.contactInfo)} />
              </div>
            );
          })()}
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center', flexWrap: 'wrap' }}>
            <button className="btn btn-primary" onClick={() => navigate('my-events')}>
              {t('myevents.title')}
            </button>
            <button className="btn btn-secondary" onClick={() => navigate('register')}>
              {t('reg.registeranother')}
            </button>
          </div>
        </div>
      </div>
    );
  }

  // v26.76: gemeinsame Auswahl-Logik für die stellvertretende Anmeldung —
  // wird sowohl im Wizard als auch (Fallback) inline genutzt. Übernimmt Name/
  // E-Mail aus dem Suchtreffer, lädt das Profil nach und prüft Doppel-Anmeldung
  // + Gästekreis (thirdPartyCheck).
  const pickProxyUser = (u: { email: string; displayName: string; location?: string; jobTitle?: string }): void => {
    let uFirstName = '';
    let uSurname = '';
    if (u.displayName.includes(',')) {
      const parts = u.displayName.split(',').map(s => s.trim());
      uSurname = parts[0] || '';
      uFirstName = parts[1] || '';
    } else {
      const parts = u.displayName.split(' ');
      uFirstName = parts[0] || '';
      uSurname = parts.slice(1).join(' ') || '';
    }
    setFirstName(uFirstName);
    setSurname(uSurname);
    setEmail(u.email);
    setUserSearch(u.displayName);
    setUserResults([]);
    setPickedUserProfile({ jobTitle: u.jobTitle || '', location: u.location || '' });
    searchUser(u.email).then(p => {
      if (p) {
        setPickedUserProfile({
          jobTitle: p.jobTitle || u.jobTitle || '',
          department: p.department || '',
          location: p.location || u.location || '',
          mobilePhone: p.mobilePhone || '',
          company: p.company || '',
        });
      }
    }).catch(() => { /* silent */ });
    setThirdPartyCheck(null);
    if (event) {
      (async () => {
        const existing = await checkRegistrationByEmail(event.id, u.email).catch(() => null);
        const alreadyRegistered = !!existing && existing.Status !== 'Abgemeldet';
        const notInAudience = !isEventVisibleForUser(event, u.email, u.location || '', [], u.jobTitle || '');
        setThirdPartyCheck({
          alreadyRegistered,
          notInAudience,
          registeredName: (existing && (existing.ParticipantName || `${existing.Vorname || ''} ${existing.Nachname || ''}`.trim())) || u.displayName || '',
          registeredDate: (existing && existing.RegistrationDate) || '',
        });
      })();
    }
  };
  // v11.5: Custom-Field-Renderer extrahiert — wird zweimal verwendet:
  // einmal direkt in der Gruppen-Auswahl-Box für Felder mit
  // onlyForGroup-Constraint, einmal im Eventspez-2-Spalten-Grid für
  // alle anderen Felder. Die Filter-Logik dazu unten in den
  // groupSpecificFields- bzw. generalFields-Konstanten.
  // v13.2: fRaw jetzt typsicher als EventSpecificField (vorher any).
  const renderRegField = (fRaw: EventSpecificField, store?: Record<string, string>, setStore?: (next: Record<string, string>) => void, rowIndex?: number, rowList?: EventSpecificField[]): React.ReactElement => {
    // v18.12: optionaler Wert-Store — für die Custom-Fields pro Team-Mitglied.
    // Default = eventSpecific/setEventSpecific (Lead bzw. Solo-Anmeldung).
    const vals = store || eventSpecific;
    const setVals = setStore || setEventSpecific;
    // Dynamisch Required erzwingen: bei aktivem Infoservice ist die
    // Mobilnummer Pflicht.
    let field: EventSpecificField = fRaw;
    // Mobilnummer bei aktiviertem Infoservice dynamisch zur Pflicht
    if (fRaw.id === 'b2run_mobilnummer' && vals['b2run_infoservice'] === 'true') {
      field = { ...field, required: true };
    }
    // v26.96: Die frühere automatische Injektion von AGB-/Datenschutz-Links
    // für b2run_datenschutz wurde entfernt — der Organizer hinterlegt Links
    // jetzt selbst (z.B. direkt in der Feld-Beschreibung mit dem Editor).
    // Laufshirt/T-Shirt-Feld bei B2Run ist Pflicht (falls in alten Events
    // noch nicht so markiert)
    if ((fRaw.id === 'b2run_laufshirt' || /laufshirt/i.test(fRaw.label || '')) && !fRaw.required) {
      field = { ...field, required: true };
    }
    // v17.20: vor jedem Render die EN-Variante einziehen, sofern verfügbar.
    const displayLabel = pickFieldLabel(field);
    const displayHelp = pickFieldHelp(field);
    const displayConfirmLabel = pickFieldConfirmLabel(field);
    // v18.18: 'inline' = Erklär-Text unter dem Label (nicht fett), sonst
    // weiterhin "i"-Hover-Box neben dem Label.
    const isInlineHelp = field.helpTextStyle === 'inline';
    const inlineHelpEl = (displayHelp && isInlineHelp)
      // v18.77: Inline-Hilfe reserviert mind. 2 Zeilen Höhe (minHeight). Dadurch
      // stehen die Eingaben benachbarter Felder auf gleicher Höhe, wenn sich die
      // Beschreibungen um eine Zeile unterscheiden — OHNE die Nachbar-Eingabe
      // (wie zuvor mit flexGrow) bis ganz nach unten zu ziehen, was bei Feldern
      // mit Inhalt UNTER der Eingabe (People-Picker mit „international suchen") zu
      // großen Lücken führte.
      // v26.91: Beschreibung darf **fett** + Links enthalten (renderFieldDescHtml
      // escaped alles andere — der Organizer-Text ist sicherer Origin).
      ? <div style={{ fontSize: '0.78rem', fontWeight: 400, color: 'var(--dex-gray-500)', lineHeight: 1.45, marginTop: 2, marginBottom: 6, minHeight: '2.9em' }} dangerouslySetInnerHTML={{ __html: renderFieldDescHtml(displayHelp) }} />
      : null;
    // v26.16: Felder OHNE Inline-Beschreibung bekommen einen leeren Platzhalter
    // gleicher Höhe, SOBALD irgendein Feld im Formular eine Inline-Beschreibung
    // hat — damit stehen die Eingaben benachbarter Felder im 2-Spalten-Grid auf
    // gleicher Höhe (z.B. „Dressing" auf Höhe von „Gerichtauswahl"). Nur dann,
    // damit ohne Beschreibungen keine unnötigen Lücken entstehen.
    // v26.91: Der leere Platzhalter für eine fehlende Inline-Beschreibung wird
    // jetzt PRO ZEILE entschieden: nur reservieren, wenn der NEBEN diesem Feld
    // stehende Partner in derselben 2-Spalten-Zeile eine Beschreibung hat — sonst
    // (beide Felder ohne Beschreibung) entsteht keine leere Lücke mehr. Kennt der
    // Aufrufer die Zeile nicht (andere Render-Kontexte), gilt das bisherige
    // globale Verhalten (irgendein Feld hat eine Inline-Beschreibung).
    const fieldHasInlineDesc = (ff: EventSpecificField): boolean => ff.helpTextStyle === 'inline' && !!pickFieldHelp(ff);
    let reserveHelpSpace: boolean;
    if (typeof rowIndex === 'number' && rowList) {
      const partnerIdx = rowIndex % 2 === 0 ? rowIndex + 1 : rowIndex - 1;
      const partner = rowList[partnerIdx];
      reserveHelpSpace = !!(partner && fieldHasInlineDesc(partner));
    } else {
      reserveHelpSpace = (event?.eventSpecificFields || []).some(fieldHasInlineDesc);
    }
    const inlineHelpSlot = inlineHelpEl || (reserveHelpSpace
      // v26.17: Der leere Platzhalter muss dieselben Font-Metriken (fontSize/
      // lineHeight) wie der echte Inline-Hilfetext tragen, da sich 'minHeight'
      // in 'em' auf die EIGENE font-size bezieht. Ohne fontSize erbte der
      // Platzhalter die größere .form-group-Schrift und reservierte ~8px mehr
      // Höhe → die Eingabe eines Feldes OHNE Beschreibung stand tiefer als die
      // des Nachbarfeldes MIT Beschreibung (z.B. „Food Allergies" vs. „Hotel").
      ? <div aria-hidden="true" style={{ fontSize: '0.78rem', lineHeight: 1.45, marginTop: 2, marginBottom: 6, minHeight: '2.9em' }} />
      : null);
    // v19.0: Ausgefüllte Felder bekommen die gleiche grüne Hervorhebung wie die
    // ausgewählten Event-Sections (grüner Rand + zarter grüner Hintergrund).
    const fieldVal = vals[field.id];
    const isFieldFilled = field.type === 'checkbox' ? fieldVal === 'true' : !!(fieldVal && fieldVal.trim());
    const greenFilledStyle: React.CSSProperties = { borderColor: 'var(--dex-green, #86bc25)', boxShadow: '0 0 0 1px var(--dex-green, #86bc25) inset', background: 'rgba(134,188,37,0.06)' };
    const isErrEmpty = !!(showErrors && field.required && (field.type === 'checkbox' ? fieldVal !== 'true' : !fieldVal?.trim()));
    const inputStyleGreen: React.CSSProperties = isErrEmpty ? errorBorder : (isFieldFilled ? greenFilledStyle : {});
    return (
  <div className="form-group" key={field.id}>
    {field.type !== 'checkbox' && (
      <>
      <label className="form-label">
        {field.required && <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>}
        {displayLabel}
        {/* v9.17: konsistenter InfoTooltip statt simples i-Icon —
            gibt schönes Hover-Popover mit der vom Organizer
            beim Event-Anlegen hinterlegten Beschreibung.
            v18.18: nur im 'tooltip'-Modus; 'inline' rendert darunter. */}
        {displayHelp && !isInlineHelp && <InfoTooltip text={displayHelp.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim()} />}
      </label>
      {inlineHelpSlot}
      </>
    )}
    {field.type === 'select' && field.multi ? (
      // v11.89: Multi-Select-Dropdown — gleicher Look wie Single-Select,
      // beim Aufklappen Checkboxen pro Option. Werte werden weiterhin
      // " | "-getrennt im selben Feld vals[field.id]
      // gespeichert (kompatibel mit Record<string,string>).
      (() => {
        const sep = ' | ';
        const raw = (vals[field.id] || '').trim();
        const selected = raw ? raw.split(sep).map(s => s.trim()).filter(Boolean) : [];
        const isErr = !!(showErrors && field.required && selected.length === 0);
        // v17.20: Anzeige-Labels für den EN-Modus mappen. Der gespeicherte
        // Wert bleibt weiterhin der DE-Wert (positional gemappt), damit alle
        // anderen Stellen (Mails, Excel-Export, Admin-Center) konsistent
        // bleiben — wir tauschen ausschliesslich das Display-Label.
        const displayOptions = (field.options || []).map((o, i) => ({
          value: o,
          label: pickOptionLabel(field, i, o),
        }));
        return (
          <MultiSelectDropdown
            options={field.options || []}
            optionLabels={useEnVariants ? displayOptions.map(d => d.label) : undefined}
            value={selected}
            onChange={next => setVals({ ...vals, [field.id]: next.join(sep) })}
            placeholder={tEvent('reg.pleaseselect')}
            error={isErr}
          />
        );
      })()
    ) : field.type === 'select' && field.optionCategories && field.optionCategories.some(c => (c || '').trim()) ? (
      // v26.75: Vorfilter — zuerst Kategorie wählen, dann nur die passenden
      // Optionen zeigen (kürzere Liste). Die Kategorie-Auswahl liegt transient
      // in vals['<id>__cat']; gespeichert wird nur der eigentliche Optionswert.
      (() => {
        // v26.96: EINE Kombibox statt zwei Auswahlfeldern — die Kategorien sind
        // <optgroup>-Überschriften, darunter die zugehörigen Optionen. Optionen
        // OHNE Kategorie (leer = immer sichtbar) stehen ungruppiert am Ende.
        const cats = field.optionCategories || [];
        const opts = field.options || [];
        const distinctCats = Array.from(new Set(cats.map(c => (c || '').trim()).filter(Boolean)));
        return (
          <select className="form-select" value={vals[field.id] || ''} onChange={e => setVals({ ...vals, [field.id]: e.target.value })} style={inputStyleGreen}>
            <option value="" disabled hidden>{tEvent('reg.pleaseselect')}</option>
            {distinctCats.map(cat => (
              <optgroup key={cat} label={cat}>
                {opts.map((opt, i) => ((cats[i] || '').trim() === cat && (opt || '').trim())
                  ? <option key={`${cat}-${i}`} value={`${cat} ${opt}`}>{cat} {pickOptionLabel(field, i, opt)}</option>
                  : null)}
              </optgroup>
            ))}
            {opts.map((opt, i) => ((cats[i] || '').trim() === '' && (opt || '').trim())
              ? <option key={`nocat-${i}`} value={opt}>{pickOptionLabel(field, i, opt)}</option>
              : null)}
          </select>
        );
      })()
    ) : field.type === 'select' ? (
      <select className="form-select" value={vals[field.id] || ''} onChange={e => setVals({ ...vals, [field.id]: e.target.value })} style={inputStyleGreen}>
        <option value="" disabled hidden>{tEvent('reg.pleaseselect')}</option>
        {field.options && field.options.map((opt, i) => <option key={opt} value={opt}>{pickOptionLabel(field, i, opt)}</option>)}
      </select>
    ) : field.type === 'user' || field.type === 'roommate' ? (
      // v7.17: 'roommate' nutzt denselben Picker wie 'user' — der
      // einzige Unterschied ist dass 'roommate' beim Anmelden
      // automatisch eine Zimmerpartner-Mail an die ausgewählte
      // Person triggert (siehe EventContext). 'user' ist der
      // generische Personen-Picker ohne Mail-Versand.
      <UserFieldPicker
        value={vals[field.id] || ''}
        onChange={v => setVals({ ...vals, [field.id]: v })}
        searchUsers={searchUsers}
        searchUserByEmail={searchUser}
        placeholder={tEvent('reg.userfield.placeholder')}
        errorStyle={showErrors && field.required && !vals[field.id]?.trim() ? errorBorder : {}}
        // v26.60: „Person wird benachrichtigt"-Hinweis nur, wenn die
        // Zimmerpartner-Anfrage-Mail nicht abgeschaltet wurde.
        hint={field.type === 'roommate' && field.notifyRoommate !== false ? tEvent('reg.userfield.notifyhint') : undefined}
        forcedIsDe={locale === 'de'}
      />
    ) : field.type === 'checkbox' ? (
      // v11.91: Checkbox bekommt jetzt eine ordentliche Karten-Box mit
      // gleichem Look wie die Dropdown-Inputs — vorher war die Mini-
      // Checkbox neben den Dropdowns visuell „verloren". Der Label-Text
      // sitzt oben (analog zu den anderen Feldern, damit die Zeilen
      // horizontal aligned sind), drinnen ein deutlich vergrößerter
      // Checkbox + kurzer „Ja, bestätigen"-Hinweis.
      <>
        <label className="form-label">
          {field.required && <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>}
          {displayLabel}
          {displayHelp && !isInlineHelp && <InfoTooltip text={displayHelp.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim()} />}
        </label>
        {inlineHelpSlot}
        <label
          style={{
            display: 'flex', alignItems: 'center', gap: 10,
            cursor: 'pointer',
            // v11.93: exakt gleiche Höhe wie .form-select — 12px 16px Padding,
            // 1.5px Border, 12px Radius. Vorher 44px minHeight = optisch
            // höher als die Dropdowns daneben.
            padding: '12px 16px',
            border: showErrors && field.required && vals[field.id] !== 'true'
              ? '1.5px solid var(--dex-red)'
              : '1.5px solid var(--dex-gray-200)',
            borderRadius: 12,
            background: vals[field.id] === 'true' ? 'rgba(134,188,37,0.10)' : 'var(--dex-white, #fff)',
            transition: 'background 0.12s',
          }}
        >
          <input
            type="checkbox"
            checked={vals[field.id] === 'true'}
            onChange={e => setVals({ ...vals, [field.id]: e.target.checked ? 'true' : 'false' })}
            style={{ width: 16, height: 16, accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer', flexShrink: 0 }}
          />
          <span style={{ fontSize: '0.95rem', color: 'var(--dex-gray-800)' }}>
            {/* v11.94: Organizer kann den Text neben der Checkbox im Wizard
                pro Feld setzen (field.confirmLabel). Default: „Ja, bestätigen".
                v17.20: pickFieldConfirmLabel zieht im EN-Modus den
                confirmLabelEn-Wert; fällt sonst auf den DE-Wert. */}
            {(displayConfirmLabel && displayConfirmLabel.trim())
              || (eventLocale === 'de' ? 'Ja, bestätigen' : 'Yes, confirm')}
          </span>
        </label>
        {field.externalLinks && field.externalLinks.length > 0 && (
          <div style={{ marginTop: 4, fontSize: '0.78rem' }}>
            {field.externalLinks.map((l, i) => (
              <span key={l.url}>
                {i > 0 && <span style={{ color: 'var(--dex-gray-300)', margin: '0 6px' }}>|</span>}
                <a
                  href={l.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  onClick={e => e.stopPropagation()}
                  style={{ color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'underline' }}
                >
                  {l.label}
                </a>
              </span>
            ))}
          </div>
        )}
      </>
    ) : field.type === 'document' ? (
      // v19.0: Dokument-Upload (PDF/Bild). Datei wird in pendingDocFiles
      // gehalten und nach erfolgreicher Anmeldung als Attachment hochgeladen.
      // Im Team-Mitglied-Kontext (store gesetzt) nicht unterstützt.
      store ? (
        <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
          {locale === 'de' ? 'Dokument-Upload pro Team-Mitglied wird nicht unterstützt.' : 'Per-member document upload is not supported.'}
        </div>
      ) : (() => {
        const picked = pendingDocFiles[field.id] || null;
        const docErr = !!(showErrors && field.required && !picked && !parentAlreadyRegistered);
        return (
          <div>
            {picked ? (
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', border: '1.5px solid var(--dex-green, #86bc25)', background: 'rgba(134,188,37,0.06)', borderRadius: 12 }}>
                <Icon iconName="Attach" style={{ fontSize: 16, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                <span style={{ flex: 1, fontSize: '0.88rem', wordBreak: 'break-all' }}>{picked.name}</span>
                <button
                  type="button"
                  onClick={() => setPendingDocFiles(prev => ({ ...prev, [field.id]: null }))}
                  title={locale === 'de' ? 'Datei entfernen' : 'Remove file'}
                  style={{ background: 'var(--dex-gray-200)', border: 'none', color: 'var(--dex-gray-700)', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.9rem', lineHeight: 1, flexShrink: 0 }}
                >×</button>
              </div>
            ) : (
              <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, padding: '10px 16px', border: `1.5px dashed ${docErr ? 'var(--dex-red)' : 'var(--dex-gray-300)'}`, borderRadius: 12, cursor: 'pointer', fontSize: '0.88rem', color: 'var(--dex-gray-600)' }}>
                <Icon iconName="Upload" style={{ fontSize: 16 }} />
                {locale === 'de' ? 'Datei wählen (PDF, JPG, PNG)' : 'Choose file (PDF, JPG, PNG)'}
                <input
                  type="file"
                  accept="application/pdf,image/png,image/jpeg,.pdf,.png,.jpg,.jpeg"
                  style={{ display: 'none' }}
                  onChange={e => {
                    const f = e.target.files?.[0] || null;
                    if (f && f.size > 10 * 1024 * 1024) {
                      // eslint-disable-next-line no-alert
                      showAlert(locale === 'de' ? 'Die Datei ist zu groß (max. 10 MB).' : 'The file is too large (max. 10 MB).', { variant: 'error' });
                      e.target.value = '';
                      return;
                    }
                    setPendingDocFiles(prev => ({ ...prev, [field.id]: f }));
                  }}
                />
              </label>
            )}
            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>
              {locale === 'de'
                ? 'Wird nach dem Absenden hochgeladen. Du kannst es auch später über „Meine Events" ergänzen oder ersetzen.'
                : 'Uploaded after submitting. You can also add or replace it later via „My Events".'}
            </div>
          </div>
        );
      })()
    ) : field.type === 'daterange' ? (
      // v28.63: Übernachtungs-Zeitraum — Anreise + Abreise in einem Feld, die
      // Nächte ergeben sich daraus. Ersetzt die frühere Kombination aus
      // „Hotel: ja/nein" plus „Zusatznächte"-Auswahlliste; die Hotel-Planung
      // im Organizer Center liest den Zeitraum direkt aus.
      <StayRangePicker
        value={vals[field.id] || ''}
        onChange={(next: string) => setVals({ ...vals, [field.id]: next })}
        isDe={locale === 'de'}
        rangeStart={field.rangeStart}
        rangeEnd={field.rangeEnd}
        maxNights={field.maxNights}
        required={field.required}
      />
    ) : field.type === 'date' ? (
      // v24.25: Datums-Feld — Kalender-Auswahl. Mit withTime zusätzlich Uhrzeit
      // (datetime-local). Der Wert wird als String gespeichert (wie alle
      // Custom-Field-Antworten).
      <input
        className="form-input"
        type={field.withTime ? 'datetime-local' : 'date'}
        value={vals[field.id] || ''}
        onChange={e => setVals({ ...vals, [field.id]: e.target.value })}
        style={inputStyleGreen}
      />
    ) : (
      <input className="form-input" value={vals[field.id] || ''} onChange={e => setVals({ ...vals, [field.id]: e.target.value })} placeholder={displayLabel} style={inputStyleGreen} />
    )}
  </div>
    );
  };

  return (
    <div className="page-container">
      {showLocationBanner && (
        <div style={{
          padding: '10px 16px', marginBottom: 16, borderRadius: 'var(--dex-radius-md)',
          background: 'rgba(237,139,0,0.1)', border: '1px solid var(--dex-orange)',
          color: 'var(--dex-orange)', fontSize: '0.85rem',
        }}>
          {t('reg.locationnotice')}
          {event && event.locationAudience.length > 0 && <> {t('reg.locationfilter')}: <strong>{event.locationAudience.join(', ')}</strong>.</>}
          {/* v9.17: bei Einzel-E-Mail-Whitelists in audienceFilter würden bei
              größeren Verteilern (50+ Adressen) der Banner zugekleistert.
              Statt alle Mails auflisten: nur die Anzahl + die ersten 3
              Adressen zeigen, der Rest als "+N weitere". Gruppen-/Group-Namen
              (ohne "@") werden weiterhin alle aufgeführt — die sind kurz. */}
          {event && event.audienceFilter && event.audienceFilter.length > 0 && (() => {
            const items = event.audienceFilter;
            const emails = items.filter(s => s.includes('@'));
            const groups = items.filter(s => !s.includes('@'));
            const showLabel = (() => {
              if (emails.length === 0) return groups.join(', ');
              if (emails.length <= 3) return [...groups, ...emails].join(', ');
              const head = emails.slice(0, 3).join(', ');
              const more = emails.length - 3;
              const tail = `${head} (+${more} ${t('reg.audience.more') || 'weitere E-Mail-Adressen'})`;
              return groups.length > 0 ? `${groups.join(', ')}, ${tail}` : tail;
            })();
            return <> {t('reg.audience')}: <strong>{showLabel}</strong>.</>;
          })()}
          {event && event.filterMode === 'AND' && <> ({t('reg.andmode')})</>}
          {' '}{t('reg.yourlocation')}: {currentUser.location || t('reg.unknown')}.
        </div>
      )}
      {/* Deadline-Banner für Organizer/Admin: die Registrierungsfrist ist abgelaufen,
          das Formular wird aber trotzdem angezeigt (Admin/Organizer darf weiter registrieren).
          Der Ton entspricht dem Location-Banner: "als normaler User könntest du dich nicht registrieren". */}
      {isDeadlinePassed && (isOrganizer || isAdmin) && (
        <div style={{
          padding: '10px 16px', marginBottom: 16, borderRadius: 'var(--dex-radius-md)',
          background: 'rgba(237,139,0,0.1)', border: '1px solid var(--dex-orange)',
          color: 'var(--dex-orange)', fontSize: '0.85rem',
        }}>
          {/* v22.55: Nur wenn ALLES zu ist ("kein User käme mehr rein") den
              harten Hinweis zeigen. Sind Sub-Events noch offen, kann sich ein
              normaler User weiterhin für diese anmelden — dann ein zutreffender
              Hinweis statt der irreführenden "keine Anmeldung mehr"-Meldung. */}
          {isFullyClosed ? (
            <>
              {t('reg.deadlinepassed.adminnotice')}
              {event && (event.klammerDeadline || event.registrationDeadline) && (
                <> {t('reg.deadlinepassed.date')}: <strong>{new Date(event.klammerDeadline || event.registrationDeadline).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}</strong>.</>
              )}
            </>
          ) : (
            locale === 'de' ? (
              <>Hinweis: Die Anmeldefrist des Hauptevents ist abgelaufen{event.registrationDeadline ? <> (war <strong>{new Date(event.registrationDeadline).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}</strong>)</> : ''} — die noch offenen Sub-Events sind aber weiterhin buchbar, auch für reguläre User.</>
            ) : (
              <>Note: The main event’s registration deadline has passed{event.registrationDeadline ? <> (was <strong>{new Date(event.registrationDeadline).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}</strong>)</> : ''} — but the still-open sub-events remain bookable, also for regular users.</>
            )
          )}
        </div>
      )}
      {/* v11.33: Submit-Overlay mit Spinner + Prozent + Live-Label.
          Wird während des gesamten Anmelde-Flows (Parent + alle Sub-Events
          + Bestätigungen) eingeblendet, sodass der User auch bei langen
          Submits klares Feedback bekommt. */}
      {isSubmitting && (
        <div
          role="dialog"
          aria-modal="true"
          aria-label={locale === 'de' ? 'Anmeldung läuft' : 'Submitting registration'}
          style={{
            position: 'fixed', inset: 0, zIndex: 2000,
            background: 'rgba(0,0,0,0.55)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            padding: 16,
          }}
        >
          <div style={{
            background: '#fff', borderRadius: 12, padding: '28px 32px',
            maxWidth: 460, width: '100%',
            boxShadow: '0 16px 48px rgba(0,0,0,0.35)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16,
          }}>
            {/* v15.13: Doppel-Ladebalken entfernt — die deterministische
                Progress-Bar weiter unten (0-100% mit Label) reicht aus.
                Die zusätzliche indeterministische „Pulse"-Bar war für den
                User verwirrend. */}
            <div style={{ fontWeight: 700, fontSize: '1rem', color: 'var(--dex-gray-800)' }}>
              {locale === 'de' ? 'Anmeldung läuft …' : 'Submitting registration …'}
            </div>
            <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', textAlign: 'center', minHeight: 18 }}>
              {submitProgressLabel}
            </div>
            <div style={{ width: '100%', height: 8, borderRadius: 4, background: 'var(--dex-gray-200)', overflow: 'hidden' }}>
              <div style={{
                width: `${Math.min(100, Math.max(0, displayProgress))}%`,
                height: '100%',
                background: 'var(--dex-green, #86bc25)',
                // v29.29: kurze Übergangszeit — die Anzeige wird alle 60 ms
                // nachgezogen, eine lange Transition würde hinterherhinken.
                transition: 'width 0.15s linear',
              }} />
            </div>
            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontVariantNumeric: 'tabular-nums' }}>
              {displayProgress}%
            </div>
          </div>
          <style>{`@keyframes dexProgressSlide { 0% { left: -40%; } 100% { left: 100%; } }`}</style>
        </div>
      )}
      {/* v18: Demo-Hinweis-Banner. Das Demo-Event wird ansonsten exakt wie ein
          echtes Event gerendert (so wie es in der Realität nutzbar wäre) —
          keine künstlichen Showcase-Elemente. Nur die echte Anmeldung ist
          deaktiviert. */}
      {event && event.isDemoShowcase && (
        <div style={{
          padding: '10px 16px', marginBottom: 16, borderRadius: 'var(--dex-radius, 12px)',
          background: 'rgba(0,118,168,0.08)', border: '1px solid var(--dex-blue, #0076a8)',
          color: 'var(--dex-gray-800)', fontSize: '0.85rem', lineHeight: 1.55,
          display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap',
        }}>
          <span style={{
            padding: '2px 8px', borderRadius: 999, background: 'var(--dex-blue, #0076a8)',
            color: '#fff', fontSize: '0.66rem', fontWeight: 700, letterSpacing: 1,
          }}>DEMO</span>
          {locale === 'de'
            ? <span>Dies ist ein <strong>Demo-Event</strong> — es wird genau so angezeigt wie ein echtes Event. Du kannst die Anmeldemaske ansehen, aber <strong>keine echte Anmeldung</strong> absenden.</span>
            : <span>This is a <strong>demo event</strong> — shown exactly like a real one. You can explore the registration form, but <strong>cannot submit a real registration</strong>.</span>}
        </div>
      )}
      <div className="registration-layout">
        {/* v28.2 „Geführte Schritte": Station 1 — Dein Event. Der frühere
            „Ausgewähltes Event"-Pill-Header entfällt (Step-Label ersetzt ihn). */}
        <div className="reg-step-head">
          <span className="reg-step-num">1</span>
          <span className="reg-step-label">{locale === 'de' ? 'Dein Event' : 'Your event'}</span>
        </div>
        <div
          className="registration-event"
          // v28.7: Kreis-Notch — der Kreis ragt über die Oberkante der Karte
          // hinaus. Dafür muss das overflow:hidden der Karte weichen und
          // oben Platz für den überstehenden Halbkreis geschaffen werden.
          // v28.98: DAS hier war der Grund, warum der Foto-Platzhalter oben
          // abgeschnitten wurde — er nutzt seit v28.91 dasselbe Notch-Layout,
          // stand aber nicht in dieser Bedingung. Ohne den zusätzlichen
          // Platz und mit dem overflow:hidden der Karte wird der überstehende
          // Halbkreis schlicht weggeschnitten. (Meine früheren Erklärungen —
          // negativer Rand, Vorschau-Banner — waren beide falsch.)
          style={(imgCircleNotch || showOrbPlaceholder) ? { overflow: 'visible', marginTop: circleSize / 2 } : undefined}
        >
          <div
            className="registration-event__card"
            style={{
              display: 'flex',
              // v28.3 („Geführte Schritte"): Desktop-Standard = Bild kompakt
              // links + Inhalt rechts. v28.5: Der Organizer kann im Wizard
              // stattdessen „Banner"-Layout wählen (event.imageBanner) — dann
              // liegt das Bild in voller Kartenbreite ÜBER den Infos (gut für
              // breite Querformat-Fotos). Handy: immer Bild oben.
              // v28.6: Infos LINKS, Bild RECHTS (row-reverse — das Bild steht
              // im DOM zuerst, wird aber rechts gerendert). Banner/Mobil: Bild oben.
              // v28.7: Kreis-Bilder → Spalte, der Kreis sitzt oben mittig.
              flexDirection: (isMobile || event.imageBanner || imgCircleNotch || showOrbPlaceholder) ? 'column' : 'row-reverse',
              gap: 16,
              alignItems: (isMobile || event.imageBanner || imgCircleNotch || showOrbPlaceholder) ? 'stretch' : 'flex-start',
            }}
          >
            {/* v28.3: Bild-Slot nur rendern, wenn das Event ein Bild hat —
                sonst stünde links ein leerer 300px-Block.
                v28.19: … und erst, wenn die Bildform-Analyse fertig ist
                (imgAspectReady) — sonst startete das Bild kurz im Seiten-Slot
                rechts und sprang dann in den Kreis. Banner-Layout hängt nicht
                von der Form ab und rendert sofort. */}
            {/* v28.90: Ohne Event-Foto blieb der Bild-Slot leer und die Karte
                sah anders aus als bei Events mit Bild — der erste Eindruck der
                Anmeldeseite hing damit daran, ob jemand ein Foto hochgeladen
                hat. Statt Leerraum steht dort jetzt das DEX-Logo. Nur auf dem
                Desktop: Auf dem Handy liegt das Bild ÜBER den Infos und würde
                Titel und Datum nach unten drücken. Kein Zoom-Knopf — es gibt
                nichts zu vergrößern. */}
            {showOrbPlaceholder && (
              <div
                className="registration-event__image"
                title={locale === 'de' ? 'Für dieses Event ist kein Bild hinterlegt.' : 'No image is set for this event.'}
                style={{
                  background: '#fff',
                  position: 'relative',
                  width: circleSize, height: circleSize, flex: '0 0 auto',
                  borderRadius: '50%',
                  border: '1px solid var(--dex-gray-200)',
                  boxShadow: '0 2px 12px rgba(0,0,0,0.10)',
                  alignSelf: 'center',
                  // v28.97: Exakt dasselbe Layout wie ein rundes EVENT-Bild
                  // (imgCircleNotch): der Kreis haengt zur Haelfte in der
                  // Kartenkante. In v28.95 hatte ich den negativen Rand
                  // herausgenommen, weil der Kreis oben abgeschnitten wirkte —
                  // damit sah der Platzhalter aber als EINZIGER anders aus als
                  // alle anderen Kreis-Bilder, mit einer Luecke darunter. Zwei
                  // Darstellungen fuer dieselbe Stelle sind schlechter als
                  // eine; deshalb zurueck auf das gemeinsame Layout. Sollte
                  // der Zuschnitt wieder auftreten, liegt die Ursache im
                  // Container darueber und gehoert dort behoben, nicht hier.
                  marginTop: -(circleSize / 2 + 16),
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  padding: 14,
                  boxSizing: 'border-box',
                }}
              >
                <img
                  src={getCachedOrbBase64() || DEX_ORB_PNG}
                  alt=""
                  style={{ display: 'block', width: '100%', height: '100%', objectFit: 'contain' }}
                />
              </div>
            )}
            {heroImgUrl && (event.imageBanner || imgAspectReady) && (
            <div
              className="registration-event__image"
              // v28.12: Hover zeigt das Lupen-Icon; die Großansicht öffnet
              // erst der Klick darauf.
              onMouseEnter={() => setImgHovered(true)}
              onMouseLeave={() => setImgHovered(false)}
              style={{
                position: 'relative',
                // v11.91: Hintergrund auf Weiß gesetzt — PNGs mit Transparenz
                // zeigten vorher den hellgrauen Hintergrund durch, was wie ein
                // unsauberer „grauer Rand" um Logos aussah.
                background: '#fff',
                borderRadius: 'var(--dex-radius)',
                overflow: 'hidden',
                // v28.3/v28.4: Desktop = fester Bild-Slot links, contain (kein
                // Crop — Event-Bilder sind oft Poster mit Text). Querformat
                // bekommt den breiteren 420er-Slot und sitzt vertikal mittig
                // neben den Infos; Hochkant/Quadrat den kompakten 300er-Slot.
                // Handy = volle Breite mit begrenzter Höhe.
                // v28.7: Kreis-/Quadrat-Bilder = eigener Kreis oben mittig,
                // ragt zur Hälfte über die Oberkante der Karte hinaus
                // (negativer marginTop gegen Karten-Padding + Halbkreis).
                ...(imgCircleNotch
                  ? {
                    width: circleSize, height: circleSize, flex: '0 0 auto',
                    borderRadius: '50%',
                    border: '1px solid var(--dex-gray-200)',
                    boxShadow: '0 2px 12px rgba(0,0,0,0.10)',
                    alignSelf: 'center',
                    marginTop: -(circleSize / 2 + 16),
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                  }
                  : isMobile
                  ? { width: '100%', maxHeight: 200, display: 'flex', justifyContent: 'center' }
                  : event.imageBanner
                  // v28.5: Banner-Layout — volle Kartenbreite, Höhe begrenzt,
                  // contain (kein Crop bei Postern mit Text).
                  ? { width: '100%', maxHeight: 320, display: 'flex', alignItems: 'center', justifyContent: 'center' }
                  : {
                    flex: `0 0 ${imgSlotW}px`,
                    maxWidth: imgSlotW,
                    maxHeight: imgSlotH,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    alignSelf: 'center',
                  }),
              }}
            >
              {heroImgUrl && (
                <img
                  src={cachedImage}
                  alt={event.title}
                  // v11.56: 'contain', damit das Bild vollständig sichtbar
                  // bleibt (kein Crop). v28.3: Desktop einheitlich auf den
                  // kompakten 300er-Slot begrenzt; die Pro-Ansicht-Hero-
                  // Einstellung (Zoom/Höhe) wirkt weiter, gedeckelt auf den
                  // Slot, damit Station 1 kompakt bleibt.
                  // v28.6: borderRadius direkt am <img> — bei contain liegt die
                  // sichtbare Bildkante INNERHALB des Containers, dessen
                  // overflow-Rundung griff daher nicht. Kreis-PNGs (transparente
                  // Ecken) bleiben unverändert rund.
                  // v28.7: Im Kreis-Notch füllt das Bild den Kreis komplett
                  // (cover) — beim typischen Kreis-Zuschnitt (Quadrat mit
                  // transparenten Ecken) liegt die Bildkante exakt am Rand.
                  // v29.13: Ein als Rückfall gezeigtes MAIL-LOGO wird im Kreis
                  // nicht beschnitten (contain + Innenabstand) — Logos haben
                  // Ränder und Schrift, die ein cover-Zuschnitt anschneidet.
                  style={imgCircleNotch
                    ? (usesMailImage
                      ? { width: '100%', height: '100%', objectFit: 'contain', display: 'block', padding: 14, boxSizing: 'border-box' }
                      : { width: '100%', height: '100%', objectFit: 'cover', display: 'block' })
                    : isMobile
                    ? { width: '100%', maxHeight: 200, height: 'auto', objectFit: 'cover', display: 'block', borderRadius: 'var(--dex-radius)' }
                    : event.imageBanner
                    ? { maxWidth: '100%', maxHeight: 320, width: 'auto', height: 'auto', objectFit: 'contain', display: 'block', margin: '0 auto', borderRadius: 'var(--dex-radius)' }
                    : event.imageDisplay?.hero
                    ? { display: 'block', margin: '0 auto', maxWidth: '100%', maxHeight: Math.min(event.imageDisplay.hero.height ?? imgSlotH, imgSlotH), width: 'auto', height: 'auto', objectFit: 'contain', transform: `scale(${Math.min(event.imageDisplay.hero.zoom || 1, 1.5)})`, transformOrigin: 'center center', borderRadius: 'var(--dex-radius)' }
                    : { maxWidth: '100%', maxHeight: imgSlotH, width: 'auto', height: 'auto', objectFit: 'contain', display: 'block', borderRadius: 'var(--dex-radius)' }
                  }
                />
              )}
              {/* v11.91: Info-Button entfernt — die Beschreibung ist jetzt
                  immer ausgeklappt, kein Toggle mehr nötig. */}
              {/* v28.12: Lupen-Icon beim Hover (auf dem Handy immer sichtbar,
                  dort gibt es kein Hover) — Klick öffnet die Lightbox.
                  Mittig unten platziert, damit es auch im Kreis-Notch
                  (borderRadius 50% + overflow hidden) sichtbar bleibt. */}
              {(imgHovered || isMobile) && (
                <button
                  type="button"
                  onClick={() => setImgZoomed(true)}
                  title={locale === 'de' ? 'Bild vergrößern' : 'Enlarge image'}
                  aria-label={locale === 'de' ? 'Bild vergrößern' : 'Enlarge image'}
                  style={{
                    position: 'absolute', bottom: 8, left: '50%', transform: 'translateX(-50%)',
                    width: 30, height: 30, borderRadius: '50%', border: 'none',
                    background: 'rgba(0,0,0,0.55)', color: '#fff',
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    cursor: 'pointer', zIndex: 2,
                  }}
                >
                  <Icon iconName="ZoomIn" style={{ fontSize: 14 }} />
                </button>
              )}
              {/* v28.12: Lightbox — Klick irgendwo (oder aufs X) schließt.
                  fixed, entkommt dem overflow:hidden des Containers. */}
              {imgZoomed && (
                <div
                  onClick={() => setImgZoomed(false)}
                  style={{
                    position: 'fixed', inset: 0, zIndex: 3000,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    background: 'rgba(0,0,0,0.6)', cursor: 'zoom-out', padding: 24,
                  }}
                >
                  <img
                    src={cachedZoomImage}
                    alt={event.title}
                    style={{
                      maxWidth: '82vw', maxHeight: '80vh', width: 'auto', height: 'auto',
                      objectFit: 'contain', background: '#fff', borderRadius: 12,
                      boxShadow: '0 16px 56px rgba(0,0,0,0.4)', padding: 8,
                    }}
                  />
                  <button
                    type="button"
                    onClick={() => setImgZoomed(false)}
                    aria-label={locale === 'de' ? 'Schließen' : 'Close'}
                    style={{
                      position: 'absolute', top: 18, right: 22,
                      width: 36, height: 36, borderRadius: '50%', border: 'none',
                      background: 'rgba(255,255,255,0.92)', color: 'var(--dex-gray-800)',
                      display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                      cursor: 'pointer', fontSize: '1.1rem', fontWeight: 700,
                    }}
                  >×</button>
                </div>
              )}
            </div>
            )}
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 10, padding: '4px 4px 4px 0' }}>
              <h4 style={{ fontSize: '1rem', margin: 0 }}>{event.title}</h4>
              {/* v11.91: Datum + Ort als prominente Badges mit Icon.
                  v11.93: Datum einzeilig (nowrap) — Box wächst auf
                  natürliche Breite. Der Ort-Kasten streckt sich auf
                  dieselbe Breite, damit beide Boxen visuell aligniert
                  sind. inline-flex + alignItems:stretch sorgt für gleiche
                  Breite ohne festen Wert. */}
              {/* v11.94: alignSelf:stretch + maxWidth:100% damit die Box
                  nicht über den Card-Rand rausragt; gleichzeitig wächst
                  sie auf die natürliche Breite des längeren Inhalts und
                  beide Boxen sind gleich breit. */}
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'stretch', gap: 6, maxWidth: '100%' }}>
                <div style={{
                  display: 'flex', alignItems: 'center', gap: 8,
                  padding: '8px 12px', borderRadius: 8,
                  background: 'rgba(134,188,37,0.10)', color: 'var(--dex-green-dark, #4a7c1f)',
                  fontSize: '0.88rem', fontWeight: 600,
                }}>
                  <Icon iconName="Calendar" style={{ fontSize: 16, flexShrink: 0 }} />
                  <span>
                    {/* v11.94: kompaktes Range-Format („–" statt „until"),
                        bei gleichem Tag nur einmal Datum + „HH:MM - HH:MM". */}
                    {formatDateRange(event.startDate, event.endDate)}
                  </span>
                </div>
                {(event.location || (event.locationAddress && (event.locationAddress.street || event.locationAddress.city))) && (() => {
                  const addr = event.locationAddress;
                  const hasAddr = !!(addr && (addr.street || addr.city));
                  const hasStreet = !!(addr && addr.street);
                  const cityLine = addr ? [addr.zip, addr.city].filter(Boolean).join(' ') : '';
                  // v26.88: Nur Name + Stadt (KEINE Straße) → einzeilig „Name, Stadt"
                  // (z.B. „RheinEnergieSTADION, Köln") statt Name/Stadt untereinander.
                  const nameCityInline = !!(event.location && cityLine && !hasStreet);
                  // v26.82: Pin-Icon nur bei echter Mehrzeiler-Adresse oben
                  // ausrichten; bei einer Zeile (inkl. „Name, Stadt") zentrieren.
                  const multiLine = hasAddr && !nameCityInline;
                  return (
                  <div style={{
                    display: 'flex', alignItems: multiLine ? 'flex-start' : 'center', gap: 8,
                    padding: '8px 12px', borderRadius: 8,
                    background: 'rgba(0,86,166,0.08)', color: '#0a3766',
                    fontSize: '0.88rem',
                  }}>
                    <Icon iconName="POI" style={{ fontSize: 16, marginTop: multiLine ? 2 : 0, flexShrink: 0 }} />
                    <span>
                      {nameCityInline ? (
                        <>
                          <span style={{ fontWeight: 700 }}>{event.location}</span>
                          <span style={{ fontWeight: 400 }}>{`, ${cityLine}`}</span>
                        </>
                      ) : (
                        <>
                          {event.location && (
                            <span style={{ fontWeight: 700 }}>{event.location}</span>
                          )}
                          {addr && (addr.street || addr.city) && (
                            <>
                              <br />
                              <span style={{ fontWeight: 400 }}>
                                {[addr.street, addr.houseNo].filter(Boolean).join(' ')}
                                {(addr.zip || addr.city) && <br />}
                                {[addr.zip, addr.city].filter(Boolean).join(' ')}
                              </span>
                            </>
                          )}
                        </>
                      )}
                    </span>
                  </div>
                  );
                })()}
              </div>
              {/* v26.89: Reihenfolge getauscht — ANSPRECHPARTNER steht jetzt VOR
                  dem ORGANIZER-Block (vorher umgekehrt). */}
              {/* v10.16: Optionaler Ansprechpartner — frei eingegebene Person
                  außerhalb des App-User-Pools. Reines Anzeige-Feld; Mailto-Link
                  wenn Email gesetzt. Wird nur gerendert wenn mindestens Name
                  oder Email gepflegt sind. */}
              {(event.contactName || event.contactEmail || event.contactInfo) && (
                <div style={{ marginTop: 12 }}>
                  {/* v28.4: Überschrift AUSSERHALB der Box — gleiche Optik und
                      Position wie das ORGANIZER-Label darunter. */}
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6, color: 'var(--dex-gray-600)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 6, fontWeight: 600, fontSize: '0.85rem' }}>
                    <span style={{ display: 'inline-flex', flexShrink: 0 }}><Mail size={15} /></span>
                    <span>{locale === 'de' ? 'Ansprechpartner' : 'Contact'}</span>
                  </div>
                  <div style={{ padding: '10px 12px', background: 'var(--dex-gray-50, #f7f7f7)', borderRadius: 8, border: '1px solid var(--dex-gray-200)' }}>
                  {/* v28.5: Schriftgrößen wie in den Datums-/Ort-Boxen (0.88rem). */}
                  {event.contactName && (
                    <div style={{ fontSize: '0.88rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>{event.contactName}</div>
                  )}
                  {event.contactEmail && (
                    <div style={{ fontSize: '0.88rem', marginTop: 2 }}>
                      <a href={`mailto:${event.contactEmail}`} style={{ color: 'var(--dex-green, #86bc25)', textDecoration: 'none' }}>{event.contactEmail}</a>
                    </div>
                  )}
                  {event.contactInfo && (
                    <div style={{ fontSize: '0.88rem', color: 'var(--dex-gray-700)', marginTop: 4, whiteSpace: 'pre-wrap', lineHeight: 1.4 }}>{event.contactInfo}</div>
                  )}
                  </div>
                </div>
              )}
              {(() => {
                // v24.15: „Organizer ausblenden" ohne Einzel-Modus = ALLE aus.
                if (event.hideOrganizer && !event.hideOrganizerIndividualOnly) return null;
                // Organizer als Chips mit Foto (Hover-Enlarge). Namen werden von "Nachname, Vorname"
                // in "Vorname Nachname" normalisiert. v11.91: Label + Chip größer für bessere Lesbarkeit.
                const orgs = event.organizers.reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map(o => {
                  const trimmed = o.trim();
                  const parts = trimmed.split(',').map(s => s.trim());
                  return parts.length === 2 ? `${parts[1]} ${parts[0]}` : trimmed;
                }).filter(Boolean);
                if (orgs.length === 0) return null;
                // v26.89: Gibt es einen expliziten Ansprechpartner, blenden wir den
                // „Bei Fragen wende dich gerne an:"-Kopf im Organizer-Hover aus —
                // für Rückfragen ist dann ausdrücklich der Ansprechpartner zuständig.
                const hasExplicitContact = !!(event.contactName || event.contactEmail || event.contactInfo);
                return (
                  <div style={{ marginTop: 6 }}>
                    <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 6, fontWeight: 600 }}>Organizer</div>
                    <OrganizerList names={orgs} emails={event.organizerEmails} hiddenEmails={(event.hideOrganizer && event.hideOrganizerIndividualOnly) ? event.hiddenOrganizerEmails : []} forceIsDe={locale === 'de'} size="md" display={event.organizerDisplayLarge ? 'card' : 'chip'} nameFontSize="1.05rem" hideContactPrompt={hasExplicitContact} fullWidth contactEmail={event.contactOrganizerEmail || undefined} />
                  </div>
                );
              })()}
              {/* v23.25: Die „X / Y Plätze frei"-Anzeige steht jetzt direkt
                  über dem Registrieren-Button (siehe registration-actions). */}
            </div>
          </div>
          {/* v11.91: Beschreibung immer ausgeklappt — kein Toggle mehr. */}
          {event.description && (
            // v9.25: Beschreibung darf HTML enthalten (RichText-Editor im
            // EventCreation/Edit). Wir rendern als HTML statt Plain-Text,
            // damit Formatierung wie Listen, Links, Fett etc. funktioniert.
            // Die Description kommt aus dem eigenen Tenant — sicherer Origin.
            <div
              className="dex-event-desc"
              style={{
                padding: '12px 16px', color: 'var(--dex-gray-700)',
                background: 'var(--dex-gray-50)', borderRadius: '0 0 var(--dex-radius) var(--dex-radius)',
                borderTop: '1px solid var(--dex-gray-200)',
                wordBreak: 'break-word',
              }}
              dangerouslySetInnerHTML={{
                __html: (() => {
                  // v11.91: Email-Adressen in der Beschreibung automatisch
                  // in mailto-Links umwandeln. Funktioniert sowohl für
                  // Plain-Text als auch für HTML — Emails in bereits
                  // verlinktem Text (innerhalb von href="...") werden
                  // übersprungen.
                  const raw = event.description || '';
                  const isHtml = /<[a-z][\s\S]*>/i.test(raw);
                  const base = isHtml
                    ? raw
                    : raw.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
                  const EMAIL_RE = /\b([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})\b/g;
                  // Skip emails that already sit inside href="..." (already linked)
                  return base.replace(EMAIL_RE, (match, _email, offset, full) => {
                    const beforeWindow = full.slice(Math.max(0, offset - 80), offset);
                    if (/href\s*=\s*["'][^"']*$/.test(beforeWindow)) return match;
                    if (/>$/.test(beforeWindow) && /<a [^>]*$/i.test(beforeWindow)) return match;
                    // v11.95: Deloitte-Grün (--dex-green #86bc25) statt dem
                    // dunkleren Olive — gleiche Farbe wie die Card-Header.
                    return `<a href="mailto:${match}" style="color:#86bc25;text-decoration:underline">${match}</a>`;
                  });
                })(),
              }}
            />
          )}
          {/* v24.59: Der frühere rote „Alle Plätze belegt …"-Text unter der
              Event-Karte ist entfernt — die Info steht jetzt im Badge über den
              Buttons (Status „Alle Plätze belegt") und in der Button-Beschriftung
              (Warteliste + aktuelle Anzahl). */}
          {/* v10.20: Sessions-/Hauptevent-Auswahl ist nun in die rechte Spalte
              ('registration-specific') eingebettet — siehe weiter unten unter
              dem section-header "reg.eventinfo". Vorher stand der Block hier
              in der linken Event-Karten-Spalte; das Layout war optisch
              ungleich (links wuchs unbegrenzt, rechts gar nichts) und der
              User musste zwischen Spalten hin- und her-springen. Jetzt sind
              alle Anmelde-Inputs (Sessions, Starter-Typ, Custom-Fields) in
              einer Spalte gebündelt. */}
          {false && (
            <div style={{ marginTop: 16, border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 16 }}>
              <h4 style={{ marginTop: 0, marginBottom: 4, fontSize: '0.95rem' }}>{t('reg.selection.title') || 'Wofür möchtest du dich anmelden?'}</h4>
              <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                {t('reg.selection.hint') || 'Haupt-Event und Sessions können unabhängig voneinander an- oder abgewählt werden.'}
              </p>

              {/* v15.3.1: Haupt-Event-Checkbox nur zeigen wenn Anmeldung
                  fürs Hauptevent überhaupt möglich ist. Im subEventsOnlyMode
                  läuft die Anmeldung exklusiv über Sub-Events. */}
              {!event.subEventsOnlyMode && (
              <label style={{
                display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
                borderRadius: 8,
                border: `1px solid ${registerForParent && !parentAlreadyRegistered && !parentRegBlocked ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                background: registerForParent && !parentAlreadyRegistered && !parentRegBlocked ? 'rgba(134,188,37,0.06)' : '#fff',
                cursor: (parentAlreadyRegistered || parentRegBlocked) ? 'default' : 'pointer',
                opacity: parentRegBlocked ? 0.6 : 1,
              }}>
                <input
                  type="checkbox"
                  checked={parentAlreadyRegistered ? true : (parentRegBlocked ? false : registerForParent)}
                  disabled={parentAlreadyRegistered || parentRegBlocked}
                  onChange={e => setRegisterForParent(e.target.checked)}
                  style={{ marginTop: 2 }}
                />
                <div style={{ flex: 1 }}>
                  {(() => { const lbl = resolveMainEventLabel(t('reg.selection.mainevent') || 'Haupt-Event'); return (
                    <div style={{ fontWeight: 700 }}>{lbl ? `${lbl}: ` : ''}{event.title}</div>
                  ); })()}
                  {parentAlreadyRegistered && (
                    <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                      {t('reg.selection.alreadyregistered') || 'Du bist bereits für das Haupt-Event angemeldet.'}
                    </div>
                  )}
                  {parentRegBlocked && !parentAlreadyRegistered && (
                    <div style={{ fontSize: '0.75rem', color: parentFullNoWaitlist ? 'var(--dex-red, #c00)' : 'var(--dex-orange, #ed8b00)', marginTop: 2 }}>
                      {parentFullNoWaitlist
                        ? (locale === 'de' ? 'Alle Plätze sind belegt — die Warteliste ist für dieses Event deaktiviert.' : 'All seats are taken — the waitlist is disabled for this event.')
                        : (t('reg.subevents.deadlinepassed') || 'Anmeldefrist abgelaufen — nur noch die offenen Sub-Events sind wählbar.')}
                    </div>
                  )}
                </div>
              </label>
              )}

              {/* Sessions */}
              {childEvents.length > 0 && (
                <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', fontWeight: 600 }}>{childTermPlural || t('reg.selection.sessions') || 'Sessions'}</div>
                  {/* v14.5: Pflicht-Hinweis wenn Organizer requireSubEventSelection aktiv hat. */}
                  {event && event.requireSubEventSelection && (
                    <div style={{
                      padding: '8px 12px', borderRadius: 8,
                      background: 'rgba(237,139,0,0.10)',
                      border: '1px solid var(--dex-orange, #ed8b00)',
                      fontSize: '0.82rem', color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 600,
                    }}>
                      {locale === 'de'
                        ? `Pflicht: bitte mindestens ${childOneDe} auswählen.`
                        : `Required: please pick at least one ${childTermSingular || 'sub-event'}.`}
                    </div>
                  )}
                  {childEvents.map(ce => {
                    const meta = sessionMeta[ce.id] || { count: 0, wasRegistered: false };
                    const isSel = selectedSessions.has(ce.id);
                    const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
                    const isSessionFull = hasCap && meta.count >= (ce.maxParticipants || 0);
                    const deadlinePassed = !!(ce.registrationDeadline && new Date(ce.registrationDeadline) < new Date());
                    // v29.28: Organizer/Admins dürfen — wie beim Haupt-Event
                    // (parentRegBlocked) und wie der Wizard es ausdrücklich
                    // verspricht — auch nach der Frist anmelden. Die
                    // Kapazitäts-Sperre bleibt für alle.
                    const deadlineLocked = deadlinePassed && !isOrganizer && !isAdmin;
                    const disabled = (isSessionFull && !isSel) || (deadlineLocked && !isSel);
                    // Erbt vom Haupt-Event wenn gleichzeitig angemeldet wird.
                    const inheritsStarter = isSplitGroup && (willRegisterParent || registerForOther);
                    const sType = sessionStarterType[ce.id] || '';

                    return (
                      <div key={ce.id} style={{
                        padding: 10, borderRadius: 8,
                        border: `1px solid ${isSel ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                        background: isSel ? 'rgba(134,188,37,0.06)' : '#fff',
                      }}>
                        <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: disabled ? 'not-allowed' : 'pointer' }}>
                          <input
                            type="checkbox"
                            checked={isSel}
                            disabled={disabled}
                            onChange={e => {
                              if (e.target.checked) {
                                // v29.27: direkt selektieren — die Fragen des
                                // Sub-Events erscheinen INLINE in der Karte
                                // (renderSubEventInlineFields), nicht mehr im
                                // Bestätigen-Modal. Pflicht prüft der Submit.
                                const next = new Set(selectedSessions);
                                next.add(ce.id);
                                setSelectedSessions(next);
                              } else {
                                // Uncheck: Session entfernen + gespeicherte Field-Werte
                                // wegräumen damit beim erneuten Checken ein frisches
                                // Modal kommt (kein Stale-State).
                                const next = new Set(selectedSessions);
                                next.delete(ce.id);
                                setSelectedSessions(next);
                                setSessionFieldValues(prev => {
                                  const copy = { ...prev };
                                  delete copy[ce.id];
                                  return copy;
                                });
                              }
                            }}
                            style={{ marginTop: 2 }}
                          />
                          <div style={{ flex: 1 }}>
                            <div style={{ fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                              {ce.title || tEvent('reg.subevents.untitled')}
                              {ce.mandatoryRegistration && (
                                <span style={{ fontSize: '0.68rem', fontWeight: 700, color: '#fff', background: 'var(--dex-orange, #ed8b00)', borderRadius: 999, padding: '2px 8px' }}>
                                  {locale === 'de' ? 'Pflicht' : 'Required'}
                                </span>
                              )}
                            </div>
                          </div>
                          {/* v29.28: Thumbnail wandert in die Titelzeile —
                              der Rest der Karte steht jetzt AUSSERHALB des
                              <label> linksbündig am Kartenrand. */}
                          {ce.imageUrl && (
                            <img
                              src={ce.imageUrl}
                              alt=""
                              style={{ width: 84, height: 60, objectFit: 'cover', borderRadius: 6, flexShrink: 0, background: 'var(--dex-gray-100)' }}
                            />
                          )}
                        </label>
                        {/* v29.28: Beschreibung, Zeit/Ort, Plätze und Hinweise
                            linksbündig auf Kartenbreite — vorher rückte die
                            Checkbox-Spalte des Labels alles ein. */}
                        <div style={{ marginTop: 4 }}>
                            {ce.description && (
                              // v11.97: gleiche Schriftgröße wie der Titel
                              // (Standard-Body). Vorher 0.78rem klein.
                              // v29.27: als sanitisiertes HTML statt rohem Text
                              // (Rich-Text-Editor-Beschreibungen, s. subEventDescHtml).
                              <div style={{ color: 'var(--dex-gray-600)', marginTop: 2, wordBreak: 'break-word' }} dangerouslySetInnerHTML={{ __html: subEventDescHtml(ce.description) }} />
                            )}
                            {/* v11.94: Datum + Ort mit Icons (analog zum
                                Haupt-Event-Header), damit Sub-Events visuell
                                konsistent sind und auf einen Blick lesbar. */}
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, marginTop: 4, color: 'var(--dex-gray-600)' }}>
                              {ce.startDate && (
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                  {/* v11.97: Icon-Größe an Standard-Body angepasst. */}
                                  <Icon iconName="Calendar" style={{ fontSize: 15, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                                  {formatDate(ce.startDate)}
                                </span>
                              )}
                              {ce.location && (
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                  <Icon iconName="POI" style={{ fontSize: 15, color: '#0a3766' }} />
                                  {ce.location}
                                </span>
                              )}
                            </div>
                            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {hasCap && (() => {
                                // v9.8: Klartext-Anzeige damit der User auf einen Blick
                                // sieht, wie viele Plätze noch frei sind. Vorher stand
                                // dort nur "0/25" ohne Label, was die User-Frage
                                // "warum steht da 0/25?" ausgelöst hat.
                                const sessionFree = Math.max(0, (ce.maxParticipants || 0) - (meta.count || 0));
                                return (
                                  <> · <span style={{ color: isSessionFull ? 'var(--dex-red)' : 'inherit', fontWeight: 600 }}>
                                    {/* v19.19: belegt-Zahl bei der Kapazität deckeln,
                                        damit eine Überbuchung (count > cap) auf der
                                        Anmeldeseite NICHT sichtbar wird — die echte
                                        Überbuchungszahl sieht nur der Organizer/Admin. */}
                                    {Math.min(meta.count, ce.maxParticipants || 0)}/{ce.maxParticipants} {t('reg.subevents.taken')}
                                  </span>
                                  {!isSessionFull && (
                                    <span style={{ color: 'var(--dex-green-dark)' }}> — {sessionFree} {t('reg.free')}</span>
                                  )}
                                  </>
                                );
                              })()}
                            </div>
                            {deadlinePassed && !isSel && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-orange)', marginTop: 2 }}>
                                {deadlineLocked
                                  ? t('reg.subevents.deadlinepassed')
                                  : (locale === 'de'
                                    ? 'Anmeldefrist abgelaufen — als Organizer/Admin trotzdem wählbar.'
                                    : 'Registration deadline passed — still selectable as organizer/admin.')}
                              </div>
                            )}
                            {isSessionFull && !isSel && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-red)', marginTop: 2 }}>
                                {t('reg.subevents.sessionfull')}
                              </div>
                            )}
                            {/* v10.20: Gruppen-Auswahl pro Session — nur wenn NICHT vom Parent geerbt.
                                Dynamische Labels splitLabelA / splitLabelB. */}
                            {isSel && isSplitGroup && !inheritsStarter && (
                              <div style={{ marginTop: 8, display: 'flex', gap: 10, fontSize: '0.8rem' }}>
                                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}>
                                  <input
                                    type="radio"
                                    name={`starter-${ce.id}`}
                                    checked={sType === 'Durchstarter'}
                                    onChange={() => setSessionStarterType({ ...sessionStarterType, [ce.id]: 'Durchstarter' })}
                                  />
                                  {splitLabelA}
                                </label>
                                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}>
                                  <input
                                    type="radio"
                                    name={`starter-${ce.id}`}
                                    checked={sType === 'Funstarter'}
                                    onChange={() => setSessionStarterType({ ...sessionStarterType, [ce.id]: 'Funstarter' })}
                                  />
                                  {splitLabelB}
                                </label>
                              </div>
                            )}
                            {/* v10.25: B2Run-spezifischer Vererbungs-Hinweis
                                entfernt — bei generischer Split-Capacity
                                (z.B. "Vormittag/Nachmittag") wirkt der
                                Hinweis "Starter-Typ wird vom Haupt-Event
                                übernommen" verwirrend. Die Logik bleibt
                                (Sub-Event übernimmt die Gruppen-Wahl des
                                Parents), nur die explizite UI-Zeile ist
                                weg. */}
                        </div>
                        {/* v29.27: Fragen dieses Sub-Events direkt in der
                            Karte — BEWUSST außerhalb des <label>, sonst
                            würde jeder Klick in ein Feld die Checkbox
                            togglen. */}
                        {isSel && renderSubEventInlineFields(ce)}
                      </div>
                    );
                  })}
                </div>
              )}

              {isSessionsOnlyMode && selectedSessions.size > 0 && !event.subEventsOnlyMode && (
                <div style={{
                  marginTop: 12, padding: '8px 10px', borderRadius: 6,
                  background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                  color: 'var(--dex-orange)', fontSize: '0.78rem',
                }}>
                  {childTermPlural
                    ? (locale === 'de'
                        ? `Du meldest dich ausschließlich für ${childTermPlural} an — NICHT für das Haupt-Event.`
                        : `You are registering exclusively for ${childTermPlural} — NOT for the main event.`)
                    : (t('reg.selection.sessionsonlyhint') || 'Du meldest dich ausschließlich für Sessions an — NICHT für das Haupt-Event.')}
                </div>
              )}
            </div>
          )}
        </div>

        {/* v18.73: Die „Offene Teams"-Box ist nach UNTEN gewandert (unter die
            „Ich melde mich + mein Team an"-Karte) — siehe weiter unten. Oben
            steht jetzt immer zuerst die persönliche Daten-Karte, dann die
            event-spezifischen Infos. */}

        {/* v28.2: Station 2 — Deine Daten. */}
        <div className="reg-step-head">
          <span className="reg-step-num">2</span>
          <span className="reg-step-label">{locale === 'de' ? 'Deine Daten — automatisch aus M365' : 'Your details — automatically from M365'}</span>
        </div>
        {/* Persönliche Daten */}
        <div className="registration-form">
          {/* v11.97: Section-Header + Register-for-other-Toggle in einer
              Zeile (grünes Section-Header-Pill links, Toggle als Link
              rechts daneben). Vorher saß der Toggle unter dem Header
              im Body — wenig auffällig. „* = Required field"-Legende
              ist hier weg und sitzt jetzt am Event-Specific-Header. */}
          <CollapsibleSection
            isMobile={isMobile}
            icon="ContactInfo"
            title={t('reg.personalinfo')}
            headerExtra={(canRegisterForOther || (registerForOther && canCreateEvents)) ? (
            <>
            {canRegisterForOther && (
              <button
                type="button"
                onClick={() => {
                  setRegisterForOther(!registerForOther);
                  setThirdPartyCheck(null);
                  setPickedUserProfile(null);
                  setOtherConsentConfirmed(false);
                  setExternalPerson(false); // v18.74: Extern-Modus beim Wechsel zurücksetzen
                  // v19.6: CC-Frage-Entscheidung beim Moduswechsel zurücksetzen.
                  ccSelfDecidedRef.current = false;
                  ccSelfRef.current = false;
                  if (!registerForOther) {
                    setFirstName(''); setSurname(''); setEmail(''); setUserSearch(''); setUserResults([]);
                    // v26.76: geführten Wizard öffnen (Person suchen → Zustimmung).
                    setProxyStep(1);
                  } else {
                    setFirstName(currentUser.firstName); setSurname(currentUser.surname); setEmail(currentUser.email); setUserSearch(''); setUserResults([]);
                    setProxyStep(0);
                  }
                }}
                style={{
                  // v26.82: Als „angedockte" Tab-Optik neben dem grünen
                  // „Persönliche Informationen"-Header. Standard (für andere
                  // anmelden) = GRAU (inaktiver Tab, klarer Farbunterschied);
                  // aktiv (im Fremd-Modus, „zurück zur Selbst-Anmeldung") = grün.
                  // v26.89: Der Tab dockt jetzt SPIEGELBILDLICH zum grünen
                  // „Persönliche Informationen"-Header in die obere RECHTE Ecke
                  // — bündig an Ober- und Rechtskante (alignSelf: stretch +
                  // oben abgerundete Ecken wie der grüne Tab, unten eckig).
                  display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                  marginLeft: 'auto', // an die rechte Ecke schieben
                  alignSelf: 'stretch', boxSizing: 'border-box',
                  padding: '7px 18px', borderRadius: 'var(--dex-radius) var(--dex-radius) 0 0',
                  fontSize: '0.85rem', fontWeight: 600, cursor: 'pointer',
                  transition: 'background 0.15s ease, color 0.15s ease',
                  ...(registerForOther
                    ? { background: 'var(--dex-green, #86bc25)', border: '1.5px solid var(--dex-green, #86bc25)', color: '#fff' }
                    : { background: 'var(--dex-gray-100, #eef0f2)', border: '1.5px solid var(--dex-gray-300, #cfd4d9)', color: 'var(--dex-gray-600, #5a6470)' }),
                }}
              >
                <Icon iconName={registerForOther ? 'Contact' : 'AddFriend'} style={{ fontSize: 14 }} />
                {registerForOther ? t('reg.registerself') : t('reg.registerother')}
              </button>
            )}
            {/* v18.13: Massenimport — nur Organizer/Admin im „Für andere"-Modus. */}
            {registerForOther && canCreateEvents && (
              <button
                type="button"
                onClick={() => { setMassImportResult(null); setMassImportRows([]); setMassImportStep('input'); setMassImportOpen(true); }}
                style={{
                  background: 'none', border: 'none', padding: '4px 12px',
                  color: 'var(--dex-blue, #0076a8)', fontSize: '0.85rem',
                  textDecoration: 'underline', cursor: 'pointer', fontWeight: 600,
                }}
              >
                {locale === 'de' ? 'Massenimport' : 'Bulk import'}
              </button>
            )}
            </>
            ) : undefined}
          >
          <div style={{ padding: '24px 20px' }}>
            {canRegisterForOther && (
              <>
                {registerForOther && isAssistant && !canCreateEvents && (
                  <div style={{
                    padding: '8px 12px', marginBottom: 12, borderRadius: 'var(--dex-radius-md)',
                    background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                    color: 'var(--dex-orange)', fontSize: '0.8rem',
                  }}>
                    As an Assistant you can only register <strong>Partners</strong> or <strong>Directors</strong> for this event.
                  </div>
                )}
                {/* v23.4: Der „Person außerhalb Deloitte"-Umschalter ist nach
                    UNTEN gewandert (direkt unter die „@deloitte.com"-Such-Zeile,
                    siehe weiter unten) und in der Schriftgröße an diese Zeile
                    angeglichen. */}
                {/* v27.6: Kompakte Zusammenfassung der stellvertretend
                    anzumeldenden Person. Auswahl/Extern/Zustimmung passieren
                    komplett im Wizard-Modal — „Ändern" öffnet ihn wieder. */}
                {registerForOther && proxyStep === 0 && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20, padding: '10px 12px', borderRadius: 'var(--dex-radius-md)', background: 'rgba(134,188,37,0.08)', border: '1px solid var(--dex-green, #86bc25)' }}>
                    <Icon iconName="Contact" style={{ fontSize: 18, color: 'var(--dex-green-dark, #4a7c1f)', flexShrink: 0 }} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.03em' }}>{locale === 'de' ? 'Du meldest an' : 'You are registering'}</div>
                      <div style={{ fontWeight: 700, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{`${firstName} ${surname}`.trim() || email || (locale === 'de' ? '— keine Person gewählt —' : '— no person selected —')}</div>
                      {!!email && <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email}</div>}
                    </div>
                    <button type="button" onClick={() => setProxyStep(1)} style={{ background: 'none', border: 'none', padding: 0, color: 'var(--dex-blue, #0076a8)', textDecoration: 'underline', cursor: 'pointer', fontSize: '0.8rem', fontWeight: 600, flexShrink: 0 }}>{locale === 'de' ? 'Ändern' : 'Change'}</button>
                  </div>
                )}
                {showInlineProxyPicker && !externalPerson && proxyStep === 0 && (
                  <div className="form-group" style={{ position: 'relative', marginBottom: 20 }}>
                    {/* v11.97: Label entfernt — Suche ist selbsterklärend (Placeholder). */}
                    <input
                      className="form-input"
                      value={userSearch}
                      onChange={e => {
                        const val = e.target.value;
                        setUserSearch(val);
                        if (searchTimerRef.current) clearTimeout(searchTimerRef.current);
                        if (val.length >= 2) {
                          searchTimerRef.current = setTimeout(async () => {
                            // v23.4: Wird eine vollständige EXTERNE E-Mail-Adresse
                            // eingegeben (nicht @deloitte.de/.com), automatisch in
                            // den Extern-Modus wechseln und die Adresse übernehmen
                            // — der Organizer muss den Toggle nicht mehr von Hand
                            // setzen. Greift erst, wenn die Eingabe eine plausible
                            // komplette Adresse ist (also nach dem Tippen).
                            const v = val.trim();
                            // v27.11: JEDE Member-Firm-Domain zählt als intern.
                            const isDeloitte = isDeloitteInternalEmail(v);
                            if (canCreateEvents && isPlausibleEmail(v) && !isDeloitte) {
                              setExternalPerson(true);
                              setUserSearch(''); setUserResults([]); setPickedUserProfile(null);
                              setThirdPartyCheck(null); setOtherConsentConfirmed(false);
                              setFirstName(''); setSurname(''); setEmail(v);
                              return;
                            }
                            setIsSearchingUser(true);
                            const results = await searchUsers(val, userSearchIncludeIntl);
                            setUserResults(results);
                            setIsSearchingUser(false);
                          }, 300);
                        } else {
                          setUserResults([]);
                        }
                      }}
                      placeholder={t('reg.searchplaceholder') || 'Name oder E-Mail eingeben...'}
                    />
                    <InternationalSearchToggle
                      query={userSearch}
                      checked={userSearchIncludeIntl}
                      onChange={async next => {
                        setUserSearchIncludeIntl(next);
                        const val = userSearch.trim();
                        if (val.length >= 2) {
                          setIsSearchingUser(true);
                          try {
                            const results = await searchUsers(val, next);
                            setUserResults(results);
                          } catch { /* */ }
                          setIsSearchingUser(false);
                        }
                      }}
                      isDe={locale === 'de'}
                    />
                    {isSearchingUser && (
                      <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>Suche...</div>
                    )}
                    {userResults.length > 0 && (
                      <div style={{
                        position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
                        background: '#fff', border: '1px solid var(--dex-gray-200)',
                        borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                        maxHeight: 200, overflowY: 'auto',
                      }}>
                        {userResults.map(u => {
                          // "Nachname, Vorname" oder "Vorname Nachname" Format
                          let uFirstName = '';
                          let uSurname = '';
                          if (u.displayName.includes(',')) {
                            const parts = u.displayName.split(',').map(s => s.trim());
                            uSurname = parts[0] || '';
                            uFirstName = parts[1] || '';
                          } else {
                            const parts = u.displayName.split(' ');
                            uFirstName = parts[0] || '';
                            uSurname = parts.slice(1).join(' ') || '';
                          }
                          // Assistant-Einschränkung: User darf nur Partner/Director auswählen,
                          // andere Treffer werden grau + nicht-klickbar angezeigt.
                          const assistantOnly = isAssistant && !canCreateEvents;
                          const targetAllowed = !assistantOnly || isAllowedTargetForAssistant(u.jobTitle);
                          return (
                            <div
                              key={u.email}
                              style={{
                                padding: '8px 12px', cursor: targetAllowed ? 'pointer' : 'not-allowed', fontSize: '0.85rem',
                                borderBottom: '1px solid var(--dex-gray-100)',
                                opacity: targetAllowed ? 1 : 0.45,
                              }}
                              onMouseEnter={e => { if (targetAllowed) (e.currentTarget as HTMLElement).style.background = 'var(--dex-gray-50)'; }}
                              onMouseLeave={e => { (e.currentTarget as HTMLElement).style.background = '#fff'; }}
                              onMouseDown={() => {
                                if (!targetAllowed) return;
                                setFirstName(uFirstName);
                                setSurname(uSurname);
                                setEmail(u.email);
                                setUserSearch(u.displayName);
                                setUserResults([]);
                                // v11.97: Profil mit Department + Mobile nachladen
                                // — searchUsers liefert nur jobTitle+location.
                                setPickedUserProfile({
                                  jobTitle: u.jobTitle || '',
                                  location: u.location || '',
                                });
                                searchUser(u.email).then(p => {
                                  if (p) {
                                    setPickedUserProfile({
                                      jobTitle: p.jobTitle || u.jobTitle || '',
                                      department: p.department || '',
                                      location: p.location || u.location || '',
                                      mobilePhone: p.mobilePhone || '',
                                      company: p.company || '',
                                    });
                                  }
                                }).catch(() => { /* silent */ });
                                // Früh-Check: bereits angemeldet? Im Verteiler?
                                setThirdPartyCheck(null);
                                if (event) {
                                  (async () => {
                                    const existing = await checkRegistrationByEmail(event.id, u.email).catch(() => null);
                                    const alreadyRegistered = !!existing && existing.Status !== 'Abgemeldet';
                                    // v24.92: Audience-Check über die KANONISCHE Sichtbarkeits-
                                    // Logik (isEventVisibleForUser) statt einer Teil-Reimplementierung.
                                    // Bugfix: Der alte Code prüfte nur rohe Filter-Einträge (E-Mail
                                    // direkt / Standort-Code) und kannte KEINE Verteiler-/Gruppen-
                                    // Mitgliedschaft — daher wurde JEDE über einen Verteiler sichtbare
                                    // Person fälschlich als „nicht im Gästekreis" gemeldet. Die beim
                                    // Event-Save aufgelösten Verteiler-Mitglieder stehen in
                                    // event.audienceResolvedEmails und werden von isEventVisibleForUser
                                    // berücksichtigt. Die Graph-Gruppen der Zielperson kennen wir nicht
                                    // (leeres groupEmails) — die aufgelösten Mitglieder decken den
                                    // DL-Fall aber ab.
                                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                    const notInAudience = !isEventVisibleForUser(event, u.email, (u as any).location || '', [], (u as any).jobTitle || '');
                                    setThirdPartyCheck({
                                      alreadyRegistered,
                                      notInAudience,
                                      // v19.8: Name + Anmeldedatum der bestehenden
                                      // Registrierung für eine konkrete Hinweis-Box.
                                      registeredName: (existing && (existing.ParticipantName || `${existing.Vorname || ''} ${existing.Nachname || ''}`.trim())) || u.displayName || '',
                                      registeredDate: (existing && existing.RegistrationDate) || '',
                                    });
                                  })();
                                }
                              }}
                              title={targetAllowed ? '' : 'Assistants can only register Partners or Directors for events.'}
                            >
                              {/* v11.3: People-Picker-Reihe mit Foto — analog
                                  zum Wizard-Organizer-Picker. SP-userphoto.aspx
                                  liefert das Profilbild zum E-Mail-Account. */}
                              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                                <img
                                  src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(u.email)}&size=S`}
                                  alt={u.displayName}
                                  onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                  style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                                />
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontWeight: 600 }}>{u.displayName}</div>
                                  <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                    {u.email}
                                    {u.jobTitle ? ` · ${u.jobTitle}` : ''}
                                    {u.location ? ` · ${u.location}` : ''}
                                  </div>
                                </div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                )}
                {/* v23.4: „Person außerhalb Deloitte"-Umschalter — jetzt UNTER
                    der Such-Zeile (inkl. „@deloitte.com"-Hinweis), gleiche
                    Schriftgröße (12px) wie der International-Toggle. Nur für
                    Organizer/Admins (nicht Assistenten). Blendet den People-
                    Picker aus und macht Vorname/Nachname/E-Mail frei eintragbar. */}
                {showInlineProxyPicker && canCreateEvents && (
                  <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, marginTop: 4, marginBottom: 16, cursor: 'pointer', fontSize: 12, color: 'var(--dex-gray-700, #444)', userSelect: 'none' }}>
                    <input
                      type="checkbox"
                      checked={externalPerson}
                      onChange={e => {
                        const on = e.target.checked;
                        setExternalPerson(on);
                        // Picker- und Feld-State frisch zurücksetzen (manuelle Eingabe).
                        setUserSearch(''); setUserResults([]); setPickedUserProfile(null);
                        setThirdPartyCheck(null); setOtherConsentConfirmed(false);
                        setFirstName(''); setSurname(''); setEmail('');
                      }}
                    />
                    <span>
                      {locale === 'de'
                        ? <>Person <strong>außerhalb Deloitte</strong> anmelden (externe E-Mail-Adresse)</>
                        : <>Register a person <strong>outside Deloitte</strong> (external email address)</>}
                    </span>
                  </label>
                )}
                {/* v15.16: Pflicht-Bestätigungs-Box bei „Für andere
                    registrieren" — analog zur Team-Anmeldung. Erscheint
                    sobald eine Person ausgewählt ist (E-Mail gefüllt) und
                    enthält zusätzlich einen Spezial-Hinweis für externe
                    Adressen (kein Outlook-Termin, Mail zur Weiterleitung
                    an die Organizer). */}
                {showInlineProxyPicker && proxyStep === 0 && (() => {
                  // v15.22: Hinweis-Box bereits anzeigen, sobald „Für andere
                  // registrieren" aktiv ist — nicht erst wenn die E-Mail
                  // gefüllt ist. Der User soll sofort sehen, dass eine
                  // Zustimmung nötig ist.
                  // v18.74/v27.11: Extern, sobald der Extern-Modus aktiv ist ODER
                  // die eingegebene E-Mail kein Deloitte-Postfach ist (beliebige
                  // Member Firm zählt als intern).
                  const isExternal = externalPerson || isExternalEmail(email);
                  const pickedName = `${firstName} ${surname}`.trim();
                  return (
                    <div style={{
                      marginBottom: 16,
                      padding: '14px 16px',
                      background: 'rgba(237,139,0,0.10)',
                      border: '2px solid var(--dex-orange, #ed8b00)',
                      borderRadius: 'var(--dex-radius-md)',
                      color: '#7a4a00',
                      fontSize: '0.88rem',
                      lineHeight: 1.55,
                    }}>
                      <div style={{ fontWeight: 700, marginBottom: 8, fontSize: '0.95rem' }}>
                        {locale === 'de'
                          ? (isExternal ? 'Einladung & Datenschutz-Rückmeldung laufen über dich' : 'Vorab die Zustimmung der Person einholen')
                          : (isExternal ? 'You send the invitation & collect the privacy confirmation' : 'Get the person\'s consent up front')}
                      </div>
                      {externalPerson && (
                        <div style={{ marginBottom: 8, fontWeight: 600 }}>
                          {locale === 'de'
                            ? <>Trage <strong>Vorname, Nachname und E-Mail-Adresse</strong> der externen Person unten selbst ein.</>
                            : <>Enter the external person&rsquo;s <strong>first name, last name and email address</strong> in the fields below.</>}
                        </div>
                      )}
                      {/* v26.98: Die ausführliche Ablauf-Erklärung ist jetzt
                          eingeklappt — auf der Anmeldeseite bleibt nur der kurze
                          Hinweis (Überschrift) + der Pflicht-Haken sichtbar. */}
                      <details style={{ marginBottom: 4 }}>
                        <summary style={{ cursor: 'pointer', fontSize: '0.82rem', fontWeight: 600, color: '#7a4a00', userSelect: 'none' }}>
                          {locale === 'de' ? 'So läuft das ab (Details anzeigen)' : 'How it works (show details)'}
                        </summary>
                      <div style={{ marginBottom: 8, marginTop: 8 }}>
                        {locale === 'de'
                          ? (isExternal
                              ? <>Mit dem Absenden setzt du {pickedName ? <><strong>{pickedName}</strong></> : <>die externe Person</>} mit dem Status <strong>&bdquo;Angemeldet (Datenschutzrückmeldung offen)&ldquo;</strong> auf die Teilnehmerliste. Final wird die Anmeldung erst, wenn die Person auf deine Einladung geantwortet hat und du die Rückmeldung in der Teilnehmerliste bestätigt hast.</>
                              : <>Mit dem Absenden meldest du {pickedName ? <><strong>{pickedName}</strong></> : <>die ausgewählte Person</>} stellvertretend an. Bitte stelle sicher, dass die Person ihrer Anmeldung <strong>VORHER zugestimmt</strong> hat — eine Anmeldung ohne Einverständnis ist nicht erlaubt.</>)
                          : (isExternal
                              ? <>Submitting puts {pickedName ? <><strong>{pickedName}</strong></> : <>the external person</>} on the participant list with the status <strong>&ldquo;Registered (privacy confirmation pending)&rdquo;</strong>. The registration only becomes final once the person has replied to your invitation and you have confirmed the response in the participant list.</>
                              : <>By submitting you register {pickedName ? <><strong>{pickedName}</strong></> : <>the selected person</>} on their behalf. Please make sure the person has <strong>consented up front</strong> — registering people without their consent is not allowed.</>)}
                      </div>
                      {isExternal && (
                        <div style={{
                          marginTop: 10, padding: '10px 12px',
                          background: '#fff', border: '1px dashed var(--dex-orange, #ed8b00)',
                          borderRadius: 6,
                        }}>
                          <div style={{ fontWeight: 700, marginBottom: 4 }}>
                            {locale === 'de'
                              ? 'Externe Person — so läuft die Einladung'
                              : 'External person — how the invitation works'}
                          </div>
                          {locale === 'de'
                            ? <>{email.trim() ? <>Die Adresse <strong>{email}</strong> gehört nicht zum Deloitte-Deutschland-Tenant. </> : <>Eine externe Adresse gehört nicht zum Deloitte-Deutschland-Tenant. </>}Die App kann an externe Adressen <strong>weder E-Mails noch Outlook-Termine</strong> versenden. Deshalb: Nach dem Absenden bekommst <strong>du eine E-Mail mit den nächsten Schritten</strong> — in der Teilnehmerliste lädst du den <strong>fertigen Einladungs-Entwurf</strong> herunter und verschickst ihn aus deinem eigenen Postfach (nur noch auf &bdquo;Senden&ldquo; klicken). Sobald die Person per Antwort-Mail zusagt, bestätigst du die <strong>Datenschutz-Rückmeldung</strong> per Klick in der Teilnehmerliste.</>
                            : <>{email.trim() ? <>The address <strong>{email}</strong> is not part of the Deloitte Germany tenant. </> : <>An external address is not part of the Deloitte Germany tenant. </>}The app can send <strong>neither emails nor Outlook invites</strong> to external addresses. Therefore: after submitting, <strong>you receive an email with the next steps</strong> — download the <strong>ready-made invitation draft</strong> from the participant list and send it from your own mailbox (just click &ldquo;Send&rdquo;). Once the person confirms by reply, confirm the <strong>privacy response</strong> with one click in the participant list.</>}
                        </div>
                      )}
                      </details>
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginTop: 12, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={otherConsentConfirmed}
                          onChange={e => setOtherConsentConfirmed(e.target.checked)}
                          style={{ marginTop: 3 }}
                        />
                        <span style={{ flex: 1, color: 'var(--dex-gray-800)' }}>
                          <span style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                          {locale === 'de'
                            ? (isExternal
                                ? 'Ich habe verstanden, dass ich die Einladung selbst an die Person versende und ihre Datenschutz-Rückmeldung abwarte, bevor die Anmeldung final wird.'
                                : 'Ich bestätige, dass die Person ihrer stellvertretenden Anmeldung zugestimmt hat.')
                            : (isExternal
                                ? 'I understand that I will send the invitation to the person myself and wait for their privacy confirmation before the registration becomes final.'
                                : 'I confirm that the person has consented to this registration on their behalf.')}
                        </span>
                      </label>
                    </div>
                  );
                })()}
                {registerForOther && thirdPartyCheck && (thirdPartyCheck.alreadyRegistered || thirdPartyCheck.notInAudience) && (
                  <div style={{
                    padding: '10px 14px', marginBottom: 16, borderRadius: 'var(--dex-radius-md)',
                    background: thirdPartyCheck.alreadyRegistered ? 'rgba(200,30,30,0.07)' : 'rgba(237,139,0,0.08)',
                    border: `1px solid ${thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)'}`,
                    color: thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)',
                    fontSize: '0.85rem',
                  }}>
                    {thirdPartyCheck.alreadyRegistered && (() => {
                      // v19.8: Konkrete Meldung mit Name + Anmeldedatum statt
                      // generischem „Diese Person ist bereits angemeldet".
                      const nm = (`${firstName} ${surname}`.trim()) || thirdPartyCheck.registeredName || (locale === 'de' ? 'Diese Person' : 'This person');
                      const d = thirdPartyCheck.registeredDate ? new Date(thirdPartyCheck.registeredDate) : null;
                      const dateStr = d && !isNaN(d.getTime())
                        ? d.toLocaleDateString(locale === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' })
                        : '';
                      return (
                        <div>
                          <strong>
                            {nm}{locale === 'de' ? ' ist bereits für das Event angemeldet' : ' is already registered for this event'}
                            {dateStr ? ` (${locale === 'de' ? 'Anmeldedatum' : 'Registered'}: ${dateStr})` : ''}.
                          </strong>
                          <div style={{ marginTop: 4, fontWeight: 400 }}>
                            {locale === 'de'
                              ? 'Eine erneute Anmeldung ist nicht möglich. Bitte wähle eine andere Person.'
                              : 'Registering this person again is not possible. Please pick a different person.'}
                          </div>
                        </div>
                      );
                    })()}
                    {thirdPartyCheck.notInAudience && (
                      <div style={{ marginTop: thirdPartyCheck.alreadyRegistered ? 6 : 0 }}>
                        <strong>{t('reg.thirdparty.notinaudience')}</strong>
                      </div>
                    )}
                  </div>
                )}
              </>
            )}

            {/* v11.80: Anrede-Dropdown nur rendern, wenn das Event das Feld
                explizit abfragt (event.askSalutation === true). Default: aus —
                viele Events brauchen die Anrede nicht. Wenn nicht gerendert,
                bleibt salutation '' und wird so in die Teilnehmer-Zeile
                geschrieben. */}
            {event.askSalutation && (
              <div className="form-group">
                <label className="form-label"><span className="required">*</span> {t('reg.salutation')}</label>
                <select className="form-select" value={salutation} onChange={e => setSalutation(e.target.value as Salutation)} style={showErrors && !salutation ? errorBorder : {}}>
                  <option value="">{t('reg.pleaseselect')}</option>
                  <option value="Herr">{locale === 'de' ? 'Herr' : 'Mr'}</option>
                  <option value="Frau">{locale === 'de' ? 'Frau' : 'Mrs'}</option>
                  <option value="Divers">{locale === 'de' ? 'Divers' : 'Diverse'}</option>
                  <option value="Keine Angabe">{locale === 'de' ? 'Keine Angabe' : 'Prefer not to say'}</option>
                </select>
              </div>
            )}

            {/* v27.13 (Feedback E. Brenneisen): Statt der grauen Feldliste eine
                Profil-KARTE mit großem Foto, Name, Position und Standort. Ein
                Plus-Toggle klappt die vollständige Liste der automatisch aus
                dem M365-Profil übernommenen Daten auf; darunter der Hinweis auf
                den automatischen Abgleich + Ticket-Verweis bei falschen Daten.
                Gilt für die EIGENE Anmeldung UND die Anmeldung Dritter mit
                Deloitte-Profil. Externe Personen (kein M365-Profil) und der
                „noch niemand gewählt"-Zustand behalten die klassischen Felder. */}
            {(() => {
              const profile = registerForOther ? pickedUserProfile : currentUser;
              const jt = profile ? ((profile as { jobTitle?: string }).jobTitle || '') : '';
              const dept = profile ? ((profile as { department?: string }).department || '') : '';
              const loc = profile ? ((profile as { location?: string }).location || '') : '';
              // v24.29: Unternehmenszugehörigkeit / Rechtsträger read-only.
              const comp = profile ? ((profile as { company?: string }).company || '') : '';
              const displayName = `${firstName} ${surname}`.trim();
              const showProfileCard = !externalPerson && !!email.trim() && !!displayName;
              if (showProfileCard) {
                const notSet = locale === 'de' ? 'nicht hinterlegt' : 'not set';
                const initials = `${(firstName.trim()[0] || '')}${(surname.trim()[0] || '')}`.toUpperCase();
                const detailRows: Array<{ label: string; value: string }> = [
                  { label: locale === 'de' ? 'E-Mail' : 'Email', value: email },
                  { label: 'Position', value: jt },
                  { label: locale === 'de' ? 'Geschäftsbereich' : 'Business Area', value: dept },
                  { label: locale === 'de' ? 'Unternehmen' : 'Company', value: comp },
                  { label: locale === 'de' ? 'Büro' : 'Office', value: loc },
                ];
                return (
                  <div className="form-group">
                    <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: '16px 18px', background: '#fff' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
                        {/* Foto: userphoto.aspx mit Initialen-Fallback (Bild
                            liegt über dem Initialen-Kreis; bei Ladefehler
                            wird es ausgeblendet und die Initialen bleiben). */}
                        <div style={{ position: 'relative', width: 88, height: 88, borderRadius: '50%', background: 'var(--dex-gray-100)', flexShrink: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: '1.5rem', color: 'var(--dex-gray-500)', overflow: 'hidden' }}>
                          {initials || '?'}
                          <img
                            src={`/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email.trim())}`}
                            alt=""
                            style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'cover' }}
                            onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
                          />
                        </div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 700, fontSize: '1.18rem', color: 'var(--dex-gray-800)' }}>{displayName}</div>
                          {jt && (
                            <div style={{ color: 'var(--dex-gray-600)', marginTop: 2 }}>{jt}</div>
                          )}
                          {loc && (
                            <div style={{ display: 'inline-flex', alignItems: 'center', gap: 5, color: 'var(--dex-gray-500)', fontSize: '0.88rem', marginTop: 3 }}>
                              <Icon iconName="POI" style={{ fontSize: 14, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                              {loc}
                            </div>
                          )}
                        </div>
                        {/* Plus-Toggle: zeigt ALLE automatisch übernommenen Daten. */}
                        <button
                          type="button"
                          onClick={() => setProfileCardExpanded(o => !o)}
                          title={profileCardExpanded
                            ? (locale === 'de' ? 'Details einklappen' : 'Collapse details')
                            : (locale === 'de' ? 'Alle automatisch übernommenen Daten anzeigen' : 'Show all automatically applied data')}
                          aria-expanded={profileCardExpanded}
                          style={{
                            width: 36, height: 36, borderRadius: '50%', flexShrink: 0,
                            border: '1px solid var(--dex-gray-300)', background: profileCardExpanded ? 'var(--dex-gray-100)' : '#fff',
                            cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center',
                            fontSize: '1.25rem', lineHeight: 1, color: 'var(--dex-gray-600)', fontWeight: 600,
                          }}
                        >
                          {profileCardExpanded ? '−' : '+'}
                        </button>
                      </div>
                      {profileCardExpanded && (
                        <div style={{ marginTop: 14, borderTop: '1px solid var(--dex-gray-100)', paddingTop: 10 }}>
                          <div style={{ fontSize: '0.72rem', fontWeight: 700, color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 6 }}>
                            {locale === 'de' ? 'Automatisch übernommene Daten' : 'Automatically applied data'}
                          </div>
                          {detailRows.map(row => (
                            <div key={row.label} style={{ display: 'flex', gap: 10, padding: '4px 0', fontSize: '0.86rem', borderBottom: '1px solid var(--dex-gray-50, #fafafa)' }}>
                              <span style={{ width: 140, flexShrink: 0, color: 'var(--dex-gray-500)' }}>{row.label}</span>
                              <span style={{ color: row.value ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)', wordBreak: 'break-word' }}>
                                {row.value || `— ${notSet}`}
                              </span>
                            </div>
                          ))}
                        </div>
                      )}
                      {/* v28.1: Hinweis bewusst klein + kursiv, ohne
                          ServiceNow-Verweis (Profildaten-Fehler sind selten;
                          der Weg zur IT ist den Kolleg:innen bekannt). */}
                      <div style={{ marginTop: 10, fontSize: '0.68rem', fontStyle: 'italic', color: 'var(--dex-gray-400)', lineHeight: 1.45 }}>
                        {locale === 'de'
                          ? <>Diese Angaben werden automatisch mit {registerForOther ? 'dem Microsoft-Profil (M365) der ausgewählten Person' : 'deinen Microsoft-Anmeldedaten (M365-Profil)'} abgeglichen und können hier nicht bearbeitet werden.</>
                          : <>These details are automatically synced with {registerForOther ? 'the selected person’s Microsoft profile (M365)' : 'your Microsoft sign-in data (M365 profile)'} and cannot be edited here.</>}
                      </div>
                    </div>
                  </div>
                );
              }
              // v28.6: „Für andere" ohne gewählte Person → KEINE leeren
              // Alt-Felder mehr (der Wizard ist ohnehin offen); stattdessen
              // ein kompakter Hinweis. Die Profil-Karte erscheint, sobald
              // eine Person gewählt wurde.
              if (registerForOther && !externalPerson) {
                return (
                  <div style={{ padding: '16px 18px', border: '1px dashed var(--dex-gray-300)', borderRadius: 10, background: 'var(--dex-gray-50, #fafafa)', color: 'var(--dex-gray-500)', fontSize: '0.85rem', lineHeight: 1.5 }}>
                    {locale === 'de'
                      ? 'Wähle zuerst eine Person aus (Fenster „Für eine andere Person anmelden") — ihre Daten erscheinen dann hier als Profil-Karte.'
                      : 'First pick a person (window “Register another person”) — their details then appear here as a profile card.'}
                  </div>
                );
              }
              // Klassische Felder: nur noch für EXTERNE Personen.
              // v28.68: Sicherheitsnetz. Konnte der Name nicht aus M365
              // aufgelöst werden (z.B. weil in der versteckten Benutzerliste
              // der Site das Claims-Login-Token statt des Namens steht und
              // auch das Benutzerprofil nichts hergibt), waren diese Felder
              // fest deaktiviert, die Pflichtprüfung verlangte aber Vor- UND
              // Nachnamen: „Bitte alle Pflichtfelder ausfüllen" ohne ein
              // einziges ausfüllbares Feld — die Anmeldung war unmöglich.
              // Fehlt einer der beiden Namen, sind sie jetzt editierbar.
              const nameUnresolved = !firstName.trim() || !surname.trim();
              return (
                <>
                  <div className="form-group">
                    <label className="form-label">{t('reg.firstname')}</label>
                    <input className="form-input" value={firstName} onChange={e => { if (externalPerson || nameUnresolved) setFirstName(e.target.value); }} placeholder={t('reg.firstname')} disabled={!externalPerson && !nameUnresolved} style={{ background: nameUnresolved ? 'var(--dex-white, #fff)' : 'var(--dex-gray-100)', ...(showErrors && !firstName.trim() ? errorBorder : {}) }} />
                  </div>
                  <div className="form-group">
                    <label className="form-label">{t('reg.surname')}</label>
                    <input className="form-input" value={surname} onChange={e => { if (externalPerson || nameUnresolved) setSurname(e.target.value); }} placeholder={t('reg.surname')} disabled={!externalPerson && !nameUnresolved} style={{ background: nameUnresolved ? 'var(--dex-white, #fff)' : 'var(--dex-gray-100)', ...(showErrors && !surname.trim() ? errorBorder : {}) }} />
                  </div>
                  {nameUnresolved && !externalPerson && (
                    <div style={{
                      marginTop: -4, marginBottom: 12, padding: '8px 10px', borderRadius: 6,
                      fontSize: '0.78rem', lineHeight: 1.5,
                      background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
                    }}>
                      {locale === 'de'
                        ? 'Dein Name konnte nicht aus deinem M365-Profil gelesen werden. Bitte trage Vor- und Nachnamen einmal von Hand ein — danach kannst du dich ganz normal anmelden.'
                        : 'We could not read your name from your M365 profile. Please enter your first and last name once — after that you can register as usual.'}
                    </div>
                  )}
                  <div className="form-group">
                    <label className="form-label">{t('reg.email')}</label>
                    <input className="form-input" type="email" value={email} onChange={e => { if (externalPerson) { setEmail(e.target.value); externalEmailConfirmedRef.current = false; /* v18.74: Tippfehler-Check bei Änderung erneut erzwingen */ } }} placeholder={externalPerson ? 'name@firma.de' : 'email@deloitte.de'} disabled style={{ background: 'var(--dex-gray-100)', ...(showErrors && !email.trim() ? errorBorder : {}) }} />
                  </div>
                </>
              );
            })()}

            {/* v11.82: Team-Anmeldung-Toggle. Nur sichtbar wenn der Organizer
                in Schritt 4 die Team-Anmeldung aktiviert hat UND der User sich
                NICHT für eine andere Person registriert (Team-für-Andere wird
                nicht unterstützt — der Stellvertreter-Pfad ist auf eine
                Einzel-Person ausgelegt). */}
            {isTeamCapable && !registerForOther && !parentAlreadyRegistered && (
              <div className="form-group" style={{ marginTop: 16, marginBottom: 0 }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={isTeamMode}
                    onChange={e => { setIsTeamMode(e.target.checked); if (e.target.checked) setPendingJoinTeam(null); }}
                    style={{ marginTop: 3 }}
                  />
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)' }}>
                      {locale === 'de'
                        ? `Ich melde mich + mein Team an (Team-Anmeldung)`
                        : 'Register me + my team (team registration)'}
                    </div>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4, lineHeight: 1.4 }}>
                      {locale === 'de'
                        ? `Belegt bis zu ${teamSize} Plätze auf einmal. Jedes Mitglied bekommt automatisch Bestätigungsmail, Outlook-Termin und sieht das Event in „Meine Events".`
                        : `Books up to ${teamSize} seats at once. Each member automatically receives a confirmation email, an Outlook invite, and sees the event in „My Events".`}
                    </div>
                  </div>
                </label>
              </div>
            )}
          </div>
          </CollapsibleSection>
        </div>

        {/* v11.82: Team-Anmeldung-Card — separat unter „Persönliche Daten",
            nur sichtbar wenn der Toggle aktiv ist. */}
        {isTeamCapable && isTeamMode && !registerForOther && !parentAlreadyRegistered && (
          <div className="registration-form" style={{ marginTop: 24 }}>
            <CollapsibleSection
              isMobile={isMobile}
              icon="People"
              title={locale === 'de' ? 'Team-Anmeldung' : 'Team registration'}
            >
            <div style={{ padding: '24px 20px' }}>
              {/* Pflicht-Hinweis-Box ganz oben — auffällig orange. */}
              <div style={{
                marginBottom: 20,
                padding: '14px 16px',
                background: 'rgba(237,139,0,0.10)',
                border: '2px solid var(--dex-orange, #ed8b00)',
                borderRadius: 'var(--dex-radius-md)',
                color: '#7a4a00',
                fontSize: '0.88rem',
                lineHeight: 1.55,
              }}>
                <div style={{ fontWeight: 700, marginBottom: 8, fontSize: '0.95rem' }}>
                  {locale === 'de'
                    ? 'Vorab die Zustimmung jedes Teammitglieds einholen'
                    : 'Get every team member\'s consent up front'}
                </div>
                <div style={{ marginBottom: 8 }}>
                  {locale === 'de'
                    ? 'Mit dem Absenden meldest du nicht nur dich selbst an, sondern auch alle weiter unten eingetragenen Personen. Jedes Teammitglied erhält automatisch:'
                    : 'By submitting you register yourself AND every person you add below. Each team member automatically receives:'}
                </div>
                <ul style={{ margin: '0 0 8px 18px', padding: 0 }}>
                  <li>{locale === 'de' ? 'eine Anmeldebestätigung per Mail' : 'a confirmation email'}</li>
                  <li>{locale === 'de' ? 'einen Outlook-Termin im Kalender' : 'an Outlook calendar invite'}</li>
                  <li>{locale === 'de' ? 'den Event in „Meine Events"' : 'the event in „My Events"'}</li>
                </ul>
                <div style={{ marginTop: 4 }}>
                  {locale === 'de'
                    ? <>Bitte stelle sicher, dass alle Teilnehmer ihrer Anmeldung <strong>VORHER zugestimmt</strong> haben — eine Anmeldung ohne Einverständnis ist nicht erlaubt.</>
                    : <>Please make sure every participant has <strong>consented up front</strong> — registering people without their consent is not allowed.</>}
                </div>
              </div>

              {event?.askTeamName && (
                <div className="form-group">
                  <label className="form-label">
                    <span className="required">*</span> {locale === 'de' ? 'Team-Name' : 'Team name'}
                  </label>
                  <input
                    className="form-input"
                    value={teamName}
                    onChange={e => setTeamName(e.target.value.slice(0, 60))}
                    placeholder={locale === 'de' ? 'z.B. „Die Schnellen"' : 'e.g. „The Quick Ones"'}
                    style={showErrors && isTeamMode && event?.askTeamName && !teamName.trim() ? errorBorder : {}}
                  />
                  <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>
                    {locale === 'de' ? 'Max. 60 Zeichen — wird in der Teilnehmerliste mitgespeichert.' : 'Max 60 characters — stored on the participant list.'}
                  </div>
                </div>
              )}

              {/* Member-Slots */}
              <div style={{ marginTop: 8 }}>
                {teamMembers.map((mv, idx) => {
                  const slotRequired = !teamPartialAllowed;
                  const parsed = parseTeamMember(mv);
                  const isErr = showErrors && isTeamMode && slotRequired && !parsed;
                  return (
                    <div className="form-group" key={`team-slot-${idx}`}>
                      <label className="form-label">
                        {slotRequired && <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>}
                        {locale === 'de'
                          ? `Mitglied ${idx + 2}${slotRequired ? '' : ' (optional)'}`
                          : `Member ${idx + 2}${slotRequired ? '' : ' (optional)'}`}
                      </label>
                      <UserFieldPicker
                        value={mv}
                        onChange={v => {
                          const next = [...teamMembers];
                          next[idx] = v;
                          setTeamMembers(next);
                        }}
                        searchUsers={async (q, includeIntl) => {
                          const results = await searchUsers(q, includeIntl);
                          return results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location, jobTitle: r.jobTitle }));
                        }}
                        searchUserByEmail={searchUser}
                        placeholder={locale === 'de' ? 'Name oder E-Mail eingeben...' : 'Type a name or email...'}
                        errorStyle={isErr ? errorBorder : {}}
                        forcedIsDe={locale === 'de'}
                      />
                      {/* v18.12: Custom-Fields pro Team-Mitglied — erscheinen,
                          sobald die Person ausgewählt ist (z.B. Essenspräferenz). */}
                      {parsed && teamMemberApplicableFields.length > 0 && (
                        <div style={{ marginTop: 8, marginLeft: 8, paddingLeft: 12, borderLeft: '2px solid var(--dex-gray-200)' }}>
                          {teamMemberApplicableFields
                            .filter(f => {
                              if (!f.showIf) return true;
                              const mv = teamMemberFields[idx] || {};
                              const v = mv[f.showIf.fieldId] || '';
                              const parts = v.split(' | ').map(s => s.trim());
                              return f.showIf.values.some(x => v === x || parts.indexOf(x) >= 0);
                            })
                            .map(f => renderRegField(
                              f,
                              teamMemberFields[idx] || {},
                              next => setTeamMemberFields(prev => ({ ...prev, [idx]: next }))
                            ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* Pflicht-Bestätigungs-Checkbox */}
              <div className="form-group" style={{ marginTop: 18, marginBottom: 0 }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={teamConsentConfirmed}
                    onChange={e => setTeamConsentConfirmed(e.target.checked)}
                    style={{ marginTop: 3 }}
                  />
                  <div style={{ flex: 1, fontSize: '0.88rem', color: 'var(--dex-gray-800)', lineHeight: 1.5 }}>
                    <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                    {locale === 'de'
                      ? 'Ich bestätige, dass alle eingetragenen Teammitglieder ihrer Anmeldung zugestimmt haben.'
                      : 'I confirm that every listed team member has consented to this registration.'}
                  </div>
                </label>
              </div>
            </div>
            </CollapsibleSection>
          </div>
        )}

        {/* v18.73: Offene Teams — sichtbar wenn der Organizer „Offene Slots
            öffentlich sichtbar" aktiviert hat und es Teams gibt, denen Plätze
            fehlen. Steht jetzt UNTER der „Ich melde mich + mein Team an"-Karte.
            Klick auf „Vormerken" wählt ein Team nur vor — die eigentliche
            Anmeldung (inkl. der oben/unten ausgefüllten persönlichen +
            event-spezifischen Felder) passiert erst über den „Anmelden"-Button.
            */}
        {event && event.teamRegistrationEnabled && event.teamOpenSlotsVisible && !registerForOther && openTeamsLoaded && openTeams.length > 0 && !parentAlreadyRegistered && (
          <div className="registration-form" style={{ marginBottom: 16 }}>
            <CollapsibleSection
              isMobile={isMobile}
              icon="People"
              title={locale === 'de' ? 'Offene Teams — einem unvollständigen Team beitreten' : 'Open teams — join an incomplete team'}
            >
            <div style={{ padding: '20px' }}>
              <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)', marginTop: 0, marginBottom: 12, lineHeight: 1.5 }}>
                {locale === 'de'
                  ? 'Andere Personen haben Teams angemeldet, denen noch Plätze fehlen. Du kannst eines vormerken — fülle dann oben deine persönlichen Daten und unten die event-spezifischen Angaben aus und klicke auf „Anmelden", um beizutreten.'
                  : 'Other people have registered teams with open slots. Pre-select one — then fill in your personal details above and the event-specific information below, and click „Register" to join.'}
                {event.teamJoinRequiresApproval && (
                  <> {locale === 'de'
                    ? <><br /><strong>Hinweis:</strong> der Team-Kapitän muss deinen Beitritt erst bestätigen.</>
                    : <><br /><strong>Note:</strong> the team lead has to approve your join.</>}
                  </>
                )}
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {openTeams.map(t => {
                  const free = t.teamSize - t.activeCount;
                  const isPicked = !!pendingJoinTeam && pendingJoinTeam.teamId === t.teamId;
                  return (
                    <div key={t.teamId} style={{
                      display: 'flex', alignItems: 'center', gap: 12,
                      padding: '10px 12px',
                      background: isPicked ? 'rgba(134,188,37,0.10)' : 'var(--dex-gray-50, #f7f7f7)',
                      borderRadius: 6,
                      border: isPicked ? '2px solid var(--dex-green, #86bc25)' : '1px solid var(--dex-gray-200)',
                    }}>
                      <Icon iconName="Group" style={{ fontSize: 16, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                      <div style={{ flex: 1, fontSize: '0.88rem' }}>
                        <div style={{ fontWeight: 600 }}>
                          {locale === 'de'
                            ? `Team „${t.teamName || 'ohne Namen'}"`
                            : `Team „${t.teamName || 'unnamed'}"`}
                        </div>
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                          {locale === 'de'
                            ? `${t.activeCount}/${t.teamSize} belegt — ${free} Slot${free === 1 ? '' : 's'} frei`
                            : `${t.activeCount}/${t.teamSize} taken — ${free} slot${free === 1 ? '' : 's'} free`}
                        </div>
                      </div>
                      <button
                        type="button"
                        className={isPicked ? 'btn btn-secondary' : 'btn btn-primary'}
                        onClick={() => togglePendingJoinTeam(t.teamId, t.teamName)}
                        style={{ fontSize: '0.82rem', padding: '6px 12px' }}
                      >
                        {isPicked
                          ? (locale === 'de' ? 'Vorgemerkt ✓ — entfernen' : 'Pre-selected ✓ — remove')
                          : (locale === 'de' ? 'Vormerken' : 'Pre-select')}
                      </button>
                    </div>
                  );
                })}
              </div>
              <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 10, marginBottom: 0, lineHeight: 1.4 }}>
                {locale === 'de'
                  ? 'Mitgliedernamen werden aus Privatsphäre-Gründen nicht angezeigt.'
                  : 'Member names are hidden for privacy reasons.'}
              </p>
              {pendingJoinTeam && (
                <div style={{
                  marginTop: 12,
                  padding: '10px 12px',
                  borderRadius: 6,
                  background: 'rgba(134,188,37,0.10)',
                  border: '1px solid var(--dex-green, #86bc25)',
                  color: 'var(--dex-green-dark, #3f5f10)',
                  fontSize: '0.85rem',
                  lineHeight: 1.5,
                }}>
                  {locale === 'de'
                    ? <>Team <strong>„{pendingJoinTeam.teamName || 'ohne Namen'}“</strong> ist vorgemerkt. Fülle deine Angaben aus und klicke unten auf <strong>„Anmelden“</strong>, um den Beitritt abzuschließen.</>
                    : <>Team <strong>“{pendingJoinTeam.teamName || 'unnamed'}”</strong> is pre-selected. Fill in your details and click <strong>“Register”</strong> below to complete your join.</>}
                </div>
              )}
            </div>
            </CollapsibleSection>
          </div>
        )}

        {/* Eventspezifische Felder (inkl. Split-Capacity Starter-Typ-Auswahl wenn
            beide Kapazitäten > 0; bei nur einem verfügbaren Typ wird dieser
            automatisch gesetzt und gar nicht angezeigt). v10.20: Sessions-/
            Hauptevent-Auswahl ist hierher gewandert (vorher links unter der
            Event-Karte). */}
        {/* v18.73: Die „Event-spezifische Informationen"-Karte nur anzeigen,
            wenn es dort tatsächlich etwas auszufüllen/auszuwählen gibt — also
            Custom-Felder, eine Gruppen-Auswahl (Split) ODER eine Sub-Event-
            Auswahl. Sonst (leeres „Keine zusätzlichen Informationen
            erforderlich") wird die Karte komplett ausgeblendet. */}
        {/* v28.2: Station 3 — Anmeldung abschließen (immer sichtbar; die
            Event-Felder-Karte darunter nur, wenn es etwas auszufüllen gibt). */}
        <div className="reg-step-head">
          <span className="reg-step-num">3</span>
          <span className="reg-step-label">{locale === 'de' ? 'Anmeldung abschließen' : 'Complete your registration'}</span>
        </div>
        {/* v29.9: …und die Karte muss auch dann erscheinen, wenn für diese
            Person KEIN Programmpunkt sichtbar ist — sonst fehlt der Hinweis
            darunter genau in dem Fall, für den er gedacht ist. */}
        {(event.eventSpecificFields.length > 0 || isSplitGroup || childEvents.length > 0
          || (hiddenChildCount > 0 && event.subEventsOnlyMode)) && (
        <div className="registration-specific">
          {/* v11.97: Section-Header + „* = Required field"-Legende in
              einer Zeile. Legende mit ROTEM Stern (vorher war der Stern
              in der Erklärung grau, jetzt im Deloitte-Rot wie alle echten
              Required-Marker). */}
          <CollapsibleSection
            isMobile={isMobile}
            icon="EditNote"
            title={t('reg.eventinfo')}
            collapsible={false}
            headerExtra={
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 12, padding: '0 12px' }}>
              <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                <span style={{ color: 'var(--dex-red, #da291c)', fontWeight: 700, marginRight: 2 }}>*</span> = {t('reg.requiredfield')}
              </span>
            </span>
            }
          >
          <div style={{ padding: '24px 20px' }}>
            {/* v11.10: Group-Selection ist ein eigener, IMMER sichtbarer
                Block (sofern das Event Split-Capacity hat). Vorher war er
                inkorrekt INNERHALB der Sub-Events-Auswahl genistet, sodass
                Events ohne Sub-Events keine Gruppen-Buttons hatten und
                Drittpersonen-Registrierungen die Gruppen-Wahl gar nicht
                anzeigten. Sub-Events erben jetzt einfach
                preferredStarterType — keine Pro-Sub-Event-Radios mehr. */}
            {isSplitGroup && (
              <div style={{ marginBottom: 20, border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 16 }}>
                <label className="form-label" style={{ fontWeight: 700, marginBottom: 6 }}>
                  <span className="required">*</span> {(event.splitSectionTitle && event.splitSectionTitle.trim()) ? event.splitSectionTitle : (locale === 'de' ? 'Gruppen-Auswahl' : 'Group selection')}
                </label>
                {/* v26.83: Organizer-eigener Hinweistext (splitHelpText) hat
                    Vorrang; sonst der Standardsatz. whiteSpace pre-wrap, damit
                    Zeilenumbrüche aus dem Wizard erhalten bleiben. */}
                <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 10, whiteSpace: 'pre-wrap' }}>
                  {(event.splitHelpText && event.splitHelpText.trim())
                    ? event.splitHelpText
                    : (locale === 'de'
                      ? `Wähle eine der zwei Gruppen aus. Ist die Wunsch-Gruppe voll, kannst du automatisch in die andere wechseln oder auf der Warteliste warten.`
                      : 'Pick one of the two groups. If your preferred group is full, you can either switch to the other or join the waitlist.')}
                </p>
                {/* v19.19: Gesamt-Kapazitäts-Zusammenfassung — Gesamtzahl der
                    Plätze, aktuell freie Plätze (geklammert ≥ 0) und die Zahl
                    der Personen auf der Warteliste. WICHTIG: hier wird NIE eine
                    Überbuchung sichtbar — freie Plätze sind bei 0 gedeckelt,
                    bei Vollbelegung steht „ausgebucht". Die echte
                    Überbuchungszahl ist ausschließlich dem Organizer/Admin
                    im Admin Center vorbehalten. */}
                {(() => {
                  const totalCap = durchCap + funCap;
                  const dActive = starterCounts?.durch ?? 0;
                  const fActive = starterCounts?.fun ?? 0;
                  const totalFree = Math.max(0, durchCap - dActive) + Math.max(0, funCap - fActive);
                  const totalWait = (starterCounts?.durchWait ?? 0) + (starterCounts?.funWait ?? 0);
                  return (
                    <div style={{
                      display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap',
                      marginBottom: 12, padding: '8px 12px', borderRadius: 8,
                      background: 'var(--dex-gray-50, #f7f7f7)', border: '1px solid var(--dex-gray-200)',
                      fontSize: '0.8rem',
                    }}>
                      <span style={{ color: 'var(--dex-gray-700)', fontWeight: 700 }}>
                        {locale === 'de' ? 'Gesamtkapazität:' : 'Total capacity:'} {totalCap} {locale === 'de' ? 'Plätze' : 'seats'}
                      </span>
                      {/* v26.72: „X frei" in der Gesamt-Zeile entfernt — die
                          Verfügbarkeit steht bereits pro Gruppe in den Karten.
                          Nur bei komplett ausgebucht bleibt ein Hinweis. */}
                      {totalFree <= 0 && (
                        <>
                          <span style={{ color: 'var(--dex-gray-300)' }}>·</span>
                          <span style={{ color: 'var(--dex-red, #c00)', fontWeight: 700 }}>
                            {locale === 'de' ? 'ausgebucht' : 'fully booked'}
                          </span>
                        </>
                      )}
                      {totalWait > 0 && (
                        <>
                          <span style={{ color: 'var(--dex-gray-300)' }}>·</span>
                          <span style={{ color: 'var(--dex-gray-600)' }}>
                            {totalWait} {locale === 'de'
                              ? (totalWait === 1 ? 'Person auf der Warteliste' : 'Personen auf der Warteliste')
                              : (totalWait === 1 ? 'person on the waitlist' : 'people on the waitlist')}
                            {event.splitSharedWaitlist ? (locale === 'de' ? ' (gemeinsam)' : ' (shared)') : ''}
                          </span>
                        </>
                      )}
                    </div>
                  );
                })()}
                <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                  {(() => {
                    // v26.72: Beschreibung pro Gruppe frei konfigurierbar
                    // (splitDescA/B aus dem Wizard); Fallback auf den B2Run-
                    // Standardtext nur bei den Default-Labels.
                    const optA = { id: 'Durchstarter', label: splitLabelA, desc: (event?.splitDescA && event.splitDescA.trim()) || (splitLabelA === 'Durchstarter' ? t('reg.starter.durch.desc') : ''), cap: durchCap, count: starterCounts?.durch ?? 0, wait: starterCounts?.durchWait ?? 0 };
                    const optB = { id: 'Funstarter', label: splitLabelB, desc: (event?.splitDescB && event.splitDescB.trim()) || (splitLabelB === 'Funstarter' ? t('reg.starter.fun.desc') : ''), cap: funCap, count: starterCounts?.fun ?? 0, wait: starterCounts?.funWait ?? 0 };
                    // v11.25: pure UI-Reihenfolge — bei reversed wird Karte B
                    // zuerst gerendert. Interne IDs/Capacities/StarterType der
                    // Anmeldungen bleiben unangetastet.
                    return event.splitDisplayOrderReversed ? [optB, optA] : [optA, optB];
                  })().map(opt => {
                    const free = opt.cap - opt.count;
                    const isFull = free <= 0;
                    const isActive = preferredStarterType === opt.id;
                    // v26.72: gewählte Box grün, nicht-gewählte grau (vorher A grün / B orange).
                    const accent = isActive ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-500, #6b7280)';
                    return (
                      <button
                        key={opt.id}
                        type="button"
                        onClick={() => setPreferredStarterType(opt.id)}
                        style={{
                          padding: 14, textAlign: 'left',
                          borderRadius: 'var(--dex-radius, 12px)',
                          border: isActive ? `2px solid ${accent}` : '2px solid var(--dex-gray-200)',
                          // v26.88: Standard-Grün (wie die Geschlecht-/Feld-Füllung,
                          // greenFilledStyle) statt Off-Brand #f0fdf4.
                          background: isActive ? 'rgba(134,188,37,0.06)' : '#fff',
                          cursor: 'pointer', transition: 'all 0.15s',
                          position: 'relative',
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                          <strong style={{ color: accent, fontSize: '0.95rem' }}>{opt.label}</strong>
                          {isActive && <span style={{ color: accent, fontSize: '0.8rem' }}>✓</span>}
                        </div>
                        {opt.desc && <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 6, whiteSpace: 'pre-wrap' }}>{opt.desc}</div>}
                        <div style={{ fontSize: '0.78rem' }}>
                          {isFull ? (
                            <span style={{ color: 'var(--dex-red, #c00)', fontWeight: 600 }}>{t('reg.starter.full')}</span>
                          ) : (
                            // v19.19: nie negativ — bei Überbuchung greift der
                            // isFull-Zweig oben (zeigt „Voll"), die echte
                            // Überbuchungszahl bleibt dem Organizer/Admin vorbehalten.
                            <span style={{ color: accent }}>{`${Math.max(0, free)} / ${opt.cap} ${t('reg.starter.free')}`}</span>
                          )}
                          {/* v19.19: Warteliste pro Gruppe — nur bei GETRENNTEN
                              Wartelisten. Bei gemeinsamer Warteliste steht die
                              Zahl gesammelt in der Kapazitäts-Zusammenfassung. */}
                          {!event.splitSharedWaitlist && opt.wait > 0 && (
                            <span style={{ display: 'block', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {opt.wait} {locale === 'de'
                                ? (opt.wait === 1 ? 'Person auf der Warteliste' : 'Personen auf der Warteliste')
                                : (opt.wait === 1 ? 'person on the waitlist' : 'people on the waitlist')}
                            </span>
                          )}
                        </div>
                      </button>
                    );
                  })}
                </div>

                {event.durchstarterRequiresProof && preferredStarterType === 'Durchstarter' && (
                  <div style={{ marginTop: 12, padding: '10px 12px', background: 'rgba(237,139,0,0.06)', border: '1px solid var(--dex-orange)', borderRadius: 8 }}>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.85rem' }}>
                      <input
                        type="checkbox"
                        checked={eventSpecific['b2run_leistungsnachweis'] === 'true'}
                        onChange={e => setEventSpecific({ ...eventSpecific, b2run_leistungsnachweis: e.target.checked ? 'true' : 'false' })}
                        style={{ marginTop: 3 }}
                      />
                      <span>
                        <strong>{t('reg.starter.proof') || 'Leistungsnachweis vorhanden'} <span className="required">*</span></strong>
                        <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                          {t('reg.starter.proof.hint') || 'Ich bestätige, dass ein entsprechender Leistungsnachweis (z.B. Wettkampfergebnis, Trainingsnachweis) vorliegt.'}
                        </span>
                      </span>
                    </label>
                    {showErrors && eventSpecific['b2run_leistungsnachweis'] !== 'true' && (
                      <div style={{ marginTop: 6, fontSize: '0.75rem', color: 'var(--dex-red)' }}>
                        {t('reg.starter.proof.required') || 'Bitte Leistungsnachweis bestätigen.'}
                      </div>
                    )}
                  </div>
                )}
                {/* v11.12: Custom-Fields mit onlyForGroup-Constraint
                    direkt INNERHALB der Gruppen-Auswahl-Box rendern.
                    Klappt erst auf, wenn der User eine Gruppe gewählt
                    hat und das Feld der gewählten Gruppe entspricht.
                    Gleicher orange-getönter Style wie der Legacy-
                    Leistungsnachweis-Block — sodass jede Gruppen-
                    spezifische Abfrage optisch klar als „Folge der
                    Gruppen-Wahl" erkennbar ist. */}
                {preferredStarterType && (() => {
                  const groupSpec = event.eventSpecificFields
                    .filter(f => f.id !== 'b2run_mobilnummer' || eventSpecific['b2run_infoservice'] === 'true')
                    .filter(f => !(f.id === 'b2run_startblock' && hasStarterBlockMapping))
                    .filter(f => {
                      if (!f.showIf || !f.showIf.fieldId) return true;
                      const raw = (eventSpecific[f.showIf.fieldId] || '').trim();
                      if (!raw) return false;
                      const answers = raw.indexOf(' | ') >= 0
                        ? raw.split(' | ').map(s => s.trim()).filter(Boolean)
                        : [raw];
                      return answers.some(a => f.showIf!.values.indexOf(a) >= 0);
                    })
                    .filter(f => {
                      const grp = f.onlyForGroup;
                      if (!grp || grp === 'all') return false;
                      if (grp === 'A') return preferredStarterType === 'Durchstarter';
                      if (grp === 'B') return preferredStarterType === 'Funstarter';
                      return false;
                    });
                  if (groupSpec.length === 0) return null;
                  const labelA = (event.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
                  const labelB = (event.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
                  const grpLabel = preferredStarterType === 'Durchstarter' ? labelA : labelB;
                  return (
                    <div style={{
                      marginTop: 12, padding: '12px 14px',
                      background: 'rgba(237,139,0,0.06)',
                      border: '1px solid var(--dex-orange)',
                      borderRadius: 8,
                      display: 'flex', flexDirection: 'column', gap: 12,
                    }}>
                      <div style={{ fontSize: '0.78rem', fontWeight: 700, color: 'var(--dex-orange, #ed8b00)' }}>
                        {locale === 'de'
                          ? `Zusätzliche Angaben für „${grpLabel}"`
                          : `Additional details for „${grpLabel}"`}
                      </div>
                      {groupSpec.map(f => renderRegField(f))}
                    </div>
                  );
                })()}
              </div>
            )}

            {/* v29.9: Sackgasse sichtbar machen — die Klammer ist nicht
                buchbar, und von ihren Programmpunkten ist für diese Person
                keiner freigegeben. Ohne diesen Hinweis steht da eine
                Anmeldeseite ohne irgendetwas zum Anklicken, und der Grund
                zeigt sich erst beim Klick auf „Registrieren".
                Bewusst NUR in diesem Fall: Ist die Klammer selbst buchbar,
                ist ein enger gefasstes Sub-Event Absicht des Organizers und
                geht die Person nichts an. */}
            {childEvents.length === 0 && hiddenChildCount > 0 && event.subEventsOnlyMode && (
              <div style={{
                marginTop: 12, padding: '10px 14px', borderRadius: 8,
                background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)',
                fontSize: '0.82rem', color: 'var(--dex-orange-dark, #b35a00)', lineHeight: 1.5,
              }}>
                {locale === 'de'
                  ? 'Die Anmeldung läuft hier ausschließlich über die einzelnen Programmpunkte. Für dich ist aktuell keiner davon freigegeben — wenn du teilnehmen möchtest, wende dich bitte an die Organizer.'
                  : 'Registration here runs exclusively via the individual programme items. None of them is currently released for you — if you would like to attend, please contact the organizers.'}
              </div>
            )}

            {/* v11.10: Sub-Events-Auswahl als eigener Block.
                v18.37: jetzt AUCH im Stellvertreter-Modus („Für andere
                registrieren") sichtbar — der Submit-Pfad meldet ausgewählte
                Sub-Events längst für die Zielperson an (registerForEvent mit
                participantEmail), nur die Auswahl-UI war versehentlich hinter
                !registerForOther versteckt. Die Belegungszahlen werden im
                Stellvertreter-Modus über getAllRegistrations geladen, ohne
                Self-Vorbelegung. Der Haupt-Event-Checkbox wird im
                Stellvertreter-Modus ausgeblendet (das Haupt-Event wird dort
                ohnehin immer mit angemeldet). Sub-Events erben
                preferredStarterType vom Group-Selection-Block oben. */}
            {childEvents.length > 0 && (
              <div style={{ marginBottom: 20, border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 16 }}>
                {/* v15.11: im subEventsOnlyMode ist die Hauptevent-Anmeldung
                    deaktiviert — Überschrift + Hinweis entsprechend
                    anpassen, sonst lesen sich „Haupt-Event und … können
                    unabhängig" widersprüchlich. */}
                <h4 style={{ marginTop: 0, marginBottom: 4, fontSize: '0.95rem' }}>
                  {event.subEventsOnlyMode
                    ? (childTermPlural
                        ? (locale === 'de' ? `${childTermPlural} auswählen` : `Select ${childTermPlural}`)
                        : (locale === 'de' ? 'Sub-Events auswählen' : 'Select sub-events'))
                    : (tEvent('reg.selection.title') || 'Wofür möchtest du dich anmelden?')}
                </h4>
                {/* v28.97: Sagen, dass nur eines geht — sonst wundert man sich,
                    warum die vorherige Auswahl verschwindet. */}
                {event.subEventSingleChoice && (
                  <p style={{ fontSize: '0.82rem', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600, marginTop: 0, marginBottom: 8 }}>
                    {locale === 'de'
                      ? `Du kannst genau ${childOneDe} auswählen — ein neuer Klick ersetzt die bisherige Wahl.`
                      : 'You can pick exactly one — a new click replaces your previous choice.'}
                  </p>
                )}
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                  {event.subEventsOnlyMode
                    ? (locale === 'de'
                        ? `Bitte wähle mindestens ${childOneDe} aus, um dich anzumelden.`
                        : `Please pick at least one ${childTermSingular || 'sub-event'} you want to register for.`)
                    : registerForOther
                      ? (locale === 'de'
                          ? `Die Person wird für das Haupt-Event angemeldet. Wähle zusätzlich die gewünschten ${childTermPlural || 'Sub-Events'} aus.`
                          : `The person will be registered for the main event. Additionally pick the desired ${childTermPlural || 'sub-events'}.`)
                    : (childTermPlural
                        ? (locale === 'de'
                            ? `Haupt-Event und ${childTermPlural} können unabhängig voneinander an- oder abgewählt werden.`
                            : `Main event and ${childTermPlural} can be selected or deselected independently.`)
                        : (tEvent('reg.selection.hint') || 'Haupt-Event und Sessions können unabhängig voneinander an- oder abgewählt werden.'))}
                </p>

                {/* v15.7: Hauptevent-Card auch hier ausblenden bei
                    subEventsOnlyMode — gleicher Fix wie der primäre Pfad
                    weiter oben. Vorher wurde dieser Render-Pfad (Register
                    for someone else) übersehen.
                    v18.37: im Stellvertreter-Modus ebenfalls ausblenden — die
                    Person wird dort immer für das Haupt-Event angemeldet, ein
                    steuerbarer Haken wäre irreführend. */}
                {!event.subEventsOnlyMode && !registerForOther && (
                <label style={{
                  display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
                  borderRadius: 8,
                  border: `1px solid ${registerForParent && !parentAlreadyRegistered && !parentRegBlocked ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                  background: registerForParent && !parentAlreadyRegistered && !parentRegBlocked ? 'rgba(134,188,37,0.06)' : '#fff',
                  cursor: (parentAlreadyRegistered || parentRegBlocked) ? 'default' : 'pointer',
                  opacity: parentRegBlocked ? 0.6 : 1,
                }}>
                  <input
                    type="checkbox"
                    checked={parentAlreadyRegistered ? true : (parentRegBlocked ? false : registerForParent)}
                    disabled={parentAlreadyRegistered || parentRegBlocked}
                    onChange={e => setRegisterForParent(e.target.checked)}
                    style={{ marginTop: 2 }}
                  />
                  <div style={{ flex: 1 }}>
                    {(() => { const lbl = resolveMainEventLabel(tEvent('reg.selection.mainevent') || 'Haupt-Event'); return (
                      <div style={{ fontWeight: 700 }}>{lbl ? `${lbl}: ` : ''}{event.title}</div>
                    ); })()}
                    {parentAlreadyRegistered && (
                      <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                        {tEvent('reg.selection.alreadyregistered') || 'Du bist bereits für das Haupt-Event angemeldet.'}
                      </div>
                    )}
                    {parentRegBlocked && !parentAlreadyRegistered && (
                      <div style={{ fontSize: '0.75rem', color: parentFullNoWaitlist ? 'var(--dex-red, #c00)' : 'var(--dex-orange, #ed8b00)', marginTop: 2 }}>
                        {parentFullNoWaitlist
                          ? (locale === 'de' ? 'Alle Plätze sind belegt — die Warteliste ist für dieses Event deaktiviert.' : 'All seats are taken — the waitlist is disabled for this event.')
                          : (tEvent('reg.subevents.deadlinepassed') || 'Anmeldefrist abgelaufen — nur noch die offenen Sub-Events sind wählbar.')}
                      </div>
                    )}
                  </div>
                </label>
                )}
                {/* v29.28: Die Fragen zum Haupt-Event stehen DIREKT unter
                    seiner Kachel — dort, wo sie hingehören; die Sub-Event-
                    Fragen stecken in deren Karten (v29.27). Im Stellvertreter-
                    Modus (Haupt-Kachel ausgeblendet, Anmeldung fürs Haupt-
                    Event läuft trotzdem) stehen sie an derselben Stelle über
                    den Sub-Events. Bei einer Klammer rendert der Block an der
                    alten Stelle unter der Auswahl (übergreifende Fragen). */}
                {!event.subEventsOnlyMode && (event.eventSpecificFields.length > 0 || isSplitGroup) && (
                  <div style={{ margin: '8px 0 4px' }}>{renderMainFieldsSection()}</div>
                )}

                {/* v28.91: Termin-Kalender statt Liste — nur, wenn der
                    Organizer die Sub-Events ausdrücklich als Termine angelegt
                    hat (subEventCalendar). Bei neun Tagen ist eine Liste aus
                    neun Funkbuttons kaum zu erfassen; im Kalender sieht man
                    Wochenstruktur, Lücken und freie Plätze auf einen Blick.
                    Es sind dieselben Sub-Events und dasselbe selectedSessions —
                    nur eine andere Darstellung derselben Auswahl. */}
                {!!event.subEventCalendar && (() => {
                  type DayEntry = { ce: typeof childEvents[0]; key: string };
                  const dayOf = (iso?: string): string => {
                    if (!iso) return '';
                    const d = new Date(iso);
                    if (isNaN(d.getTime())) return '';
                    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
                  };
                  const entries: DayEntry[] = childEvents
                    .map(ce => ({ ce, key: dayOf(ce.startDate) }))
                    .filter(e => !!e.key);
                  if (entries.length === 0) return null;
                  const byDay: Record<string, DayEntry> = {};
                  entries.forEach(e => { byDay[e.key] = e; });
                  // Monate, in denen Termine liegen — jeder als eigenes Raster.
                  const monthKeys: string[] = [];
                  entries.forEach(e => {
                    const mk = e.key.slice(0, 7);
                    if (monthKeys.indexOf(mk) < 0) monthKeys.push(mk);
                  });
                  monthKeys.sort();
                  const weekdays = locale === 'de'
                    ? ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So']
                    : ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
                  const pickDay = (ce: typeof childEvents[0], isSel: boolean, disabled: boolean): void => {
                    if (disabled) return;
                    if (!isSel) {
                      // Hat der Termin eigene Abfragefelder, läuft die Auswahl
                      // über denselben Modal-Flow wie in der Listen-Ansicht.
                      if ((ce.eventSpecificFields || []).length > 0) {
                        setPendingSubEventModal({ subEventId: ce.id, draftValues: { ...(sessionFieldValues[ce.id] || {}) } });
                        return;
                      }
                      // v28.97: „Genau eines" — die neue Wahl ERSETZT die alte,
                      // statt sich danebenzulegen. Sonst müsste der Teilnehmer
                      // erst abwählen und würde bei jedem Wechsel scheitern.
                      const next = event.subEventSingleChoice ? new Set<string>() : new Set(selectedSessions);
                      next.add(ce.id);
                      setSelectedSessions(next);
                      if (event.subEventSingleChoice) {
                        setSessionFieldValues(prev => {
                          const keep: typeof prev = {};
                          if (prev[ce.id]) keep[ce.id] = prev[ce.id];
                          return keep;
                        });
                      }
                      return;
                    }
                    const next = new Set(selectedSessions);
                    next.delete(ce.id);
                    setSelectedSessions(next);
                    setSessionFieldValues(prev => { const c = { ...prev }; delete c[ce.id]; return c; });
                  };
                  return (
                    <div style={{ marginTop: 12 }}>
                      <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', fontWeight: 600, marginBottom: 8 }}>
                        {locale === 'de'
                          ? 'Termine auswählen — angebotene Tage sind hervorgehoben.'
                          : 'Pick your dates — offered days are highlighted.'}
                      </div>
                      {monthKeys.map(mk => {
                        const [my, mm] = mk.split('-').map(n => parseInt(n, 10));
                        const first = new Date(my, mm - 1, 1);
                        const daysInMonth = new Date(my, mm, 0).getDate();
                        // Montag als erster Wochentag (getDay: So=0).
                        const lead = (first.getDay() + 6) % 7;
                        const cells: Array<string | null> = [];
                        for (let i = 0; i < lead; i++) cells.push(null);
                        for (let d = 1; d <= daysInMonth; d++) {
                          cells.push(`${my}-${String(mm).padStart(2, '0')}-${String(d).padStart(2, '0')}`);
                        }
                        return (
                          <div key={mk} style={{ marginBottom: 14 }}>
                            <div style={{ fontWeight: 700, fontSize: '0.9rem', marginBottom: 6, color: 'var(--dex-gray-700, #444)' }}>
                              {first.toLocaleDateString(locale === 'de' ? 'de-DE' : 'en-GB', { month: 'long', year: 'numeric' })}
                            </div>
                            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4 }}>
                              {weekdays.map(w => (
                                <div key={w} style={{ fontSize: '0.68rem', fontWeight: 700, color: 'var(--dex-gray-400)', textAlign: 'center', padding: '2px 0' }}>{w}</div>
                              ))}
                              {cells.map((key, i) => {
                                if (!key) return <div key={`e${i}`} />;
                                const entry = byDay[key];
                                const dayNum = parseInt(key.slice(8), 10);
                                if (!entry) {
                                  return (
                                    <div key={key} style={{
                                      textAlign: 'center', padding: '8px 0', borderRadius: 8,
                                      fontSize: '0.8rem', color: 'var(--dex-gray-300, #ccc)',
                                    }}>{dayNum}</div>
                                  );
                                }
                                const ce = entry.ce;
                                const meta = sessionMeta[ce.id] || { count: 0, wasRegistered: false };
                                const isSel = selectedSessions.has(ce.id);
                                const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
                                const isFull = hasCap && meta.count >= (ce.maxParticipants || 0);
                                const deadlinePassed = !!(ce.registrationDeadline && new Date(ce.registrationDeadline) < new Date());
                                // v29.28: Frist-Bypass für Organizer/Admins (s. Listen-Pfad).
                                const deadlineLocked = deadlinePassed && !isOrganizer && !isAdmin;
                                const disabled = (isFull && !isSel) || (deadlineLocked && !isSel);
                                const free = hasCap ? Math.max(0, (ce.maxParticipants || 0) - meta.count) : -1;
                                const title = [
                                  ce.title || '',
                                  hasCap
                                    ? (locale === 'de' ? `${free} von ${ce.maxParticipants} Plätzen frei` : `${free} of ${ce.maxParticipants} seats free`)
                                    : (locale === 'de' ? 'Unbegrenzte Plätze' : 'Unlimited seats'),
                                  deadlinePassed
                                    ? (deadlineLocked
                                      ? (locale === 'de' ? 'Anmeldefrist abgelaufen' : 'Registration deadline passed')
                                      : (locale === 'de' ? 'Anmeldefrist abgelaufen — als Organizer/Admin trotzdem wählbar' : 'Registration deadline passed — still selectable as organizer/admin'))
                                    : '',
                                  ce.mandatoryRegistration ? (locale === 'de' ? 'Pflichttermin' : 'Mandatory date') : '',
                                ].filter(Boolean).join(' · ');
                                return (
                                  <button
                                    key={key}
                                    type="button"
                                    onClick={() => pickDay(ce, isSel, disabled)}
                                    disabled={disabled}
                                    title={title}
                                    aria-pressed={isSel}
                                    style={{
                                      display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
                                      gap: 1, padding: '6px 0 5px', borderRadius: 8, minHeight: 46,
                                      border: `1px solid ${isSel ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                      background: isSel ? 'var(--dex-green, #86bc25)' : '#fff',
                                      color: isSel ? '#fff' : (disabled ? 'var(--dex-gray-400)' : 'var(--dex-gray-800, #333)'),
                                      cursor: disabled ? 'not-allowed' : 'pointer',
                                      opacity: disabled ? 0.55 : 1,
                                      fontWeight: 700, fontSize: '0.82rem',
                                    }}
                                  >
                                    <span>{dayNum}</span>
                                    <span style={{ fontSize: '0.6rem', fontWeight: 600, opacity: 0.85 }}>
                                      {deadlineLocked
                                        ? (locale === 'de' ? 'zu' : 'closed')
                                        : isFull
                                        ? (locale === 'de' ? 'voll' : 'full')
                                        : hasCap
                                        ? (locale === 'de' ? `${free} frei` : `${free} free`)
                                        : '—'}
                                    </span>
                                  </button>
                                );
                              })}
                            </div>
                          </div>
                        );
                      })}
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                        {selectedSessions.size === 0
                          ? (locale === 'de' ? 'Noch kein Termin gewählt.' : 'No date picked yet.')
                          : (locale === 'de'
                            ? `${selectedSessions.size} ${selectedSessions.size === 1 ? 'Termin' : 'Termine'} gewählt.`
                            : `${selectedSessions.size} ${selectedSessions.size === 1 ? 'date' : 'dates'} picked.`)}
                      </div>
                    </div>
                  );
                })()}

                {/* Sessions */}
                {!event.subEventCalendar && (
                <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', fontWeight: 600 }}>{childTermPlural || tEvent('reg.selection.sessions') || 'Sessions'}</div>
                  {childEvents.map(ce => {
                    const meta = sessionMeta[ce.id] || { count: 0, wasRegistered: false };
                    const isSel = selectedSessions.has(ce.id);
                    const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
                    const isSessionFull = hasCap && meta.count >= (ce.maxParticipants || 0);
                    const deadlinePassed = !!(ce.registrationDeadline && new Date(ce.registrationDeadline) < new Date());
                    // v29.28: Organizer/Admins dürfen — wie beim Haupt-Event
                    // (parentRegBlocked) und wie der Wizard es ausdrücklich
                    // verspricht — auch nach der Frist anmelden. Die
                    // Kapazitäts-Sperre bleibt für alle.
                    const deadlineLocked = deadlinePassed && !isOrganizer && !isAdmin;
                    const disabled = (isSessionFull && !isSel) || (deadlineLocked && !isSel);

                    return (
                      <div key={ce.id} style={{
                        padding: 10, borderRadius: 8,
                        border: `1px solid ${isSel ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                        background: isSel ? 'rgba(134,188,37,0.06)' : '#fff',
                      }}>
                        <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: disabled ? 'not-allowed' : 'pointer' }}>
                          <input
                            type="checkbox"
                            checked={isSel}
                            disabled={disabled}
                            onChange={e => {
                              if (e.target.checked) {
                                // v29.27: direkt selektieren — Fragen inline in
                                // der Karte (s. Listen-Pfad oben).
                                // v28.97: siehe Kalender — bei „genau eines"
                                // ersetzt die neue Wahl die bisherige.
                                const next = event.subEventSingleChoice ? new Set<string>() : new Set(selectedSessions);
                                next.add(ce.id);
                                setSelectedSessions(next);
                              } else {
                                const next = new Set(selectedSessions);
                                next.delete(ce.id);
                                setSelectedSessions(next);
                                setSessionFieldValues(prev => {
                                  const copy = { ...prev };
                                  delete copy[ce.id];
                                  return copy;
                                });
                              }
                            }}
                            style={{ marginTop: 2 }}
                          />
                          <div style={{ flex: 1 }}>
                            <div style={{ fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                              {ce.title || tEvent('reg.subevents.untitled')}
                              {ce.mandatoryRegistration && (
                                <span style={{ fontSize: '0.68rem', fontWeight: 700, color: '#fff', background: 'var(--dex-orange, #ed8b00)', borderRadius: 999, padding: '2px 8px' }}>
                                  {locale === 'de' ? 'Pflicht' : 'Required'}
                                </span>
                              )}
                            </div>
                          </div>
                        </label>
                        {/* v29.28: Karteninhalt linksbündig (s. Listen-Pfad). */}
                        <div style={{ marginTop: 4 }}>
                            {ce.description && (
                              // v11.97: gleiche Schriftgröße wie der Titel
                              // (Standard-Body). Vorher 0.78rem klein.
                              // v29.27: als sanitisiertes HTML statt rohem Text
                              // (s. subEventDescHtml beim Listen-Pfad oben).
                              <div style={{ color: 'var(--dex-gray-600)', marginTop: 2, wordBreak: 'break-word' }} dangerouslySetInnerHTML={{ __html: subEventDescHtml(ce.description) }} />
                            )}
                            {/* v11.94: gleiches Icon-Layout wie oben (anderer
                                Render-Pfad für Team-Modus). */}
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, marginTop: 4, color: 'var(--dex-gray-600)' }}>
                              {ce.startDate && (
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                  {/* v11.97: Icon-Größe an Standard-Body angepasst. */}
                                  <Icon iconName="Calendar" style={{ fontSize: 15, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                                  {formatDate(ce.startDate)}
                                </span>
                              )}
                              {ce.location && (
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                  <Icon iconName="POI" style={{ fontSize: 15, color: '#0a3766' }} />
                                  {ce.location}
                                </span>
                              )}
                            </div>
                            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {hasCap && (() => {
                                const sessionFree = Math.max(0, (ce.maxParticipants || 0) - (meta.count || 0));
                                return (
                                  <> · <span style={{ color: isSessionFull ? 'var(--dex-red)' : 'inherit', fontWeight: 600 }}>
                                    {/* v19.19: belegt-Zahl bei der Kapazität deckeln (s.o.) —
                                        Überbuchung nie auf der Anmeldeseite anzeigen. */}
                                    {Math.min(meta.count, ce.maxParticipants || 0)}/{ce.maxParticipants} {tEvent('reg.subevents.taken')}
                                  </span>
                                  {!isSessionFull && (
                                    <span style={{ color: 'var(--dex-green-dark)' }}> — {sessionFree} {tEvent('reg.free')}</span>
                                  )}
                                  </>
                                );
                              })()}
                            </div>
                            {deadlinePassed && !isSel && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-orange)', marginTop: 2 }}>
                                {tEvent('reg.subevents.deadlinepassed')}
                              </div>
                            )}
                            {isSessionFull && !isSel && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-red)', marginTop: 2 }}>
                                {tEvent('reg.subevents.sessionfull')}
                              </div>
                            )}
                            {/* v11.10: Hardcoded Sub-Event-Gruppen-Radios entfernt.
                                Sub-Events erben jetzt grundsätzlich
                                preferredStarterType vom Group-Selection-Block
                                oben. Pro-Sub-Event-Gruppe ist konzeptionell
                                Quatsch — die Gruppe gehört zum Teilnehmer
                                (z.B. „Vormittag/Nachmittag"), nicht zur
                                Session. */}
                        </div>
                        {/* v29.27: Fragen inline in der Karte (s. Listen-Pfad). */}
                        {isSel && renderSubEventInlineFields(ce)}
                      </div>
                    );
                  })}
                </div>
                )}

                {isSessionsOnlyMode && selectedSessions.size > 0 && !event.subEventsOnlyMode && (
                  <div style={{
                    marginTop: 12, padding: '8px 10px', borderRadius: 6,
                    background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                    color: 'var(--dex-orange)', fontSize: '0.78rem',
                  }}>
                    {childTermPlural
                      ? (locale === 'de'
                          ? `Du meldest dich ausschließlich für ${childTermPlural} an — NICHT für das Haupt-Event.`
                          : `You are registering exclusively for ${childTermPlural} — NOT for the main event.`)
                      : (tEvent('reg.selection.sessionsonlyhint') || 'Du meldest dich ausschließlich für Sessions an — NICHT für das Haupt-Event.')}
                  </div>
                )}
              </div>
            )}
            {/* v29.28: Bei Events MIT Sub-Events rendert der Felder-Block
                oben in der Auswahl-Box direkt unter der Haupt-Event-Kachel
                (renderMainFieldsSection) — hier nur noch ohne Sub-Events
                oder bei einer Klammer (dort gelten die Fragen übergreifend
                und es gibt keine Haupt-Event-Kachel). */}
            {(childEvents.length === 0 || !!event.subEventsOnlyMode) && renderMainFieldsSection()}
          </div>
          </CollapsibleSection>
        </div>
        )}
      </div>

      {/* v11.4: Fehlermeldung + Action-Buttons stehen jetzt direkt unter
          dem registration-layout (also unter der Eventspez-Karte) und
          NICHT mehr unterhalb des Datenschutz-Hinweises. Der Hinweis ist
          eine Fußnote und gehört ans Seitenende — die Aktions-Buttons
          gehören thematisch zur Anmelde-Maske. */}

      {/* Fehlermeldung */}
      {error && (
        <div className="mt-16" style={{ padding: '12px 16px', background: 'rgba(239,68,68,0.1)', border: '1px solid var(--dex-red)', borderRadius: 'var(--dex-radius-md)', color: 'var(--dex-red)', fontSize: '0.9rem' }}>
          {error}
        </div>
      )}

      {/* v24.48: Die „Meine Assistenz"-Abfrage ist von inline auf ein Modal
          beim Register-Klick umgestellt (siehe assistantModalOpen unten). */}

      {/* Buttons + Platz-Badge (v24.57) — Badge über den zentrierten Buttons. */}
      <div style={{ maxWidth: 1100, margin: '24px auto 0', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
        {/* v24.57: Badge mit Icon — freie Plätze ODER (wenn voll + Warteliste
            aktiv) „Warteliste". Bei unbegrenzter Teilnehmerzahl gar nichts. */}
        {event.maxParticipants > 0 && (() => {
          // v24.72: Wartelisten-Anzahl von den freien Plätzen abziehen. Ein frei
          // gewordener Platz geht IMMER zuerst an die Warteliste — er ist also
          // nicht „frei" für neue Anmeldungen. Das verhindert auch das kurze,
          // fälschliche „1 freier Platz" während des Nachrückens.
          // v24.73: Aktiv-/Warteliste-Zahl bevorzugt aus dem für alle lesbaren
          // Counter (liveStats) — sonst sieht ein normaler Teilnehmer wegen der
          // Item-Level-Security der Teilnehmerliste keine korrekte Zahl.
          const effActive = liveStats ? liveStats.active : (event.currentParticipants || 0);
          const effWaitlist = (liveStats && liveStats.waitlist >= 0) ? liveStats.waitlist : (event.waitlistCount || 0);
          const free = Math.max(0, event.maxParticipants - effActive - effWaitlist);
          const isFull = free <= 0;
          // v27.11: voll + Warteliste deaktiviert → NICHT mehr stumm bleiben,
          // sondern rote Badge zeigen (vorher: return null; die Anmeldung lief
          // dann trotzdem still auf die abgeschaltete Warteliste).
          if (isFull && !event.waitlistEnabled) {
            return (
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 7, background: 'rgba(204,0,0,0.08)', color: 'var(--dex-red, #c00)', border: '1px solid var(--dex-red, #c00)', borderRadius: 999, padding: '5px 14px', fontSize: '0.82rem', fontWeight: 700 }}>
                <Icon iconName="People" style={{ fontSize: 15 }} />
                {locale === 'de' ? 'Alle Plätze belegt — keine Warteliste' : 'All seats taken — no waitlist'}
              </span>
            );
          }
          const waitlist = isFull && !!event.waitlistEnabled;
          const nearlyFull = !isFull && free <= Math.max(1, Math.round(event.maxParticipants * 0.1));
          const isTeamEvent = !!(event.teamRegistrationEnabled && event.teamSize && event.teamSize > 1);
          const teamsFree = isTeamEvent ? Math.floor(free / (event.teamSize || 1)) : 0;
          const orange = waitlist || nearlyFull;
          const bg = orange ? 'rgba(237,139,0,0.12)' : 'rgba(134,188,37,0.14)';
          const fg = orange ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)';
          const bd = orange ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)';
          return (
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 7, background: bg, color: fg, border: `1px solid ${bd}`, borderRadius: 999, padding: '5px 14px', fontSize: '0.82rem', fontWeight: 700 }}>
              <Icon iconName={waitlist ? 'Clock' : 'People'} style={{ fontSize: 15 }} />
              {waitlist
                ? (() => {
                    const wc = effWaitlist;
                    return locale === 'de'
                      ? `Alle Plätze belegt | Warteliste aktuell ${wc} ${wc === 1 ? 'Person' : 'Personen'}`
                      : `All places taken | Waitlist currently ${wc} ${wc === 1 ? 'person' : 'people'}`;
                  })()
                : (
                  <span>
                    {free} / {event.maxParticipants} {locale === 'de' ? 'freie Plätze' : 'available'}
                    {isTeamEvent && <span style={{ fontWeight: 500, marginLeft: 6 }}>({teamsFree} {teamsFree === 1 ? (locale === 'de' ? 'Team' : 'team') : (locale === 'de' ? 'Teams' : 'teams')} {locale === 'de' ? 'frei' : 'available'})</span>}
                  </span>
                )}
            </span>
          );
        })()}
        <div className="registration-actions" style={{ alignItems: 'center' }}>
        {(() => {
          // v15.11: im subEventsOnlyMode (Hauptevent nicht anmeldbar) muss
          // mindestens ein Sub-Event ausgewählt sein, sonst Button ausgrauen
          // + Hinweis statt „Registrieren (Haupt-Event)" zeigen.
          const isSubOnly = !!(event && event.subEventsOnlyMode) && !registerForOther;
          const nothingPicked = isSubOnly && selectedSessions.size === 0;
          // v15.16: Consent-Pflicht bei „Für andere registrieren".
          const needsOtherConsent = registerForOther && !!email.trim() && !otherConsentConfirmed;
          // v19.8: Bei stellvertretender Anmeldung den Button sperren, wenn die
          // ausgewählte Person bereits angemeldet ist — vorher konnte man
          // trotz Hinweis auf „Registrieren" klicken (und es kam danach noch
          // die CC-Frage). Jetzt klare Blockade direkt am Button.
          const targetAlreadyRegistered = registerForOther && !!(thirdPartyCheck && thirdPartyCheck.alreadyRegistered);
          // v18: Demo-Event — Register-Button ist bewusst NICHT auswählbar
          // (keine echte Anmeldung; reine Showcase-Ansicht).
          const isDemo = !!(event && event.isDemoShowcase);
          // v28.88: Bereits angemeldet und nichts (mehr) auszuwählen → der
          // Klick konnte ohnehin nichts bewirken und endete in einer
          // Fehlermeldung. Jetzt sagt der Button selbst, dass die Anmeldung
          // schon steht. (nothingToSubmit deckt auch Abwahl-Änderungen ab —
          // wer Sub-Events abmeldet, kommt weiterhin durch.)
          const alreadyDone = parentAlreadyRegistered && nothingToSubmit;
          const isDisabled = isDemo || isSubmitting || (isTeamMode && !teamValidation.ok) || nothingPicked || needsOtherConsent || targetAlreadyRegistered || alreadyDone;
          const titleAttr = isDemo
            ? (locale === 'de' ? 'Demo-Event — eine echte Anmeldung ist nicht möglich.' : 'Demo event — real registration is not possible.')
            : (alreadyDone
            ? (locale === 'de' ? 'Du bist für dieses Event bereits angemeldet. Abmelden kannst du dich über „Meine Events".' : 'You are already registered for this event. You can cancel via „My events".')
            : (targetAlreadyRegistered
            ? (locale === 'de' ? 'Diese Person ist bereits für das Event angemeldet.' : 'This person is already registered for this event.')
            : (isTeamMode && !teamValidation.ok
            ? (teamValidation.reason || '')
            : (nothingPicked
                ? (locale === 'de'
                    ? `Bitte mindestens ${childOneDe} auswählen.`
                    : `Please pick at least one ${childTermSingular || 'sub-event'}.`)
                : (needsOtherConsent
                    ? (locale === 'de'
                        ? 'Bitte bestätige die Zustimmung der Person.'
                        : 'Please confirm the person\'s consent.')
                    : '')))));
          return (
            <button
              className="btn btn-primary"
              onClick={handleSubmit}
              disabled={isDisabled}
              title={titleAttr}
            >
              {/* v24.94: Label in EINEN Span wickeln. Sonst werden „Register" und
                  das „(Warteliste)"-Suffix-Span zu separaten Flex-Items des
                  Buttons (.btn ist inline-flex mit gap:8px) → der Flex-Gap PLUS
                  das Leerzeichen ergaben einen doppelten Abstand. */}
              <Send size={16} /> <span>{(() => {
                if (isSubmitting) return t('reg.submitting');
                // v28.88: Bestehende Anmeldung, nichts zu ändern — der Button
                // sagt das jetzt selbst, statt „Registrieren" anzubieten und
                // beim Klick zu meckern.
                if (alreadyDone) return locale === 'de' ? 'Bereits angemeldet' : 'Already registered';
                // v24.62: Wenn das Hauptevent voll ist und eine Warteliste hat,
                // landet die Anmeldung auf der Warteliste — im Button steht das als
                // kurzer, NICHT fetter Zusatz „(Warteliste)" (die aktuelle Anzahl
                // steht im Badge über dem Button).
                const mfActive = liveStats ? liveStats.active : (event.currentParticipants || 0);
                const mfWaitlist = (liveStats && liveStats.waitlist >= 0) ? liveStats.waitlist : (event.waitlistCount || 0);
                const mainFull = event.maxParticipants > 0
                  && Math.max(0, event.maxParticipants - mfActive - mfWaitlist) <= 0
                  && !!event.waitlistEnabled;
                const waitlistSuffixNode: React.ReactNode = mainFull
                  ? <span style={{ fontWeight: 400 }}> ({locale === 'de' ? 'Warteliste' : 'waitlist'})</span>
                  : null;
                // v18.73: Vorgemerkter Team-Beitritt — eigener Button-Text.
                if (pendingJoinTeam) {
                  return event?.teamJoinRequiresApproval
                    ? (locale === 'de' ? 'Beitritt anfragen' : 'Request to join')
                    : (locale === 'de' ? 'Team beitreten & anmelden' : 'Join team & register');
                }
                // v11.82: Team-Modus — eigener Button-Text mit Personen-Zahl.
                if (isTeamMode) {
                  const n = 1 + teamMembersParsed.filter(Boolean).length;
                  return locale === 'de'
                    ? `Team anmelden (${n} ${n === 1 ? 'Person' : 'Personen'})`
                    : `Register team (${n} ${n === 1 ? 'person' : 'people'})`;
                }
                if (nothingPicked) {
                  return locale === 'de'
                    ? `Bitte mindestens ${childOneDe} auswählen`
                    : `Please pick at least one ${childTermSingular || 'sub-event'}`;
                }
                if (registerForOther) return <>{t('reg.register')}{waitlistSuffixNode}</>;
                // v7.3: Kein Selection-Block → einfacher "Registrieren"-Text ohne
                // Parantheses-Info. Erst wenn Sub-Events existieren, zeigen wir
                // detailliert an, was gerade submittet wird.
                if (childEvents.length === 0) return <>{t('reg.register')}{waitlistSuffixNode}</>;
                const parts: string[] = [];
                if (willRegisterParent) parts.push(resolveMainEventLabel(t('reg.selection.mainevent') || 'Haupt-Event') || event.title);
                if (selectedSessions.size > 0) {
                  parts.push(`${selectedSessions.size} ${selectedSessions.size === 1 ? (childTermSingular || t('reg.selection.sessioncount.one') || 'Session') : (childTermPlural || t('reg.selection.sessioncount.many') || 'Sessions')}`);
                }
                if (parts.length === 0) return <>{t('reg.register')}{waitlistSuffixNode}</>;
                // Bei gleichzeitiger Hauptevent-Anmeldung den Warteliste-Hinweis anhängen.
                return <>{t('reg.register')} ({parts.join(' + ')}){willRegisterParent ? waitlistSuffixNode : null}</>;
              })()}</span>
            </button>
          );
        })()}
        {/* v18.11: „Ich nehme nicht teil" — proaktive Absage. Nur bei
            Selbst-Anmeldung (nicht „für andere", nicht Team-Modus, kein
            Demo-Event). Braucht keine Pflichtfelder. */}
        {!registerForOther && !isTeamMode && !pendingJoinTeam && !(event && event.isDemoShowcase) && (
          <button
            type="button"
            className="btn btn-secondary"
            onClick={handleDecline}
            disabled={isDeclining || isSubmitting}
            title={locale === 'de'
              ? (childEvents.length > 0
                ? `Melde zurück, dass du nicht teilnehmen wirst — gilt für das gesamte Event inklusive aller ${childTermPlural || 'Sub-Events'}. Eine Auswahl ist dafür nicht nötig.`
                : 'Melde zurück, dass du nicht teilnehmen wirst (keine Anmeldung).')
              : (childEvents.length > 0
                ? 'Let us know you will not attend — applies to the whole event including all sub-events. No selection needed.'
                : 'Let us know you will not attend (no registration).')}
            style={{ color: 'var(--dex-gray-700, #444)' }}
          >
            <X size={16} /> {isDeclining
              ? (locale === 'de' ? 'Wird gesendet…' : 'Submitting…')
              : (locale === 'de' ? 'Ich nehme nicht teil' : 'I will not attend')}
          </button>
        )}
        </div>
      </div>

      {/* Datenschutz-Hinweis als Fußnote ganz unten.
          v11.93: Breite auf 1100px begrenzt + zentriert (analog
          .registration-layout), damit der Text nicht über die ganze
          App-Breite läuft. */}
      <div
        className="footer-disclaimer mt-24"
        style={{ borderRadius: 'var(--dex-radius-lg)', maxWidth: 1100, margin: '24px auto 0' }}
      >
        <p>
          {/* Datenverarbeitungs-Einwilligung. „{link}" wird als Anchor auf die
              Deloitte-Datenschutzhinweise gerendert; der Rest ist reiner Text
              mit {title}-Ersetzung. */}
          {(() => {
            const raw = t('reg.privacy.data').replace('{title}', event.title);
            const parts = raw.split('{link}');
            const linkLabel = t('reg.privacy.data.link');
            return (
              <>
                {parts[0]}
                <a
                  href="https://www.deloitte.com/de/de/legal/privacy.html"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  {linkLabel}
                </a>
                {parts[1] || ''}
              </>
            );
          })()}
        </p>
        <p>
          {/* Bild-/Video-Einwilligung. privacy@deloitte.de wird als
              mailto-Link gerendert (gleiche Adresse in DE und EN). */}
          {(() => {
            const raw = t('reg.privacy').replace('{title}', event.title);
            const mail = 'privacy@deloitte.de';
            const parts = raw.split(mail);
            return (
              <>
                {parts[0]}
                {parts.length > 1 && (
                  <>
                    <a href={`mailto:${mail}`}>{mail}</a>
                    {parts.slice(1).join(mail)}
                  </>
                )}
              </>
            );
          })()}
        </p>
      </div>

      {/* Fallback-Dialog (seit v6.5): Wunsch-Starter-Typ voll, aber Alternative frei.
          User entscheidet explizit zwischen Umsteigen oder Warteliste. */}
      {fallbackDialog && (
        <Modal
          open={true}
          onClose={() => setFallbackDialog(null)}
          maxWidth={480}
          padding={24}
          ariaLabel="Plätze voll"
        >
            {(() => {
              // v10.20: Label-Mapping für die freie Bezeichnung — wunsch/alt
              // sind interne IDs ('Durchstarter' / 'Funstarter'); die Anzeige
              // nimmt splitLabelA / splitLabelB.
              const wunschLabel = fallbackDialog.wunsch === 'Durchstarter' ? splitLabelA : splitLabelB;
              const altLabel = fallbackDialog.alt === 'Durchstarter' ? splitLabelA : splitLabelB;
              return (
                <>
                  {/* v17.22: Attendee-facing → bilingual. Vorher war dieser
                      Fallback-Dialog (Wunsch-Gruppe voll) rein deutsch. */}
                  <h3 style={{ margin: 0, marginBottom: 10 }}>
                    {locale === 'de' ? `${wunschLabel}-Plätze sind voll` : `${wunschLabel} is full`}
                  </h3>
                  <p style={{ color: 'var(--dex-gray-700)', lineHeight: 1.5, marginBottom: 8 }}>
                    {locale === 'de'
                      ? <>Für <strong>{wunschLabel}</strong> gibt es aktuell keine freien Plätze mehr.</>
                      : <>There are currently no free spots left for <strong>{wunschLabel}</strong>.</>}
                  </p>
                  <p style={{ color: 'var(--dex-gray-700)', lineHeight: 1.5, marginBottom: 20 }}>
                    {locale === 'de'
                      ? <>Es sind allerdings noch <strong>{fallbackDialog.altFree}</strong> Plätze als <strong>{altLabel}</strong> frei. Möchtest du stattdessen als <strong>{altLabel}</strong> starten?</>
                      : <>However, there are still <strong>{fallbackDialog.altFree}</strong> spots available as <strong>{altLabel}</strong>. Would you like to join as <strong>{altLabel}</strong> instead?</>}
                  </p>
                  <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                    <button
                      className="btn btn-secondary"
                      style={{ fontSize: '0.9rem' }}
                      onClick={async () => {
                        const wunsch = fallbackDialog.wunsch;
                        setFallbackDialog(null);
                        // Wunsch beibehalten → landet auf Warteliste für den Wunsch-Typ.
                        await performRegistration(wunsch);
                      }}
                    >
                      {locale === 'de' ? `Auf ${wunschLabel}-Warteliste` : `Join ${wunschLabel} waitlist`}
                    </button>
                    <button
                      className="btn btn-primary"
                      style={{ fontSize: '0.9rem' }}
                      onClick={async () => {
                        const alt = fallbackDialog.alt;
                        setFallbackDialog(null);
                        // Preferred auf den Alt-Typ setzen, damit sowohl Anzeige
                        // als auch das Register-Payload den neuen Wunsch nutzen.
                        setPreferredStarterType(alt);
                        await performRegistration(alt);
                      }}
                    >
                      {locale === 'de' ? `Als ${altLabel} starten` : `Join as ${altLabel}`}
                    </button>
                  </div>
                </>
              );
            })()}
        </Modal>
      )}

      {/* v9.22: Modal für externe Email-Anmeldung */}
      {/* v18.13: Massenimport-Modal. */}
      {/* v26.76: Geführter Wizard für die stellvertretende Anmeldung (interner
          Fall): Schritt 1 Person suchen (mit Foto), Schritt 2 Zustimmung. Nach
          „OK" ist die Person übernommen und die persönlichen Felder vorbefüllt.
          Externe Person / Massenimport bleiben als eigene Wege erhalten. */}
      {proxyStep > 0 && (() => {
        const cancelWizard = (): void => {
          setRegisterForOther(false);
          setProxyStep(0);
          setFirstName(currentUser.firstName); setSurname(currentUser.surname); setEmail(currentUser.email);
          setUserSearch(''); setUserResults([]); setPickedUserProfile(null);
          setThirdPartyCheck(null); setOtherConsentConfirmed(false); setExternalPerson(false);
        };
        const clearPick = (): void => {
          setFirstName(''); setSurname(''); setEmail(''); setUserSearch(''); setUserResults([]);
          setThirdPartyCheck(null); setPickedUserProfile(null);
        };
        const linkBtn: React.CSSProperties = { background: 'none', border: 'none', padding: 0, color: 'var(--dex-blue, #0076a8)', textDecoration: 'underline', cursor: 'pointer', fontSize: '0.78rem', fontWeight: 600 };
        const picked = !!email.trim();
        const blocked = !!(thirdPartyCheck && thirdPartyCheck.alreadyRegistered);
        const pName = `${firstName} ${surname}`.trim() || email;
        return (
          <Modal
            open={proxyStep > 0}
            onClose={cancelWizard}
            maxWidth={560}
            padding={24}
            ariaLabel={locale === 'de' ? 'Für eine andere Person anmelden' : 'Register another person'}
          >
            <h3 style={{ margin: '0 0 2px', fontSize: '1.1rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {locale === 'de' ? 'Für eine andere Person anmelden' : 'Register another person'}
            </h3>
            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginBottom: 14 }}>
              {locale === 'de' ? `Schritt ${proxyStep} von 2 — ${proxyStep === 1 ? 'Person suchen' : 'Zustimmung'}` : `Step ${proxyStep} of 2 — ${proxyStep === 1 ? 'find person' : 'consent'}`}
            </div>

            {proxyStep === 1 && (
              <>
                {!externalPerson && !picked && (
                  <div style={{ position: 'relative' }}>
                    <input
                      className="form-input"
                      autoFocus
                      value={userSearch}
                      onChange={e => {
                        const val = e.target.value;
                        setUserSearch(val);
                        if (searchTimerRef.current) clearTimeout(searchTimerRef.current);
                        if (val.length >= 2) {
                          searchTimerRef.current = setTimeout(async () => {
                            setIsSearchingUser(true);
                            const results = await searchUsers(val, userSearchIncludeIntl);
                            setUserResults(results);
                            setIsSearchingUser(false);
                          }, 300);
                        } else { setUserResults([]); }
                      }}
                      placeholder={t('reg.searchplaceholder') || 'Name oder E-Mail eingeben...'}
                    />
                    <InternationalSearchToggle
                      query={userSearch}
                      checked={userSearchIncludeIntl}
                      onChange={async next => {
                        setUserSearchIncludeIntl(next);
                        const val = userSearch.trim();
                        if (val.length >= 2) { setIsSearchingUser(true); try { setUserResults(await searchUsers(val, next)); } catch { /* */ } setIsSearchingUser(false); }
                      }}
                    />
                    {isSearchingUser && <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 8 }}>{locale === 'de' ? 'Wird gesucht…' : 'Searching…'}</p>}
                    {userResults.length > 0 && (
                      <div style={{ maxHeight: 260, overflowY: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 8, marginTop: 8 }}>
                        {userResults.map(u => {
                          const assistantOnly = isAssistant && !canCreateEvents;
                          const targetAllowed = !assistantOnly || isAllowedTargetForAssistant(u.jobTitle);
                          return (
                            <div
                              key={u.email}
                              onClick={() => { if (targetAllowed) pickProxyUser(u); }}
                              title={targetAllowed ? '' : 'Assistants can only register Partners or Directors for events.'}
                              style={{ padding: '8px 12px', cursor: targetAllowed ? 'pointer' : 'not-allowed', opacity: targetAllowed ? 1 : 0.45, borderBottom: '1px solid var(--dex-gray-100)', display: 'flex', alignItems: 'center', gap: 10 }}
                            >
                              <img src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(u.email)}&size=S`} alt={u.displayName} onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }} style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }} />
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontWeight: 600, fontSize: '0.85rem' }}>{u.displayName}</div>
                                <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.email}{u.jobTitle ? ` · ${u.jobTitle}` : ''}{u.location ? ` · ${u.location}` : ''}</div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                    {canCreateEvents && (
                      <div style={{ marginTop: 14, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                        {locale === 'de' ? 'Person außerhalb Deloitte oder mehrere auf einmal? ' : 'External person or several at once? '}
                        <button type="button" style={linkBtn} onClick={() => { setExternalPerson(true); clearPick(); setOtherConsentConfirmed(false); }}>{locale === 'de' ? 'Externe Person' : 'External person'}</button>
                        {' · '}
                        <button type="button" style={linkBtn} onClick={() => { setProxyStep(0); setMassImportResult(null); setMassImportRows([]); setMassImportStep('input'); setMassImportOpen(true); }}>{locale === 'de' ? 'Massenimport' : 'Bulk import'}</button>
                      </div>
                    )}
                  </div>
                )}
                {!externalPerson && picked && (
                  <>
                    <div style={{ padding: '10px 12px', border: '1px solid var(--dex-green, #86bc25)', borderRadius: 8, background: 'rgba(134,188,37,0.06)', display: 'flex', alignItems: 'center', gap: 10 }}>
                      <img src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=S`} alt={pName} onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }} style={{ width: 40, height: 40, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }} />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 700 }}>{pName}</div>
                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email}{pickedUserProfile?.jobTitle ? ` · ${pickedUserProfile.jobTitle}` : ''}</div>
                      </div>
                      <button type="button" style={linkBtn} onClick={clearPick}>{locale === 'de' ? 'Ändern' : 'Change'}</button>
                    </div>
                    {thirdPartyCheck && (thirdPartyCheck.alreadyRegistered || thirdPartyCheck.notInAudience) && (
                      <div style={{ marginTop: 8, padding: '8px 12px', borderRadius: 8, fontSize: '0.8rem', background: thirdPartyCheck.alreadyRegistered ? 'rgba(200,30,30,0.07)' : 'rgba(237,139,0,0.08)', border: `1px solid ${thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)'}`, color: thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)' }}>
                        {thirdPartyCheck.alreadyRegistered
                          ? (locale === 'de' ? 'Diese Person ist bereits für dieses Event angemeldet.' : 'This person is already registered for this event.')
                          : (locale === 'de' ? 'Hinweis: Diese Person ist nicht im Gästekreis dieses Events — die Anmeldung ist trotzdem möglich.' : 'Note: this person is not in this event’s audience — registration is still possible.')}
                      </div>
                    )}
                  </>
                )}
                {/* v26.85: Externe Person direkt IM Wizard erfassen (statt unten
                    im Formular). Vor-/Nachname + E-Mail hier eingeben, „Weiter"
                    führt zur Zustimmung. */}
                {externalPerson && (
                  <div>
                    <div style={{ padding: '10px 12px', marginBottom: 12, borderRadius: 8, background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange, #ed8b00)', fontSize: '0.82rem', color: 'var(--dex-orange-dark, #b35a00)', lineHeight: 1.5 }}>
                      {/* v27.12: Wording-Feinschliff (Feedback Datenschutz-Review). */}
                      {locale === 'de'
                        ? 'Person außerhalb von Deloitte (externe E-Mail-Adresse). Trage Vorname, Nachname und E-Mail-Adresse ein. Nach der Zustimmung meldest du die Person stellvertretend an — die Einladung und die Datenschutz-Rückmeldung laufen anschließend über dich.'
                        : 'Person outside Deloitte (external email address). Enter first name, last name and email address. After consent you register the person on their behalf — the invitation and the privacy confirmation are then handled through you.'}
                    </div>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 10 }}>
                      <div>
                        <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, marginBottom: 3 }}>{t('reg.firstname') || 'Vorname'}</label>
                        <input className="form-input" autoFocus value={firstName} onChange={e => setFirstName(e.target.value)} placeholder={t('reg.firstname') || 'Vorname'} />
                      </div>
                      <div>
                        <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, marginBottom: 3 }}>{t('reg.surname') || 'Nachname'}</label>
                        <input className="form-input" value={surname} onChange={e => setSurname(e.target.value)} placeholder={t('reg.surname') || 'Nachname'} />
                      </div>
                    </div>
                    <div style={{ marginBottom: 6 }}>
                      <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, marginBottom: 3 }}>E-Mail</label>
                      <input className="form-input" type="email" value={email} onChange={e => { setEmail(e.target.value); externalEmailConfirmedRef.current = false; setThirdPartyCheck(null); /* v27.11: Duplikat-Check bei Adress-Änderung zurücksetzen */ }} placeholder="name@firma.de" />
                    </div>
                    <button type="button" style={linkBtn} onClick={() => { setExternalPerson(false); clearPick(); }}>{locale === 'de' ? '← Zurück zur Personensuche' : '← Back to search'}</button>
                  </div>
                )}
                <div style={{ display: 'flex', justifyContent: 'space-between', gap: 10, marginTop: 18 }}>
                  <button type="button" className="btn btn-secondary" onClick={cancelWizard}>{locale === 'de' ? 'Abbrechen' : 'Cancel'}</button>
                  {externalPerson ? (
                    <button
                      type="button"
                      className="btn btn-primary"
                      disabled={!(firstName.trim() && surname.trim() && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email.trim()))}
                      onClick={() => {
                        // v27.11 (Bug „Externe können mehrfach angemeldet
                        // werden"): Duplikat-Check jetzt auch für externe
                        // Personen — vorher lief er NUR beim Personen-Picker
                        // (interne), Externe rutschten ungeprüft durch.
                        // thirdPartyCheck aktiviert zugleich die bestehende
                        // Submit-Sperre + den Button-Disable am Formular.
                        (async () => {
                          const existing = await checkRegistrationByEmail(event.id, email.trim()).catch(() => null);
                          const alreadyRegistered = !!existing && existing.Status !== 'Abgemeldet';
                          setThirdPartyCheck({
                            alreadyRegistered,
                            notInAudience: false,
                            registeredName: (existing && (existing.ParticipantName || `${existing.Vorname || ''} ${existing.Nachname || ''}`.trim())) || `${firstName} ${surname}`.trim(),
                            registeredDate: (existing && existing.RegistrationDate) || '',
                          });
                          if (alreadyRegistered) {
                            showAlert(locale === 'de'
                              ? `${email.trim()} ist bereits für dieses Event angemeldet — eine erneute Anmeldung ist nicht möglich.`
                              : `${email.trim()} is already registered for this event — registering again is not possible.`, { variant: 'error' });
                            return;
                          }
                          setProxyStep(2);
                        })().catch(() => setProxyStep(2));
                      }}
                    >{locale === 'de' ? 'Weiter' : 'Next'}</button>
                  ) : (
                    <button type="button" className="btn btn-primary" disabled={!picked || blocked} onClick={() => setProxyStep(2)}>{locale === 'de' ? 'Weiter' : 'Next'}</button>
                  )}
                </div>
              </>
            )}

            {proxyStep === 2 && (
              <>
                {/* v26.98: Die ausführliche Ablauf-Erklärung lebt jetzt HIER im
                    Wizard-Schritt „Zustimmung" (statt als große Box auf der
                    Anmeldeseite). Auf der Anmeldeseite bleibt danach nur ein
                    kurzer Hinweis + der Pflicht-Haken. */}
                <div style={{ padding: '12px 14px', background: 'rgba(237,139,0,0.10)', border: '2px solid var(--dex-orange, #ed8b00)', borderRadius: 8, color: '#7a4a00', fontSize: '0.86rem', lineHeight: 1.55 }}>
                  {locale === 'de'
                    ? <>Mit dem Absenden meldest du <strong>{pName}</strong> stellvertretend an. Bitte stelle sicher, dass die Person ihrer Anmeldung <strong>vorher zugestimmt</strong> hat — eine Anmeldung ohne Einverständnis ist nicht erlaubt.</>
                    : <>By submitting you register <strong>{pName}</strong> on their behalf. Please make sure the person has <strong>consented up front</strong> — registering people without their consent is not allowed.</>}
                  <div style={{ marginTop: 8 }}>
                    {/* v27.12: Wording-Feinschliff (Feedback Datenschutz-Review). */}
                    {locale === 'de'
                      ? <>Die Person erscheint anschließend regulär in der Teilnehmerliste. Falls sie doch nicht teilnehmen kann, lässt sich die Anmeldung jederzeit stornieren — bitte gib in dem Fall kurz Bescheid, damit Wartelisten-Plätze nachrücken können.</>
                      : <>The person then appears in the participant list as usual. If they are unable to attend after all, the registration can be cancelled at any time — please let us know in that case so waitlist spots can be filled.</>}
                  </div>
                </div>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginTop: 12, cursor: 'pointer' }}>
                  <input type="checkbox" checked={otherConsentConfirmed} onChange={e => setOtherConsentConfirmed(e.target.checked)} style={{ marginTop: 3 }} />
                  <span style={{ flex: 1, color: 'var(--dex-gray-800)' }}>
                    <span style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                    {locale === 'de'
                      ? 'Ich bestätige, dass die Person ihrer stellvertretenden Anmeldung zugestimmt hat.'
                      : 'I confirm that the person has consented to this registration on their behalf.'}
                  </span>
                </label>
                <div style={{ display: 'flex', justifyContent: 'space-between', gap: 10, marginTop: 18 }}>
                  <button type="button" className="btn btn-secondary" onClick={() => setProxyStep(1)}>{locale === 'de' ? 'Zurück' : 'Back'}</button>
                  <button type="button" className="btn btn-primary" disabled={!otherConsentConfirmed} onClick={() => setProxyStep(0)}>{locale === 'de' ? 'OK, Person übernehmen' : 'OK, take over person'}</button>
                </div>
              </>
            )}
          </Modal>
        );
      })()}
      {massImportOpen && (
        <Modal
          open={massImportOpen}
          onClose={() => { if (!massImportBusy) setMassImportOpen(false); }}
          maxWidth={600}
          padding={24}
          dismissable={!massImportBusy}
          ariaLabel={locale === 'de' ? 'Teilnehmer-Massenimport' : 'Bulk participant import'}
        >
          <h3 style={{ margin: '0 0 8px', fontSize: '1.1rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
            {locale === 'de' ? 'Teilnehmer-Massenimport' : 'Bulk participant import'}
          </h3>

          {massImportStep === 'input' && (
            <>
              <p style={{ margin: '0 0 12px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {locale === 'de'
                  ? <>Namen und/oder E-Mail-Adressen einfügen — <strong>eine Person pro Zeile</strong>. Das Tool gleicht jede Zeile mit dem Deloitte-Verzeichnis ab und zeigt dir danach eine <strong>Vorschau-Tabelle</strong> (Vorname, Nachname, Position, Standort, E-Mail) zum Prüfen, bevor angemeldet wird.</>
                  : <>Paste names and/or email addresses — <strong>one person per line</strong>. The tool matches each line against the Deloitte directory and then shows a <strong>preview table</strong> (first name, last name, position, location, email) to review before registering.</>}
              </p>
              <textarea
                className="form-input"
                value={massImportText}
                onChange={e => setMassImportText(e.target.value)}
                disabled={massImportResolving}
                rows={8}
                placeholder={'Mustermann, Max\nerika.musterfrau@deloitte.de\nMax Mustermann; max.mustermann@deloitte.de'}
                style={{ width: '100%', fontFamily: 'monospace', fontSize: '0.82rem', resize: 'vertical' }}
              />
              {massImportResolving && (
                <div style={{ marginTop: 10, fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                  {locale === 'de' ? 'Verzeichnis-Abgleich läuft…' : 'Matching against the directory…'} {massImportProgress}
                </div>
              )}
              <div style={{ marginTop: 18, display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
                <button className="btn btn-secondary" onClick={() => setMassImportOpen(false)} disabled={massImportResolving}>
                  {locale === 'de' ? 'Abbrechen' : 'Cancel'}
                </button>
                <button className="btn btn-primary" onClick={resolveMassImport} disabled={massImportResolving || !massImportText.trim()}>
                  {massImportResolving ? (locale === 'de' ? 'Wird abgeglichen…' : 'Matching…') : (locale === 'de' ? 'Abgleichen & Vorschau' : 'Match & preview')}
                </button>
              </div>
            </>
          )}

          {massImportStep === 'preview' && (() => {
            const okCount = massImportRows.filter(r => r.status === 'ok').length;
            const dupCount = massImportRows.filter(r => r.status === 'duplicate').length;
            const nfCount = massImportRows.filter(r => r.status === 'notfound').length;
            const removeRow = (idx: number): void => setMassImportRows(prev => prev.filter((_, i) => i !== idx));
            const tdStyle: React.CSSProperties = { padding: '6px 8px', fontSize: '0.8rem', borderBottom: '1px solid var(--dex-gray-100)' };
            const thStyle: React.CSSProperties = { padding: '6px 8px', fontSize: '0.75rem', textAlign: 'left', color: 'var(--dex-gray-500)', borderBottom: '2px solid var(--dex-gray-200)', textTransform: 'uppercase', letterSpacing: 0.4 };
            return (
              <>
                <p style={{ margin: '0 0 10px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {locale === 'de'
                    ? <><strong>{okCount}</strong> bereit zum Anmelden{dupCount > 0 ? `, ${dupCount} Duplikat(e)` : ''}{nfCount > 0 ? `, ${nfCount} nicht gefunden` : ''}. Prüfe die Tabelle — nicht passende Zeilen kannst du entfernen.</>
                    : <><strong>{okCount}</strong> ready to register{dupCount > 0 ? `, ${dupCount} duplicate(s)` : ''}{nfCount > 0 ? `, ${nfCount} not found` : ''}. Review the table — remove rows that don&apos;t fit.</>}
                </p>
                <div style={{ maxHeight: 320, overflowY: 'auto', overflowX: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 8 }}>
                  <table style={{ width: '100%', minWidth: 560, borderCollapse: 'collapse' }}>
                    <thead>
                      <tr>
                        <th style={thStyle}>{locale === 'de' ? 'Vorname' : 'First name'}</th>
                        <th style={thStyle}>{locale === 'de' ? 'Nachname' : 'Last name'}</th>
                        <th style={thStyle}>{locale === 'de' ? 'Position' : 'Position'}</th>
                        <th style={thStyle}>{locale === 'de' ? 'Standort' : 'Location'}</th>
                        <th style={thStyle}>E-Mail</th>
                        <th style={thStyle}>{locale === 'de' ? 'Status' : 'Status'}</th>
                        <th style={thStyle} />
                      </tr>
                    </thead>
                    <tbody>
                      {massImportRows.map((r, idx) => (
                        <tr key={`${r.email || r.raw}-${idx}`} style={{ opacity: r.status === 'ok' ? 1 : 0.6 }}>
                          <td style={tdStyle}>{r.firstName || '–'}</td>
                          <td style={tdStyle}>{r.lastName || '–'}</td>
                          <td style={{ ...tdStyle, color: 'var(--dex-gray-600)' }}>{r.jobTitle || '–'}</td>
                          <td style={{ ...tdStyle, color: 'var(--dex-gray-600)' }}>{r.location || '–'}</td>
                          <td style={{ ...tdStyle, color: 'var(--dex-gray-600)' }}>{r.email || <span style={{ color: 'var(--dex-red, #c00)' }}>{r.raw}</span>}</td>
                          <td style={tdStyle}>
                            {r.status === 'ok' && <span style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-green-dark, #4a7c1f)' }}>{locale === 'de' ? 'OK' : 'OK'}</span>}
                            {r.status === 'duplicate' && <span style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-orange-dark, #b35a00)' }}>{locale === 'de' ? 'Duplikat' : 'Duplicate'}</span>}
                            {r.status === 'notfound' && <span style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-red, #c00)' }}>{locale === 'de' ? 'Nicht gefunden' : 'Not found'}</span>}
                          </td>
                          <td style={tdStyle}>
                            <button type="button" onClick={() => removeRow(idx)} disabled={massImportBusy} title={locale === 'de' ? 'Zeile entfernen' : 'Remove row'} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-400)', padding: 2 }}>
                              <X size={14} />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div style={{ marginTop: 14, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
                    {locale === 'de' ? 'Benachrichtigung' : 'Notification'}
                  </div>
                  {([
                    { v: 'mail', de: 'Bestätigungsmail senden (+ Outlook-Termin)', en: 'Send confirmation email (+ Outlook invite)' },
                    { v: 'nomail', de: 'Ohne Bestätigungsmail (aber Outlook-Termin)', en: 'No confirmation email (but Outlook invite)' },
                    { v: 'silent', de: 'Stille Anmeldung (keine Mail, kein Kalendereintrag)', en: 'Silent registration (no email, no calendar invite)' },
                  ] as const).map(opt => (
                    <label key={opt.v} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                      <input type="radio" name="massImportMode" checked={massImportMode === opt.v} onChange={() => setMassImportMode(opt.v)} disabled={massImportBusy} />
                      {locale === 'de' ? opt.de : opt.en}
                    </label>
                  ))}
                </div>

                {massImportBusy && (
                  <div style={{ marginTop: 12, fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                    {locale === 'de' ? 'Anmeldung läuft…' : 'Registering…'} {massImportProgress}
                  </div>
                )}
                {massImportResult && (
                  <div style={{
                    marginTop: 12, padding: '10px 12px', borderRadius: 8, fontSize: '0.82rem',
                    background: massImportResult.failed.length > 0 ? 'rgba(237,139,0,0.08)' : 'rgba(134,188,37,0.10)',
                    border: `1px solid ${massImportResult.failed.length > 0 ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)'}`,
                    color: 'var(--dex-gray-700)', lineHeight: 1.5,
                  }}>
                    {locale === 'de' ? <><strong>{massImportResult.ok}</strong> Person(en) angemeldet.</> : <><strong>{massImportResult.ok}</strong> person(s) registered.</>}
                    {massImportResult.failed.length > 0 && (
                      <div style={{ marginTop: 4 }}>
                        {locale === 'de' ? 'Nicht angemeldet (bereits angemeldet / Fehler): ' : 'Not registered (already registered / error): '}
                        {massImportResult.failed.join(', ')}
                      </div>
                    )}
                  </div>
                )}

                <div style={{ marginTop: 18, display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
                  {massImportResult ? (
                    <button className="btn btn-primary" onClick={() => setMassImportOpen(false)}>
                      {locale === 'de' ? 'Schließen' : 'Close'}
                    </button>
                  ) : (
                    <>
                      <button className="btn btn-secondary" onClick={() => { setMassImportStep('input'); setMassImportResult(null); }} disabled={massImportBusy}>
                        {locale === 'de' ? 'Zurück' : 'Back'}
                      </button>
                      <button className="btn btn-primary" onClick={runMassImport} disabled={massImportBusy || okCount === 0}>
                        {massImportBusy ? (locale === 'de' ? 'Läuft…' : 'Running…') : (locale === 'de' ? `${okCount} anmelden` : `Register ${okCount}`)}
                      </button>
                    </>
                  )}
                </div>
              </>
            );
          })()}
        </Modal>
      )}

      {/* v18.75: Sicherheitshinweis-Dialog vor dem Absenden (pro Event). */}
      {confirmDialogOpen && event && (() => {
        const isFree = event.confirmDialogMode === 'freetext';
        // v18.76: ALLE Sub-Events zeigen (auch nicht ausgewählte), damit der
        // Teilnehmer im Dialog ab- UND zuwählen kann.
        const allChildren = childEvents;
        // v23.9: Im Klammer-Modus (subEventsOnlyMode) ist das Hauptevent nicht
        // buchbar — es wird nur als Schatten mitgeführt. Deshalb NICHT als
        // wählbare „(Haupt-Event)"-Zeile im Bestätigungs-Dialog zeigen (auch
        // nicht im Stellvertreter-Modus, wo es sonst über registerForOther
        // fälschlich auftauchte).
        const showParent = (willRegisterParent || registerForOther) && !(event && event.subEventsOnlyMode);
        const parentEditable = willRegisterParent && !registerForOther; // proxy: Parent fix
        const canConfirm = isFree
          ? confirmDialogAck
          : (confirmDraftParent || confirmDraftSessions.size > 0 || (showParent && !parentEditable));
        // v18.76: Datum + Uhrzeit pro Eintrag anzeigen.
        const fmtDT = (iso?: string): string => {
          if (!iso) return '';
          try { return new Date(iso).toLocaleString(locale === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); } catch { return ''; }
        };
        const dtRange = (s?: string, e?: string): string => {
          const a = fmtDT(s); const b = fmtDT(e);
          return a && b ? `${a} – ${b}` : (a || b);
        };
        return (
          <Modal
            open={confirmDialogOpen}
            onClose={() => setConfirmDialogOpen(false)}
            maxWidth={560}
            padding={24}
            ariaLabel={locale === 'de' ? 'Anmeldung bestätigen' : 'Confirm registration'}
          >
            <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {locale === 'de' ? 'Bitte bestätigen' : 'Please confirm'}
            </h3>
            {isFree ? (
              <>
                <div style={{
                  margin: '0 0 14px', padding: '12px 14px', whiteSpace: 'pre-wrap',
                  background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)',
                  borderRadius: 8, fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-800)',
                }}>
                  {(event.confirmDialogText || '').trim() || (locale === 'de' ? 'Bitte bestätige deine Anmeldung.' : 'Please confirm your registration.')}
                </div>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', marginBottom: 4 }}>
                  <input type="checkbox" checked={confirmDialogAck} onChange={e => setConfirmDialogAck(e.target.checked)} style={{ marginTop: 3 }} />
                  <span style={{ flex: 1, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
                    {locale === 'de' ? 'Ich habe den Hinweis gelesen und bestätige.' : 'I have read and acknowledge the note.'}
                  </span>
                </label>
              </>
            ) : (() => {
              // v19.0: statt generischem „Punkte/items" den konfigurierten
              // Section-Begriff verwenden (Default „Event-Sections").
              const sectionTerm = childTermPlural || (locale === 'de' ? 'Event-Sections' : 'event sections');
              return (
              <>
                <p style={{ margin: '0 0 12px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
                  {locale === 'de'
                    ? `Du meldest dich für die angehakten ${sectionTerm} an. Du kannst vor dem Absenden einzelne ${sectionTerm} ab- oder zuwählen:`
                    : `You are registering for the checked ${sectionTerm}. You can de-/select ${sectionTerm} before submitting:`}
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 14 }}>
                  {showParent && (
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: parentEditable ? 'pointer' : 'default', padding: '8px 10px', background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)', borderRadius: 6 }}>
                      <input
                        type="checkbox"
                        checked={parentEditable ? confirmDraftParent : true}
                        disabled={!parentEditable}
                        onChange={e => setConfirmDraftParent(e.target.checked)}
                        style={{ marginTop: 2 }}
                      />
                      <span style={{ flex: 1 }}>
                        <span style={{ fontSize: '0.88rem', fontWeight: 600, display: 'block' }}>{event.title}{(() => { const lbl = resolveMainEventLabel(locale === 'de' ? 'Haupt-Event' : 'main event'); return lbl ? <> <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>({lbl})</span></> : null; })()}</span>
                        {dtRange(event.startDate, event.endDate) && (
                          <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', display: 'block', marginTop: 1 }}>{dtRange(event.startDate, event.endDate)}</span>
                        )}
                      </span>
                    </label>
                  )}
                  {allChildren.map(ce => (
                    <label key={ce.id} style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', padding: '8px 10px', background: confirmDraftSessions.has(ce.id) ? 'rgba(134,188,37,0.06)' : 'var(--dex-gray-50, #f7f7f5)', border: `1px solid ${confirmDraftSessions.has(ce.id) ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`, borderRadius: 6 }}>
                      <input
                        type="checkbox"
                        checked={confirmDraftSessions.has(ce.id)}
                        style={{ marginTop: 2 }}
                        onChange={e => {
                          if (e.target.checked) {
                            // v18.76: Sub-Event mit eigenen Pflichtfeldern erst über
                            // das Sub-Event-Modal erfassen, damit keine leeren
                            // Pflicht-Antworten entstehen. Dialog schließen, Modal
                            // öffnen; nach dem Ausfüllen erscheint der Dialog erneut.
                            const hasCF = (ce.eventSpecificFields || []).length > 0;
                            if (hasCF && !sessionFieldValues[ce.id] && !selectedSessions.has(ce.id)) {
                              confirmDialogConfirmedRef.current = false;
                              setConfirmDialogOpen(false);
                              setPendingSubEventModal({ subEventId: ce.id, draftValues: { ...(sessionFieldValues[ce.id] || {}) } });
                            } else {
                              setConfirmDraftSessions(prev => { const n = new Set(prev); n.add(ce.id); return n; });
                            }
                          } else {
                            setConfirmDraftSessions(prev => { const n = new Set(prev); n.delete(ce.id); return n; });
                          }
                        }}
                      />
                      <span style={{ flex: 1 }}>
                        <span style={{ fontSize: '0.88rem', display: 'block' }}>{ce.title}</span>
                        {dtRange(ce.startDate, ce.endDate) && (
                          <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', display: 'block', marginTop: 1 }}>{dtRange(ce.startDate, ce.endDate)}</span>
                        )}
                      </span>
                    </label>
                  ))}
                </div>
                {!canConfirm && (
                  <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-red, #c00)' }}>
                    {locale === 'de' ? `Bitte mindestens eine ${sectionTerm} auswählen.` : `Please select at least one of the ${sectionTerm}.`}
                  </p>
                )}
              </>
            );
            })()}
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
              <button className="btn btn-secondary" onClick={() => setConfirmDialogOpen(false)} style={{ fontSize: '0.85rem' }}>
                {locale === 'de' ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                className="btn btn-primary"
                disabled={!canConfirm}
                onClick={() => {
                  // v18.75: In der Auswahl-Übersicht die (ggf. angepasste)
                  // Auswahl in den echten State übernehmen, dann Submit erneut
                  // anstoßen (Ref überspringt den Dialog).
                  if (!isFree) {
                    if (parentEditable) setRegisterForParent(confirmDraftParent);
                    setSelectedSessions(new Set(confirmDraftSessions));
                  }
                  confirmDialogConfirmedRef.current = true;
                  setConfirmDialogOpen(false);
                  setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 60);
                }}
                style={{ fontSize: '0.85rem' }}
              >
                {locale === 'de' ? 'Anmeldung bestätigen' : 'Confirm registration'}
              </button>
            </div>
          </Modal>
        );
      })()}

      {externalEmailWarning && (
        <Modal
          open={externalEmailWarning}
          onClose={() => setExternalEmailWarning(false)}
          maxWidth={540}
          padding={24}
          ariaLabel={locale === 'de' ? 'E-Mail-Adresse prüfen' : 'Check the email address'}
        >
            {/* v18.74: Tippfehler-Gegenlesen — die externe Adresse groß
                anzeigen und zur Bestätigung auffordern. */}
            <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem', color: 'var(--dex-orange-dark, #b35a00)' }}>
              {locale === 'de' ? 'E-Mail-Adresse prüfen' : 'Check the email address'}
            </h3>
            <p style={{ margin: '0 0 10px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
              {locale === 'de'
                ? <>Du meldest eine <strong>externe Person</strong> an. Bitte lies die Adresse genau gegen — an externe Adressen lässt sich ein <strong>Tippfehler nachträglich nicht korrigieren</strong>:</>
                : <>You are registering an <strong>external person</strong>. Please read the address carefully — a <strong>typo cannot be corrected afterwards</strong> for external addresses:</>}
            </p>
            <div style={{
              margin: '0 0 12px', padding: '12px 14px', textAlign: 'center',
              background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)',
              borderRadius: 8, fontSize: '1.05rem', fontWeight: 700, wordBreak: 'break-all',
              color: 'var(--dex-gray-900, #222)',
            }}>
              {email}
            </div>
            <p style={{ margin: '0 0 12px', fontSize: '0.82rem', lineHeight: 1.55, color: 'var(--dex-gray-600)' }}>
              {locale === 'de'
                ? <>Die <strong>Anmeldebestätigung</strong> geht direkt an diese Adresse, mit den <strong>Organizern auf CC</strong>. Ein <strong>Outlook-Termin</strong> wird an externe Adressen nicht versendet.</>
                : <>The <strong>confirmation email</strong> is sent directly to this address, with the <strong>organizers on CC</strong>. An <strong>Outlook invite</strong> is not sent to external addresses.</>}
            </p>
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setExternalEmailWarning(false)}
                style={{ fontSize: '0.85rem' }}
              >
                {locale === 'de' ? 'Zurück, korrigieren' : 'Back, edit'}
              </button>
              <button
                className="btn btn-primary"
                onClick={() => {
                  externalEmailConfirmedRef.current = true;
                  setExternalEmailWarning(false);
                  // Re-trigger submit via short timeout
                  setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
                }}
                style={{ fontSize: '0.85rem' }}
              >
                {locale === 'de' ? 'Adresse ist korrekt' : 'Address is correct'}
              </button>
            </div>
        </Modal>
      )}

      {/* v19.6: CC-Frage bei stellvertretender INTERNER Anmeldung. Erscheint
          nach dem „Anmelden"-Klick und vor der eigentlichen Anmeldung — der/die
          Anmeldende (Organizer, Co-Organizer oder Assistenz) entscheidet, ob er/
          sie selbst auf CC der Bestätigungs-Mail soll. */}
      {ccSelfModalOpen && (
        <Modal
          open={ccSelfModalOpen}
          onClose={() => setCcSelfModalOpen(false)}
          maxWidth={520}
          padding={24}
          ariaLabel={locale === 'de' ? 'Auf Kopie der Bestätigung?' : 'Copy on the confirmation?'}
        >
          <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
            {locale === 'de' ? 'Möchtest du eine Kopie der Bestätigung?' : 'Do you want a copy of the confirmation?'}
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
            {locale === 'de'
              ? <>Du meldest {`${firstName} ${surname}`.trim() ? <strong>{`${firstName} ${surname}`.trim()}</strong> : <>die ausgewählte Person</>} stellvertretend an. Möchtest du selbst auf <strong>CC der Bestätigungs-Mail</strong> gesetzt werden? Du bekommst dann eine Kopie der Anmeldebestätigung.<br /><br />Der <strong>Outlook-Termin</strong> wird davon nicht berührt — die CC gilt nur für die Bestätigungs-Mail.</>
              : <>You are registering {`${firstName} ${surname}`.trim() ? <strong>{`${firstName} ${surname}`.trim()}</strong> : <>the selected person</>} on their behalf. Would you like to be added to the <strong>CC of the confirmation email</strong>? You will then receive a copy of the confirmation.<br /><br />The <strong>Outlook invite</strong> is not affected — the CC only applies to the confirmation email.</>}
          </p>
          <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
            <button
              className="btn btn-secondary"
              onClick={() => {
                ccSelfRef.current = false;
                ccSelfDecidedRef.current = true;
                setCcSelfModalOpen(false);
                setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
              }}
              style={{ fontSize: '0.85rem' }}
            >
              {locale === 'de' ? 'Nein, ohne CC' : 'No, without CC'}
            </button>
            <button
              className="btn btn-primary"
              onClick={() => {
                ccSelfRef.current = true;
                ccSelfDecidedRef.current = true;
                setCcSelfModalOpen(false);
                setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
              }}
              style={{ fontSize: '0.85rem' }}
            >
              {locale === 'de' ? 'Ja, mich auf CC setzen' : 'Yes, add me to CC'}
            </button>
          </div>
        </Modal>
      )}

      {/* v24.48: Assistenz-Abfrage — erscheint nach „Register"-Klick für
          Partner/Director. People-Picker (Suche nach Name/E-Mail). */}
      {assistantModalOpen && (
        <Modal
          open={assistantModalOpen}
          onClose={() => setAssistantModalOpen(false)}
          maxWidth={560}
          padding={24}
          ariaLabel={locale === 'de' ? 'Assistenz informieren?' : 'Inform assistant?'}
        >
          <div style={{ display: 'inline-block', background: 'var(--dex-green, #86bc25)', color: '#fff', fontSize: '0.72rem', fontWeight: 700, padding: '4px 12px', borderRadius: 999, marginBottom: 12, letterSpacing: 0.4 }}>
            {locale === 'de' ? 'Für Partner & Directoren' : 'For Partners & Directors'}
          </div>
          <h3 style={{ margin: '0 0 10px', fontSize: '1.1rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
            {locale === 'de' ? 'Möchtest du deine Assistenz informieren?' : 'Do you want to inform your assistant?'}
          </h3>
          <p style={{ margin: '0 0 14px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
            {locale === 'de'
              ? 'Deine Assistenz bekommt eine Kopie der Bestätigung und sieht deine Anmeldung in der App — so bleibt sie auf dem Laufenden.'
              : 'Your assistant gets a copy of the confirmation and can see your registration in the app — so they stay in the loop.'}
          </p>
          <label style={{ display: 'block', fontWeight: 600, fontSize: '0.85rem', marginBottom: 6 }}>
            {locale === 'de' ? 'Assistenz auswählen (Name oder E-Mail)' : 'Select assistant (name or email)'}
          </label>
          <UserFieldPicker
            value={delegateAssistValue}
            onChange={setDelegateAssistValue}
            searchUsers={searchUsers}
            searchUserByEmail={searchUser}
            placeholder={locale === 'de' ? 'Vorname Nachname, Nachname Vorname oder E-Mail…' : 'First last, last first or email…'}
            errorStyle={{}}
            forcedIsDe={locale === 'de'}
          />
          <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap', marginTop: 18 }}>
            <button
              className="btn btn-secondary"
              onClick={() => {
                // Ohne Assistenz weiter.
                delegateChoiceRef.current = { enabled: false, value: '' };
                setDelegateAssistEnabled(false);
                setDelegateAssistValue('');
                assistantModalDecidedRef.current = true;
                setAssistantModalOpen(false);
                setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
              }}
              style={{ fontSize: '0.85rem' }}
            >
              {locale === 'de' ? 'Ohne Assistenz anmelden' : 'Register without assistant'}
            </button>
            <button
              className="btn btn-primary"
              disabled={!parsedDelegateAssist}
              title={!parsedDelegateAssist ? (locale === 'de' ? 'Bitte zuerst eine Assistenz auswählen.' : 'Please select an assistant first.') : ''}
              onClick={() => {
                delegateChoiceRef.current = { enabled: true, value: delegateAssistValue };
                setDelegateAssistEnabled(true);
                assistantModalDecidedRef.current = true;
                setAssistantModalOpen(false);
                setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
              }}
              style={{ fontSize: '0.85rem' }}
            >
              {locale === 'de' ? 'Mit Assistenz anmelden' : 'Register with assistant'}
            </button>
          </div>
        </Modal>
      )}

      {/* v10.12: Sub-Event Custom-Fields Modal — wird geöffnet wenn ein
          Sub-Event mit eigenen Custom-Fields angecheckt wird. Der User muss die
          Antworten ausfüllen + bestätigen, dann wandert die Session in
          selectedSessions und die Werte in sessionFieldValues. Beim „Abbrechen"
          wird die Session NICHT angecheckt. */}
      {pendingSubEventModal && (() => {
        const ce = childEvents.find(c => c.id === pendingSubEventModal.subEventId);
        if (!ce) return null;
        const draft = pendingSubEventModal.draftValues;
        // v24.16 BUG-FIX: showIf (Sichtbarkeitsbedingung) auch im Sub-Event-
        // Modal anwenden — bedingte Fragen wurden vorher IMMER angezeigt und
        // blockierten als Pflichtfeld die Bestätigung. Quell-Antwort steht im
        // Sub-Event-eigenen `draft`.
        const fields = (ce.eventSpecificFields || [])
          .filter(f => f && f.label)
          .filter(f => {
            if (!f.showIf || !f.showIf.fieldId) return true;
            const raw = (draft[f.showIf.fieldId] || '').trim();
            if (!raw) return false;
            const answers = raw.indexOf(' | ') >= 0
              ? raw.split(' | ').map(s => s.trim()).filter(Boolean)
              : [raw];
            return answers.some(a => f.showIf!.values.indexOf(a) >= 0);
          });
        const setDraft = (next: Record<string, string>): void => {
          setPendingSubEventModal(prev => prev ? { ...prev, draftValues: next } : prev);
        };
        const updateFieldValue = (fieldId: string, value: string): void => {
          setDraft({ ...draft, [fieldId]: value });
        };
        // v17.22: EN-Varianten auch im Sub-Event-Modal respektieren — geknüpft
        // an die Bilingual-Einstellung DES Sub-Events (ce), nicht des Parents.
        const useEnHere = locale === 'en' && !!ce.bilingualFields;
        const fLabel = (f: EventSpecificField): string =>
          (useEnHere && f.labelEn && f.labelEn.trim()) ? f.labelEn : f.label;
        const fHelp = (f: EventSpecificField): string | undefined =>
          (useEnHere && f.helpTextEn && f.helpTextEn.trim()) ? f.helpTextEn : f.helpText;
        const fOpt = (f: EventSpecificField, opt: string, idx: number): string =>
          (useEnHere && f.optionsEn && f.optionsEn[idx] && f.optionsEn[idx].trim()) ? f.optionsEn[idx] : opt;
        const requiredMissing = fields.filter(f => f.required && !((draft[f.id] || '').trim())).map(f => fLabel(f));
        const canSubmit = requiredMissing.length === 0;
        const onConfirm = (): void => {
          if (!canSubmit) return;
          setSessionFieldValues(prev => ({ ...prev, [ce.id]: { ...draft } }));
          setSelectedSessions(prev => {
            const next = new Set(prev);
            next.add(ce.id);
            return next;
          });
          setPendingSubEventModal(null);
        };
        const onCancel = (): void => setPendingSubEventModal(null);

        return (
          <Modal
            open={true}
            onClose={onCancel}
            maxWidth={520}
            padding={24}
            ariaLabel={ce.title || childTermSingular || (locale === 'de' ? 'Sub-Event' : 'Sub-event')}
          >
              <h3 style={{ margin: '0 0 6px', fontSize: '1.1rem' }}>
                {ce.title || childTermSingular || (locale === 'de' ? 'Sub-Event' : 'Sub-event')}
              </h3>
              <p style={{ margin: '0 0 18px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {locale === 'de'
                  ? `Bitte beantworte die Fragen für dieses ${childTermSingular || 'Sub-Event'}:`
                  : `Please answer the questions for this ${childTermSingular || 'sub-event'}:`}
              </p>

              <div style={{ display: 'flex', flexDirection: 'column', gap: 14, marginBottom: 20 }}>
                {fields.map(f => {
                  const val = draft[f.id] || '';
                  return (
                    <div key={f.id}>
                      <label className="form-label" style={{ display: 'block', fontSize: '0.85rem', marginBottom: 4 }}>
                        {fLabel(f)}
                        {f.required && <span style={{ color: 'var(--dex-red, #c00)', marginLeft: 4 }}>*</span>}
                        {/* v11.16: konsistenter InfoTooltip statt grauer
                            Inline-Beschreibung — gleicher Look wie auf
                            der Haupt-Register-Page. */}
                        {fHelp(f) && <InfoTooltip text={fHelp(f)} />}
                      </label>
                      {f.type === 'select' && f.multi ? (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          {(f.options || []).map((opt, optIdx) => {
                            const current = val.split(' | ').map(s => s.trim()).filter(Boolean);
                            const checked = current.indexOf(opt) >= 0;
                            return (
                              <label key={opt} style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={e => {
                                    const next = e.target.checked
                                      ? [...current, opt]
                                      : current.filter(x => x !== opt);
                                    updateFieldValue(f.id, next.join(' | '));
                                  }}
                                />
                                {fOpt(f, opt, optIdx)}
                              </label>
                            );
                          })}
                        </div>
                      ) : f.type === 'select' ? (
                        <select
                          className="form-input"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                          style={{ width: '100%', fontSize: '0.9rem' }}
                        >
                          <option value="">{locale === 'de' ? '— bitte wählen —' : '— please select —'}</option>
                          {(f.options || []).map((opt, optIdx) => <option key={opt} value={opt}>{fOpt(f, opt, optIdx)}</option>)}
                        </select>
                      ) : f.type === 'checkbox' ? (
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.85rem' }}>
                          <input
                            type="checkbox"
                            checked={val === 'true'}
                            onChange={e => updateFieldValue(f.id, e.target.checked ? 'true' : 'false')}
                          />
                          {locale === 'de' ? 'Ja' : 'Yes'}
                        </label>
                      ) : f.type === 'number' ? (
                        <input
                          type="number"
                          className="form-input"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                          style={{ width: '100%', fontSize: '0.9rem' }}
                        />
                      ) : (
                        <input
                          type="text"
                          className="form-input"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                          style={{ width: '100%', fontSize: '0.9rem' }}
                        />
                      )}
                    </div>
                  );
                })}
              </div>

              {!canSubmit && requiredMissing.length > 0 && (
                <div style={{ fontSize: '0.78rem', color: 'var(--dex-red, #c00)', marginBottom: 12 }}>
                  {locale === 'de' ? 'Pflichtfelder fehlen: ' : 'Required fields missing: '}{requiredMissing.join(', ')}
                </div>
              )}

              <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button type="button" className="btn btn-secondary" onClick={onCancel}>
                  {locale === 'de' ? 'Abbrechen' : 'Cancel'}
                </button>
                <button type="button" className="btn btn-primary" onClick={onConfirm} disabled={!canSubmit}>
                  {locale === 'de' ? 'Bestätigen' : 'Confirm'}
                </button>
              </div>
          </Modal>
        );
      })()}
    </div>
  );
}

