/**
 * Admin / Organizer Seite
 *
 * Zeigt alle Events des Admins. Nach Auswahl eines Events:
 * - Event bearbeiten (Daten ändern)
 * - Teilnehmerliste anzeigen
 * - Teilnehmerliste in SharePoint öffnen
 * - Neues Event erstellen
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { DeloitteEvent } from '../types';
import { SPRegistration } from '../services/EventService';
import { Plus, Users, FileText, Trash2, Copy, Mail, Send, Download, Pencil, ExternalLink, AlertCircle, Hash, Columns, Wrench, RefreshCw, X, Check, Link2, ChevronUp, ChevronDown, QrCode, Search, Info } from './Icons';
import { downloadSelfCheckInPdf } from '../utils/selfCheckInPdf';
import { isEventOver } from '../utils/eventFormat';
// v20.1: Self-Check-in jederzeit aktivierbar (Token-Erzeugung beim Klick).
// v20.2: + statische Check-in-URL für die QR-Kachel im Event-Detail.
// v20.3: + Default-Zeitfenster (2 Std. vor Start bis Event-Ende) zur Vorbelegung.
import { generateSelfCheckInToken, buildStaticCheckInUrl, defaultCheckInWindow } from '../utils/selfCheckIn';
// v20.0 (Audit): xlsx + qrcode werden nicht mehr statisch importiert, sondern
// erst beim tatsächlichen Gebrauch (Export-Klick / QR-Vorschau) als eigener
// Chunk nachgeladen — spart ~1 MB im Haupt-Bundle.
import { EventService } from '../services/EventService';
import { qrCodeEmail, qrEmailDefaults, buildQrBlockHtml, QrEmailOverride, cancellationEmail, promotionEmail, wrapTemplate, replacePlaceholders, buildEmailFromTemplate, getCachedLogoBase64, getCachedOrbBase64, injectIntoEmailContent } from '../services/EmailTemplates';
import { applyEventTemplateOverride, formatOrganizerList } from '../context/EventContext';
import { HtmlEditorModal } from './HtmlEditorModal';
import { InfoTooltip } from './InfoTooltip';
import { MultiSelectDropdown } from './MultiSelectDropdown';
import Modal from './Modal';
import { Icon } from '@fluentui/react/lib/Icon';
import InternationalSearchToggle from './InternationalSearchToggle';
// v20.4: moderne Confirm-/Alert-Modals statt window.confirm/alert.
import { useDialog } from '../context/DialogContext';

function formatDate(iso: string): string {
  if (!iso) return '-';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '-';
  return d.toLocaleDateString('de-DE', {
    day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit',
  });
}

function getStatusColor(status: string): string {
  switch (status) {
    case 'Active': return 'var(--dex-green)';
    case 'Completed': return 'var(--dex-gray-400)';
    case 'Cancelled': return 'var(--dex-red)';
    default: return 'var(--dex-orange)';
  }
}

// v11.41: Einladungsmail-Empfänger-Blocker. Die Einladungsmail darf NIE an
// komplette Standort-Verteiler ('de.duesseldorf@...', 'duesseldorf@...' etc.)
// oder an pauschale 'all'-Listen ('deall@...', 'all@...', 'alldeloitte@...')
// gehen. Hintergrund: solche Aussendungen sind ohne CMC-/Marketing-Freigabe
// nicht erlaubt — kleinere explizite Verteilergruppen (Team-Mailboxen,
// Funktions-Accounts) bleiben aber zulässig.
const DEX_LOCATION_TOKENS: string[] = [
  'berlin', 'dresden', 'duesseldorf', 'dusseldorf', 'düsseldorf',
  'frankfurt', 'goerlitz', 'görlitz', 'halle', 'hamburg', 'hannover',
  'koeln', 'köln', 'cologne', 'leipzig', 'magdeburg', 'mannheim',
  'muenchen', 'münchen', 'munich', 'nuernberg', 'nürnberg', 'nuremberg',
  'stuttgart', 'walldorf',
];

/** Wenn die Adresse als unerlaubter Massen-Verteiler erkannt wird, gibt den
 *  Block-Grund zurück — sonst null. Heuristik bewusst konservativ: matched
 *  nur, wenn der Local-Part (bzw. der gesamte Token, falls kein '@' vorhanden)
 *  eindeutig ein Standort-/All-Verteiler ist. Team-Mailboxen wie
 *  'frankfurt-event-team@' bleiben erlaubt.
 *
 *  v11.44: Auch reine Tokens ohne '@' werden geprüft — der Mailverteiler
 *  kann Einträge wie 'All' oder 'Duesseldorf' enthalten, die direkt aus dem
 *  Standort-/Location-Picker stammen. Vorher wurden die durchgelassen, weil
 *  der Parser an `at <= 0` zurückkehrte. */
function getBlockedInviteReason(email: string): string | null {
  const lc = (email || '').trim().toLowerCase();
  if (!lc) return null;
  const at = lc.indexOf('@');
  // Mit '@': Local-Part vor dem '@' prüfen. Ohne '@': gesamten Token prüfen
  // (z.B. wenn die Sichtbarkeit per Location-Picker auf 'All' gesetzt war
  // und 'All' so im audienceFilter landet).
  const local = at > 0 ? lc.slice(0, at) : lc;
  // Token-Split: Local-Part nach .-_ tokenisieren.
  const tokens = local.split(/[._-]/).filter(Boolean);
  // (1) deall / de.all / de-all / alldeloitte etc. — globaler DE-Verteiler.
  if (local === 'deall' || local === 'alldeloitte' || tokens.includes('deall') || tokens.includes('alldeloitte')) {
    return 'globaler Deloitte-DE-Verteiler';
  }
  // (2) 'all' als eigenständiger Token oder Local-Part — pauschale Liste.
  if (local === 'all' || tokens.includes('all')) {
    return 'globaler "all"-Verteiler';
  }
  // (3) Standort-Verteiler: Local-Part ist exakt eine Stadt ODER beginnt /
  //     endet mit 'de.' / 'de-' und enthält eine Stadt als Token.
  for (const loc of DEX_LOCATION_TOKENS) {
    if (local === loc) return `Standort-Verteiler (${loc})`;
    // 'de.<loc>' / 'de-<loc>' / '<loc>.de' / '<loc>-de'
    if (tokens.length === 2 && tokens.includes('de') && tokens.includes(loc)) {
      return `Standort-Verteiler (${loc})`;
    }
  }
  return null;
}

/** Liefert pro Empfänger die Block-Begründung — leeres Array = alles OK. */
function getBlockedInviteRecipients(emails: string[]): Array<{ email: string; reason: string }> {
  const out: Array<{ email: string; reason: string }> = [];
  for (const e of emails) {
    const reason = getBlockedInviteReason(e);
    if (reason) out.push({ email: e, reason });
  }
  return out;
}

// v19.11: Kurzname eines Sub-Events / einer Event-Section für die Anzeige im
// Admin Center. Sub-Events heißen per Konvention „<Hauptevent> | <Section>"
// (z.B. „P/D Meeting T&T+ | HER SPACE"). Im Admin Center reicht die Section
// („HER SPACE") — der Hauptevent-Name steht ohnehin oben. Heuristik:
//  1) Wenn ein Pipe „|" vorkommt, nur den Teil dahinter zeigen.
//  2) Sonst, falls der Titel exakt mit dem Hauptevent-Titel beginnt, diesen
//     Präfix (plus führende Trennzeichen) abschneiden.
//  3) Sonst den Titel unverändert lassen.
function shortSubEventTitle(title: string | undefined, parentTitle?: string): string {
  const t = (title || '').trim();
  if (!t) return t;
  const pipe = t.lastIndexOf('|');
  if (pipe >= 0) {
    const after = t.substring(pipe + 1).trim();
    if (after) return after;
  }
  const p = (parentTitle || '').trim();
  if (p && t.toLowerCase().startsWith(p.toLowerCase())) {
    const rest = t.substring(p.length).replace(/^[\s|:\-–—·•]+/, '').trim();
    if (rest) return rest;
  }
  return t;
}

// v9.20: EventStatus-Labels lokalisieren (DE).
// v11.89: 'Under Construction' wird transparent als 'Entwurf' angezeigt,
// solange noch Legacy-Daten existieren — neue Events nutzen IsFictive.
function localizeStatus(status: string): string {
  switch (status) {
    case 'Active': return 'Aktiv';
    case 'Under Construction': return 'Entwurf';
    case 'Completed': return 'Abgeschlossen';
    case 'Cancelled': return 'Abgesagt';
    default: return status;
  }
}

// v22.16: Heuristik für die „Hinweise"-Box bei aktiven Events — erkennt
// englischsprachigen Event-Inhalt (Beschreibung + Felder), damit die App
// empfehlen kann, die Anmeldesprache fest auf Englisch zu stellen (sonst
// mischt das Formular je nach App-Sprache des Teilnehmers Deutsch/Englisch).
function stripHtmlToText(html: string): string {
  return (html || '').replace(/<[^>]+>/g, ' ').replace(/&[a-z]+;/gi, ' ').replace(/\s+/g, ' ').trim();
}
function looksEnglishText(text: string): boolean {
  const t = ' ' + (text || '').toLowerCase().replace(/\s+/g, ' ') + ' ';
  if (t.trim().length < 40) return false; // zu wenig Text für ein Urteil
  const enWords = [' the ', ' and ', ' you ', ' your ', ' please ', ' with ', ' for ', ' our ', ' this ', ' are ', ' join ', ' we ', ' from ', ' all '];
  const deWords = [' der ', ' die ', ' das ', ' und ', ' nicht ', ' bitte ', ' wir ', ' euch ', ' dich ', ' ihr ', ' eine ', ' einen ', ' zur ', ' zum ', ' bei ', ' auf '];
  let en = 0;
  let de = 0;
  for (const w of enWords) if (t.indexOf(w) >= 0) en++;
  for (const w of deWords) if (t.indexOf(w) >= 0) de++;
  // Umlaute/ß sind ein starkes Deutsch-Signal.
  if (/[äöüß]/.test(t)) de += 2;
  return en >= 4 && en >= de * 2;
}

// Status-Werte sind in SP als deutsche Strings gespeichert ('Angemeldet',
// 'QR versendet', 'Warteliste', 'Eingecheckt', 'Abgemeldet'). Die App
// rendert sie hier in der UI-Sprache des Users, ohne den Datenbankwert
// zu aendern.
function translateStatus(status: string, isDe: boolean): string {
  if (isDe || !status) return status;
  switch (status) {
    case 'Angemeldet': return 'Registered';
    case 'QR versendet': return 'QR sent';
    case 'Warteliste': return 'Waitlist';
    case 'Eingecheckt': return 'Checked in';
    case 'Abgemeldet': return 'Cancelled';
    default: return status;
  }
}

// v11.14: Migriert hardcoded B2Run-Sonderbehandlungen aus dem Render-
// Code von RegistrationPage.tsx in echte Custom-Field-Properties:
//
// - b2run_mobilnummer ist nur sichtbar wenn b2run_infoservice='true'
//   → wird durch eine showIf-Bedingung auf dem Mobilnummer-Feld ersetzt.
//   Der Pflicht-Status bleibt dynamisch (true wenn sichtbar via showIf).
// - b2run_datenschutz hat im Render hardcoded externalLinks-Fallbacks
//   (B2Run-AGB + Datenschutz-URL) wenn die Field-Properties leer sind
//   → wird in das Field selbst persistiert.
// - b2run_laufshirt wird im Render auf required=true gezwungen
//   → wird in der Field-Property persistiert.
//
// Wird nur einmalig bei der Migration aufgerufen — die in-Memory-Field-
// Liste wird mutiert; Caller speichert das Ergebnis als CustomFields-
// JSON. Wenn keine relevanten Felder existieren, no-op.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function migrateB2RunFieldExtras(fields: any[]): { changed: boolean } {
  let changed = false;
  for (const f of fields) {
    const id = String(f.id || '').toLowerCase();
    if (id === 'b2run_mobilnummer') {
      // v11.14: showIf-Constraint auf b2run_infoservice='true'.
      // Damit übernimmt die generische Render-Logik die Sichtbarkeit
      // statt der hardcoded Sonderprüfung.
      const sf = f.showIf;
      const alreadySet = sf && sf.fieldId === 'b2run_infoservice'
        && Array.isArray(sf.values) && sf.values.indexOf('true') >= 0;
      if (!alreadySet) {
        f.showIf = { fieldId: 'b2run_infoservice', values: ['true'] };
        // Wenn der User Infoservice aktiviert, ist die Mobilnummer
        // Pflicht — über showIf gerendert ist die Pflicht-Logik
        // jetzt deterministisch (Feld sichtbar ⇒ Feld Pflicht).
        f.required = true;
        changed = true;
      }
    } else if (id === 'b2run_datenschutz') {
      // v11.14: Hardcoded B2Run-AGB- und Datenschutz-Links als
      // externalLinks-Property persistieren, sodass der Render-
      // Fallback-Path obsolet wird.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const links: any[] = Array.isArray(f.externalLinks) ? f.externalLinks : [];
      if (links.length === 0) {
        f.externalLinks = [
          { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
          { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
        ];
        changed = true;
      }
    } else if (id === 'b2run_laufshirt' || /laufshirt/i.test(String(f.label || ''))) {
      // v11.14: Hardcoded required=true persistieren.
      if (!f.required) {
        f.required = true;
        changed = true;
      }
    }
  }
  return { changed };
}

// v7.6: Wiederverwendbare Action-Kachel für den Aktionen-Bereich.
// Default in Grau, beim Hover/Focus kippt Border + Icon + Hintergrund auf
// Deloitte-Grün. Unterstützt Button (onClick), Link (href, oeffnet in neuem
// Tab) und passive Wrapper (children-Mode für Spezialfälle wie das
// Excel-Dropdown). Badge zeigt zwingend "Organizer" (grüner Tint) oder
// "Nur Admin" (oranger Tint), damit auf einen Blick klar ist, für welche
// Rolle die Aktion gedacht ist.
// v20.3: Kategorien für das Aktionen-Dropdown — die Aktionen werden nicht
// mehr als flache Alphabet-Liste gerendert, sondern als aufklappbare
// Kategorien (mehrzeilige Einträge: Titel fett, Beschreibung darunter).
type ActionCategoryKey = 'event' | 'participants' | 'mails' | 'checkin' | 'maintenance';
const ACTION_CATEGORY_ORDER: ActionCategoryKey[] = ['event', 'participants', 'mails', 'checkin', 'maintenance'];
// v20.4: pro Kategorie zusätzlich eine Kurzbeschreibung, was darin steckt —
// sichtbar direkt im zugeklappten Kategorie-Kopf.
const ACTION_CATEGORY_LABELS: Record<ActionCategoryKey, { de: string; en: string; descDe: string; descEn: string }> = {
  event: {
    de: 'Event', en: 'Event',
    descDe: 'Event bearbeiten, in SharePoint öffnen, Link teilen, Änderungsprotokoll ansehen.',
    descEn: 'Edit the event, open it in SharePoint, share the link, view the change history.',
  },
  participants: {
    de: 'Teilnehmer', en: 'Participants',
    descDe: 'Teilnehmerliste exportieren, Warteliste nachrücken, Nummern neu vergeben, Überbuchung prüfen.',
    descEn: 'Export the participant list, promote from the waitlist, renumber participants, check overbooking.',
  },
  mails: {
    de: 'E-Mails', en: 'Emails',
    descDe: 'Mails an Teilnehmer schreiben, Einladungsmail verschicken, E-Mail-Adressen kopieren.',
    descEn: 'Write emails to participants, send the invitation email, copy email addresses.',
  },
  checkin: {
    de: 'Check-in', en: 'Check-in',
    descDe: 'Check-in am Event-Tag starten, QR-Codes versenden, Self-Check-in (QR aushängen/anzeigen) einrichten.',
    descEn: 'Start check-in on event day, send QR codes, set up self check-in (post/show the QR).',
  },
  maintenance: {
    de: 'Wartung & Reparatur', en: 'Maintenance & repair',
    descDe: 'Werkzeuge für Sonderfälle: Daten reparieren, Zähler korrigieren, alte Events umstellen.',
    descEn: 'Tools for edge cases: repair data, correct counters, migrate old events.',
  },
};

interface ActionTileProps {
  icon: React.ReactNode;
  title: string;
  desc: string;
  badge: 'organizer' | 'admin';
  onClick?: () => void;
  href?: string;
  disabled?: boolean;
  busy?: boolean;
  result?: string | null;
  resultIsError?: boolean;
  // v9.19: filled-Variante für Highlight-Aktionen (z.B. Event aktivieren).
  // accent='green' = grün gefüllt, accent='red' = rot gefüllt.
  accent?: 'green' | 'red';
  // v20.3: Kategorie + optionale Unterkategorie (z.B. „Self-Check-in"
  // innerhalb von Check-in) für das gruppierte Aktionen-Dropdown.
  category?: ActionCategoryKey;
  subCategory?: string;
  // children: zusätzlicher Inhalt, der unterhalb der Standard-Tile-Inhalte
  // gerendert wird (z.B. das Excel-Dropdown-Menü).
  children?: React.ReactNode;
}
function ActionTile(props: ActionTileProps): React.ReactElement | null {
  // v12.7: Wenn ActionTile innerhalb eines ActionsRegistryProvider gerendert
  // wird, registriert er sich dort (title, desc, onClick) statt eine
  // Kachel zu zeichnen. Children-Mode (Excel-Sub-Dropdown) bleibt
  // sichtbar — sonst gingen Modals/Dropdowns verloren.
  const registry = React.useContext(ActionsRegistryContext);
  const registered = !!registry && !props.children;
  const [hover, setHover] = React.useState(false);
  // v22.6: NUR die (stabilen) register/unregister-Funktionen als Effekt-Deps —
  // nicht das ganze Context-Objekt. Das war vorher bei jedem Provider-Render ein
  // neues Objekt und ließ den Effekt endlos neu feuern (Render-Schleife → das
  // Suchfeld im Aktionen-Dropdown war dadurch unbeschreibbar).
  const registryRegister = registry?.register;
  const registryUnregister = registry?.unregister;
  React.useEffect(() => {
    if (!registryRegister || !registryUnregister || !registered) return undefined;
    const key = props.title;
    registryRegister({
      key,
      title: props.title,
      desc: props.desc,
      badge: props.badge,
      onClick: props.onClick,
      href: props.href,
      disabled: props.disabled || props.busy,
      // v20.3: Kategorie-Zuordnung fürs gruppierte Dropdown (Fallback: Event).
      category: props.category || 'event',
      subCategory: props.subCategory,
    });
    return () => registryUnregister(key);
  }, [registryRegister, registryUnregister, registered, props.title, props.desc, props.badge, props.onClick, props.href, props.disabled, props.busy, props.category, props.subCategory]);
  if (registered) return null;
  const isInteractive = !props.disabled && !props.busy;
  const greenAccent = isInteractive && hover;
  // v9.19/v9.20: filled-Look — Tile dezent eingefärbt für
  // Highlight-Aktionen. Pastell statt voll gesättigt, damit nicht
  // alarmierend wirkt.
  const isFilled = !!props.accent;
  const filledBg = props.accent === 'green' ? '#e3f0c5' : props.accent === 'red' ? '#ffe5e5' : '';
  const filledBorder = props.accent === 'green' ? 'var(--dex-green, #86bc25)' : props.accent === 'red' ? 'var(--dex-red, #da291c)' : '';
  const borderColor = isFilled ? filledBorder : (greenAccent ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200, #e5e7eb)');
  const bg = isFilled ? filledBg : (greenAccent ? 'rgba(134,188,37,0.06)' : '#fff');
  // v9.20: bei pastell-gefüllten Tiles Text/Icon dunkel halten — auf
  // hellem Pastell-Hintergrund gut lesbar (im Gegensatz zum vorherigen
  // weiss auf saturated-Color).
  const filledIconColor = props.accent === 'green' ? 'var(--dex-green-dark, #4a7c1f)' : props.accent === 'red' ? '#a01e15' : 'var(--dex-gray-500, #6b7280)';
  const iconColor = isFilled ? filledIconColor : (greenAccent ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-500, #6b7280)');
  const filledTextColor = isFilled
    ? (props.accent === 'green' ? 'var(--dex-green-dark, #3f5f10)' : props.accent === 'red' ? '#a01e15' : 'var(--dex-gray-800, #1f2937)')
    : 'var(--dex-gray-800, #1f2937)';
  const badgeLabel = props.badge === 'admin' ? 'Nur Admin' : 'Organizer';
  const badgeColors = props.badge === 'admin'
    ? { bg: 'rgba(237,139,0,0.12)', fg: 'var(--dex-orange, #ed8b00)' }
    : { bg: 'rgba(134,188,37,0.12)', fg: 'var(--dex-green-dark, #4a7c1f)' };
  const sharedStyle: React.CSSProperties = {
    textAlign: 'left', textDecoration: 'none', color: 'inherit',
    background: bg, border: `1px solid ${borderColor}`,
    borderRadius: 12, padding: 14,
    cursor: isInteractive ? 'pointer' : 'not-allowed',
    opacity: isInteractive ? 1 : 0.55,
    display: 'flex', flexDirection: 'column', gap: 8,
    fontFamily: 'inherit', fontSize: 'inherit',
    transition: 'all 0.15s ease',
    boxShadow: greenAccent ? '0 4px 12px rgba(134,188,37,0.18)' : 'none',
    position: 'relative',
    // width:100% sorgt dafür, dass die Kachel auch in einem flex-Wrapper
    // (z.B. Excel-Export hat einen <div display:flex>-Wrapper für das
    // Dropdown-Positioning) auf die volle Grid-Zellen-Breite gestreckt
    // wird — sonst sieht sie schmaler aus als die direkten Grid-Geschwister.
    width: '100%',
    boxSizing: 'border-box',
  };
  const inner = (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8, color: iconColor, transition: 'color 0.15s ease' }}>
          {props.icon}
          <span style={{ fontWeight: 600, fontSize: '0.88rem', color: filledTextColor }}>{props.title}</span>
        </span>
        <span style={{
          fontSize: '0.65rem', padding: '2px 8px', borderRadius: 999,
          // v9.20: Badge auf pastell Tiles in normalem badge-Look (auf hellem
          // Hintergrund gut sichtbar, im Gegensatz zur vorherigen
          // semi-transparenten weissen Variante auf saturated bg).
          background: badgeColors.bg,
          color: badgeColors.fg, fontWeight: 600,
          whiteSpace: 'nowrap', flexShrink: 0, letterSpacing: '0.02em',
        }}>{badgeLabel}</span>
      </div>
      {props.result && (
        <p style={{
          margin: 0, fontSize: '0.72rem',
          color: props.resultIsError ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)',
          fontStyle: 'italic',
        }}>{props.result}</p>
      )}
      {props.children}
      {hover && props.desc && (
        <div
          role="tooltip"
          style={{
            position: 'absolute',
            top: 'calc(100% + 8px)',
            left: 0,
            right: 0,
            zIndex: 50,
            background: 'var(--dex-gray-900, #1f2937)',
            color: '#fff',
            padding: '10px 12px',
            borderRadius: 8,
            fontSize: '0.76rem',
            lineHeight: 1.45,
            fontWeight: 400,
            boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
            pointerEvents: 'none',
            whiteSpace: 'normal',
          }}
        >
          {props.desc}
        </div>
      )}
    </>
  );
  if (props.href) {
    return (
      <a
        href={props.href}
        target="_blank"
        rel="noopener noreferrer"
        style={sharedStyle}
        onMouseEnter={() => setHover(true)}
        onMouseLeave={() => setHover(false)}
      >
        {inner}
      </a>
    );
  }
  return (
    <button
      type="button"
      disabled={!isInteractive}
      onClick={props.onClick}
      style={sharedStyle}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
    >
      {inner}
    </button>
  );
}

// v11.98: Pill-Toggle für die Aktiv-Teilnehmer-Tabelle bei Split-Kapazität.
// Default 'split' = getrennte Tabellen pro Gruppe. 'merged' = einzelne
// Tabelle (alter Look).
function SplitMergeToggle(props: {
  view: 'split' | 'merged';
  setView: (v: 'split' | 'merged') => void;
  isDe: boolean;
}): React.ReactElement {
  const pill = (active: boolean): React.CSSProperties => ({
    padding: '5px 12px',
    borderRadius: 999,
    fontSize: '0.78rem',
    fontWeight: 600,
    cursor: 'pointer',
    border: `1px solid ${active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
    background: active ? 'rgba(134,188,37,0.10)' : '#fff',
    color: active ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
    transition: 'all 0.12s ease',
  });
  return (
    <div style={{ display: 'inline-flex', gap: 4, alignItems: 'center' }}>
      <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginRight: 6 }}>
        {props.isDe ? 'Ansicht:' : 'View:'}
      </span>
      <button type="button" onClick={() => props.setView('split')} style={pill(props.view === 'split')}>
        {props.isDe ? 'Getrennt' : 'Split'}
      </button>
      <button type="button" onClick={() => props.setView('merged')} style={pill(props.view === 'merged')}>
        {props.isDe ? 'Zusammen' : 'Merged'}
      </button>
    </div>
  );
}

// v12.7: Sammel-Card-Wrapper aus v12.6 entfernt — Aktionen leben jetzt
// als alphabetische Dropdown-Liste innerhalb der Event-Detail-Card
// (siehe ActionsDropdown weiter unten). Diese Komponente bleibt im
// Code für Backward-Compat, ihre Children werden display:none gerendert
// damit React-State + onClick-Handler weiterhin funktionieren.
function ActionsCollapsibleCard(props: {
  isDe: boolean;
  children: React.ReactNode;
}): React.ReactElement {
  // v12.7: nicht mehr in eigener Card — wir verstecken die ganze Box
  // (display:none) und die ActionTiles registrieren sich via Context
  // im ActionsDropdown.
  void props.isDe;
  return (
    <div style={{ display: 'none' }}>
      {props.children}
    </div>
  );
}

// v12.7: Action-Registry — ActionTile-Instanzen melden sich beim Mount
// hier an. Der ActionsDropdown unten in der Event-Detail-Card liest den
// registry-State und rendert alle Einträge als alphabetisch sortierte
// Dropdown-Liste mit Hover-Tooltip (desc).
interface RegisteredAction {
  key: string;
  title: string;
  desc: string;
  badge: 'organizer' | 'admin';
  onClick?: () => void;
  href?: string;
  disabled?: boolean;
  // v20.3: Kategorie + optionale Unterkategorie fürs gruppierte Dropdown.
  category: ActionCategoryKey;
  subCategory?: string;
}
const ActionsRegistryContext = React.createContext<{
  register: (_a: RegisteredAction) => void;
  unregister: (_key: string) => void;
  actions: RegisteredAction[];
} | null>(null);

function ActionsRegistryProvider(props: { children: React.ReactNode }): React.ReactElement {
  const [actions, setActions] = React.useState<RegisteredAction[]>([]);
  const register = React.useCallback((a: RegisteredAction) => {
    setActions(prev => {
      const filtered = prev.filter(x => x.key !== a.key);
      return [...filtered, a];
    });
  }, []);
  const unregister = React.useCallback((key: string) => {
    setActions(prev => prev.filter(x => x.key !== key));
  }, []);
  // v22.6: Context-Value memoisieren — sonst entsteht bei jedem Render ein neues
  // Objekt, das die ActionTile-Register-Effekte erneut feuern lässt → Render-
  // Schleife (machte zuvor das Suchfeld im Aktionen-Dropdown unbeschreibbar).
  const value = React.useMemo(() => ({ register, unregister, actions }), [register, unregister, actions]);
  return React.createElement(ActionsRegistryContext.Provider, { value }, props.children);
}

// v22.50: Sprung aus der globalen Header-Suche in eine konkrete Aktion. Die
// Suche legt den Aktions-Key in localStorage ab; hier öffnen wir das Dropdown
// und filtern es auf den passenden Begriff vor. DE-/EN-Seed muss ein
// Teilstring des registrierten Aktions-Titels sein (Substring-Filter).
const ACTION_FOCUS_SEED: Record<string, { de: string; en: string }> = {
  export: { de: 'Excel-Export', en: 'Excel export' },
  qr: { de: 'QR-Codes versenden', en: 'Send QR codes' },
  massmail: { de: 'E-Mail versenden', en: 'Send email' },
  invite: { de: 'Einladungsmail', en: 'Invitation email' },
  audit: { de: 'Audit-Log', en: 'Audit log' },
  selfcheckin: { de: 'Self-Check-in', en: 'self check-in' },
  idreorder: { de: 'IDs neu vergeben', en: 'Reassign IDs' },
  overbook: { de: 'Überbuchung', en: 'overbooking' },
  accessfix: { de: 'Zugriff reparieren', en: 'repair access' },
  fixcols: { de: 'Spalten fixen', en: 'Fix columns' },
};

function ActionsDropdown(props: { isDe: boolean }): React.ReactElement | null {
  const ctx = React.useContext(ActionsRegistryContext);
  const [open, setOpen] = React.useState(false);
  // v20.3: aufklappbare Kategorien + Unterkategorien (z.B. „Self-Check-in"
  // unter Check-in) statt flacher Alphabet-Liste. Einträge sind mehrzeilig:
  // Titel fett, Beschreibung darunter — der frühere Hover-Tooltip entfällt.
  const [expanded, setExpanded] = React.useState<Set<string>>(new Set());
  // v22.5: Freitext-Suche über alle Aktionen (Titel + Beschreibung). Solange
  // etwas eingetippt ist, werden alle Kategorien automatisch aufgeklappt.
  const [query, setQuery] = React.useState('');
  const rootRef = React.useRef<HTMLDivElement | null>(null);
  const focusSeededRef = React.useRef(false);
  React.useEffect(() => {
    if (!open) return undefined;
    const onDocClick = (e: MouseEvent): void => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', onDocClick);
    return () => document.removeEventListener('mousedown', onDocClick);
  }, [open]);
  // v22.5: Suchfeld leeren, sobald das Dropdown geschlossen wird.
  React.useEffect(() => { if (!open) setQuery(''); }, [open]);
  // v22.50: Auto-Open + Vorfilter, wenn die Header-Suche eine Aktion angesteuert
  // hat. Einmalig, sobald Aktionen registriert sind.
  React.useEffect(() => {
    if (focusSeededRef.current) return;
    if (!ctx || ctx.actions.length === 0) return;
    let hint = '';
    try { hint = window.localStorage.getItem('dex_search_focus_action') || ''; } catch { /* */ }
    if (!hint) return;
    try { window.localStorage.removeItem('dex_search_focus_action'); } catch { /* */ }
    const seed = ACTION_FOCUS_SEED[hint];
    focusSeededRef.current = true;
    if (seed) { setQuery(props.isDe ? seed.de : seed.en); setOpen(true); }
  }, [ctx, props.isDe]);
  if (!ctx || ctx.actions.length === 0) return null;
  const lang = props.isDe ? 'de' : 'en';
  // v22.5: Kategorien alphabetisch nach lokalisiertem Label sortieren.
  const sortedCats = ACTION_CATEGORY_ORDER.slice().sort((a, b) => {
    const la = props.isDe ? ACTION_CATEGORY_LABELS[a].de : ACTION_CATEGORY_LABELS[a].en;
    const lb = props.isDe ? ACTION_CATEGORY_LABELS[b].de : ACTION_CATEGORY_LABELS[b].en;
    return la.localeCompare(lb, lang);
  });
  // v22.5: Aktiver Suchbegriff (klein geschrieben) + Treffer-Filter.
  const q = query.trim().toLowerCase();
  const matchesQuery = (a: RegisteredAction): boolean =>
    !q || a.title.toLowerCase().indexOf(q) >= 0 || (!!a.desc && a.desc.toLowerCase().indexOf(q) >= 0);
  const visibleActions = ctx.actions.filter(matchesQuery);
  const toggleKey = (k: string): void => {
    setExpanded(prev => {
      const next = new Set(prev);
      if (next.has(k)) next.delete(k); else next.add(k);
      return next;
    });
  };
  const runAction = (a: RegisteredAction): void => {
    if (a.disabled) return;
    setOpen(false);
    if (a.href) {
      window.open(a.href, '_blank', 'noopener,noreferrer');
    } else if (a.onClick) {
      a.onClick();
    }
  };
  const renderActionRow = (a: RegisteredAction, indent: number): React.ReactElement => {
    const adminOnly = a.badge === 'admin';
    return (
      <div
        key={a.key}
        onClick={() => runAction(a)}
        onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = 'rgba(134,188,37,0.07)'; }}
        onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = '#fff'; }}
        style={{
          padding: `9px 12px 9px ${indent}px`,
          cursor: a.disabled ? 'not-allowed' : 'pointer',
          borderBottom: '1px solid var(--dex-gray-100)',
          opacity: a.disabled ? 0.5 : 1,
          background: '#fff',
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
          <span style={{ fontWeight: 700, fontSize: '0.87rem', color: 'var(--dex-gray-800)' }}>{a.title}</span>
          <span style={{
            fontSize: '0.68rem', padding: '2px 8px', borderRadius: 999,
            background: adminOnly ? 'rgba(237,139,0,0.12)' : 'rgba(134,188,37,0.12)',
            color: adminOnly ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green-dark, #4a7c1f)',
            fontWeight: 700, whiteSpace: 'nowrap', flexShrink: 0,
          }}>
            {adminOnly ? (props.isDe ? 'Nur Admin' : 'Admin only') : 'Organizer'}
          </span>
        </div>
        {a.desc && (
          <div style={{ marginTop: 3, fontSize: '0.76rem', color: 'var(--dex-gray-500)', lineHeight: 1.45 }}>
            {a.desc}
          </div>
        )}
      </div>
    );
  };
  return (
    <div ref={rootRef} style={{ position: 'relative', marginTop: 12 }}>
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        style={{
          width: '100%', textAlign: 'left', padding: '10px 14px',
          // v19.27: grün hinterlegt, damit die Aktionen-Auswahl deutlich auffällt.
          border: '1.5px solid var(--dex-green, #86bc25)', borderRadius: 10,
          background: 'rgba(134,188,37,0.12)', color: 'var(--dex-green-dark, #4a7c1f)',
          fontSize: '0.92rem', fontWeight: 700, cursor: 'pointer',
          display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12,
        }}
      >
        <span>{props.isDe ? `Aktion auswählen (${ctx.actions.length})` : `Pick an action (${ctx.actions.length})`}</span>
        <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.85rem', transform: open ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.15s ease' }}>▾</span>
      </button>
      {open && (
        <div style={{
          position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0,
          background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 10,
          boxShadow: '0 8px 24px rgba(0,0,0,0.12)', zIndex: 50,
          maxHeight: 480, overflowY: 'auto',
        }}>
          {/* v22.5: Suchfeld — filtert alle Aktionen quer über die Kategorien. */}
          <div style={{
            position: 'sticky', top: 0, zIndex: 2, background: '#fff',
            padding: 10, borderBottom: '1px solid var(--dex-gray-200)',
          }}>
            <div style={{ position: 'relative' }}>
              <span style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--dex-gray-400)', display: 'inline-flex', pointerEvents: 'none' }}>
                <Search size={15} />
              </span>
              <input
                type="text"
                autoFocus
                value={query}
                onChange={e => setQuery(e.target.value)}
                placeholder={props.isDe ? 'Aktion suchen…' : 'Search action…'}
                style={{
                  width: '100%', boxSizing: 'border-box', padding: '8px 30px 8px 32px',
                  border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.85rem',
                  outline: 'none',
                }}
              />
              {query && (
                <button
                  type="button"
                  onClick={() => setQuery('')}
                  aria-label={props.isDe ? 'Suche leeren' : 'Clear search'}
                  style={{ position: 'absolute', right: 6, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-400)', display: 'inline-flex', padding: 4 }}
                >
                  <X size={14} />
                </button>
              )}
            </div>
          </div>
          {q && visibleActions.length === 0 && (
            <div style={{ padding: '18px 14px', textAlign: 'center', color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
              {props.isDe ? 'Keine Aktion gefunden.' : 'No action found.'}
            </div>
          )}
          {sortedCats.map(catKey => {
            const inCat = visibleActions.filter(a => a.category === catKey);
            if (inCat.length === 0) return null;
            const catLabel = props.isDe ? ACTION_CATEGORY_LABELS[catKey].de : ACTION_CATEGORY_LABELS[catKey].en;
            const catDesc = props.isDe ? ACTION_CATEGORY_LABELS[catKey].descDe : ACTION_CATEGORY_LABELS[catKey].descEn;
            // v22.5: bei aktiver Suche alle Treffer-Kategorien automatisch öffnen.
            const catOpen = q ? true : expanded.has(catKey);
            const direct = inCat.filter(a => !a.subCategory).slice().sort((a, b) => a.title.localeCompare(b.title, lang));
            const subNames = Array.from(new Set(inCat.filter(a => !!a.subCategory).map(a => a.subCategory as string))).sort((a, b) => a.localeCompare(b, lang));
            return (
              <div key={catKey}>
                <div
                  onClick={() => toggleKey(catKey)}
                  onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = 'rgba(134,188,37,0.10)'; }}
                  onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = 'var(--dex-gray-50, #fafafa)'; }}
                  style={{
                    padding: '10px 12px', cursor: 'pointer', userSelect: 'none',
                    display: 'flex', alignItems: 'center', gap: 8,
                    background: 'var(--dex-gray-50, #fafafa)',
                    borderBottom: '1px solid var(--dex-gray-200)',
                    transition: 'background 0.12s ease',
                  }}
                >
                  <span style={{ width: 14, color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.8rem', flexShrink: 0 }}>{catOpen ? '▾' : '▸'}</span>
                  <span style={{ minWidth: 0, flex: 1 }}>
                    <span style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <span style={{ fontWeight: 800, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>{catLabel}</span>
                      <span style={{
                        fontSize: '0.68rem', padding: '1px 7px', borderRadius: 999,
                        background: 'rgba(134,188,37,0.12)', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700,
                      }}>{inCat.length}</span>
                    </span>
                    {/* v20.4: Kurzbeschreibung, was in der Kategorie steckt. */}
                    <span style={{ display: 'block', fontSize: '0.74rem', color: 'var(--dex-gray-500)', fontWeight: 400, marginTop: 1, lineHeight: 1.4 }}>
                      {catDesc}
                    </span>
                  </span>
                </div>
                {catOpen && direct.map(a => renderActionRow(a, 30))}
                {catOpen && subNames.map(sub => {
                  const subKey = `${catKey}::${sub}`;
                  const subOpen = q ? true : expanded.has(subKey);
                  const subActions = inCat.filter(a => a.subCategory === sub).slice().sort((a, b) => a.title.localeCompare(b.title, lang));
                  return (
                    <div key={subKey}>
                      <div
                        onClick={() => toggleKey(subKey)}
                        onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = 'rgba(134,188,37,0.12)'; }}
                        onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = 'rgba(134,188,37,0.05)'; }}
                        style={{
                          padding: '8px 12px 8px 30px', cursor: 'pointer', userSelect: 'none',
                          display: 'flex', alignItems: 'center', gap: 8,
                          background: 'rgba(134,188,37,0.05)',
                          borderBottom: '1px solid var(--dex-gray-100)',
                          transition: 'background 0.12s ease',
                        }}
                      >
                        <span style={{ width: 14, color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.75rem' }}>{subOpen ? '▾' : '▸'}</span>
                        <span style={{ fontWeight: 700, fontSize: '0.84rem', color: 'var(--dex-gray-700)' }}>{sub}</span>
                        <span style={{
                          fontSize: '0.66rem', padding: '1px 6px', borderRadius: 999,
                          background: 'rgba(134,188,37,0.12)', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700,
                        }}>{subActions.length}</span>
                      </div>
                      {subOpen && subActions.map(a => renderActionRow(a, 46))}
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// v12.6: Sammel-Card für alle Aktionen + Quick-Actions unter „Currently
// registered". v12.7: ersetzt durch ActionsDropdown + ActionsRegistry —
// die folgende Funktion bleibt aus Kompatibilitätsgründen als no-op.

// v14.11 / v19.30: Aggregierte Zeile im konsolidierten View (Hauptevent mit
// Sub-Events), eine pro Person. Auf Modul-Ebene definiert, damit auch
// Handler außerhalb des Render-Bodys (z.B. das Abmelde-/Edit-Modal von
// Feature A/B) den Typ referenzieren können.
type ConsolidatedRow = {
  emailKey: string;
  email: string;
  vorname: string;
  nachname: string;
  jobTitle: string;
  location: string;
  teilnehmerId: number | null;
  /** v15.23: Früheste RegistrationDate über alle Sub-Event-
   *  Registrierungen der Person — Default-Sortierschlüssel im
   *  konsolidierten View (chronologisch nach erster Anmeldung). */
  earliestRegistrationTs: number;
  perChild: Record<string, SPRegistration | undefined>;
  activeCount: number;
};

export default function AdminPage(): React.ReactElement {
  const { navigate, selectedEventId } = useNavigation();
  // v14.11: zusätzlich `events` (alle Events inkl. Sub-Events) als `allEvents`
  // für die Parent-Lookup-Logik im konsolidierten View + im Sub-Event-Detail.
  const { events: allEvents, topLevelEvents: events, childEventsOf, isEventsLoading, getAllRegistrations, deleteEvent, updateEvent, refreshEvents, addTeamMember, assignTeamlessToTeam, notifyExistingTeamMembers, transferTeamLead } = useEvents();
  const [isRefreshing, setIsRefreshing] = React.useState(false);
  const handleRefresh = async (): Promise<void> => {
    if (isRefreshing) return;
    setIsRefreshing(true);
    try {
      await refreshEvents();
      // Wenn ein Event gerade selektiert ist, auch dessen Registrations neu laden
      if (selectedEvent) {
        try {
          const regs = await getAllRegistrations(selectedEvent.id);
          setRegistrations(regs);
        } catch { /* */ }
      }
    } finally { setIsRefreshing(false); }
  };
  const { currentUser } = useCurrentUser();
  const { isAdmin, siteUrl, currentUserRole, searchUser, searchUsers, isImpersonating } = useRoles();
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  // v20.4: App-Modals statt nativer Browser-Dialoge.
  const { confirmDialog, showAlert } = useDialog();
  const [selectedEvent, setSelectedEvent] = React.useState<DeloitteEvent | null>(null);
  const [registrations, setRegistrations] = React.useState<SPRegistration[]>([]);
  // v22.7: Konten-Aktiv-Check — Adressen (lowercase) der Teilnehmer, deren
  // Deloitte-Konto nicht mehr aktiv ist (Person hat womöglich das Unternehmen
  // verlassen). Wird im Hintergrund max. 1×/Tag pro Event geprüft.
  const [inactiveAccounts, setInactiveAccounts] = React.useState<string[]>([]);
  // v23.2: Doppel-Anmeldungen erkennen. Eine E-Mail mit ≥2 NICHT-abgemeldeten
  // Zeilen in derselben Teilnehmerliste = Duplikat (z.B. dieselbe Person in
  // zwei Teams). Wird oben in einer Hinweis-Box surfaced + die Zeilen werden
  // rot markiert. duplicateEmails = Set der betroffenen Adressen (lowercase).
  const duplicateEmails = React.useMemo<Set<string>>(() => {
    const counts: Record<string, number> = {};
    for (const r of registrations) {
      if ((r.Status || '') === 'Abgemeldet') continue;
      const em = (r.ParticipantEmail || '').trim().toLowerCase();
      if (!em) continue;
      counts[em] = (counts[em] || 0) + 1;
    }
    const dup = new Set<string>();
    Object.keys(counts).forEach(em => { if (counts[em] > 1) dup.add(em); });
    return dup;
  }, [registrations]);
  // v23.2: Duplikat-Abmelde-Modal — gezogene Zeile + Entscheidung still löschen
  // (Duplikat entfernen, keine Mail/Outlook/Nachrücken) vs. normal abmelden.
  const [dupCancelReg, setDupCancelReg] = React.useState<SPRegistration | null>(null);
  const [dupCancelBusy, setDupCancelBusy] = React.useState(false);

  // v23.2: Standard-Abmeldung einer Teilnehmer-Zeile (extrahiert aus dem
  // Abmelden-Button, damit das Duplikat-Modal denselben „normal abmelden"-Pfad
  // wiederverwenden kann). Enthält KEINEN Confirm — der Aufrufer bestätigt.
  // Spiegelt das bisherige Inline-Verhalten 1:1 (vergangenes Event → still,
  // sonst Abmelde-Mail + Outlook-Ausladen + Nachrücken + ID-Reorder).
  const performStandardCancel = async (reg: SPRegistration): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
    const eventWasOver = isEventOver(selectedEvent);
    setAdminToast({ kind: 'cancelling', name });
    const cancelledStarterType = reg.StarterType || '';
    await eventServiceRef.cancelRegistration(selectedEvent.subsiteUrl, reg.Id, `${currentUser.firstName} ${currentUser.surname}`.trim(), currentUser.email);
    if (reg.ParticipantEmail && !eventWasOver) {
      if (!selectedEvent.disableEmails && !selectedEvent.disableCancellationEmail) {
        const emailData = cancellationEmail(name, selectedEvent.title);
        eventServiceRef.queueEmail(
          emailData.subject, reg.ParticipantEmail, name, emailData.body,
          'Abmeldung', selectedEvent.title, selectedEvent.id
        ).catch(err => console.warn('[DEX]', err));
      }
      if (!selectedEvent.disableOutlook) {
        eventServiceRef.queueOutlookEvent(
          reg.ParticipantEmail, selectedEvent.id, selectedEvent.title, 'Ausladen'
        ).catch(err => console.warn('[DEX]', err));
      }
    }
    if (reg.ParticipantEmail && selectedEvent.eventNumber) {
      eventServiceRef.removeParticipantEvent(
        reg.ParticipantEmail, selectedEvent.eventNumber
      ).catch(err => console.warn('[DEX]', err));
    }
    const isSplitEvent = typeof selectedEvent.durchstarterCapacity === 'number'
      && typeof selectedEvent.funstarterCapacity === 'number'
      && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
    const useTypeFilter = isSplitEvent && !selectedEvent.splitSharedWaitlist;
    if (!eventWasOver) {
      try {
        const promoted = await eventServiceRef.promoteFirstWaitlistItem(
          selectedEvent.subsiteUrl,
          cancelledStarterType || undefined,
          selectedEvent.maxParticipants,
          (useTypeFilter && cancelledStarterType) ? cancelledStarterType : undefined,
          { itemId: reg.Id, participantEmail: reg.ParticipantEmail || '' },
        );
        if (promoted && promoted.success && promoted.email) {
          setAdminToast({ kind: 'promoted', name: promoted.name || promoted.email, email: promoted.email, type: cancelledStarterType || undefined });
          if (!selectedEvent.disableEmails) {
            try {
              const lang = selectedEvent.emailLanguage || 'EN';
              const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
              const promoteVars = {
                Name: promotedFirstName,
                EventTitle: selectedEvent.title,
                Organizer: formatOrganizerList(selectedEvent.organizers, lang),
                AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
                WaitlistPosition: '',
              };
              let emailData: { subject: string; body: string };
              const spTplRaw = await eventServiceRef.getEmailTemplate('Nachruecken', lang).catch(() => null);
              const spTpl = applyEventTemplateOverride(spTplRaw, selectedEvent.emailTemplateOverrides, 'Nachruecken');
              if (spTpl) { emailData = buildEmailFromTemplate(spTpl, promoteVars); }
              else { emailData = promotionEmail(promotedFirstName, selectedEvent.title); }
              await eventServiceRef.queueEmail(
                emailData.subject, promoted.email, promoted.name || '', emailData.body,
                'Nachruecken', selectedEvent.title, selectedEvent.id
              );
            } catch (err) { console.warn('[DEX] promote-email failed:', err); }
          }
          if (!selectedEvent.disableOutlook) {
            try { await eventServiceRef.queueOutlookEvent(promoted.email, selectedEvent.id, selectedEvent.title, 'Einladen'); }
            catch (err) { console.warn('[DEX] promote-outlook failed:', err); }
          }
        } else {
          setAdminToast({ kind: 'no-promote', name });
        }
      } catch (err) {
        console.warn('[DEX] promoteFirstWaitlistItem failed:', err);
        setAdminToast({ kind: 'no-promote', name });
      }
    }
    if (selectedEvent.subsiteUrl && !eventWasOver) {
      try {
        const ok = await eventServiceRef.queueIDReorder(
          selectedEvent.id, selectedEvent.eventNumber || 0,
          selectedEvent.subsiteUrl, selectedEvent.title, name, reg.ParticipantEmail || undefined
        );
        if (!ok) {
          console.warn('[DEX] queueIDReorder returned false');
          showAlert(isDe ? 'Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.' : 'Cancellation successful, but the ID reorder entry could not be written to the queue. Please click "Reassign IDs" once.');
        }
      } catch (err) {
        console.warn('[DEX] queueIDReorder threw:', err);
        showAlert('Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.');
      }
    }
    const regs = await getAllRegistrations(selectedEvent.id);
    setRegistrations(regs);
  };

  // v23.2: Stilles Löschen einer doppelten Anmelde-Zeile. Anders als die
  // normale Abmeldung wird die Zeile HART gelöscht (kein „Abgemeldet"-Status,
  // der die Abmeldungs-Liste aufblähen würde) und es laufen KEINE Seiteneffekte
  // (keine Abmelde-Mail, kein Outlook-Ausladen, kein Nachrücken, kein
  // ID-Reorder, kein DEX_Participants-Cleanup) — die Person bleibt über ihre
  // andere Zeile regulär angemeldet. Sitzplatz-Counter wird nachgezogen.
  const performSilentDuplicateDelete = async (reg: SPRegistration): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
    setAdminToast({ kind: 'cancelling', name });
    try {
      await eventServiceRef.deleteRegistration(selectedEvent.subsiteUrl, reg.Id);
      try {
        await eventServiceRef.writeChangeLog({
          action: 'RegistrationDeleted',
          targetType: 'Participant',
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          targetId: ((reg as any).ParticipantEmail || '') + '#' + reg.Id,
          targetName: name,
          eventId: selectedEvent.id,
          eventTitle: selectedEvent.title,
          details: { note: 'Doppel-Anmeldung still entfernt (Duplikat). Person bleibt über die zweite Zeile angemeldet.' },
        });
      } catch (err) { console.warn('[DEX] writeChangeLog (dup delete) failed:', err); }
      try {
        const isSplit = typeof selectedEvent.durchstarterCapacity === 'number'
          && typeof selectedEvent.funstarterCapacity === 'number'
          && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
        await eventServiceRef.syncSeatsToActiveCount(selectedEvent.subsiteUrl, { isSplit });
      } catch { /* best-effort */ }
    } catch (err) {
      console.warn('[DEX] performSilentDuplicateDelete failed:', err);
    }
    const regs = await getAllRegistrations(selectedEvent.id);
    setRegistrations(regs);
    setAdminToast(null);
  };
  // v22.16: „Hinweise"-Box für aktive Events — Busy-State für den 1-Klick-
  // Sprach-Fix + Tick, damit „Ausblenden" (localStorage) sofort re-rendert.
  const [hintLangBusy, setHintLangBusy] = React.useState(false);
  const [hintsDismissTick, setHintsDismissTick] = React.useState(0);
  // v11.97/v11.98: bei Events mit Split-Kapazität (zwei Gruppen) wird die
  // Aktiv-Teilnehmer-Tabelle standardmäßig nach Gruppe getrennt angezeigt
  // (kleinere Gruppe zuerst). Per Toggle umschaltbar auf zusammengeführte
  // Sicht. Default: 'split'. Bei Events ohne Split-Kapazität ohne Wirkung.
  const [splitParticipantsView, setSplitParticipantsView] = React.useState<'split' | 'merged'>('split');
  // v14.11: subEventsOnlyMode-Konsolidierung. Wenn das selektierte Hauptevent
  // im „Nur Sub-Events"-Modus ist und Sub-Events hat, hat das Hauptevent
  // selbst keine direkten Teilnehmer — stattdessen werden alle Sub-Event-
  // Teilnehmer pro Person zu einer Zeile aggregiert (Matrix-View mit einer
  // X-Spalte pro Sub-Event). Hier halten wir die rohen Registrierungen je
  // Sub-Event.
  const [subEventRegsByEventId, setSubEventRegsByEventId] = React.useState<Record<string, SPRegistration[]>>({});
  const [isLoadingSubEventRegs, setIsLoadingSubEventRegs] = React.useState(false);
  // v22.59: manueller Reload-Trigger für die Sub-Event-Regs (z.B. nach dem
  // Löschen einer konsolidierten Abmeldung).
  const [subRegReloadTick, setSubRegReloadTick] = React.useState(0);
  const [expandedConsolidatedEmail, setExpandedConsolidatedEmail] = React.useState<string | null>(null);
  // v14.11: eigene Sort-States für den Matrix-View. `consolidatedSort` kann
  // 'id' | 'vorname' | 'nachname' | 'email' | 'jobTitle' | 'location' |
  // 'child:<eventId>' sein. Default: 'nachname' aufsteigend.
  // v15.23: Default-Sort im konsolidierten View jetzt chronologisch
  // nach erster Anmeldung (früheste zuerst), nicht mehr alphabetisch
  // nach Nachname. Damit ist die # in der Liste die Reihenfolge der
  // Anmeldung, nicht die Alphabet-Position.
  const [consolidatedSort, setConsolidatedSort] = React.useState<string>('id');
  const [consolidatedSortAsc, setConsolidatedSortAsc] = React.useState<boolean>(true);
  // v11.0: Bei Events mit Teilnehmer-Upload alle Attachment-Listen
  // einmalig laden, sobald sich registrations oder das ausgewählte
  // Event ändern. Damit zeigt der „Anhang"-Button in der Action-Spalte
  // sofort die korrekte Anzahl.
  React.useEffect(() => {
    // v19.0: Attachments auch laden, wenn das Event ein Dokument-Custom-Feld hat
    // (nicht nur beim generischen Attendee-Upload).
    const hasDocField = (selectedEvent?.eventSpecificFields || []).some(f => f.type === 'document');
    if (!selectedEvent || (!selectedEvent.allowAttendeeUpload && !hasDocField) || !eventServiceRef || !selectedEvent.subsiteUrl) {
      setAttachmentsByReg({});
      return;
    }
    const subsiteUrl = selectedEvent.subsiteUrl;
    const ids = registrations.map(r => r.Id).filter(Boolean);
    let cancelled = false;
    (async () => {
      const map: Record<number, Array<{ fileName: string; serverRelativeUrl: string }>> = {};
      for (const id of ids) {
        try {
          const list = await eventServiceRef!.listRegistrationAttachments(subsiteUrl, id);
          if (list.length > 0) map[id] = list;
        } catch { /* */ }
      }
      if (!cancelled) setAttachmentsByReg(map);
    })().catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, selectedEvent?.allowAttendeeUpload, registrations.length]);
  // v9.29: Header-Refresh-Button triggert ein globales Event — wir hooken uns ein.
  React.useEffect(() => {
    const onRefresh = (): void => { void handleRefresh(); };
    window.addEventListener('dex-refresh-page', onRefresh);
    return () => window.removeEventListener('dex-refresh-page', onRefresh);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent]);
  // v14.11: subEventsOnlyMode — alle Sub-Event-Anmeldungen einsammeln, um
  // den konsolidierten Matrix-View pro Person zu rendern. Nur aktiv, wenn
  // das selektierte Event tatsächlich Hauptevent ohne eigene Anmeldungen
  // (subEventsOnlyMode) ist und es Sub-Events gibt.
  React.useEffect(() => {
    if (!selectedEvent || !selectedEvent.subEventsOnlyMode) {
      setSubEventRegsByEventId({});
      return;
    }
    const children = childEventsOf(selectedEvent.id);
    if (children.length === 0) {
      setSubEventRegsByEventId({});
      return;
    }
    let cancelled = false;
    setIsLoadingSubEventRegs(true);
    (async () => {
      const map: Record<string, SPRegistration[]> = {};
      for (const ch of children) {
        try {
          const regs = await getAllRegistrations(ch.id);
          map[ch.id] = regs;
        } catch {
          map[ch.id] = [];
        }
      }
      if (!cancelled) {
        setSubEventRegsByEventId(map);
        setIsLoadingSubEventRegs(false);
      }
    })().catch(() => { if (!cancelled) setIsLoadingSubEventRegs(false); });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, selectedEvent?.subEventsOnlyMode, subRegReloadTick]);
  // v15.14: Wenn ein Sub-Event direkt selektiert wurde, laden wir die
  // Registrierungen des Parent-Events mit, damit die Pastel-A-Spalten
  // (Custom-Fields des Hauptevents) pro Teilnehmer-Zeile mit den
  // tatsächlichen Antworten aus der Parent-Registrierung gefüllt
  // werden können. Vorher waren diese Zellen leer („-"), weil die
  // Sub-Event-Registrierung diese Antworten nicht enthält.
  const [parentRegsByEmail, setParentRegsByEmail] = React.useState<Record<string, SPRegistration>>({});
  React.useEffect(() => {
    if (!selectedEvent || !selectedEvent.parentEventId) {
      setParentRegsByEmail({});
      return;
    }
    let cancelled = false;
    (async () => {
      try {
        const parentRegs = await getAllRegistrations(selectedEvent.parentEventId!);
        const map: Record<string, SPRegistration> = {};
        for (const r of parentRegs) {
          const key = (r.ParticipantEmail || '').toLowerCase().trim();
          if (key) map[key] = r;
        }
        if (!cancelled) setParentRegsByEmail(map);
      } catch {
        if (!cancelled) setParentRegsByEmail({});
      }
    })().catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, selectedEvent?.parentEventId]);
  const [isLoadingRegs, setIsLoadingRegs] = React.useState(false);
  const [regLoadError, setRegLoadError] = React.useState('');
  // v18.24: beim Event-/Tab-Wechsel die aktuelle Höhe der Detail-Card
  // „einfrieren", solange die Teilnehmer neu geladen werden — sonst klappt
  // die Card auf die „Lade..."-Zeile zusammen und springt danach wieder auf
  // (klein→groß-Flackern). null = keine Reservierung aktiv.
  const detailCardRef = React.useRef<HTMLDivElement>(null);
  const [reservedDetailHeight, setReservedDetailHeight] = React.useState<number | undefined>(undefined);
  const [deletingId, setDeletingId] = React.useState<string | null>(null);
  // v9.0: Danger-Zone-Modal — User muss den Event-Titel exakt (lowercase)
  // eintippen bevor der Lösch-Button aktiv wird. Schutz gegen versehentliche
  // Löschungen (früher: Click-to-Confirm-Pattern, war zu schwach).
  const [confirmDeleteEvent, setConfirmDeleteEvent] = React.useState<DeloitteEvent | null>(null);
  const [confirmDeleteText, setConfirmDeleteText] = React.useState('');
  // v9.0: ChangeLog-Modal — Admin/Organizer sehen den Audit-Log aller
  // Event- und Teilnehmer-Aenderungen (DEX_ChangeLog).
  const [showChangeLogModal, setShowChangeLogModal] = React.useState(false);
  const [changeLogEntries, setChangeLogEntries] = React.useState<Array<{
    Id: number; Created: string; Action: string; TargetType: string;
    TargetId: string; TargetName: string; EventId: string; EventTitle: string;
    ActorName: string; ActorEmail: string; Details: string;
  }>>([]);
  const [changeLogLoading, setChangeLogLoading] = React.useState(false);
  const [changeLogFilterAction, setChangeLogFilterAction] = React.useState('');
  const [changeLogFilterEvent, setChangeLogFilterEvent] = React.useState('');
  const [changeLogFilterActor, setChangeLogFilterActor] = React.useState('');
  // Self-Actions (User registriert/storniert sich selbst) sind Datenrauschen
  // im Audit-Log — Admins/Organizer wollen normalerweise nur Aktionen sehen,
  // die jemand AUF einen anderen User angewendet hat. Marker im Details-JSON:
  // `"asActor":"self"` bei selbst durchgeführten Registrierungen/Stornos.
  const [changeLogHideSelf, setChangeLogHideSelf] = React.useState(true);
  const openChangeLog = async (): Promise<void> => {
    if (!eventServiceRef) return;
    setShowChangeLogModal(true);
    setChangeLogLoading(true);
    try {
      const entries = await eventServiceRef.readChangeLog({ top: 500 });
      setChangeLogEntries(entries);
    } finally {
      setChangeLogLoading(false);
    }
  };
  const [isDeleting, setIsDeleting] = React.useState(false);
  const [copiedEmails, setCopiedEmails] = React.useState(false);
  const [copiedDeepLink, setCopiedDeepLink] = React.useState(false);
  const [isSendingQR, setIsSendingQR] = React.useState(false);
  const [qrSentCount, setQrSentCount] = React.useState(0);
  // v9.15: QR-Code-Versand-Modal mit Test-/Volldurchlauf. v20.7: der
  // Auto-Send-Toggle ist entfallen — Auto-Send ist immer aktiv.
  const [qrSendModalOpen, setQrSendModalOpen] = React.useState(false);
  const [qrSendResult, setQrSendResult] = React.useState<string | null>(null);
  // v9.37: Vorschau der QR-Code-Mail (analog zur Live-Vorschau im Event-Wizard
  // unter „Kommunikation"). Der Organizer sieht damit vorab genau die Mail, die
  // beim Versand rausgeht — inklusive echtem QR-Code für ihn selbst als Empfänger.
  const [qrPreviewOpen, setQrPreviewOpen] = React.useState(false);
  const [qrPreviewHtml, setQrPreviewHtml] = React.useState('');
  const [qrPreviewSubject, setQrPreviewSubject] = React.useState('');
  const [qrPreviewLoading, setQrPreviewLoading] = React.useState(false);
  // v22.18: QR-Mail-Text pro Event anpassbar (HtmlEditorModal mit Live-
  // Vorschau, gespeichert im Event → gilt auch für den Auto-Versand).
  const [qrEditOpen, setQrEditOpen] = React.useState(false);
  const [qrEditSubject, setQrEditSubject] = React.useState('');
  const [qrEditHeading, setQrEditHeading] = React.useState('');
  const [qrEditSubheading, setQrEditSubheading] = React.useState('');
  const [qrEditBody, setQrEditBody] = React.useState('');
  const [qrEditSaving, setQrEditSaving] = React.useState(false);
  const [qrEditSampleBlock, setQrEditSampleBlock] = React.useState('');
  const [searchQuery, setSearchQuery] = React.useState('');
  const [sortColumn, setSortColumn] = React.useState<'id' | 'anrede' | 'vorname' | 'nachname' | 'email' | 'status' | 'date'>('id');
  const [sortAsc, setSortAsc] = React.useState(true);
  // v17.8: Sortierung der Warteliste (analog Teilnehmerliste). Default 'pos'
  // = Reihenfolge nach TeilnehmerID asc (FIFO-Position der Warteliste).
  const [waitlistSortColumn, setWaitlistSortColumn] = React.useState<'pos' | 'vorname' | 'nachname' | 'email' | 'jobtitle' | 'location' | 'date'>('pos');
  const [waitlistSortAsc, setWaitlistSortAsc] = React.useState(true);
  // v18.11: Sortierung der Abmeldungs-Liste (gleiche Spalten wie Teilnehmer/Warteliste).
  const [cancelledSortColumn, setCancelledSortColumn] = React.useState<'vorname' | 'nachname' | 'email' | 'jobtitle' | 'location' | 'type' | 'date'>('date');
  const [cancelledSortAsc, setCancelledSortAsc] = React.useState(false);
  const [isReorderingIDs, setIsReorderingIDs] = React.useState(false);
  const [reorderResult, setReorderResult] = React.useState<string | null>(null);
  // v18.70: Manueller Nachrück-Button (freien Platz mit erstem Wartelistler füllen)
  const [isPromoting, setIsPromoting] = React.useState(false);
  const [promoteResult, setPromoteResult] = React.useState<string | null>(null);
  const [isResettingCounter, setIsResettingCounter] = React.useState(false);
  const [resetCounterResult, setResetCounterResult] = React.useState<string | null>(null);
  const [isFixingColumns, setIsFixingColumns] = React.useState(false);
  const [fixColumnsResult, setFixColumnsResult] = React.useState<string | null>(null);
  // v11.36: Überbuchungs-Bereinigung
  const [isDetectingOverbook, setIsDetectingOverbook] = React.useState(false);
  const [detectOverbookResult, setDetectOverbookResult] = React.useState<string | null>(null);
  // Modal-State für „Bestätigen" (einzeln oder Sammel) bzw. „Platz behalten".
  // mode: 'confirm' = auf Warteliste; 'keep' = Platz behalten.
  // targets: betroffene Registrierungen (1 = einzeln, n = Sammel „Alle").
  const [overbookModal, setOverbookModal] = React.useState<{
    mode: 'confirm' | 'keep';
    targets: SPRegistration[];
  } | null>(null);
  const [obWithMail, setObWithMail] = React.useState(true);
  const [obMailSubject, setObMailSubject] = React.useState('');
  const [obMailBody, setObMailBody] = React.useState('');
  const [obMailLang, setObMailLang] = React.useState<'DE' | 'EN'>('DE');
  const [obRemoveCalendar, setObRemoveCalendar] = React.useState(true);
  const [obKeepVariant, setObKeepVariant] = React.useState<'active' | 'firstWaitlist'>('firstWaitlist');
  const [obBusy, setObBusy] = React.useState(false);
  // v11.36: Fortschritts-Overlay für die ID-Neuvergabe (0..100 %, null = aus).
  const [reorderProgress, setReorderProgress] = React.useState<number | null>(null);
  const [reorderProgressLabel, setReorderProgressLabel] = React.useState('');
  const [isFixingFields, setIsFixingFields] = React.useState(false);
  const [fixFieldsResult, setFixFieldsResult] = React.useState<string | null>(null);
  // v11.84: Teams-Section (Admin Center Team-Management).
  const [teamsCollapsed, setTeamsCollapsed] = React.useState<boolean>(false);
  // Add-Member-Modal pro Team — gleiche Mechanik wie MyEventsPage.
  const [adminAddMemberDialog, setAdminAddMemberDialog] = React.useState<{
    teamId: string;
    teamName: string;
    freeSlots: number;
    /** v17.1: true wenn dieser Dialog im „Neues Team anlegen"-Flow geöffnet
     *  wurde — dann zeigen wir ein optionales Team-Name-Eingabefeld
     *  und uebernehmen den eingegebenen Namen beim Insert. */
    isNewTeam?: boolean;
  } | null>(null);
  const [adminAddMemberPick, setAdminAddMemberPick] = React.useState<{ email: string; displayName: string } | null>(null);
  const [adminAddMemberQuery, setAdminAddMemberQuery] = React.useState('');
  const [adminAddMemberResults, setAdminAddMemberResults] = React.useState<Array<{ email: string; displayName: string }>>([]);
  const [adminAddMemberSearching, setAdminAddMemberSearching] = React.useState(false);
  const [adminAddMemberConsent, setAdminAddMemberConsent] = React.useState(false);
  // v17.4: Multi-Select aus teamlosen Personen + Lead-Auswahl + Mail-Opt-in.
  const [adminAddTeamlessPicks, setAdminAddTeamlessPicks] = React.useState<Set<number>>(new Set());
  const [adminAddLeadRegId, setAdminAddLeadRegId] = React.useState<number | null>(null);
  const [adminAddSendMail, setAdminAddSendMail] = React.useState<boolean>(false);
  // v22.42: Organizer kann die Bestätigungs-/Info-Mail der Team-Zuordnung
  // optional als Kopie (CC) an sich selbst bekommen.
  const [adminAddCcOrganizer, setAdminAddCcOrganizer] = React.useState<boolean>(false);
  // v22.49: Kommunikation an die ÜBRIGEN Team-Mitglieder (optional) — Reichweite
  // alle vs. nur Lead. Plus: bei ganz neuer Person ist die Anmeldebestätigung
  // (+ Outlook) an die Person ebenfalls optional (Default an).
  const [adminAddNotifyOthers, setAdminAddNotifyOthers] = React.useState<boolean>(false);
  const [adminAddNotifyScope, setAdminAddNotifyScope] = React.useState<'all' | 'lead'>('all');
  const [adminAddNewPersonMail, setAdminAddNewPersonMail] = React.useState<boolean>(true);
  const [adminAddMemberBusy, setAdminAddMemberBusy] = React.useState(false);
  const [adminAddMemberError, setAdminAddMemberError] = React.useState('');
  const [adminAddMemberIncludeIntl, setAdminAddMemberIncludeIntl] = React.useState(false);
  const adminAddMemberQueryTimer = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  React.useEffect(() => {
    const q = adminAddMemberQuery.trim();
    if (q.length >= 2 && !adminAddMemberPick) {
      (async () => {
        setAdminAddMemberSearching(true);
        try {
          const res = await searchUsers(q, adminAddMemberIncludeIntl);
          setAdminAddMemberResults(res.map(r => ({ email: r.email, displayName: r.displayName })));
        } catch { setAdminAddMemberResults([]); }
        setAdminAddMemberSearching(false);
      })().catch(() => {});
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [adminAddMemberIncludeIntl]);
  // Lead-Transfer-Dropdown: pro Team ein offener Dropdown-Index (TeamId-Key).
  const [leadTransferOpenFor, setLeadTransferOpenFor] = React.useState<string | null>(null);
  // v22.45: Pro-Team „Anpassen"-Modus — erst dann erscheinen die
  // „Entfernen"-Buttons pro Mitglied (nicht dauerhaft an jedem Namen).
  const [teamEditOpenFor, setTeamEditOpenFor] = React.useState<string | null>(null);
  const [leadTransferBusy, setLeadTransferBusy] = React.useState(false);
  // v23.0: Drag&Drop-Zuordnung in der Teams-Sektion. dragRegId = gezogene
  // Registrierung, dragOverTid = aktuelles Drop-Ziel ('' = „ohne Team").
  const [dragRegId, setDragRegId] = React.useState<number | null>(null);
  const [dragOverTid, setDragOverTid] = React.useState<string | null>(null);
  // v23.0: Per-Team-Info-Mail (z.B. Teams-Einwahllink je Break-Out-Session).
  const [teamMailOpen, setTeamMailOpen] = React.useState(false);
  const [teamMailSubject, setTeamMailSubject] = React.useState('');
  const [teamMailBody, setTeamMailBody] = React.useState('');
  const [teamMailInfoByTid, setTeamMailInfoByTid] = React.useState<Record<string, string>>({});
  const [teamMailSending, setTeamMailSending] = React.useState(false);
  // Toast nach erfolgreicher Aktion in der Teams-Section.
  const [teamsToast, setTeamsToast] = React.useState<string>('');
  const [isRefreshingProfiles, setIsRefreshingProfiles] = React.useState(false);
  const [refreshProfilesResult, setRefreshProfilesResult] = React.useState<string | null>(null);
  // Globale Reparatur: Organizer-Email-Mismatch über alle Events fixen
  const [isRepairingOrganizers, setIsRepairingOrganizers] = React.useState(false);
  const [repairOrganizersResult, setRepairOrganizersResult] = React.useState<string | null>(null);
  // v20.6: Reparatur "Fremd-Anmeldungen: Zugriff" über alle aktiven Events —
  // prüft pro Teilnehmerliste die "nur eigene Elemente"-Sicherheit und setzt
  // bei Anmeldungen durch Dritte den Zeilen-Autor auf den Teilnehmer.
  const [isRepairingAccess, setIsRepairingAccess] = React.useState(false);
  const [repairAccessResult, setRepairAccessResult] = React.useState<string | null>(null);
  // v20.7: Fortschritts-Modal für die Zugriffs-Reparatur (Event i/N +
  // Eintrag x/y + Abschluss-Summary). running=false ⇒ Summary + Schließen.
  const [accessFixModal, setAccessFixModal] = React.useState<{
    running: boolean;
    evIdx: number; evTotal: number; evTitle: string;
    itemDone: number; itemTotal: number;
    summary: string[] | null;
  } | null>(null);
  // Email Compose Modal
  const [showEmailModal, setShowEmailModal] = React.useState(false);
  // v17.10: Massmail-Target-Picker. Erst Zielgruppe wählen, dann den
  // RichText-Editor oeffnen. Mode = 'closed' | 'pick' | 'paste' | 'editor'.
  const [massmailMode, setMassmailMode] = React.useState<'closed' | 'pick' | 'paste' | 'editor'>('closed');
  type MassmailAudience = 'active' | 'activePlusWait' | 'waitOnly' | 'nachruecker' | 'custom';
  const [massmailAudience, setMassmailAudience] = React.useState<MassmailAudience>('active');
  // v22.9: Eigene Status-Auswahl ('custom') — welche Status die Mail bekommen.
  const [massmailStatuses, setMassmailStatuses] = React.useState<Set<string>>(new Set(['Angemeldet', 'QR versendet', 'Eingecheckt']));
  // Für 'nachruecker': der eingefügte Rohtext + die nach Extraktion
  // verbleibenden Teilnehmer (= angemeldete Personen, die NICHT in der
  // eingefügten Liste stehen).
  const [massmailPasteRaw, setMassmailPasteRaw] = React.useState<string>('');
  const [emailSubject, setEmailSubject] = React.useState('');
  const [emailHeading, setEmailHeading] = React.useState('');
  const [emailBody, setEmailBody] = React.useState('');
  const [emailSending, setEmailSending] = React.useState(false);
  // v22.9: Massenmail-Entwurf pro Event speichern (wie Einladungsmail) +
  // Testmail an die Organizer.
  const [massmailDraftSaved, setMassmailDraftSaved] = React.useState(false);
  const massmailHydratingRef = React.useRef(false);
  const [massmailTesting, setMassmailTesting] = React.useState(false);
  const [massmailTestMsg, setMassmailTestMsg] = React.useState<string | null>(null);
  // v22.11: Unter-Überschrift der Massenmail editierbar (Parität zur
  // Einladungsmail; vorher war das Feld sichtbar, aber nicht angebunden).
  const [massmailSubheading, setMassmailSubheading] = React.useState('');
  // v11.40: Einladungsmail-Modal — Mail mit Anmelde-Link an Organizer (zum
  // Weiterleiten) oder direkt an den hinterlegten Mailverteiler des Events.
  const [showInviteModal, setShowInviteModal] = React.useState(false);
  const [inviteSubject, setInviteSubject] = React.useState('');
  const [inviteHeading, setInviteHeading] = React.useState('');
  const [inviteBody, setInviteBody] = React.useState('');
  const [inviteTarget, setInviteTarget] = React.useState<'organizer' | 'audience'>('organizer');
  const [inviteSending, setInviteSending] = React.useState(false);
  // v22.5: Unter-Überschrift der Einladungsmail (vorher nicht erfasst) + Entwurf-
  // Speicherung pro Event in localStorage, damit ein angefangener Text beim
  // Schließen + erneuten Öffnen erhalten bleibt.
  const [inviteSubheading, setInviteSubheading] = React.useState('');
  // verhindert, dass das Auto-Speichern den gerade geladenen Entwurf sofort
  // wieder überschreibt, bevor der State gesetzt ist.
  const inviteHydratingRef = React.useRef(false);
  // v22.5: kurzes „Gespeichert"-Feedback nach Klick auf den Speichern-Button.
  const [inviteDraftSaved, setInviteDraftSaved] = React.useState(false);
  const [showExportMenu, setShowExportMenu] = React.useState(false);
  // v17.12: Zielgruppen-Picker für Excel-Export.
  const [excelTargetModal, setExcelTargetModal] = React.useState<null | { mode: 'deloitte' | 'b2run' }>(null);
  const [excelAudience, setExcelAudience] = React.useState<'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled'>('active');
  // v20.4: Excel-Export im Klammer-Modus — konsolidierte Matrix (eine Zeile
  // pro Person, Spalten pro Sub-Event) und/oder einzelne Sub-Event-Blätter
  // sind im Export-Modal wählbar.
  const [excelIncludeMatrix, setExcelIncludeMatrix] = React.useState(true);
  const [excelSubIds, setExcelSubIds] = React.useState<Set<string>>(new Set());
  // Outlook-Decline-Check (Admin only): zeigt Teilnehmer, die in Outlook
  // abgesagt haben, aber in der Teilnehmerliste noch aktiv gelistet sind.
  const [isCheckingDeclines, setIsCheckingDeclines] = React.useState(false);
  const [declineResult, setDeclineResult] = React.useState<{
    declinedAndRegistered: Array<{ email: string; name: string; reg: SPRegistration }>;
    declinedTotal: number;
    error: string | null;
  } | null>(null);
  const [showDeclineModal, setShowDeclineModal] = React.useState(false);
  const [declineCopied, setDeclineCopied] = React.useState(false);

  // Admin-Toast für Abmelde-/Nachrück-Feedback (seit v6.8):
  //  - 'cancelling': während die Abmeldung + Nachrück-Suche läuft (orange, Spinner)
  //  - 'promoted'  : erfolgreicher Nachrücker mit Namen + Typ (grün)
  //  - 'no-promote': Abmeldung ok, aber keiner auf der Warteliste (grau)
  type AdminToast =
    | { kind: 'cancelling'; name: string }
    | { kind: 'promoted'; name: string; email: string; type?: string }
    | { kind: 'no-promote'; name: string };
  const [adminToast, setAdminToast] = React.useState<AdminToast | null>(null);

  // v8.0: In-App-Edit-Modal für Teilnehmer (Admin/Organizer kann jeden
  // Teilnehmer-Eintrag direkt aus der Liste editieren — Anrede, Name, Email,
  // Phone, Department, Location, JobTitle, Status, plus alle Custom-Felder).
  // Beim Save wird eine Audit-Zeile in ChangeLog geschrieben (wer/wann/was)
  // und LastModifiedDate gesetzt — kein direkter SP-Edit mehr nötig, was
  // gleichzeitig das deutsche Datumsformat-Problem in SP umgeht.
  const [editingReg, setEditingReg] = React.useState<SPRegistration | null>(null);
  const [editForm, setEditForm] = React.useState<Record<string, string>>({});
  const [isSavingEdit, setIsSavingEdit] = React.useState(false);
  const [editError, setEditError] = React.useState('');
  // v11.0: Attachment-Modal pro Teilnehmer-Reihe — nur wenn das Event
  // den Teilnehmer-Upload erlaubt (event.allowAttendeeUpload). Modal
  // listet alle Item-Attachments der jeweiligen Reg-Zeile + bietet
  // Download-Link + Lösch-Button (Admin/Organizer kann auch fremde
  // Uploads löschen). Map(regId → attachments) wird beim Laden der
  // Teilnehmerliste einmalig befüllt, damit der Anhang-Button die
  // Anzahl direkt anzeigen kann.
  const [attachmentsByReg, setAttachmentsByReg] = React.useState<Record<number, Array<{ fileName: string; serverRelativeUrl: string }>>>({});
  const [attachmentsModalReg, setAttachmentsModalReg] = React.useState<SPRegistration | null>(null);
  const [attachmentsBusy, setAttachmentsBusy] = React.useState(false);
  // v19.30 — Feature A: Im konsolidierten View (Hauptevent mit Sub-Events) die
  // Custom-Felder des Hauptevents („Felder des Hauptevents") pro Teilnehmer
  // editierbar machen. Die Antworten leben in der Registrierung der Person auf
  // der Hauptevent-Subsite (`selectedEvent.subsiteUrl`). `mainFieldsEditReg`
  // hält diese Parent-Registrierung; `mainFieldsEditForm` die editierten Werte
  // (Field-ID → String). `mainFieldsEditName` nur zur Anzeige im Modal-Titel.
  const [mainFieldsEditReg, setMainFieldsEditReg] = React.useState<SPRegistration | null>(null);
  const [mainFieldsEditName, setMainFieldsEditName] = React.useState('');
  const [mainFieldsEditForm, setMainFieldsEditForm] = React.useState<Record<string, string>>({});
  const [mainFieldsEditSaving, setMainFieldsEditSaving] = React.useState(false);
  const [mainFieldsEditError, setMainFieldsEditError] = React.useState('');
  // v19.30 — Feature B: Abmeldung eines Teilnehmers aus dem konsolidierten
  // View mit Sub-Event-Auswahl. Der Modal listet alle Sub-Events, für die die
  // Person aktiv angemeldet ist (Status angemeldet/QR/eingecheckt/Warteliste),
  // je mit Checkbox. `deregModal` hält die betroffene Person + die abmeldbaren
  // Sub-Event-Registrierungen; `deregSelected` die angehakten Sub-Event-IDs.
  const [deregModal, setDeregModal] = React.useState<{
    emailKey: string;
    name: string;
    email: string;
    items: Array<{ child: DeloitteEvent; reg: SPRegistration }>;
  } | null>(null);
  const [deregSelected, setDeregSelected] = React.useState<Set<string>>(new Set());
  const [deregBusy, setDeregBusy] = React.useState(false);
  const openEditModal = (reg: SPRegistration): void => {
    setEditError('');
    setEditingReg(reg);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const r = reg as any;
    const initial: Record<string, string> = {
      Anrede: r.Anrede || '',
      Vorname: r.Vorname || '',
      Nachname: r.Nachname || '',
      ParticipantEmail: r.ParticipantEmail || '',
      Phone: r.Phone || '',
      Department: r.Department || '',
      Location: r.Location || '',
      JobTitle: r.JobTitle || '',
      Status: r.Status || '',
      // v10.13+: B2Run-Felder ins Edit-Form aufnehmen damit das B2Run-Modul
      // im Edit-Modal die aktuellen Werte vorbefüllen kann. Strings (auch
      // wenn leer) — bei Nicht-B2Run-Events werden die Felder im Modal
      // sowieso nicht angezeigt.
      StarterType: r.StarterType || '',
      PreferredStarterType: r.PreferredStarterType || '',
    };
    // Custom-Field-Werte aus dem reg laden (sie sind als SP-Spalten gespeichert)
    if (selectedEvent?.eventSpecificFields) {
      for (const f of selectedEvent.eventSpecificFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = (f as any).spInternalName || '';
        if (sp) initial[sp] = (r[sp] !== undefined && r[sp] !== null) ? String(r[sp]) : '';
      }
    }
    setEditForm(initial);
  };
  const closeEditModal = (): void => {
    setEditingReg(null);
    setEditForm({});
    setEditError('');
  };
  const saveEdit = async (): Promise<void> => {
    if (!editingReg || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setIsSavingEdit(true);
    setEditError('');
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const r = editingReg as any;

      // Stamm-Daten (Vorname, Nachname, E-Mail) werden seit v9.7 ebenfalls
      // editierbar gemacht — z.B. um Tippfehler nach manueller Anlage zu
      // korrigieren. Validierung:
      //   1. E-Mail muss eine Deloitte-Deutschland-Adresse sein (@deloitte.de).
      //      Die Plattform ist nur für DEALL freigeschaltet — auch @deloitte.com
      //      (US/Global) zählt als extern. Sonst Abbruch mit Fehler.
      //   2. Person muss in M365 existieren (searchUserByEmail). Sonst
      //      Abbruch mit "Tippfehler"-Hinweis.
      // Die uebrigen Profil-Felder (Phone, Department, Location, JobTitle)
      // bleiben read-only — sie kommen aus dem M365-Profil.
      const oldVorname = String(r.Vorname || '');
      const oldNachname = String(r.Nachname || '');
      const oldEmail = String(r.ParticipantEmail || '');
      const newVorname = (editForm.Vorname || '').trim();
      const newNachname = (editForm.Nachname || '').trim();
      const newEmail = (editForm.ParticipantEmail || '').trim();
      const stammChanged = newVorname !== oldVorname || newNachname !== oldNachname || newEmail !== oldEmail;

      const profileFields: { Department?: string; Location?: string; JobTitle?: string } = {};
      if (stammChanged) {
        // Plausibilität: nicht-leer
        if (!newVorname || !newNachname || !newEmail) {
          setEditError(isDe
            ? 'Vorname, Nachname und E-Mail dürfen nicht leer sein.'
            : 'First name, last name and email must not be empty.');
          return;
        }
        // Domain-Check: nur Deloitte-Adressen zulassen
        const lower = newEmail.toLowerCase();
        const isDeloitte = /@(.*\.)?deloitte\.de$/.test(lower);
        if (!isDeloitte) {
          setEditError(isDe
            ? `Externe E-Mail-Adresse — nicht erlaubt. Die Plattform ist nur für Deloitte Deutschland (@deloitte.de) freigeschaltet.`
            : `External email address — not allowed. The platform is only available for Deloitte Germany (@deloitte.de).`);
          return;
        }
        // Existenz-Check via M365-Profile (UPN!=SMTP-aware). Wenn wir hier
        // nichts finden, ist es entweder ein Tippfehler oder ein Account
        // der gar nicht (mehr) im Tenant ist — beides nicht akzeptabel.
        if (newEmail.toLowerCase() !== oldEmail.toLowerCase()) {
          const profile = await searchUser(newEmail);
          if (!profile || !profile.displayName) {
            setEditError(isDe
              ? `Person mit E-Mail "${newEmail}" wurde im Deloitte-Tenant nicht gefunden. Bitte Adresse prüfen (Tippfehler?).`
              : `No person found in the Deloitte tenant for "${newEmail}". Please check the address (typo?).`);
            return;
          }
          // Profil-Daten gleich mit-uebernehmen, damit der Eintrag konsistent
          // bleibt (Department / Location / JobTitle passen zum neuen User).
          profileFields.Department = ''; // searchUser liefert displayName/location/jobTitle
          profileFields.Location = profile.location || '';
          profileFields.JobTitle = profile.jobTitle || '';
        }
      }

      const oldValues: Record<string, unknown> = {};
      const patch: Record<string, unknown> = {};
      const fieldLabelMap: Record<string, string> = {};

      if (stammChanged) {
        if (newVorname !== oldVorname) {
          oldValues.Vorname = oldVorname; patch.Vorname = newVorname; fieldLabelMap.Vorname = isDe ? 'Vorname' : 'First name';
        }
        if (newNachname !== oldNachname) {
          oldValues.Nachname = oldNachname; patch.Nachname = newNachname; fieldLabelMap.Nachname = isDe ? 'Nachname' : 'Last name';
        }
        if (newEmail !== oldEmail) {
          oldValues.ParticipantEmail = oldEmail; patch.ParticipantEmail = newEmail; fieldLabelMap.ParticipantEmail = 'E-Mail';
          // Profil-Daten mit aktualisieren (nur wenn ueberhaupt was zurückkam)
          if (profileFields.Location) {
            oldValues.Location = String(r.Location || ''); patch.Location = profileFields.Location;
            fieldLabelMap.Location = isDe ? 'Standort' : 'Location';
          }
          if (profileFields.JobTitle) {
            oldValues.JobTitle = String(r.JobTitle || ''); patch.JobTitle = profileFields.JobTitle;
            fieldLabelMap.JobTitle = 'Job Title';
          }
        }
      }

      // Custom-Felder des Events. v10.15+: nur Felder ins Patch aufnehmen die
      // sich tatsächlich geändert haben — sonst sendet ein unverändertes
      // Choice-Feld ohne ausgewählten Wert einen leeren String an SP, der
      // mit HTTP 400 'Invalid choice' kippt und das ganze Update abbricht.
      if (selectedEvent?.eventSpecificFields) {
        for (const f of selectedEvent.eventSpecificFields) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const sp = (f as any).spInternalName || '';
          if (!sp) continue;
          const oldVal = r[sp] !== undefined && r[sp] !== null ? String(r[sp]) : '';
          const newVal = editForm[sp] || '';
          if (newVal === oldVal) continue;  // unverändert → skip
          fieldLabelMap[sp] = f.label;
          oldValues[sp] = oldVal;
          patch[sp] = newVal;
        }
      }

      // v10.13+: B2Run-Felder explizit ins Patch aufnehmen, wenn sich was
      // geändert hat. Sind keine regulären customFields, daher werden sie
      // oben in der eventSpecificFields-Loop nicht abgeholt. Nur bei
      // Split-Capacity-Events relevant.
      const isSplitEvent = !!selectedEvent
        && (selectedEvent.durchstarterCapacity || 0) > 0
        && (selectedEvent.funstarterCapacity || 0) > 0;
      if (isSplitEvent) {
        const oldStarter = String(r.StarterType || '');
        const newStarter = (editForm.StarterType || '').trim();
        if (newStarter !== oldStarter) {
          oldValues.StarterType = oldStarter;
          patch.StarterType = newStarter;
          fieldLabelMap.StarterType = isDe ? 'Starter-Typ' : 'Starter type';
        }
        const oldPref = String(r.PreferredStarterType || '');
        const newPref = (editForm.PreferredStarterType || '').trim();
        if (newPref !== oldPref) {
          oldValues.PreferredStarterType = oldPref;
          patch.PreferredStarterType = newPref;
          fieldLabelMap.PreferredStarterType = isDe ? 'Wunsch-Starter-Typ' : 'Preferred starter type';
        }
      }
      if (Object.keys(patch).length === 0) {
        // Keine Aenderung — nichts zu tun.
        closeEditModal();
        return;
      }
      const actor = {
        name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email,
        email: currentUser.email,
      };
      const ok = await eventServiceRef.adminUpdateRegistration(
        selectedEvent.subsiteUrl, editingReg.Id, patch, actor, oldValues, fieldLabelMap
      );
      if (!ok) {
        // Häufigste 400-Ursache: eine SP-Spalte aus dem Patch existiert nicht
        // auf dieser Teilnehmerliste (z.B. StarterType auf einem v9-Event ohne
        // B2Run-Schema, oder ein neu hinzugefügtes Custom-Field ohne 'Spalten
        // fixen'-Run). Hilfreicher Hinweis auf den Repair-Button.
        setEditError(isDe
          ? 'Speichern fehlgeschlagen — vermutlich fehlt eine SP-Spalte in der Teilnehmerliste. Klicke einmal „Spalten fixen" im Toolbox-Bereich des Events, dann erneut versuchen.'
          : 'Save failed — likely a missing SP column on the participant list. Click „Fix columns" in the event toolbox once, then retry.');
        return;
      }
      // v9.0: Audit-Log mit Diff der geänderten Felder
      try {
        const changes: Record<string, { old: unknown; new: unknown }> = {};
        for (const k of Object.keys(patch)) {
          if (oldValues[k] !== patch[k]) changes[k] = { old: oldValues[k], new: patch[k] };
        }
        await eventServiceRef.writeChangeLog({
          action: 'ParticipantUpdated',
          targetType: 'Participant',
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          targetId: ((editingReg as any).ParticipantEmail || '') + '#' + editingReg.Id,
          targetName: `${editingReg.Vorname || ''} ${editingReg.Nachname || ''}`.trim(),
          eventId: selectedEvent.id,
          eventTitle: selectedEvent.title,
          details: { changes },
        });
      } catch { /* */ }
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
      closeEditModal();
    } catch (err) {
      console.warn('[DEX] saveEdit error:', err);
      setEditError(isDe
        ? 'Unerwarteter Fehler beim Speichern.'
        : 'Unexpected error while saving.');
    } finally {
      setIsSavingEdit(false);
    }
  };

  // v19.30 — Feature A: Edit-Modal für die Hauptevent-Custom-Felder einer
  // konsolidierten Zeile öffnen. Die Antworten stehen in der Registrierung der
  // Person auf der Hauptevent-Subsite. Wir suchen sie per E-Mail in
  // `registrations` (das ist die Teilnehmerliste des selektierten Hauptevents).
  const openMainFieldsEdit = (emailKey: string, displayName: string): void => {
    if (!selectedEvent) return;
    setMainFieldsEditError('');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parentReg = registrations.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === emailKey) || null;
    setMainFieldsEditReg(parentReg);
    setMainFieldsEditName(displayName);
    const initial: Record<string, string> = {};
    if (parentReg) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyReg = parentReg as any;
      let cd: Record<string, unknown> = {};
      if (anyReg.CustomData) { try { cd = JSON.parse(anyReg.CustomData); } catch { /* */ } }
      const parentFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'document' && f.label && f.label.trim());
      for (const f of parentFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = (f as any).spInternalName || '';
        let v: unknown = sp ? anyReg[sp] : undefined;
        if (v === undefined || v === null || v === '') v = cd[f.id];
        initial[f.id] = (v === undefined || v === null) ? '' : String(v);
      }
    }
    setMainFieldsEditForm(initial);
  };
  const closeMainFieldsEdit = (): void => {
    setMainFieldsEditReg(null);
    setMainFieldsEditName('');
    setMainFieldsEditForm({});
    setMainFieldsEditError('');
  };
  // v19.30 — Feature A: Speichern der Hauptevent-Custom-Felder. Persistiert
  // über dasselbe `adminUpdateRegistration` wie das reguläre Teilnehmer-Edit
  // (schreibt die SP-Spalten der Hauptevent-Teilnehmerliste) und legt eine
  // Audit-Zeile 'ParticipantUpdated' mit dem Vorher/Nachher-Diff an. Es werden
  // nur geänderte Felder ins Patch aufgenommen — sonst kippt ein unverändertes
  // Choice-Feld den ganzen Save (HTTP 400 'Invalid choice').
  const saveMainFieldsEdit = async (): Promise<void> => {
    if (!mainFieldsEditReg || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setMainFieldsEditSaving(true);
    setMainFieldsEditError('');
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyReg = mainFieldsEditReg as any;
      let cd: Record<string, unknown> = {};
      if (anyReg.CustomData) { try { cd = JSON.parse(anyReg.CustomData); } catch { /* */ } }
      const parentFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'document' && f.label && f.label.trim());
      const patch: Record<string, unknown> = {};
      const oldValues: Record<string, unknown> = {};
      const fieldLabelMap: Record<string, string> = {};
      const nextCd: Record<string, unknown> = { ...cd };
      let cdChanged = false;
      for (const f of parentFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = (f as any).spInternalName || '';
        const newVal = mainFieldsEditForm[f.id] || '';
        let oldVal = '';
        let oldFromSp: unknown = sp ? anyReg[sp] : undefined;
        if (oldFromSp === undefined || oldFromSp === null || oldFromSp === '') oldFromSp = cd[f.id];
        if (oldFromSp !== undefined && oldFromSp !== null) oldVal = String(oldFromSp);
        if (newVal === oldVal) continue; // unverändert → überspringen
        fieldLabelMap[sp || f.id] = f.label;
        oldValues[sp || f.id] = oldVal;
        // SP-Spalte patchen, falls vorhanden; sonst nur CustomData mitschreiben.
        if (sp) patch[sp] = newVal;
        nextCd[f.id] = newVal;
        cdChanged = true;
      }
      if (!cdChanged && Object.keys(patch).length === 0) {
        closeMainFieldsEdit();
        return;
      }
      // CustomData immer mitschreiben, damit der konsolidierte View (der bei
      // fehlender SP-Spalte auf CustomData zurückfällt) konsistent bleibt.
      if (cdChanged) patch['CustomData'] = JSON.stringify(nextCd);
      const actor = {
        name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email,
        email: currentUser.email,
      };
      const ok = await eventServiceRef.adminUpdateRegistration(
        selectedEvent.subsiteUrl, mainFieldsEditReg.Id, patch, actor, oldValues, fieldLabelMap
      );
      if (!ok) {
        setMainFieldsEditError(isDe
          ? 'Speichern fehlgeschlagen — vermutlich fehlt eine SP-Spalte in der Hauptevent-Teilnehmerliste. Klicke einmal „Spalten fixen" für das Hauptevent, dann erneut versuchen.'
          : 'Save failed — likely a missing SP column on the main-event participant list. Click „Fix columns" for the main event once, then retry.');
        return;
      }
      // Audit-Log mit Diff der geänderten Felder (analog saveEdit).
      try {
        const changes: Record<string, { old: unknown; new: unknown }> = {};
        for (const k of Object.keys(oldValues)) {
          changes[k] = { old: oldValues[k], new: (k in patch ? patch[k] : mainFieldsEditForm[k]) };
        }
        await eventServiceRef.writeChangeLog({
          action: 'ParticipantUpdated',
          targetType: 'Participant',
          targetId: (mainFieldsEditReg.ParticipantEmail || '') + '#' + mainFieldsEditReg.Id,
          targetName: `${mainFieldsEditReg.Vorname || ''} ${mainFieldsEditReg.Nachname || ''}`.trim() || mainFieldsEditName,
          eventId: selectedEvent.id,
          eventTitle: selectedEvent.title,
          details: { changes, scope: 'mainEventFields' },
        });
      } catch { /* */ }
      // Hauptevent-Teilnehmerliste neu laden, damit die Pastel-A-Spalten
      // (Hauptevent-Felder) die neuen Werte zeigen.
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
      closeMainFieldsEdit();
    } catch (err) {
      console.warn('[DEX] saveMainFieldsEdit error:', err);
      setMainFieldsEditError(isDe
        ? 'Unerwarteter Fehler beim Speichern.'
        : 'Unexpected error while saving.');
    } finally {
      setMainFieldsEditSaving(false);
    }
  };

  // v19.30 — Feature B: Sub-Event-Registrierungen neu laden (nach einer
  // Abmeldung im konsolidierten View). Spiegelt den Lade-Effekt von oben.
  const reloadSubEventRegs = async (): Promise<void> => {
    if (!selectedEvent || !selectedEvent.subEventsOnlyMode) return;
    const children = childEventsOf(selectedEvent.id);
    const map: Record<string, SPRegistration[]> = {};
    for (const ch of children) {
      try { map[ch.id] = await getAllRegistrations(ch.id); }
      catch { map[ch.id] = []; }
    }
    setSubEventRegsByEventId(map);
  };
  // v19.30 — Feature B: Abmelde-Modal für eine konsolidierte Zeile öffnen.
  // Sammelt alle Sub-Events, in denen die Person aktiv angemeldet ist.
  const openDeregModal = (row: ConsolidatedRow): void => {
    if (!selectedEvent) return;
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    const items: Array<{ child: DeloitteEvent; reg: SPRegistration }> = [];
    for (const ch of childEventsOf(selectedEvent.id)) {
      const r = row.perChild[ch.id];
      if (r && ACTIVE.indexOf(r.Status) >= 0) items.push({ child: ch, reg: r });
    }
    setDeregModal({
      emailKey: row.emailKey,
      name: `${row.vorname} ${row.nachname}`.trim() || row.email,
      email: row.email,
      items,
    });
    // Default: alle Sub-Events vorausgewählt — der häufigste Fall ist „ganz
    // abmelden". Der Organizer kann einzelne wieder abwählen.
    setDeregSelected(new Set(items.map(i => i.child.id)));
  };
  const closeDeregModal = (): void => {
    setDeregModal(null);
    setDeregSelected(new Set());
  };
  // v19.30 — Feature B: Abmeldung pro gewähltem Sub-Event durchführen. Spiegelt
  // exakt die Nebenwirkungen des Einzel-Event-Abmeldens (Abmelde-Mail +
  // Outlook 'Ausladen' + DEX_Participants-Cleanup + Nachrücken + ID-Reorder)
  // und schreibt zusätzlich pro Abmeldung eine 'RegistrationCancelled'-
  // Audit-Zeile (die der Einzel-Pfad nicht setzt).
  const runDeregModal = async (): Promise<void> => {
    if (!deregModal || !eventServiceRef) return;
    setDeregBusy(true);
    const actorName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
    const actorEmail = currentUser.email;
    const chosen = deregModal.items.filter(i => deregSelected.has(i.child.id));
    for (const { child, reg } of chosen) {
      const sub = child.subsiteUrl;
      if (!sub) continue;
      const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
      const cancelledStarterType = reg.StarterType || '';
      try {
        await eventServiceRef.cancelRegistration(sub, reg.Id, actorName, actorEmail);
        // Audit-Zeile (Feature D: Abmeldungen sollen im Event-Log auftauchen).
        try {
          await eventServiceRef.writeChangeLog({
            action: 'RegistrationCancelled',
            targetType: 'Participant',
            targetId: (reg.ParticipantEmail || '') + '#' + reg.Id,
            targetName: name,
            eventId: child.id,
            eventTitle: child.title,
            details: { asActor: 'organizer', via: 'consolidatedDeregister' },
          });
        } catch { /* */ }
        // Abmelde-Mail + Outlook 'Ausladen' (event-weite Schalter respektieren).
        // v22.22: Vergangenes Sub-Event → stille Abmeldung (keine Mail, kein
        // Outlook, kein Nachrücken, kein ID-Reorder).
        const childWasOver = isEventOver(child);
        if (reg.ParticipantEmail && !childWasOver) {
          if (!child.disableEmails && !child.disableCancellationEmail) {
            try {
              const emailData = cancellationEmail(name, child.title);
              await eventServiceRef.queueEmail(
                emailData.subject, reg.ParticipantEmail, name, emailData.body,
                'Abmeldung', child.title, child.id
              );
            } catch (err) { console.warn('[DEX]', err); }
          }
          if (!child.disableOutlook) {
            try {
              await eventServiceRef.queueOutlookEvent(
                reg.ParticipantEmail, child.id, child.title, 'Ausladen'
              );
            } catch (err) { console.warn('[DEX]', err); }
          }
        }
        // DEX_Participants aufräumen.
        if (reg.ParticipantEmail && child.eventNumber) {
          eventServiceRef.removeParticipantEvent(reg.ParticipantEmail, child.eventNumber)
            .catch(err => console.warn('[DEX]', err));
        }
        // Client-seitiges Nachrücken (typ-bewusst bei Split-Capacity, außer
        // splitSharedWaitlist) — identisch zum Einzel-Event-Abmelden.
        const isSplitEvent = typeof child.durchstarterCapacity === 'number'
          && typeof child.funstarterCapacity === 'number'
          && ((child.durchstarterCapacity || 0) > 0 || (child.funstarterCapacity || 0) > 0);
        const useTypeFilter = isSplitEvent && !child.splitSharedWaitlist;
        if (!childWasOver) {
        try {
          const promoted = await eventServiceRef.promoteFirstWaitlistItem(
            sub,
            cancelledStarterType || undefined,
            child.maxParticipants,
            (useTypeFilter && cancelledStarterType) ? cancelledStarterType : undefined,
            { itemId: reg.Id, participantEmail: reg.ParticipantEmail || '' },
          );
          if (promoted && promoted.success && promoted.email) {
            if (!child.disableEmails) {
              try {
                const lang = child.emailLanguage || 'EN';
                const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
                const promoteVars = {
                  Name: promotedFirstName,
                  EventTitle: child.title,
                  Organizer: formatOrganizerList(child.organizers, lang),
                  AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
                  WaitlistPosition: '',
                };
                let emailData: { subject: string; body: string };
                const spTplRaw = await eventServiceRef.getEmailTemplate('Nachruecken', lang).catch(() => null);
                const spTpl = applyEventTemplateOverride(spTplRaw, child.emailTemplateOverrides, 'Nachruecken');
                if (spTpl) emailData = buildEmailFromTemplate(spTpl, promoteVars);
                else emailData = promotionEmail(promotedFirstName, child.title);
                await eventServiceRef.queueEmail(
                  emailData.subject, promoted.email, promoted.name || '', emailData.body,
                  'Nachruecken', child.title, child.id
                );
              } catch (err) { console.warn('[DEX] promote-email failed:', err); }
            }
            if (!child.disableOutlook) {
              try {
                await eventServiceRef.queueOutlookEvent(promoted.email, child.id, child.title, 'Einladen');
              } catch (err) { console.warn('[DEX] promote-outlook failed:', err); }
            }
          }
        } catch (err) { console.warn('[DEX] promoteFirstWaitlistItem failed:', err); }
        // ID-Reorder in die Queue (Flow macht nur noch Reorder).
        try {
          await eventServiceRef.queueIDReorder(
            child.id, child.eventNumber || 0, sub, child.title, name, reg.ParticipantEmail || undefined
          );
        } catch (err) { console.warn('[DEX] queueIDReorder threw:', err); }
        }
      } catch (err) {
        console.warn('[DEX] consolidated deregister failed for child', child.id, err);
      }
    }
    try { await reloadSubEventRegs(); } catch { /* */ }
    setDeregBusy(false);
    closeDeregModal();
  };

  // v19.30 — Feature D: Audit-Log vorgefiltert auf das aktuell selektierte
  // Event öffnen (setzt den Event-/Ziel-Filter auf den Event-Titel).
  const openChangeLogForEvent = (): void => {
    setChangeLogFilterEvent(selectedEvent?.title || '');
    setChangeLogFilterAction('');
    setChangeLogFilterActor('');
    void openChangeLog();
  };

  // v11.36: Fairer Wartelisten-Rang einer Person in ihrer Gruppe — gleiche
  // Logik wie die Review-Box. Genutzt für die "neue Warteliste-Position" im
  // Mailtext (Vorschlag + Sammel-Versand).
  const getFairWaitlistRank = (reg: SPRegistration): number => {
    if (!selectedEvent) return 0;
    const ACT = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const isSplit = isSplitCapacity;
    const keyOf = (r: SPRegistration): string => isSplit ? (r.StarterType || r.PreferredStarterType || '?') : 'all';
    const capOf = (k: string): number => !isSplit
      ? (selectedEvent.maxParticipants || 0)
      : (k === 'Durchstarter' ? (selectedEvent.durchstarterCapacity || 0) : k === 'Funstarter' ? (selectedEvent.funstarterCapacity || 0) : 0);
    const k = keyOf(reg);
    const activeSorted = registrations
      .filter(r => ACT.indexOf(r.Status) >= 0 && keyOf(r) === k)
      .slice().sort((a, b) => a.Id - b.Id);
    const cap = capOf(k);
    const overCap = cap > 0 ? activeSorted.slice(cap) : [];
    const existingWl = registrations.filter(r => r.Status === 'Warteliste' && keyOf(r) === k);
    const fairWl = [...overCap, ...existingWl].sort((a, b) =>
      new Date(a.RegistrationDate).getTime() - new Date(b.RegistrationDate).getTime());
    const idx = fairWl.findIndex(x => x.Id === reg.Id);
    return idx >= 0 ? idx + 1 : 0;
  };

  // v11.36: Beim Öffnen des „Bestätigen"-Dialogs die Mail-Sprache aus dem
  // Event vorbelegen (Default DE wenn nicht explizit EN) — umschaltbar.
  React.useEffect(() => {
    if (overbookModal?.mode === 'confirm') {
      setObMailLang((selectedEvent?.emailLanguage || '').toUpperCase() === 'EN' ? 'EN' : 'DE');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [overbookModal]);

  // v11.36: Mailtext vorbefüllen — reagiert auf Dialog-Öffnen UND Sprachwahl.
  // Enthält die neue Wartelisten-Position der ersten Zielperson.
  // v13.0: buildOverbookApologyEmail ist jetzt async (lädt Template aus
  // DEX_EmailTemplates). Effect wartet auf das Promise und setzt State
  // wenn der Modal noch offen ist.
  React.useEffect(() => {
    if (overbookModal?.mode === 'confirm' && eventServiceRef && selectedEvent) {
      const t = overbookModal.targets[0];
      const nm = t ? ((t.Vorname && t.Nachname) ? `${t.Vorname} ${t.Nachname}` : t.ParticipantName) : '';
      const pos = t ? getFairWaitlistRank(t) : 0;
      let cancelled = false;
      eventServiceRef.buildOverbookApologyEmail(nm, selectedEvent.title, obMailLang, pos).then(m => {
        if (cancelled) return;
        setObMailSubject(m.subject);
        setObMailBody(m.body);
      }).catch(() => { /* */ });
      return () => { cancelled = true; };
    }
    return undefined;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [overbookModal, obMailLang]);

  // v11.36: Überbuchungs-Entscheidung ausführen (einzeln oder Sammel) und
  // danach IDs neu vergeben + Counter/Seat-Sync + Liste neu laden.
  const runOverbookResolution = async (): Promise<void> => {
    if (!overbookModal || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setObBusy(true);
    const sub = selectedEvent.subsiteUrl;
    const isBulk = overbookModal.targets.length > 1;
    try {
      for (const reg of overbookModal.targets) {
        const grp = reg.StarterType || reg.PreferredStarterType || '';
        const nm = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
        if (overbookModal.mode === 'confirm') {
          await eventServiceRef.resolveOverbookToWaitlist(sub, reg.Id, grp);
          if (obWithMail && reg.ParticipantEmail && !selectedEvent.disableEmails) {
            // Einzeln: ggf. vom Admin editierter Text. Sammel: pro Person
            // frisch personalisiert aus dem Standard-Template.
            const mail = isBulk
              ? await eventServiceRef.buildOverbookApologyEmail(nm, selectedEvent.title, obMailLang, getFairWaitlistRank(reg))
              : { subject: obMailSubject, body: obMailBody };
            try {
              await eventServiceRef.queueEmail(
                mail.subject, reg.ParticipantEmail, nm, mail.body,
                'Info', selectedEvent.title, selectedEvent.id
              );
            } catch { /* Mail-Fehler darf Korrektur nicht blockieren */ }
          }
          if (obRemoveCalendar && reg.ParticipantEmail && !selectedEvent.disableOutlook) {
            try {
              await eventServiceRef.queueOutlookEvent(
                reg.ParticipantEmail, selectedEvent.id, selectedEvent.title, 'Ausladen'
              );
            } catch { /* Kalender-Abmeldung best-effort */ }
          }
        } else {
          // Platz behalten
          if (obKeepVariant === 'active') {
            await eventServiceRef.resolveOverbookKeepActive(sub, reg.Id);
          } else {
            await eventServiceRef.resolveOverbookKeepAsFirstWaitlist(sub, reg.Id, grp);
          }
        }
      }
      // IDs neu vergeben (Aktive 1..N, Warteliste N+1..) + Counter + Seat-Sync.
      // Mit Fortschritts-Overlay, damit man bei großen Listen sieht wie weit.
      setReorderProgressLabel('IDs werden neu vergeben…');
      setReorderProgress(0);
      try { await eventServiceRef.reorderParticipantIDs(sub, pct => setReorderProgress(pct)); } catch { /* */ }
      try { await eventServiceRef.syncSeatsToActiveCount(sub, { isSplit: isSplitCapacity }); } catch { /* */ }
      setReorderProgress(null);
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch { /* einzelne Fehler werden geschluckt; Liste wird trotzdem neu geladen */ }
    setObBusy(false);
    setOverbookModal(null);
  };

  // v11.36: TeilnehmerIDs neu vergeben — gemeinsam von der Toolbox-Kachel
  // UND dem Hinweis-Modal genutzt (mit %-Fortschritts-Overlay).
  const runIdReorder = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setIsReorderingIDs(true);
    setReorderResult(null);
    setReorderProgressLabel(isDe ? 'IDs werden neu vergeben…' : 'Reassigning IDs…');
    setReorderProgress(0);
    try {
      const result = await eventServiceRef.reorderParticipantIDs(
        selectedEvent.subsiteUrl,
        pct => setReorderProgress(pct)
      );
      setReorderResult(isDe
        ? `${result.success} aktualisiert, ${result.errors} Fehler`
        : `${result.success} updated, ${result.errors} errors`);
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch {
      setReorderResult(isDe ? 'Fehler beim Neuvergeben der IDs' : 'Error reassigning IDs');
    }
    setReorderProgress(null);
    setIsReorderingIDs(false);
  };

  // v18.70: Manuelles Nachrücken — füllt einen freien Platz mit dem ersten
  // Wartelistler (nach TeilnehmerID). Nutzt dieselbe promoteFirstWaitlistItem-
  // Logik + Mail/Outlook wie der Admin-Cancel-Pfad, nur ohne Cancel.
  const runManualPromote = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setIsPromoting(true);
    setPromoteResult(null);
    try {
      // Bei getrennten Split-Wartelisten ist ohne konkreten Cancel nicht
      // bekannt, welche Gruppe frei ist → ohne Typfilter den ersten
      // Wartelistler nehmen. promoteFirstWaitlistItem prüft die Kapazität
      // (maxParticipants) und rückt nur nach, wenn wirklich ein Platz frei ist.
      const promoted = await eventServiceRef.promoteFirstWaitlistItem(
        selectedEvent.subsiteUrl,
        undefined,
        selectedEvent.maxParticipants,
        undefined,
      );
      if (promoted && promoted.success && promoted.email) {
        setAdminToast({ kind: 'promoted', name: promoted.name || promoted.email, email: promoted.email });
        if (!selectedEvent.disableEmails) {
          try {
            const lang = selectedEvent.emailLanguage || 'EN';
            const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
            const promoteVars = {
              Name: promotedFirstName,
              EventTitle: selectedEvent.title,
              Organizer: formatOrganizerList(selectedEvent.organizers, lang),
              AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
              WaitlistPosition: '',
            };
            let emailData: { subject: string; body: string };
            const spTplRaw = await eventServiceRef.getEmailTemplate('Nachruecken', lang).catch(() => null);
            const spTpl = applyEventTemplateOverride(spTplRaw, selectedEvent.emailTemplateOverrides, 'Nachruecken');
            if (spTpl) {
              emailData = buildEmailFromTemplate(spTpl, promoteVars);
            } else {
              emailData = promotionEmail(promotedFirstName, selectedEvent.title);
            }
            await eventServiceRef.queueEmail(
              emailData.subject, promoted.email, promoted.name || '', emailData.body,
              'Nachruecken', selectedEvent.title, selectedEvent.id
            );
          } catch (err) { console.warn('[DEX] manual promote-email failed:', err); }
        }
        if (!selectedEvent.disableOutlook) {
          try {
            await eventServiceRef.queueOutlookEvent(
              promoted.email, selectedEvent.id, selectedEvent.title, 'Einladen'
            );
          } catch (err) { console.warn('[DEX] manual promote-outlook failed:', err); }
        }
        // IDs neu vergeben + Counter/Seat-Sync + Liste neu laden
        await runIdReorder();
        try {
          const isSplit = typeof selectedEvent.durchstarterCapacity === 'number'
            && typeof selectedEvent.funstarterCapacity === 'number'
            && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
          await eventServiceRef.syncSeatsToActiveCount(selectedEvent.subsiteUrl, { isSplit });
        } catch { /* */ }
        setPromoteResult(isDe ? `${promoted.name || promoted.email} ist nachgerückt.` : `${promoted.name || promoted.email} moved up.`);
      } else if (promoted && promoted.skippedOverbooked) {
        setPromoteResult(isDe ? 'Kein freier Platz — Event ist voll.' : 'No free seat — event is full.');
      } else {
        setPromoteResult(isDe ? 'Niemand auf der Warteliste.' : 'Nobody on the waitlist.');
      }
    } catch (err) {
      console.warn('[DEX] runManualPromote failed:', err);
      setPromoteResult(isDe ? 'Fehler beim Nachrücken.' : 'Error promoting.');
    }
    setIsPromoting(false);
  };

  // v11.70 / v11.71: Hinweis-Box „IDs sind ggf. nicht korrekt" wird jetzt
  // an die tatsächliche TeilnehmerID-Sequenz gekoppelt — nicht mehr an
  // eine 10-Minuten-Zeit-Heuristik nach der letzten Abmeldung.
  //
  // Erwartet: alle nicht-abgemeldeten Einträge (Status in
  // Angemeldet/QR versendet/Eingecheckt/Warteliste) haben TeilnehmerIDs,
  // die nach Sortierung lückenlos 1..N durchlaufen. Sobald
  //   - eine ID fehlt (Lücke),
  //   - eine ID doppelt vorkommt,
  //   - ein nicht-abgemeldeter Eintrag keine (oder ≤0) ID hat,
  // ist der Zustand „IDs evtl. nicht korrekt". Typischer Trigger: gerade
  // erfolgte Abmeldung, der DEX_IDReorder-Flow ist noch nicht fertig.
  // Das gibt einen ehrlichen Status — die Box verschwindet automatisch,
  // sobald der Flow durch ist (statt nach willkürlichen 10 Minuten).
  const recentCancellation = (regs: SPRegistration[]): { recent: boolean; whenIso: string; detail: string } => {
    const active = regs.filter(r => r.Status !== 'Abgemeldet');
    if (active.length === 0) return { recent: false, whenIso: '', detail: '' };
    const ids: number[] = [];
    let noId = 0;
    for (const r of active) {
      const id = Number(r.TeilnehmerID);
      if (!isFinite(id) || id <= 0) { noId++; continue; }
      ids.push(id);
    }
    ids.sort((a, b) => a - b);
    // v22.12: konkrete Diagnose statt nur ja/nein — erste Lücke + Duplikate
    // zählen, damit die Box belegt, WAS in den geladenen Daten falsch ist.
    let dups = 0;
    let firstGapAt = 0;
    for (let i = 0; i < ids.length; i++) {
      if (i > 0 && ids[i] === ids[i - 1]) dups++;
      if (firstGapAt === 0 && ids[i] !== i + 1) firstGapAt = i + 1;
    }
    if (noId === 0 && dups === 0 && firstGapAt === 0) return { recent: false, whenIso: '', detail: '' };
    const parts: string[] = [];
    if (firstGapAt > 0) parts.push(`Nummern nicht durchgängig (erwartet Nr. ${firstGapAt})`);
    if (dups > 0) parts.push(`${dups} doppelte Nummer${dups === 1 ? '' : 'n'}`);
    if (noId > 0) parts.push(`${noId} Eintr${noId === 1 ? 'ag' : 'äge'} ohne Nummer`);
    return {
      recent: true,
      whenIso: latestCancelIso(regs),
      detail: `${active.length} aktive Einträge — ${parts.join(', ')}`,
    };
  };
  // Hilfsfunktion: jüngste CancellationDate aus der Liste (für den
  // optionalen Zeit-Hinweis in der Box).
  const latestCancelIso = (regs: SPRegistration[]): string => {
    let latest = 0;
    for (const r of regs) {
      if (r.Status !== 'Abgemeldet') continue;
      const t = new Date(r.CancellationDate || '').getTime();
      if (!isNaN(t) && t > latest) latest = t;
    }
    return latest > 0 ? new Date(latest).toISOString() : '';
  };
  const idFixCheckedForRef = React.useRef<string | null>(null);
  // v22.12: solange die geladenen Daten kaputte IDs zeigen, lädt die App die
  // Teilnehmerliste automatisch alle 30 Sekunden neu — sobald der
  // DEX_IDReorder-Flow durch ist, verschwindet die Warn-Box von selbst
  // (vorher musste man manuell „Aktualisieren" klicken und hielt den
  // durchgelaufenen Flow fälschlich für kaputt).
  const idRecheckBusyRef = React.useRef(false);
  const [idRecheckBusy, setIdRecheckBusy] = React.useState(false);
  const reloadRegistrationsForIdCheck = React.useCallback(async (): Promise<void> => {
    if (!selectedEvent || idRecheckBusyRef.current) return;
    idRecheckBusyRef.current = true;
    setIdRecheckBusy(true);
    try {
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch { /* best-effort */ }
    idRecheckBusyRef.current = false;
    setIdRecheckBusy(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id]);

  // v23.0: Eine Registrierung per Drag&Drop in ein Team/eine Break-Out-Session
  // (targetTid) oder zurück „ohne Team" ('') verschieben. War die Person Lead
  // ihres alten Teams und bleiben Mitglieder, rückt die früheste nach.
  const moveRegToTeam = async (reg: SPRegistration, targetTid: string, targetTeamName: string | undefined): Promise<void> => {
    if (!selectedEvent?.subsiteUrl || !eventServiceRef) return;
    const sub = selectedEvent.subsiteUrl;
    const curTid = reg.TeamId || '';
    if (curTid === (targetTid || '')) return;
    try {
      await eventServiceRef.assignRegistrationToTeam(sub, reg.Id, targetTid || '', targetTeamName || '', false);
      if (curTid && reg.TeamLead) {
        const rest = registrations.filter(x => x.Id !== reg.Id && (x.TeamId || '') === curTid && x.Status !== 'Abgemeldet');
        if (rest.length > 0) {
          rest.sort((a, b) => ((a.TeilnehmerID ?? 9_999_999) as number) - ((b.TeilnehmerID ?? 9_999_999) as number));
          const tn = rest.find(x => x.TeamName)?.TeamName || '';
          try { await eventServiceRef.assignRegistrationToTeam(sub, rest[0].Id, curTid, tn || '', true); } catch { /* */ }
        }
      }
      const nm = `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantName || reg.ParticipantEmail;
      eventServiceRef.writeChangeLog({
        action: targetTid ? 'TeamMemberAssigned' : 'TeamMemberRemoved',
        targetType: 'Participant', targetId: reg.ParticipantEmail, targetName: nm,
        eventId: selectedEvent.id, eventTitle: selectedEvent.title,
        details: { fromTeam: curTid, toTeam: targetTid, via: 'dragdrop' },
      }).catch(() => { /* */ });
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch (err) { console.warn('[DEX] moveRegToTeam failed:', err); }
  };
  // Drop-Handler: gezogene Registrierung ermitteln + verschieben.
  const onTeamDrop = (targetTid: string, targetTeamName: string | undefined): void => {
    setDragOverTid(null);
    const id = dragRegId;
    setDragRegId(null);
    if (id === null) return;
    const reg = registrations.find(r => r.Id === id);
    if (reg) moveRegToTeam(reg, targetTid, targetTeamName).catch(() => { /* */ });
  };

  // v23.0: Aktive Teams aus den geladenen Registrierungen gruppieren
  // (für die Per-Team-Mail).
  const getActiveTeams = (): Array<{ tid: string; teamName: string; members: SPRegistration[] }> => {
    const map: Record<string, SPRegistration[]> = {};
    for (const r of registrations) {
      if (r.Status === 'Abgemeldet') continue;
      const tid = r.TeamId || '';
      if (!tid) continue;
      (map[tid] = map[tid] || []).push(r);
    }
    return Object.entries(map).map(([tid, members]) => ({ tid, members, teamName: members.find(m => m.TeamName)?.TeamName || '' }));
  };
  // Mail-Dialog mit vorausgefülltem Text öffnen.
  const openTeamMailDialog = (): void => {
    if (!selectedEvent) return;
    const termS = selectedEvent.teamTermSingular || 'Team';
    setTeamMailSubject(isDe ? `Deine ${termS}: ${selectedEvent.title}` : `Your ${termS}: ${selectedEvent.title}`);
    setTeamMailBody(isDe
      ? `<p>Hallo {{Vorname}},</p>\n<p>hier sind die Infos zu deiner <strong>${termS}</strong> beim Event <strong>{{EventTitle}}</strong>:</p>\n<p><strong>{{TeamName}}</strong></p>\n<p>{{TeamInfo}}</p>\n<p>Viele Grüße<br />Dein Event-Team</p>`
      : `<p>Hi {{Vorname}},</p>\n<p>here is the info for your <strong>${termS}</strong> at <strong>{{EventTitle}}</strong>:</p>\n<p><strong>{{TeamName}}</strong></p>\n<p>{{TeamInfo}}</p>\n<p>Best regards<br />Your event team</p>`);
    const init: Record<string, string> = {};
    for (const t of getActiveTeams()) init[t.tid] = teamMailInfoByTid[t.tid] || '';
    setTeamMailInfoByTid(init);
    setTeamMailOpen(true);
  };
  // Pro Team: jedes aktive Mitglied bekommt eine eigene Mail mit team-
  // spezifischer Info (z.B. Teams-Einwahllink). Im Deloitte-Layout gewrappt.
  const sendTeamMails = async (): Promise<void> => {
    if (!selectedEvent || !eventServiceRef) return;
    if (selectedEvent.disableEmails) {
      showAlert(isDe ? 'E-Mails sind für dieses Event deaktiviert (Schritt 6 „Kommunikation").' : 'Emails are disabled for this event (step 6 “Communication”).', { variant: 'error' });
      return;
    }
    setTeamMailSending(true);
    const escHtml = (s: string): string => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    // URLs in der Team-Info klickbar machen + Zeilenumbrüche zu <br>.
    const linkify = (raw: string): string => escHtml(raw)
      .replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" style="color:#86bc25;font-weight:600;">$1</a>')
      .replace(/\n/g, '<br />');
    const termS = selectedEvent.teamTermSingular || 'Team';
    let sent = 0;
    for (const t of getActiveTeams()) {
      const infoHtml = linkify((teamMailInfoByTid[t.tid] || '').trim());
      const tName = t.teamName || termS;
      for (const m of t.members) {
        const first = (m.Vorname && m.Vorname.trim()) || (m.ParticipantName || '').split(/\s+/)[0] || '';
        const fullName = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
        const bodyFilled = teamMailBody
          .replace(/\{\{Vorname\}\}/g, escHtml(first))
          .replace(/\{\{Name\}\}/g, escHtml(fullName))
          .replace(/\{\{TeamName\}\}/g, escHtml(tName))
          .replace(/\{\{EventTitle\}\}/g, escHtml(selectedEvent.title))
          .replace(/\{\{TeamInfo\}\}/g, infoHtml || (isDe ? '<em>(keine zusätzlichen Infos)</em>' : '<em>(no additional info)</em>'));
        const subjectFilled = teamMailSubject
          .replace(/\{\{TeamName\}\}/g, tName)
          .replace(/\{\{EventTitle\}\}/g, selectedEvent.title);
        const wrapped = wrapTemplate('#86bc25', subjectFilled, tName, bodyFilled);
        try {
          const ok = await eventServiceRef.queueEmail(subjectFilled, m.ParticipantEmail, fullName, wrapped, 'TeamInfo', selectedEvent.title, selectedEvent.id);
          if (ok) sent += 1;
        } catch { /* best-effort pro Empfänger */ }
      }
    }
    setTeamMailSending(false);
    setTeamMailOpen(false);
    showAlert(isDe ? `${sent} Mail(s) in die Warteschlange gelegt — sie werden in Kürze versendet.` : `${sent} mail(s) queued — they will be sent shortly.`, { variant: 'success' });
  };
  // Max. 10 automatische Neu-Checks (≈5 Min) pro Event — wenn die Lücke dann
  // immer noch da ist, ist sie echt (Tail-Race, siehe Box-Text) und kein
  // weiteres Polling nötig.
  const idRecheckCountRef = React.useRef(0);
  React.useEffect(() => { idRecheckCountRef.current = 0; }, [selectedEvent?.id]);
  React.useEffect(() => {
    if (!selectedEvent) return undefined;
    // v22.67: kein ID-Durchgängigkeits-Polling im Klammer-Modus (Schatten-Zeilen
    // haben keine fortlaufenden Nummern — das war ein Fehlalarm).
    if (selectedEvent.subEventsOnlyMode) return undefined;
    if (!recentCancellation(registrations).recent) return undefined;
    if (idRecheckCountRef.current >= 10) return undefined;
    const timer = window.setInterval(() => {
      idRecheckCountRef.current++;
      if (idRecheckCountRef.current > 10) { window.clearInterval(timer); return; }
      reloadRegistrationsForIdCheck().catch(() => { /* */ });
    }, 30000);
    return () => window.clearInterval(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, registrations, reloadRegistrationsForIdCheck]);

  // v11.70: kein Modal mehr beim Event-Oeffnen — der Hinweis steht ab
  // jetzt direkt als Box oben in der Teilnehmerliste, solange die
  // Bedingung erfüllt ist (siehe Render-Block unten). Der Ref bleibt
  // erhalten, um in Zukunft ein erneutes „Mount-Trigger"-Verhalten
  // einbauen zu können, ohne den Save-Pfad zu touchen.
  React.useEffect(() => {
    if (!selectedEvent || isLoadingRegs || registrations.length === 0) return;
    if (idFixCheckedForRef.current === selectedEvent.id) return;
    idFixCheckedForRef.current = selectedEvent.id;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, registrations, isLoadingRegs]);

  // v6.17: Spaltenkonfiguration der Teilnehmertabelle (pro Event, lokal gespeichert).
  //  - columnOrder = geordnete Liste sichtbarer Spalten-IDs
  //  - hiddenColumns = ausgeblendete Spalten-IDs (können übers "+ Spalte"-Popover wieder zugeschaltet werden)
  // Die Spezialspalten 'id' / 'vorname' / 'nachname' / 'action' sind alwaysVisible und können nicht ausgeblendet werden.
  const [columnOrder, setColumnOrder] = React.useState<string[]>([]);
  const [hiddenColumns, setHiddenColumns] = React.useState<string[]>([]);
  const [showColumnPicker, setShowColumnPicker] = React.useState(false);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const spfxContext = (window as any).__dexSpfxContext;
  const eventServiceRef = React.useMemo(() => spfxContext ? new EventService(spfxContext) : null, []);

  // v22.7: Hintergrund-Check beim Öffnen eines Events — sind die E-Mail-Adressen
  // der Teilnehmer noch zu einem aktiven Deloitte-Konto? Ergebnis wird pro Event
  // max. 1×/Tag in localStorage gecacht (kein Graph-Call bei jedem Öffnen).
  React.useEffect(() => {
    setInactiveAccounts([]);
    if (!selectedEvent || !eventServiceRef || registrations.length === 0) return undefined;
    const emails = Array.from(new Set(registrations
      .filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt' || r.Status === 'Warteliste')
      .map(r => (r.ParticipantEmail || '').trim().toLowerCase())
      .filter(Boolean)));
    if (emails.length === 0) return undefined;
    const cacheKey = `dex_acctcheck_${selectedEvent.id}`;
    try {
      const raw = window.localStorage.getItem(cacheKey);
      if (raw) {
        const parsed = JSON.parse(raw) as { ts?: number; inactive?: string[]; checked?: string[] };
        const fresh = !!parsed && typeof parsed.ts === 'number' && (Date.now() - parsed.ts) < 24 * 60 * 60 * 1000 && Array.isArray(parsed.inactive);
        // v22.43: Cache nur nutzen, wenn ALLE aktuellen Adressen schon geprüft
        // wurden. Sind seit dem letzten Lauf neue Teilnehmer dazugekommen
        // (Adresse nicht in `checked`), wird frisch geprüft — sonst blieben
        // neu hinzugefügte Personen bis zu 24h ungeprüft (Bug v22.7–v22.42).
        const checked = Array.isArray(parsed?.checked) ? (parsed!.checked as string[]) : [];
        const coversAll = checked.length > 0 && emails.every(e => checked.indexOf(e) >= 0);
        if (fresh && coversAll) {
          setInactiveAccounts((parsed!.inactive || []).filter(e => emails.indexOf(e) >= 0));
          return undefined;
        }
      }
    } catch { /* localStorage evtl. blockiert */ }
    let cancelled = false;
    (async () => {
      try {
        const res = await eventServiceRef.checkAccountsActive(emails);
        if (cancelled || !res.ok) return;
        setInactiveAccounts(res.inactive);
        try { window.localStorage.setItem(cacheKey, JSON.stringify({ ts: Date.now(), inactive: res.inactive, checked: emails })); } catch { /* */ }
      } catch { /* best-effort */ }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, registrations.length, eventServiceRef]);

  // SuperAdmin sieht alle Events, EventAdmin nur seine + QR-Scanner-Events.
  // Zugriff wird strikt per E-Mail geprüft — NICHT per Namens-Substring
  // (hatte mehrere Jahre einen Match-per-Surname-Bug, der bei häufigen Nachnamen
  // zu False-Positives führte: z.B. eine Assistentin "Frau Müller" konnte Events
  // sehen, deren Organizer auf "Max Müller" hieß — weil "müller" in "max müller"
  // vorkommt. Seit v6.20 nur noch exakt per currentUser.email gegen
  // event.organizerEmails bzw. event.qrScannerEmails.)
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isQRScannerFor = (ev: DeloitteEvent): boolean =>
    !!currentEmailLc && !!ev.qrScannerEmails && ev.qrScannerEmails.some(e => e.toLowerCase() === currentEmailLc);
  // v9.18: Co-Organizer haben pro Event die gleichen Rechte wie der Hauptorganizer.
  // isOrganizerFor returned true sowohl für event.organizerEmails als auch
  // für event.coOrganizerEmails (per-Event-Rolle).
  const isOrganizerFor = (ev: DeloitteEvent): boolean => {
    // v18.3: Im Demo-Modus ist der (User-)Demo-Account „Organizer" des
    // synthetischen Demo-Events — so sieht er die Teilnehmer-Verwaltung
    // (read-only) im Admin-Center. Greift nur für das Demo-Event.
    if (isImpersonating && ev.isDemoShowcase) return true;
    if (!currentEmailLc) return false;
    if (ev.organizerEmails && ev.organizerEmails.some(e => e.toLowerCase() === currentEmailLc)) return true;
    if (ev.coOrganizerEmails && ev.coOrganizerEmails.some(e => (e || '').toLowerCase() === currentEmailLc)) return true;
    // v22.14: Organizer des HAUPTEVENTS gelten auch auf dessen Sub-Events als
    // Organizer. Vorher waren Sub-Event-Tabs für Parent-Organizer beschnitten
    // (Status-Badge nicht klickbar, Organizer-Aktionen ausgeblendet), wenn die
    // Organizer-Liste des Kindes nicht (mehr) identisch gepflegt war.
    if (ev.parentEventId) {
      const parent = allEvents.find(p => p.id === ev.parentEventId);
      if (parent) {
        if (parent.organizerEmails && parent.organizerEmails.some(e => (e || '').toLowerCase() === currentEmailLc)) return true;
        if (parent.coOrganizerEmails && parent.coOrganizerEmails.some(e => (e || '').toLowerCase() === currentEmailLc)) return true;
      }
    }
    return false;
  };
  const adminEvents = isAdmin
    ? events
    : events.filter(e => isOrganizerFor(e) || isQRScannerFor(e));
  // Wenn der User NUR QR-Scanner ist (nicht Organizer + nicht Admin), dann läuft die
  // Admin-Page im eingeschränkten Modus für das ausgewählte Event: nur KPI-Kacheln
  // + QR-Code-Scanner-Button sichtbar.
  const isQRScannerOnlyForSelected = !!selectedEvent && !isAdmin && !isOrganizerFor(selectedEvent) && isQRScannerFor(selectedEvent);

  // Für Admins: vergangene Events in eine einklappbare Sektion auslagern
  // (Organizer sehen nur ihre eigenen Events — dort bleiben auch abgelaufene
  // sichtbar, weil der Organizer sie für den Abschluss / CSV-Export etc.
  // evtl. direkt griffbereit braucht).
  const now = Date.now();
  const isPastEvent = (e: DeloitteEvent): boolean =>
    !!e.endDate && new Date(e.endDate).getTime() < now;
  const currentEventsRaw = isAdmin ? adminEvents.filter(e => !isPastEvent(e)) : adminEvents;
  const pastEventsRaw = isAdmin ? adminEvents.filter(isPastEvent) : [];
  const [showPastEvents, setShowPastEvents] = React.useState(false);
  // v18.2: Entwurf-Filter + Sortierung der Admin/Organizer-Event-Liste.
  // Default-Sortierung alphabetisch nach Titel; alternativ nach Startdatum
  // aufsteigend. „Entwürfe ausblenden" filtert isFictive-Events raus.
  const [hideDrafts, setHideDrafts] = React.useState(false);
  const [eventSortMode, setEventSortMode] = React.useState<'alpha' | 'date'>('alpha');
  const draftCount = adminEvents.filter(e => e.isFictive).length;
  const sortAndFilterEvents = React.useCallback((list: DeloitteEvent[]): DeloitteEvent[] => {
    let arr = list.slice();
    if (hideDrafts) arr = arr.filter(e => !e.isFictive);
    arr.sort((a, b) => {
      if (eventSortMode === 'date') {
        const am = a.startDate ? new Date(a.startDate).getTime() : Number.POSITIVE_INFINITY;
        const bm = b.startDate ? new Date(b.startDate).getTime() : Number.POSITIVE_INFINITY;
        if (am !== bm) return am - bm;
        return (a.title || '').localeCompare(b.title || '', isDe ? 'de' : 'en');
      }
      return (a.title || '').localeCompare(b.title || '', isDe ? 'de' : 'en');
    });
    return arr;
  }, [hideDrafts, eventSortMode, isDe]);
  const currentEvents = sortAndFilterEvents(currentEventsRaw);
  const pastEvents = sortAndFilterEvents(pastEventsRaw);

  // v6.17: Verfügbare Spalten der Teilnehmer-Tabelle aufbauen. MUSS vor dem
  // early return `if (!selectedEvent) return ...` stehen — sonst verletzen
  // die Hooks die Rules-of-Hooks (unterschiedliche Hook-Anzahl pro Render =
  // React Error #310).
  // v19.11: Hat dieses Event überhaupt Warteliste-/Nachrück-Aktivität? Nur dann
  // sind die Nachrück-Audit-Spalten („Nachgerückt am", „Hat ersetzt", „Wurde
  // ersetzt durch") sinnvoll. `waitlistEnabled` allein reicht NICHT, weil es per
  // Default `true` ist (e.WaitlistEnabled !== false) — Events ohne konfigurierte
  // Warteliste hätten sonst immer die leeren Audit-Spalten. „Aktiv" = jemand
  // steht auf der Warteliste ODER es gibt bereits Nachrück-Daten.
  const hasWaitlistActivity = React.useMemo(() => registrations.some(r =>
    r.Status === 'Warteliste'
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    || !!(r as any).PromotedDate || !!(r as any).ReplacedParticipantEmail || !!(r as any).ReplacedByParticipantEmail
  ), [registrations]);
  const availableColumns = React.useMemo(() => {
    const isSplit = !!selectedEvent
      && typeof selectedEvent.durchstarterCapacity === 'number'
      && typeof selectedEvent.funstarterCapacity === 'number'
      && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
    const userIds = (selectedEvent?.eventSpecificFields || [])
      .filter(f => f.type === 'user' || f.type === 'roommate')
      .map(f => f.id);
    // v14.8: Anrede-Spalte nur anbieten, wenn das Event die Anrede beim
    // Anmelden tatsächlich abfragt (askSalutation). Sonst landet eine leere
    // Spalte voller „-" im Admin-Center, die niemand braucht.
    const askSal = !!selectedEvent?.askSalutation;
    const cols: Array<{ id: string; label: string; alwaysVisible?: boolean }> = [
      { id: 'id', label: '#', alwaysVisible: true },
      ...(askSal ? [{ id: 'anrede', label: 'Anrede' }] : []),
      // v11.26: getrennte Vorname / Nachname Spalten statt der einen
      // kombinierten 'name'-Spalte. Alte localStorage-Einträge mit 'name'
      // werden im useEffect-Loader unten in 'vorname','nachname' migriert.
      { id: 'vorname', label: 'Vorname', alwaysVisible: true },
      { id: 'nachname', label: 'Nachname', alwaysVisible: true },
      { id: 'email', label: 'Email' },
      { id: 'jobTitle', label: 'Job Title' },
      { id: 'location', label: 'Standort' },
    ];
    // v11.6: bei Split-Capacity die frei wählbaren Gruppen-Labels nutzen
    // (Fallback auf 'Starter-Typ' wenn keine Labels gesetzt sind).
    if (isSplit) {
      const lblA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || '';
      const lblB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || '';
      const colLabel = (lblA && lblB) ? `${lblA} / ${lblB}` : (isDe ? 'Gruppe' : 'Group');
      cols.push({ id: 'starterType', label: colLabel });
    }
    cols.push({ id: 'status', label: 'Status' });
    cols.push({ id: 'date', label: 'Registriert am' });
    // v17.15/v17.17.1: Nachrück-Audit-Spalten — nur sichtbar wenn das
    // Event ueberhaupt eine Warteliste haben KANN (waitlistEnabled UND
    // maxParticipants > 0). Bei „Unbegrenzt"-Events kommt nie jemand auf
    // die Warteliste, deshalb sind die drei Audit-Spalten ohne Inhalt.
    // v19.11: Zusätzlich `hasWaitlistActivity` — Events OHNE echte Warteliste
    // (Default waitlistEnabled=true, aber niemand wartet/nachgerückt) zeigen die
    // leeren Audit-Spalten jetzt nicht mehr.
    if (selectedEvent?.waitlistEnabled && (selectedEvent?.maxParticipants || 0) > 0 && hasWaitlistActivity) {
      cols.push({ id: 'promotedDate', label: 'Nachgerückt am' });
      // v19.4: „Hat ersetzt" = die abgemeldete Person, deren Platz diese Person
      // übernommen hat. „Wurde ersetzt durch" wandert in die Abmeldungen-Tabelle
      // (gehört zur abgemeldeten Person, nicht zur aktiven).
      cols.push({ id: 'replaced', label: 'Hat ersetzt' });
    }
    cols.push({ id: 'registeredBy', label: 'Registriert von' });
    // v16.1: Team-Spalte — zeigt pro Teilnehmer den Team-Namen (falls Team-
    // Anmeldung aktiv und der TN in einem Team ist).
    if (selectedEvent?.teamRegistrationEnabled) {
      cols.push({ id: 'team', label: 'Team' });
    }
    if (userIds.length > 0) {
      // v11.56: Label aus dem ersten roommate-/user-Feld ableiten, statt hart
      // „Zimmerpartner" zu nennen. Wenn ein roommate-Feld existiert, nimm dessen
      // Label (User-Picker-Pairs); sonst das erste user-Feld; Fallback bleibt
      // der deutsche Default.
      const fields = selectedEvent?.eventSpecificFields || [];
      const firstRoommate = fields.filter(f => f.type === 'roommate' && f.label && f.label.trim())[0];
      const firstUser = fields.filter(f => f.type === 'user' && f.label && f.label.trim())[0];
      const roommateLabel = (firstRoommate?.label || firstUser?.label || 'Zimmerpartner').trim();
      cols.push({ id: 'roommate', label: roommateLabel });
    }
    // v14.11: Wenn ein Sub-Event selektiert ist, blenden wir die
    // Custom-Fields des Parent-Events (Pastel A) zusätzlich ein. Die
    // eigenen Sub-Event-Fields (Pastel B) folgen direkt danach. ID-
    // Präfix `cfp-` unterscheidet Parent- von Sub-Event-Feldern (`cf-`).
    const parentForCols: DeloitteEvent | null = (selectedEvent && selectedEvent.parentEventId)
      ? (allEvents.find(e => e.id === selectedEvent.parentEventId) || null)
      : null;
    if (parentForCols) {
      const ownIds = new Set((selectedEvent?.eventSpecificFields || []).map(f => f.id));
      // v19.10: 'roommate' (wie 'user') NICHT als generische Spalte ausgeben —
      // diese Felder werden bereits über die dedizierte „roommate"-Spalte (mit
      // Match-Badge) gerendert. Sonst erscheint das Feld DOPPELT (einmal mit
      // Match, einmal als roher „Name <email>"-Text).
      for (const f of (parentForCols.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'roommate' && f.label && f.label.trim())) {
        // Sub-Events erben Parent-Felder evtl. 1:1 (Wizard kopiert das beim
        // Anlegen). Nicht doppelt ausgeben, wenn das eigene Feld die
        // gleiche ID hat — in dem Fall reicht die Sub-Event-Spalte.
        if (ownIds.has(f.id)) continue;
        cols.push({ id: `cfp-${f.id}`, label: f.label });
      }
    }
    // v19.10: 'roommate'-Felder (wie 'user') hier ausschließen — sie haben
    // bereits die dedizierte „roommate"-Spalte mit Match-Badge. Vorher fehlte
    // `f.type !== 'roommate'`, deshalb erschien ein Zimmerpartner-Feld DOPPELT:
    // einmal als Match-Spalte, einmal als generische cf-Spalte mit rohem
    // „Nachname, Vorname <email>"-Text.
    for (const f of (selectedEvent?.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'roommate' && f.label && f.label.trim())) {
      cols.push({ id: `cf-${f.id}`, label: f.label });
    }
    cols.push({ id: 'action', label: 'Aktion', alwaysVisible: true });
    return cols;
  }, [
    selectedEvent?.id,
    selectedEvent?.parentEventId,
    selectedEvent?.durchstarterCapacity,
    selectedEvent?.funstarterCapacity,
    (selectedEvent?.eventSpecificFields || []).map(f => `${f.id}:${f.type}:${f.label}`).join(','),
    // v14.11: Parent-Custom-Fields als Dep
    (() => {
      if (!selectedEvent?.parentEventId) return '';
      const p = allEvents.find(e => e.id === selectedEvent.parentEventId);
      return (p?.eventSpecificFields || []).map(f => `${f.id}:${f.type}:${f.label}`).join(',');
    })(),
    // v19.11: Audit-Spalten-Sichtbarkeit hängt an der Warteliste-Aktivität.
    hasWaitlistActivity,
  ]);

  const columnStorageKey = selectedEvent ? `dex_admin_columns_${selectedEvent.id}` : '';
  // localStorage-Load beim Event-Wechsel.
  React.useEffect(() => {
    if (!selectedEvent) { setColumnOrder([]); setHiddenColumns([]); return; }
    const allIds = availableColumns.map(c => c.id);
    try {
      const raw = localStorage.getItem(columnStorageKey);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && Array.isArray(parsed.order) && Array.isArray(parsed.hidden)) {
          // v11.26: Migration alter Spaltenkonfigurationen — die zentrale
          // 'name'-Spalte wurde in 'vorname' + 'nachname' aufgeteilt. Wenn
          // ein gespeichertes Layout noch 'name' enthält, an gleicher
          // Position durch ['vorname','nachname'] ersetzen, damit der
          // User seine gewünschte Reihenfolge beibehält.
          const migratedOrder: string[] = [];
          for (const id of parsed.order as string[]) {
            if (id === 'name') {
              if (migratedOrder.indexOf('vorname') < 0) migratedOrder.push('vorname');
              if (migratedOrder.indexOf('nachname') < 0) migratedOrder.push('nachname');
            } else {
              migratedOrder.push(id);
            }
          }
          const knownOrder = migratedOrder.filter((id: string) => allIds.indexOf(id) >= 0);
          const missing = allIds.filter(id => knownOrder.indexOf(id) < 0);
          // v15.3: neu hinzugekommene Spalten (z.B. nach Custom-Field-Anlage
          // an einem bestehenden Event) VOR der „Aktion"-Spalte einreihen,
          // nicht hinten dran — sonst landen sie rechts neben den Buttons.
          const actionPos = knownOrder.indexOf('action');
          const mergedOrder = actionPos >= 0
            ? [...knownOrder.slice(0, actionPos), ...missing, ...knownOrder.slice(actionPos)]
            : [...knownOrder, ...missing];
          setColumnOrder(mergedOrder);
          // 'name' aus hidden auch herausfiltern (wenn jemals manuell hidden gesetzt wurde,
          // unwahrscheinlich da alwaysVisible — aber defensiv).
          setHiddenColumns(parsed.hidden.filter((id: string) => id !== 'name' && allIds.indexOf(id) >= 0));
          return;
        }
      }
    } catch { /* kaputte Config ignorieren */ }
    setColumnOrder(allIds);
    setHiddenColumns([]);
  }, [columnStorageKey, availableColumns.map(c => c.id).join(',')]);

  // Persistieren bei Änderungen.
  React.useEffect(() => {
    if (!columnStorageKey || columnOrder.length === 0) return;
    try {
      localStorage.setItem(columnStorageKey, JSON.stringify({ order: columnOrder, hidden: hiddenColumns }));
    } catch { /* quota exceeded oder private mode → ignorieren */ }
  }, [columnStorageKey, columnOrder.join(','), hiddenColumns.join(',')]);

  // Helper: Spalte ausblenden / wieder einblenden / verschieben.
  const hideColumn = (id: string): void => {
    const col = availableColumns.find(c => c.id === id);
    if (!col || col.alwaysVisible) return;
    if (hiddenColumns.indexOf(id) < 0) setHiddenColumns([...hiddenColumns, id]);
  };
  const showColumn = (id: string): void => {
    setHiddenColumns(hiddenColumns.filter(h => h !== id));
    if (columnOrder.indexOf(id) < 0) {
      const actionIdx = columnOrder.indexOf('action');
      const next = [...columnOrder];
      if (actionIdx >= 0) next.splice(actionIdx, 0, id); else next.push(id);
      setColumnOrder(next);
    }
  };
  const moveColumn = (id: string, direction: -1 | 1): void => {
    const idx = columnOrder.indexOf(id);
    if (idx < 0) return;
    const target = idx + direction;
    if (target < 0 || target >= columnOrder.length) return;
    if (columnOrder[target] === 'action') return;
    const next = [...columnOrder];
    [next[idx], next[target]] = [next[target], next[idx]];
    setColumnOrder(next);
  };

  /**
   * CSV Export für Teilnehmerlisten.
   * - 'deloitte': alle internen Felder (Anrede, Name, Email, Department, Location, JobTitle, Phone, Status, ...)
   * - 'b2run': Format laut B2Run Excel-Template (Nr, Anrede, Vorname, Nachname, E-Mail, Startblock, Zustimmung AGB, Anonym, Gruppe, Strasse, PLZ, Stadt, Mobilnummer, Infoservice, Altersklasse)
   */
  const exportCsv = (mode: 'deloitte' | 'b2run', audience: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled' = 'active'): void => {
    if (!selectedEvent) return;
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const audienceFilter = (r: SPRegistration): boolean => {
      if (audience === 'waitOnly') return r.Status === 'Warteliste';
      if (audience === 'activePlusWait') return ACTIVE.indexOf(r.Status) >= 0 || r.Status === 'Warteliste';
      // v20.4: alles inkl. Abgemeldete (Status-Spalte ist im Export enthalten).
      if (audience === 'withCancelled') return true;
      return ACTIVE.indexOf(r.Status) >= 0;
    };
    // v17.12: nach TeilnehmerID asc sortieren (vorher random / Status-Reihenfolge).
    const activeRegsForExport = registrations
      .filter(audienceFilter)
      .slice()
      .sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0));
    if (activeRegsForExport.length === 0) { showAlert('Keine Teilnehmer zum Exportieren.'); return; }

    // v20.0 (Audit): toter CSV-Escaper `esc` entfernt — seit dem Umstieg auf
    // natives XLSX (v8.4) wurde er nie mehr aufgerufen.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parseCustom = (json: string): Record<string, any> => {
      try { return JSON.parse(json || '{}'); } catch { return {}; }
    };

    let headers: string[] = [];
    let rows: string[][] = [];

    if (mode === 'b2run') {
      // Reihenfolge exakt wie B2Run Excel
      headers = [
        'Nr.', 'Anrede', 'Vorname', 'Nachname', 'E-Mail',
        'Startblock', 'Zustimmung AGB & Datenschutzhinweise', 'Anonym',
        'Gruppe', 'Strasse und Hausnummer (privat)', 'PLZ (privat)', 'Stadt (privat)',
        'Mobilnummer', 'Verwendung Infoservice', 'Altersklasse',
      ];
      rows = activeRegsForExport.map(r => {
        const cd = parseCustom(r.CustomData || '{}');
        const vorname = r.Vorname || (r.ParticipantName || '').split(' ').slice(0, -1).join(' ') || '';
        const nachname = r.Nachname || (r.ParticipantName || '').split(' ').slice(-1).join(' ') || '';
        return [
          String(r.TeilnehmerID || ''),
          r.Anrede || '',
          vorname,
          nachname,
          r.ParticipantEmail || '',
          cd.b2run_startblock || '',
          cd.b2run_datenschutz ? 'Ja' : 'Nein',
          cd.b2run_anonym ? 'Ja' : 'Nein',
          cd.b2run_gruppe || '',
          '', // Strasse - nicht abgefragt
          '', // PLZ - nicht abgefragt
          '', // Stadt - nicht abgefragt
          cd.b2run_mobilnummer || '',
          cd.b2run_infoservice ? 'Ja' : 'Nein',
          cd.b2run_altersklasse || '',
        ];
      });
    } else {
      // Deloitte View: alle internen Felder
      headers = [
        'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'Email',
        'Department', 'Location', 'JobTitle', 'Phone',
        'Status', 'RegistrationDate',
      ];
      // Dynamisch alle Custom Field Labels aus dem Event sammeln
      const customLabels: Array<{ id: string; label: string }> = (selectedEvent.eventSpecificFields || []).map(f => ({ id: f.id, label: f.label }));
      headers = headers.concat(customLabels.map(cf => cf.label));

      rows = activeRegsForExport.map(r => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const anyReg = r as any;
        const cd = parseCustom(r.CustomData || '{}');
        const base = [
          String(r.TeilnehmerID || ''),
          r.Anrede || '',
          r.Vorname || '',
          r.Nachname || '',
          r.ParticipantEmail || '',
          anyReg.Department || '',
          anyReg.Location || '',
          anyReg.JobTitle || '',
          anyReg.Phone || '',
          r.Status || '',
          r.RegistrationDate ? new Date(r.RegistrationDate).toLocaleString('de-DE') : '',
        ];
        const customValues = customLabels.map(cf => {
          const v = cd[cf.id];
          if (v === undefined || v === null) return '';
          if (typeof v === 'boolean') return v ? 'Ja' : 'Nein';
          return String(v);
        });
        return base.concat(customValues);
      });
    }

    const safeName = (selectedEvent.title || 'event').replace(/[^a-zA-Z0-9]/g, '_');
    // XLSX Export — natives Excel-Format, automatische Spalten-Breiten, keine
    // CSV-Escaping-Quirks. Gilt für beide Modi (Teilnehmerliste + B2Run).
    const aoa: (string | number)[][] = [headers, ...rows];
    const sheetName = mode === 'b2run' ? 'B2Run' : 'Teilnehmer';
    const filePrefix = mode === 'b2run' ? 'B2Run' : 'Teilnehmer';
    const fileName = `${filePrefix}_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`;

    // v20.0 (Audit): xlsx erst beim Export-Klick als Chunk nachladen — die
    // Bibliothek ist mit Abstand die schwerste Dependency und wird nur hier
    // gebraucht. Der .then/.catch-Pfad ersetzt das frühere try/catch.
    // v8.4: Manueller Blob-Download statt XLSX.writeFile. Im SPFx-Iframe-
    // Context ist saveAs/createObjectURL häufig blockiert (CORS / Sandbox-
    // Policies), wodurch der Download stillschweigend nicht startet. Mit
    // anchor.click() läuft das in jeder Browser-Umgebung zuverlässig.
    import('xlsx').then(XLSX => {
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const colWidths = headers.map((h, ci) => {
        const maxLen = Math.max(h.length, ...rows.map(r => String(r[ci] || '').length));
        return { wch: Math.min(40, Math.max(10, maxLen + 2)) };
      });
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = colWidths;
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 0);
    }).catch(err => {
      console.warn('[DEX] Excel-Export fehlgeschlagen:', err);
      showAlert(isDe
        ? 'Excel-Export fehlgeschlagen. Bitte Browser-Console prüfen.'
        : 'Excel export failed. Please check the browser console.');
    });
  };

  // v20.4: Excel-Export der konsolidierten Klammer-Ansicht. Baut EINE Datei
  // mit (wählbar) einem Matrix-Blatt — eine Zeile pro Person, Spalten =
  // Stammdaten + Klammer-Felder + pro Sub-Event der Status + dessen Feld-
  // Antworten — und/oder je einem eigenen Blatt pro gewähltem Sub-Event.
  // Datenquellen sind die bereits geladenen States (registrations = Klammer-
  // Zeilen, subEventRegsByEventId = Sub-Event-Listen) — kein Extra-Roundtrip.
  const exportConsolidatedExcel = (
    audience: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled',
    includeMatrix: boolean,
    subIds: string[]
  ): void => {
    if (!selectedEvent) return;
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const matches = (r: SPRegistration): boolean => {
      if (audience === 'waitOnly') return r.Status === 'Warteliste';
      if (audience === 'activePlusWait') return ACTIVE.indexOf(r.Status) >= 0 || r.Status === 'Warteliste';
      if (audience === 'withCancelled') return true;
      return ACTIVE.indexOf(r.Status) >= 0;
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parseCustom = (json: string): Record<string, any> => {
      try { return JSON.parse(json || '{}'); } catch { return {}; }
    };
    const fieldVal = (cd: Record<string, unknown>, id: string): string => {
      const v = cd[id];
      if (v === undefined || v === null) return '';
      if (typeof v === 'boolean') return v ? 'Ja' : 'Nein';
      return String(v);
    };
    const chosenChildren = consolidatedChildren.filter(c => subIds.indexOf(c.id) >= 0);
    const sheets: Array<{ name: string; headers: string[]; rows: string[][] }> = [];
    const sanitizeSheet = (s: string): string => (s || 'Blatt').replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31) || 'Blatt';

    if (includeMatrix) {
      const parentFields = (selectedEvent.eventSpecificFields || []).filter(f => f.label);
      type PersonRow = {
        vorname: string; nachname: string; email: string; jobTitle: string; location: string;
        parentCd: Record<string, unknown>;
        perChild: Record<string, SPRegistration | undefined>;
        hasMatch: boolean;
      };
      const persons: Record<string, PersonRow> = {};
      const ensurePerson = (r: SPRegistration): PersonRow => {
        const key = (r.ParticipantEmail || '').toLowerCase().trim();
        if (!persons[key]) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const anyReg = r as any;
          persons[key] = {
            vorname: r.Vorname || '', nachname: r.Nachname || '',
            email: r.ParticipantEmail || '',
            jobTitle: anyReg.JobTitle || '', location: anyReg.Location || '',
            parentCd: {}, perChild: {}, hasMatch: false,
          };
        }
        return persons[key];
      };
      for (const r of registrations) {
        const p = ensurePerson(r);
        p.parentCd = parseCustom(r.CustomData || '{}');
        if (matches(r)) p.hasMatch = true;
      }
      for (const child of consolidatedChildren) {
        const regs = subEventRegsByEventId[child.id] || [];
        for (const r of regs) {
          const p = ensurePerson(r);
          if (matches(r)) {
            p.perChild[child.id] = r;
            p.hasMatch = true;
          }
        }
      }
      const matrixHeaders: string[] = ['Vorname', 'Nachname', 'Email', 'JobTitle', 'Standort']
        .concat(parentFields.map(f => f.label));
      for (const child of consolidatedChildren) {
        const short = shortSubEventTitle(child.title, selectedEvent.title) || child.title || '?';
        matrixHeaders.push(short);
        for (const f of (child.eventSpecificFields || []).filter(ff => ff.label)) {
          matrixHeaders.push(`${short}: ${f.label}`);
        }
      }
      const matrixRows: string[][] = Object.keys(persons)
        .map(k => persons[k])
        .filter(p => p.hasMatch)
        .sort((a, b) => (a.nachname || '').localeCompare(b.nachname || '', 'de') || (a.vorname || '').localeCompare(b.vorname || '', 'de'))
        .map(p => {
          const row: string[] = [p.vorname, p.nachname, p.email, p.jobTitle, p.location]
            .concat(parentFields.map(f => fieldVal(p.parentCd, f.id)));
          for (const child of consolidatedChildren) {
            const reg = p.perChild[child.id];
            row.push(reg ? (reg.Status || '') : '');
            const cd = reg ? parseCustom(reg.CustomData || '{}') : {};
            for (const f of (child.eventSpecificFields || []).filter(ff => ff.label)) {
              row.push(reg ? fieldVal(cd, f.id) : '');
            }
          }
          return row;
        });
      sheets.push({ name: 'Konsolidiert', headers: matrixHeaders, rows: matrixRows });
    }

    for (const child of chosenChildren) {
      const regs = (subEventRegsByEventId[child.id] || [])
        .filter(matches)
        .slice()
        .sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0));
      const childFields = (child.eventSpecificFields || []).filter(f => f.label);
      const headers = ['TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'Email', 'Department', 'Location', 'JobTitle', 'Status', 'RegistrationDate']
        .concat(childFields.map(f => f.label));
      const rows = regs.map(r => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const anyReg = r as any;
        const cd = parseCustom(r.CustomData || '{}');
        return [
          String(r.TeilnehmerID || ''), r.Anrede || '', r.Vorname || '', r.Nachname || '',
          r.ParticipantEmail || '', anyReg.Department || '', anyReg.Location || '', anyReg.JobTitle || '',
          r.Status || '', r.RegistrationDate ? new Date(r.RegistrationDate).toLocaleString('de-DE') : '',
        ].concat(childFields.map(f => fieldVal(cd, f.id)));
      });
      sheets.push({ name: sanitizeSheet(shortSubEventTitle(child.title, selectedEvent.title) || child.title || 'Sub-Event'), headers, rows });
    }

    if (sheets.length === 0) { showAlert(isDe ? 'Bitte mindestens die Matrix oder ein Sub-Event auswählen.' : 'Please select at least the matrix or one sub-event.'); return; }
    const safeName = (selectedEvent.title || 'event').replace(/[^a-zA-Z0-9]/g, '_');
    const fileName = `Konsolidiert_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    import('xlsx').then(XLSX => {
      const wb = XLSX.utils.book_new();
      const usedNames = new Set<string>();
      for (const sheet of sheets) {
        const aoa: (string | number)[][] = [sheet.headers, ...sheet.rows];
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const colWidths = sheet.headers.map((h, ci) => {
          const maxLen = Math.max(h.length, ...sheet.rows.map(r => String(r[ci] || '').length));
          return { wch: Math.min(40, Math.max(10, maxLen + 2)) };
        });
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (ws as any)['!cols'] = colWidths;
        // Doppelte Blattnamen entschärfen (xlsx verlangt eindeutige Namen).
        let name = sheet.name;
        let i = 2;
        while (usedNames.has(name)) { name = `${sheet.name.slice(0, 28)}_${i}`; i++; }
        usedNames.add(name);
        XLSX.utils.book_append_sheet(wb, ws, name);
      }
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 0);
    }).catch(err => {
      console.warn('[DEX] Konsolidierter Excel-Export fehlgeschlagen:', err);
      showAlert(isDe ? 'Excel-Export fehlgeschlagen. Bitte Browser-Console prüfen.' : 'Excel export failed. Please check the browser console.', { variant: 'error' });
    });
  };

  const handleSelectEvent = async (event: DeloitteEvent): Promise<void> => {
    // v18.24: aktuelle Card-Höhe einfrieren, BEVOR der State wechselt (DOM
    // zeigt noch den alten Stand) — verhindert das Zusammenklappen während
    // die Teilnehmer des neuen Events geladen werden.
    setReservedDetailHeight(detailCardRef.current?.offsetHeight);
    setSelectedEvent(event);
    // v10.19: NavigationContext.selectedEventId mitziehen, damit Header die
    // Page-ID granular ableiten kann (admin-center vs. admin-event) und der
    // Deep-Link-Kopier-Button immer die echte Item-ID des aktuell offenen
    // Events kennt. Skip falls bereits synchron — sonst doppelter History-
    // Eintrag beim Auto-Select via Deep-Link.
    if (selectedEventId !== event.id) {
      navigate('admin', event.id);
    }
    setIsLoadingRegs(true);
    setRegLoadError('');
    try {
      const regs = await getAllRegistrations(event.id);
      setRegistrations(regs);
    } catch {
      setRegistrations([]);
      setRegLoadError('Teilnehmerliste konnte nicht geladen werden.');
    }
    setIsLoadingRegs(false);
    // Reservierung freigeben — der neue Inhalt steht jetzt, die Card nimmt
    // im selben Render die echte neue Höhe an (kein Zwischen-Kollaps).
    setReservedDetailHeight(undefined);
  };

  // v6.31: wenn navigation.selectedEventId gesetzt ist beim Mount (z.B. vom
  // Handbuch-Preview oder einem Deep-Link), direkt in die Detail-Ansicht
  // springen statt auf die Event-Auswahl-Liste.
  const didAutoSelectRef = React.useRef(false);
  React.useEffect(() => {
    if (didAutoSelectRef.current) return;
    if (!selectedEventId || selectedEvent) return;
    const match = adminEvents.find(e => e.id === selectedEventId);
    if (match) {
      didAutoSelectRef.current = true;
      handleSelectEvent(match).catch(() => { /* Fehler wird intern gesetzt */ });
    }
  }, [selectedEventId, adminEvents, selectedEvent]);

  // v22.40: Auto-Heilung stale Überbuchungs-Marker. Hat sich seit dem
  // „Überbuchung prüfen"-Lauf jemand abgemeldet, passt eine vorher als
  // überbucht markierte Person womöglich wieder in die Kapazität (oder ist
  // selbst nicht mehr aktiv). Solche `OverbookReview='Pending'`-Marker werden
  // hier still entfernt — sonst zeigt die Review-Box (und die orange Tabellen-
  // Markierung) jemanden als „über Kapazität", der längst regulär drinsteht.
  const overbookHealRef = React.useRef(false);
  React.useEffect(() => {
    if (overbookHealRef.current) { overbookHealRef.current = false; return; }
    if (!selectedEvent || !selectedEvent.subsiteUrl || !eventServiceRef) return;
    const flagged = registrations.filter(r => r.OverbookReview === 'Pending');
    if (flagged.length === 0) return;
    const isSplit = typeof selectedEvent.durchstarterCapacity === 'number'
      && typeof selectedEvent.funstarterCapacity === 'number'
      && ((selectedEvent.durchstarterCapacity || 0) > 0 || (selectedEvent.funstarterCapacity || 0) > 0);
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const groupOf = (r: SPRegistration): string => r.StarterType || r.PreferredStarterType || '';
    const keyOf = (r: SPRegistration): string => isSplit ? (groupOf(r) || '?') : 'all';
    const capOf = (key: string): number => {
      if (!isSplit) return selectedEvent.maxParticipants || 0;
      if (key === 'Durchstarter') return selectedEvent.durchstarterCapacity || 0;
      if (key === 'Funstarter') return selectedEvent.funstarterCapacity || 0;
      return 0;
    };
    const activeByGroup: Record<string, SPRegistration[]> = {};
    registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0).slice().sort((a, b) => a.Id - b.Id)
      .forEach(r => { const k = keyOf(r); (activeByGroup[k] = activeByGroup[k] || []).push(r); });
    // Stale = nicht (mehr) aktiv ODER Position passt wieder in die Kapazität.
    const stale = flagged.filter(r => {
      const k = keyOf(r); const cap = capOf(k); const bucket = activeByGroup[k] || [];
      const idx = bucket.findIndex(x => x.Id === r.Id);
      if (idx < 0) return true;
      return !(cap > 0 && (idx + 1) > cap);
    });
    if (stale.length === 0) return;
    let cancelled = false;
    (async () => {
      for (const r of stale) {
        try { await eventServiceRef.clearOverbookMark(selectedEvent.subsiteUrl, r.Id); }
        catch (err) { console.warn('[DEX] clearOverbookMark (auto-heal) failed:', err); }
      }
      if (cancelled) return;
      overbookHealRef.current = true; // nächsten Effekt-Lauf nach Reload überspringen
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    })().catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [registrations, selectedEvent?.id]);

  // Soft-Refresh-Sync: wenn `events` durch refreshEvents() aktualisiert wurde
  // (z.B. nach Event-deaktivieren, Event-aktivieren, Edit-save), den lokalen
  // `selectedEvent`-State aus der frischen Liste neu derivieren. Sonst bleibt
  // der Status-Badge (z.B. „Aktiv" vs „Entwurf") nach Toggle stale, weil
  // `selectedEvent` ein eigener useState ist und nicht aus `events` derived.
  React.useEffect(() => {
    if (!selectedEvent) return;
    const fresh = adminEvents.find(e => e.id === selectedEvent.id);
    if (fresh && fresh !== selectedEvent) {
      setSelectedEvent(fresh);
    }
  }, [adminEvents]);

  // v17.9: Map regId → Beitritts-Position (1-basiert, sortiert nach
  // v17.15: joinOrderById useMemo entfernt — wurde mit der Beitritts-#-
  // Spalte (v17.9) eingeführt, die der User in v17.10 wieder rausgeworfen
  // hat. Damit kein Hook mehr, der bei /joinOrder/ stale referenziert war.

  // v20.0 (Audit): ungenutzte Helper-Funktion getRegListUrl entfernt
  // (war seit Jahren nie aufgerufen, lieferte ohnehin nur `<base>/Lists`).

  // v20.1: Self-Check-in auto-aktivieren, falls das Event noch keinen aktiven
  // Token hat (Wizard-Toggle nie gesetzt): Token erzeugen + am Event
  // persistieren. Damit sind QR-PDF + Live-Anzeige grundsätzlich immer
  // verfügbar — der Klick auf die Aktion IST die Aktivierung.
  const ensureSelfCheckInReady = async (ev: DeloitteEvent): Promise<string | null> => {
    if (ev.selfCheckInEnabled && ev.selfCheckInToken) return ev.selfCheckInToken;
    const token = ev.selfCheckInToken || generateSelfCheckInToken();
    let ok = false;
    try {
      ok = await updateEvent(ev.id, { 'SelfCheckInEnabled': true, 'SelfCheckInToken': token });
    } catch { ok = false; }
    if (!ok) {
      showAlert(isDe
        ? 'Self-Check-in konnte nicht aktiviert werden (Speichern am Event fehlgeschlagen). Bitte erneut versuchen.'
        : 'Self check-in could not be activated (saving to the event failed). Please try again.');
      return null;
    }
    return token;
  };

  // v20.2: Self-Check-in-QR-Kachel unter dem Event-Logo + Erklär-/Einstell-Modal.
  // Die Kachel erscheint ab 5 Tagen vor Event-Start ODER sobald QR-Codes
  // versendet wurden; Klick öffnet das Modal mit großem QR, PDF-/Live-Aktionen
  // und dem editierbaren Check-in-Zeitfenster (Von/Bis).
  const [sciModalOpen, setSciModalOpen] = React.useState(false);
  const [sciModalQr, setSciModalQr] = React.useState('');
  const [sciMiniQr, setSciMiniQr] = React.useState('');
  const [sciToken, setSciToken] = React.useState('');
  const [sciFrom, setSciFrom] = React.useState('');
  const [sciTo, setSciTo] = React.useState('');
  const [sciBusy, setSciBusy] = React.useState(false);
  const [sciSaveMsg, setSciSaveMsg] = React.useState('');
  // Mini-QR für die Kachel, sobald das Event einen Token hat (lazy qrcode-Chunk).
  React.useEffect(() => {
    let cancelled = false;
    const token = selectedEvent?.selfCheckInToken;
    if (!token) { setSciMiniQr(''); return undefined; }
    import('qrcode').then(async QRCode => {
      const d = await QRCode.toDataURL(buildStaticCheckInUrl(token), { width: 220, margin: 0 });
      if (!cancelled) setSciMiniQr(d);
    }).catch(() => { /* Kachel zeigt dann das Icon-Fallback */ });
    return () => { cancelled = true; };
  }, [selectedEvent?.selfCheckInToken]);
  const isoToLocalInput = (iso?: string): string => {
    if (!iso) return '';
    const d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    const p = (n: number): string => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
  };
  const openSelfCheckInModal = async (): Promise<void> => {
    if (!selectedEvent || sciBusy) return;
    setSciBusy(true);
    try {
      const token = await ensureSelfCheckInReady(selectedEvent);
      if (!token) return;
      setSciToken(token);
      // v20.3: Von/Bis immer vorbelegen — gespeicherte Werte ODER der
      // Standard (2 Stunden vor Event-Start bis Event-Ende). Der Standard
      // gilt auch zur Laufzeit, solange nichts anderes gespeichert ist
      // (isWithinCheckInWindow) — Anzeige und Verhalten sind damit deckungsgleich.
      const def = defaultCheckInWindow(selectedEvent.startDate, selectedEvent.endDate);
      setSciFrom(isoToLocalInput(selectedEvent.selfCheckInFrom) || (def.opensAt ? isoToLocalInput(def.opensAt.toISOString()) : ''));
      setSciTo(isoToLocalInput(selectedEvent.selfCheckInTo) || (def.closesAt ? isoToLocalInput(def.closesAt.toISOString()) : ''));
      setSciSaveMsg('');
      try {
        const QRCode = await import('qrcode');
        setSciModalQr(await QRCode.toDataURL(buildStaticCheckInUrl(token), { width: 560, margin: 1 }));
      } catch { setSciModalQr(''); }
      setSciModalOpen(true);
    } finally { setSciBusy(false); }
  };
  // v20.3: Der Status-Badge neben dem Event-Titel ist klickbar — Aktiv ⇄
  // Entwurf (ersetzt den früheren Eintrag im Aktionen-Menü). Gleiche Logik
  // wie der alte v11.89-Toggle: IsFictive flippen, beim Live-Schalten
  // Legacy-EventStatus auf 'Active' setzen.
  const toggleDraftStatus = async (): Promise<void> => {
    if (!selectedEvent) return;
    const isDraft = !!selectedEvent.isFictive;
    // v22.15: Abgeschlossen/Abgesagt → zurück auf Aktiv (Reaktivierung).
    // Vorher waren diese Zustände eine Sackgasse — auch für Admins.
    const isFinalState = !isDraft && (selectedEvent.status === 'Completed' || selectedEvent.status === 'Cancelled');
    if (isFinalState) {
      const fromLabel = isDe ? localizeStatus(selectedEvent.status) : selectedEvent.status;
      if (!(await confirmDialog(
        isDe
          ? `Event von „${fromLabel}" wieder auf Aktiv setzen? Danach ist es für die Berechtigten wieder sichtbar und buchbar. Hinweis: Liegt das End-Datum in der Vergangenheit, setzt der automatische Aufräum-Lauf das Event beim nächsten App-Start erneut auf „Abgeschlossen" — dann zuerst das Datum korrigieren.`
          : `Set event from "${fromLabel}" back to Active? It will be visible and bookable for eligible users again. Note: if the end date is in the past, the automatic cleanup will set it back to "Completed" on the next app start — fix the date first in that case.`,
        { title: isDe ? 'Event reaktivieren' : 'Reactivate event', confirmLabel: isDe ? 'Auf Aktiv setzen' : 'Set to Active' },
      ))) return;
      const ok = await updateEvent(selectedEvent.id, { 'EventStatus': 'Active' });
      if (ok) {
        setSelectedEvent(prev => prev ? { ...prev, status: 'Active' } : prev);
        await refreshEvents();
      } else {
        showAlert(isDe
          ? 'Der Status konnte nicht geändert werden. Vermutlich fehlen dir Schreibrechte auf der Event-Liste — bitte einen Haupt-Organizer oder Admin den Status umschalten lassen.'
          : 'The status could not be changed. You probably lack write permission on the event list — please ask a main organizer or admin to switch the status.', { variant: 'error' });
      }
      return;
    }
    const nextIsFictive = !isDraft;
    const confirmMsg = nextIsFictive
      ? (isDe ? 'Event auf "Entwurf" zurücksetzen? Reguläre User sehen das Event danach nicht mehr.' : 'Reset event to "draft"? Regular users will no longer see the event afterwards.')
      : (isDe ? 'Event live schalten? Alle Berechtigten können sich danach anmelden.' : 'Publish event? All eligible users can register afterwards.');
    if (!(await confirmDialog(confirmMsg, { title: isDe ? 'Event-Status ändern' : 'Change event status', confirmLabel: nextIsFictive ? (isDe ? 'Auf Entwurf setzen' : 'Set to draft') : (isDe ? 'Live schalten' : 'Publish') }))) return;
    const patch: Record<string, unknown> = { 'IsFictive': nextIsFictive };
    if (!nextIsFictive) patch['EventStatus'] = 'Active';
    const ok = await updateEvent(selectedEvent.id, patch);
    if (ok) {
      // Badge sofort umschalten — selectedEvent ist lokaler State und wird
      // durch refreshEvents nicht automatisch ersetzt.
      setSelectedEvent(prev => prev ? { ...prev, isFictive: nextIsFictive, ...(nextIsFictive ? {} : { status: 'Active' }) } : prev);
      // v22.67: Beim Live-Schalten eines Events mit Sub-Events werden die
      // Sub-Events automatisch mit live geschaltet (Entwurf → Aktiv) — sonst
      // bliebe das Event sichtbar, aber die Sub-Events wären für Teilnehmer
      // nicht buchbar.
      if (!nextIsFictive) {
        for (const c of childEventsOf(selectedEvent.id)) {
          if (c.isFictive) {
            try { await updateEvent(c.id, { 'IsFictive': false, 'EventStatus': 'Active' }); } catch { /* best-effort */ }
          }
        }
      }
      await refreshEvents();
    } else {
      // v22.14: vorher scheiterte der Klick STUMM — der Organizer dachte,
      // der Status lasse sich nicht ändern, ohne zu erfahren warum.
      showAlert(isDe
        ? 'Der Status konnte nicht geändert werden. Vermutlich fehlen dir Schreibrechte auf der Event-Liste (z.B. als Co-Organizer ohne Organizer-Rolle) — bitte einen Haupt-Organizer oder Admin den Status umschalten lassen.'
        : 'The status could not be changed. You probably lack write permission on the event list (e.g. co-organizer without the organizer role) — please ask a main organizer or admin to switch the status.', { variant: 'error' });
    }
  };

  // v22.5: Einladungsmail — Default-Texte bauen, Entwurf laden/speichern
  // (localStorage pro Event), Modal öffnen, zurücksetzen.
  const inviteDraftKey = (id: string): string => `dex_invite_draft_${id}`;
  const buildInviteDefaults = (ev: DeloitteEvent): { subject: string; heading: string; subheading: string; body: string } => {
    const appUrl = `${siteUrl}/SitePages/DEX.aspx?env=WebView`;
    const linkHtml = `<a href="${appUrl}" style="color:#86bc25;font-weight:600;">${appUrl}</a>`;
    const orgList = (ev.organizers || []).map(s => (s || '').trim()).filter(Boolean);
    const teamLine = isDe ? `Das ${ev.title} Orga Team` : `The ${ev.title} Organizer Team`;
    const signatureNames = orgList.length > 0 ? `${teamLine}<br />${orgList.join('<br />')}` : teamLine;
    const body = isDe
      ? `<p>Hallo,</p>\n<p>wir laden dich herzlich zum Event <strong>${ev.title}</strong> ein.</p>\n<p>Du kannst dich ab sofort über unsere Event-Plattform anmelden:</p>\n<p>${linkHtml}</p>\n<p>Falls du dich im Nachgang doch nicht beteiligen kannst, ist eine <strong>Abmeldung jederzeit über dieselbe Plattform</strong> möglich — bitte gib uns rechtzeitig Bescheid, damit Wartelisten-Plätze nachrücken können.</p>\n<p>Bei Rückfragen meld dich gern bei uns.</p>\n<p>Viele Grüße<br />${signatureNames}</p>`
      : `<p>Hello,</p>\n<p>we would like to invite you to the event <strong>${ev.title}</strong>.</p>\n<p>You can register via our event platform:</p>\n<p>${linkHtml}</p>\n<p>If you change your mind, you can <strong>cancel anytime via the same platform</strong> — please let us know early so people on the waitlist can move up.</p>\n<p>Feel free to reach out if you have any questions.</p>\n<p>Best regards<br />${signatureNames}</p>`;
    return {
      subject: isDe ? `Einladung: ${ev.title}` : `Invitation: ${ev.title}`,
      heading: isDe ? `Einladung zu ${ev.title}` : `Invitation to ${ev.title}`,
      subheading: '',
      body,
    };
  };
  const applyInviteDraftOrDefaults = (ev: DeloitteEvent): void => {
    inviteHydratingRef.current = true;
    let loaded: { subject?: string; heading?: string; subheading?: string; body?: string; target?: string } | null = null;
    try {
      const raw = window.localStorage.getItem(inviteDraftKey(ev.id));
      if (raw) loaded = JSON.parse(raw);
    } catch { /* localStorage evtl. blockiert */ }
    const def = buildInviteDefaults(ev);
    setInviteSubject(loaded && typeof loaded.subject === 'string' ? loaded.subject : def.subject);
    setInviteHeading(loaded && typeof loaded.heading === 'string' ? loaded.heading : def.heading);
    setInviteSubheading(loaded && typeof loaded.subheading === 'string' ? loaded.subheading : def.subheading);
    setInviteBody(loaded && typeof loaded.body === 'string' ? loaded.body : def.body);
    setInviteTarget(loaded && loaded.target === 'audience' ? 'audience' : 'organizer');
    // Hydration-Flag im nächsten Tick freigeben, damit das Auto-Speichern erst
    // auf echte Nutzer-Edits reagiert (nicht auf das initiale Laden).
    window.setTimeout(() => { inviteHydratingRef.current = false; }, 0);
  };
  const openInviteModal = (): void => {
    if (!selectedEvent) return;
    applyInviteDraftOrDefaults(selectedEvent);
    setShowInviteModal(true);
  };
  const resetInviteDraft = (): void => {
    if (!selectedEvent) return;
    try { window.localStorage.removeItem(inviteDraftKey(selectedEvent.id)); } catch { /* */ }
    inviteHydratingRef.current = true;
    const def = buildInviteDefaults(selectedEvent);
    setInviteSubject(def.subject);
    setInviteHeading(def.heading);
    setInviteSubheading(def.subheading);
    setInviteBody(def.body);
    setInviteDraftSaved(false);
    window.setTimeout(() => { inviteHydratingRef.current = false; }, 0);
  };
  // v22.5/v22.6: expliziter „Entwurf speichern"-Klick — schreibt den aktuellen
  // Stand sofort in localStorage und zeigt kurz „Gespeichert".
  const saveInviteDraft = (): void => {
    if (!selectedEvent) return;
    try {
      window.localStorage.setItem(inviteDraftKey(selectedEvent.id), JSON.stringify({
        subject: inviteSubject, heading: inviteHeading, subheading: inviteSubheading,
        body: inviteBody, target: inviteTarget,
      }));
      setInviteDraftSaved(true);
      window.setTimeout(() => setInviteDraftSaved(false), 2500);
    } catch { /* localStorage evtl. blockiert */ }
  };
  // Auto-Speichern, solange das Modal offen ist — beim nächsten Öffnen wird der
  // Entwurf wiederhergestellt.
  React.useEffect(() => {
    if (!showInviteModal || !selectedEvent || inviteHydratingRef.current) return;
    try {
      window.localStorage.setItem(inviteDraftKey(selectedEvent.id), JSON.stringify({
        subject: inviteSubject, heading: inviteHeading, subheading: inviteSubheading,
        body: inviteBody, target: inviteTarget,
      }));
    } catch { /* */ }
  }, [showInviteModal, selectedEvent, inviteSubject, inviteHeading, inviteSubheading, inviteBody, inviteTarget]);

  // v22.9: Massenmail-Entwurf — Default-Texte, laden/speichern (localStorage pro
  // Event), Picker öffnen, zurücksetzen, Testmail an die Organizer.
  const massmailDraftKey = (id: string): string => `dex_massmail_draft_${id}`;
  const buildMassmailDefaults = (ev: DeloitteEvent): { subject: string; heading: string; body: string } => ({
    subject: `${ev.title} - Info`,
    heading: ev.title,
    body: '',
  });
  const applyMassmailDraftOrDefaults = (ev: DeloitteEvent): void => {
    massmailHydratingRef.current = true;
    let loaded: { subject?: string; heading?: string; subheading?: string; body?: string } | null = null;
    try {
      const raw = window.localStorage.getItem(massmailDraftKey(ev.id));
      if (raw) loaded = JSON.parse(raw);
    } catch { /* localStorage evtl. blockiert */ }
    const def = buildMassmailDefaults(ev);
    setEmailSubject(loaded && typeof loaded.subject === 'string' ? loaded.subject : def.subject);
    setEmailHeading(loaded && typeof loaded.heading === 'string' ? loaded.heading : def.heading);
    setMassmailSubheading(loaded && typeof loaded.subheading === 'string' ? loaded.subheading : '');
    setEmailBody(loaded && typeof loaded.body === 'string' ? loaded.body : def.body);
    window.setTimeout(() => { massmailHydratingRef.current = false; }, 0);
  };
  const openMassmailPicker = (): void => {
    if (selectedEvent) applyMassmailDraftOrDefaults(selectedEvent);
    setMassmailAudience('active');
    setMassmailPasteRaw('');
    setMassmailTestMsg(null);
    setMassmailMode('pick');
  };
  const resetMassmailDraft = (): void => {
    if (!selectedEvent) return;
    try { window.localStorage.removeItem(massmailDraftKey(selectedEvent.id)); } catch { /* */ }
    massmailHydratingRef.current = true;
    const def = buildMassmailDefaults(selectedEvent);
    setEmailSubject(def.subject);
    setEmailHeading(def.heading);
    setMassmailSubheading('');
    setEmailBody(def.body);
    setMassmailDraftSaved(false);
    window.setTimeout(() => { massmailHydratingRef.current = false; }, 0);
  };
  const saveMassmailDraft = (): void => {
    if (!selectedEvent) return;
    try {
      window.localStorage.setItem(massmailDraftKey(selectedEvent.id), JSON.stringify({
        subject: emailSubject, heading: emailHeading, subheading: massmailSubheading, body: emailBody,
      }));
      setMassmailDraftSaved(true);
      window.setTimeout(() => setMassmailDraftSaved(false), 2500);
    } catch { /* */ }
  };
  // Auto-Speichern, solange der Massenmail-Editor offen ist.
  React.useEffect(() => {
    if (massmailMode !== 'editor' || !showEmailModal || !selectedEvent || massmailHydratingRef.current) return;
    try {
      window.localStorage.setItem(massmailDraftKey(selectedEvent.id), JSON.stringify({
        subject: emailSubject, heading: emailHeading, subheading: massmailSubheading, body: emailBody,
      }));
    } catch { /* */ }
  }, [massmailMode, showEmailModal, selectedEvent, emailSubject, emailHeading, massmailSubheading, emailBody]);
  // Testmail mit dem aktuellen Stand an die Organizer (zur Kontrolle vor dem
  // echten Massenversand). Geht NICHT an die Teilnehmer.
  const sendMassmailTestToOrganizers = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent) return;
    const orgEmails = (selectedEvent.organizerEmails || []).filter(Boolean);
    const to = orgEmails.length > 0 ? orgEmails.join(';') : (currentUser.email || '');
    if (!to) {
      setMassmailTestMsg(isDe ? 'Keine Organizer-E-Mail hinterlegt — Test nicht möglich.' : 'No organizer email available — test not possible.');
      return;
    }
    if (!emailSubject.trim() || !emailBody.trim()) {
      setMassmailTestMsg(isDe ? 'Bitte Betreff und Text ausfüllen.' : 'Please fill in subject and body.');
      return;
    }
    setMassmailTesting(true);
    setMassmailTestMsg(null);
    try {
      const previewVars: Record<string, string> = { EventTitle: selectedEvent.title, Organizer: (selectedEvent.organizers || []).join(', ') };
      const resolvedSubject = `[TEST] ${replacePlaceholders(emailSubject, previewVars)}`;
      const resolvedHeading = replacePlaceholders(emailHeading, previewVars);
      const resolvedBody = replacePlaceholders(emailBody, previewVars);
      const resolvedSub = massmailSubheading.trim() ? replacePlaceholders(massmailSubheading, previewVars) : `Event ${selectedEvent.title}`;
      const fullBody = wrapTemplate('#86bc25', resolvedHeading, resolvedSub, resolvedBody);
      await eventServiceRef.queueEmail(resolvedSubject, to, 'Organizer (Test)', fullBody, 'Massenmail', selectedEvent.title, selectedEvent.id);
      setMassmailTestMsg(isDe ? `Testmail an die Organizer (${to.split(';').length}) verschickt — bitte Postfach prüfen.` : `Test email sent to the organizers (${to.split(';').length}) — please check the mailbox.`);
    } catch (err) {
      setMassmailTestMsg((isDe ? 'Fehler beim Test-Versand: ' : 'Error during test send: ') + (err instanceof Error ? err.message : String(err)));
    }
    setMassmailTesting(false);
  };

  // v22.6: QR-Versand-Aktionen als benannte Funktionen (vorher inline im Modal) —
  // macht das neue kompakte Querformat-Layout lesbar. Verhalten unverändert.
  // v22.18: pro-Event angepasster QR-Mail-Text — Override-Key 'QRCode' im
  // EmailTemplateOverrides-JSON des Events (übersteht den Wizard-Roundtrip,
  // weil Nicht-Unterstrich-Keys dort erhalten bleiben). QR-Block bleibt fix.
  const getQrMailOverride = (ev: DeloitteEvent | null): QrEmailOverride | undefined => {
    if (!ev) return undefined;
    try {
      const all = JSON.parse(ev.emailTemplateOverrides || '{}');
      const ov = all && all['QRCode'];
      if (ov && (ov.subject || ov.heading || ov.subheading || ov.bodyHtml)) return ov as QrEmailOverride;
    } catch { /* kein Override */ }
    return undefined;
  };
  // Editor öffnen: Felder aus Override (falls vorhanden) oder den Standard-
  // Texten vorbelegen + Beispiel-QR-Block für die Live-Vorschau erzeugen.
  const openQrMailEditor = async (): Promise<void> => {
    if (!selectedEvent) return;
    const ov = getQrMailOverride(selectedEvent);
    const def = qrEmailDefaults(selectedEvent.emailLanguage || 'EN');
    setQrEditSubject((ov && ov.subject) || def.subject);
    setQrEditHeading((ov && ov.heading) || def.heading);
    setQrEditSubheading((ov && ov.subheading) || def.subheading);
    setQrEditBody((ov && ov.bodyHtml) || def.body);
    // Beispiel-QR (eigene Daten) für die Vorschau — gleicher Aufbau wie im Versand.
    const myName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
    const qrData = `DEX|${selectedEvent.eventNumber}|${currentUser.email}`;
    const qrImageHtml = await buildQrImageHtml(qrData);
    setQrEditSampleBlock(buildQrBlockHtml(qrImageHtml, myName, selectedEvent.title));
    // v22.19: Versand-Modal schließen — der Editor zeigt die Versand-Aktionen
    // in einer eigenen linken Spalte (nebeneinander statt übereinander).
    // Beim Schließen des Editors öffnet das Versand-Modal wieder.
    setQrSendModalOpen(false);
    setQrEditOpen(true);
  };
  const closeQrMailEditor = (): void => {
    setQrEditOpen(false);
    setQrSendModalOpen(true);
  };
  // Speichern: Override in das EmailTemplateOverrides-JSON des Events mergen
  // (andere Keys + Piggybacks bleiben erhalten). Entspricht alles den
  // Standard-Texten, wird der Key entfernt (= zurück auf Standard).
  const saveQrMailOverride = async (): Promise<void> => {
    if (!selectedEvent || qrEditSaving) return;
    setQrEditSaving(true);
    try {
      const def = qrEmailDefaults(selectedEvent.emailLanguage || 'EN');
      const isDefault = qrEditSubject.trim() === def.subject.trim()
        && qrEditHeading.trim() === def.heading.trim()
        && qrEditSubheading.trim() === def.subheading.trim()
        && qrEditBody.trim() === def.body.trim();
      let all: Record<string, unknown> = {};
      try { all = JSON.parse(selectedEvent.emailTemplateOverrides || '{}') || {}; } catch { all = {}; }
      if (isDefault) {
        delete all['QRCode'];
      } else {
        all['QRCode'] = { subject: qrEditSubject, heading: qrEditHeading, subheading: qrEditSubheading, bodyHtml: qrEditBody };
      }
      const json = JSON.stringify(all);
      const ok = await updateEvent(selectedEvent.id, { 'EmailTemplateOverrides': json });
      if (ok) {
        setSelectedEvent(prev => prev ? { ...prev, emailTemplateOverrides: json } : prev);
        await refreshEvents();
        showAlert(isDe
          ? (isDefault
            ? 'QR-Mail auf den Standardtext zurückgesetzt.'
            : 'QR-Mail-Text gespeichert — gilt ab jetzt für Vorschau, Versand UND den automatischen QR-Versand bei neuen Anmeldungen.')
          : (isDefault ? 'QR email reset to the default text.' : 'QR email text saved — now used for preview, sending AND the automatic QR send for new registrations.'), { variant: 'success' });
      } else {
        showAlert(isDe ? 'Speichern fehlgeschlagen — vermutlich fehlen Schreibrechte auf der Event-Liste.' : 'Saving failed — you probably lack write permission on the event list.', { variant: 'error' });
      }
    } finally { setQrEditSaving(false); }
  };
  const buildQrImageHtml = async (qrData: string): Promise<string> => {
    let qrImageHtml = `<p style="font-family:monospace;font-size:1.2rem;background:#f5f5f5;padding:12px;border-radius:8px;text-align:center;">${qrData}</p>`;
    try {
      const QRCode = await import('qrcode');
      const qrDataUrl = await QRCode.toDataURL(qrData, { width: 300, margin: 2 });
      qrImageHtml = `<img src="${qrDataUrl}" alt="QR-Code" style="width:300px;max-width:100%;height:auto;" />`;
    } catch { /* */ }
    return qrImageHtml;
  };
  const qrPreviewAction = async (): Promise<void> => {
    if (!selectedEvent) return;
    setQrPreviewLoading(true);
    try {
      const orgEmail = currentUser.email;
      const orgFullName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || orgEmail;
      const orgFirstName = currentUser.firstName || orgFullName.split(/\s+/)[0] || orgFullName;
      const qrData = `DEX|${selectedEvent.eventNumber}|${orgEmail}`;
      const qrImageHtml = await buildQrImageHtml(qrData);
      const emailData = qrCodeEmail(orgFirstName, selectedEvent.title, qrImageHtml, selectedEvent.emailLanguage || 'EN', orgFullName, getQrMailOverride(selectedEvent));
      let eventOrb = '';
      try {
        const ov = JSON.parse(selectedEvent.emailTemplateOverrides || '{}');
        if (ov && typeof ov._eventLogo === 'string') eventOrb = ov._eventLogo;
      } catch { /* */ }
      const previewBody = emailData.body.replace(/\{\{ORB_URL\}\}/g, eventOrb || getCachedOrbBase64() || '');
      setQrPreviewSubject(emailData.subject);
      setQrPreviewHtml(previewBody);
      setQrPreviewOpen(true);
    } finally { setQrPreviewLoading(false); }
  };
  // v22.19: optionaler liveOverride — der Test-Versand aus dem Mail-Editor
  // nutzt den AKTUELLEN (ggf. ungespeicherten) Editor-Text, damit Test = Vorschau.
  const qrTestSendAction = async (liveOverride?: QrEmailOverride): Promise<void> => {
    if (!eventServiceRef || !selectedEvent) return;
    setIsSendingQR(true); setQrSendResult(null); setQrSentCount(0);
    try {
      const orgEmail = currentUser.email;
      const orgFullName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || orgEmail;
      const orgFirstName = currentUser.firstName || orgFullName.split(/\s+/)[0] || orgFullName;
      const qrData = `DEX|${selectedEvent.eventNumber}|${orgEmail}`;
      const qrImageHtml = await buildQrImageHtml(qrData);
      const emailData = qrCodeEmail(orgFirstName, selectedEvent.title, qrImageHtml, selectedEvent.emailLanguage || 'EN', orgFullName, liveOverride || getQrMailOverride(selectedEvent));
      await eventServiceRef.queueEmail(emailData.subject, orgEmail, orgFullName, emailData.body, 'QRCode', selectedEvent.title, selectedEvent.id);
      setQrSendResult(isDe
        ? `Test-Mail an ${orgEmail} verschickt — bitte in deinem Postfach prüfen.`
        : `Test email sent to ${orgEmail} — please check your mailbox.`);
    } catch (err) {
      setQrSendResult((isDe ? 'Fehler beim Test-Versand: ' : 'Error during test send: ') + (err instanceof Error ? err.message : String(err)));
    }
    setIsSendingQR(false);
  };
  const qrFullSendAction = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent) return;
    const eligible = registrations.filter(r => r.Status === 'Angemeldet');
    if (eligible.length === 0) {
      setQrSendResult(isDe ? 'Alle Teilnehmer haben bereits einen QR-Code — nichts zu senden.' : 'All participants already have a QR code — nothing to send.');
      return;
    }
    if (!(await confirmDialog(isDe ? `QR-Code an ${eligible.length} Teilnehmer ohne Code senden?` : `Send the QR code to ${eligible.length} participants without a code?`, { confirmLabel: isDe ? 'Senden' : 'Send' }))) return;
    setIsSendingQR(true); setQrSendResult(null); setQrSentCount(0);
    let sent = 0; let extCount = 0;
    for (const reg of eligible) {
      const qrData = `DEX|${selectedEvent.eventNumber}|${reg.ParticipantEmail}`;
      const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
      const firstName = reg.Vorname || (reg.ParticipantName || '').trim().split(/\s+/)[0] || name;
      const qrImageHtml = await buildQrImageHtml(qrData);
      const emailData = qrCodeEmail(firstName, selectedEvent.title, qrImageHtml, selectedEvent.emailLanguage || 'EN', name, getQrMailOverride(selectedEvent));
      const isExternal = !!reg.ParticipantEmail && !/@(.*\.)?deloitte\.de$/i.test(reg.ParticipantEmail);
      if (isExternal) {
        const orgEmails = (selectedEvent.organizerEmails || []).filter(Boolean);
        const orgRecipient = orgEmails.length > 0 ? orgEmails.join(';') : currentUser.email;
        const orgSubject = `[Externer Teilnehmer] QR-Code für ${name} — ${selectedEvent.title}`;
        const qrExternalHint = `<div style="margin:0 0 16px;padding:12px 16px;background:#fff3e0;border:1px solid #ed8b00;border-radius:8px;font-size:13px;line-height:1.55;color:#7a4a00;">`
          + `<strong>QR-Code für externen Teilnehmer.</strong><br>`
          + `Eigentlich für <strong>${reg.ParticipantEmail}</strong> (${name}). Da externe Adressen keinen Mail-Versand bekommen, landet der QR-Code bei dir — drucke ihn aus oder leite die Mail intern an den Empfänger weiter (Datenschutzrichtlinien Deloitte Deutschland beachten).`
          + `</div>`;
        const qrBody = injectIntoEmailContent(emailData.body, qrExternalHint);
        await eventServiceRef.queueEmail(orgSubject, orgRecipient, 'Organizer', qrBody, 'QRCode', selectedEvent.title, selectedEvent.id);
        extCount++;
      } else {
        await eventServiceRef.queueEmail(emailData.subject, reg.ParticipantEmail, name, emailData.body, 'QRCode', selectedEvent.title, selectedEvent.id);
      }
      if (selectedEvent.subsiteUrl) {
        await eventServiceRef.setQRSentStatus(selectedEvent.subsiteUrl, reg.Id);
      }
      sent++; setQrSentCount(sent);
    }
    // v21: Erster Massen-Versand startet die QR-Phase (AutoSendQRCode=true).
    try { await eventServiceRef.updateEvent(parseInt(selectedEvent.id, 10), { AutoSendQRCode: true }); } catch { /* */ }
    const regs = await getAllRegistrations(selectedEvent.id);
    setRegistrations(regs);
    setIsSendingQR(false);
    setQrSendResult(extCount > 0
      ? (isDe
        ? `${sent} QR-Codes verschickt (davon ${extCount} an dich/Organizer umgeleitet — externe Adressen).`
        : `${sent} QR codes sent (${extCount} of them redirected to you/the organizer — external addresses).`)
      : (isDe ? `${sent} QR-Codes verschickt.` : `${sent} QR codes sent.`));
  };

  const saveSelfCheckInWindow = async (): Promise<void> => {
    if (!selectedEvent || sciBusy) return;
    if (sciFrom && sciTo && new Date(sciFrom).getTime() >= new Date(sciTo).getTime()) {
      showAlert(isDe ? '„Bis" muss zeitlich nach „Von" liegen.' : '"Until" must be after "From".');
      return;
    }
    setSciBusy(true);
    try {
      const ok = await updateEvent(selectedEvent.id, {
        'SelfCheckInFrom': sciFrom ? new Date(sciFrom).toISOString() : null,
        'SelfCheckInTo': sciTo ? new Date(sciTo).toISOString() : null,
      });
      setSciSaveMsg(ok
        ? (isDe ? 'Zeitfenster gespeichert.' : 'Time window saved.')
        : (isDe ? 'Speichern fehlgeschlagen — bitte erneut versuchen.' : 'Saving failed — please try again.'));
    } finally { setSciBusy(false); }
  };

  // Danger-Zone-Modal als gemeinsames Element — wird in BEIDEN Render-Branches
  // (Event-Liste und Event-Detail) eingehängt, sonst läuft der Löschen-Klick auf
  // der Event-Liste ins Leere (Bug v9.x: Modal war nur im Detail-Branch gerendert).
  const dangerZoneModal: React.ReactElement | null = confirmDeleteEvent ? (() => {
    const expected = (confirmDeleteEvent.title || '').trim().toLowerCase();
    const typed = confirmDeleteText.trim().toLowerCase();
    const matches = !!expected && expected === typed;
    const close = (): void => { setConfirmDeleteEvent(null); setConfirmDeleteText(''); };
    return (
      <div
        style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1300,
          display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
        }}
        onClick={() => { if (!isDeleting) close(); }}
      >
        <div
          className="card"
          style={{ width: '100%', maxWidth: 560, padding: 24, borderRadius: 16, background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.3)', borderTop: '4px solid var(--dex-red, #c00)' }}
          onClick={e => e.stopPropagation()}
        >
          <div className="flex-between mb-16">
            <h3 style={{ margin: 0, color: 'var(--dex-red, #c00)', display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Trash2 size={20} /> {isDe ? 'Danger Zone — Event löschen' : 'Danger Zone — Delete event'}
            </h3>
            <button
              onClick={close}
              disabled={isDeleting}
              style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: isDeleting ? 'not-allowed' : 'pointer', color: 'var(--dex-gray-500)' }}
            ><X size={20} /></button>
          </div>
          <p style={{ margin: '0 0 12px', fontSize: '0.88rem', lineHeight: 1.55 }}>
            {isDe
              ? <>Du bist dabei das Event <strong>&bdquo;{confirmDeleteEvent.title}&ldquo;</strong> zu löschen.</>
              : <>You are about to delete the event <strong>&bdquo;{confirmDeleteEvent.title}&ldquo;</strong>.</>}
          </p>
          <ul style={{ margin: '0 0 16px', fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.55, paddingLeft: 18 }}>
            <li>{isDe ? 'Subsite (inkl. Teilnehmerliste) und Event-Item wandern in den SharePoint-Papierkorb.' : 'Subsite (incl. attendee list) and event item move to the SharePoint recycle bin.'}</li>
            <li>{isDe ? 'Wiederherstellung durch einen Admin innerhalb von 93 Tagen möglich (zweistufig).' : 'A site collection admin can restore within 93 days (two-stage).'}</li>
            <li>{isDe ? 'Outlook-Termin wird über den Power-Automate-Flow gelöscht.' : 'Outlook calendar event will be deleted via the Power Automate flow.'}</li>
            <li>{isDe ? 'Diese Aktion wird im DEX_ChangeLog mit deinem Namen + Datum protokolliert.' : 'This action is logged in DEX_ChangeLog with your name + date.'}</li>
          </ul>
          <div style={{ background: 'rgba(218,41,28,0.06)', border: '1px solid var(--dex-red, #c00)', padding: 12, borderRadius: 8, marginBottom: 16 }}>
            <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 600, marginBottom: 6 }}>
              {isDe
                ? <>Tippe zur Bestätigung den Event-Titel <strong>kleingeschrieben</strong> ein:</>
                : <>Type the event title <strong>in lowercase</strong> to confirm:</>}
            </label>
            <code style={{ display: 'inline-block', padding: '4px 8px', background: '#fff', borderRadius: 4, fontSize: '0.85rem', marginBottom: 8, wordBreak: 'break-all' }}>{expected}</code>
            <input
              className="form-input"
              value={confirmDeleteText}
              onChange={e => setConfirmDeleteText(e.target.value)}
              placeholder={isDe ? 'Event-Titel kleingeschrieben…' : 'Event title in lowercase…'}
              disabled={isDeleting}
              autoFocus
              style={{ width: '100%' }}
            />
          </div>
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
            <button
              type="button"
              className="btn btn-outline"
              onClick={close}
              disabled={isDeleting}
            >{isDe ? 'Abbrechen' : 'Cancel'}</button>
            <button
              type="button"
              className="btn btn-danger"
              disabled={!matches || isDeleting}
              style={{
                background: matches && !isDeleting ? 'var(--dex-red, #c00)' : 'var(--dex-gray-300)',
                color: '#fff',
                border: 'none',
                cursor: matches && !isDeleting ? 'pointer' : 'not-allowed',
                padding: '8px 16px',
              }}
              onClick={async () => {
                if (!matches || !confirmDeleteEvent) return;
                setIsDeleting(true);
                setDeletingId(confirmDeleteEvent.id);
                try {
                  await deleteEvent(confirmDeleteEvent.id);
                } finally {
                  setIsDeleting(false);
                  setDeletingId(null);
                  close();
                }
              }}
            >
              <Trash2 size={14} /> {isDeleting ? (isDe ? 'Wird gelöscht…' : 'Deleting…') : (isDe ? 'Endgültig löschen' : 'Delete')}
            </button>
          </div>
        </div>
      </div>
    );
  })() : null;

  // ChangeLog-/Audit-Log-Modal als gemeinsames Element — wie das Danger-Zone-
  // Modal muss auch dieses in BEIDEN Render-Branches verfügbar sein, sonst
  // öffnet sich der "Audit log"-Button auf der Event-Liste ins Leere.
  const changeLogModal: React.ReactElement | null = showChangeLogModal ? (() => {
    const fa = changeLogFilterAction.toLowerCase().trim();
    const fe = changeLogFilterEvent.toLowerCase().trim();
    const fac = changeLogFilterActor.toLowerCase().trim();
    // Self-Action-Erkennung: Marker im Details-JSON ODER (als Fallback)
    // Actor-E-Mail == Target-E-Mail (User hat sich selbst registriert/abgemeldet).
    const isSelfAction = (e: typeof changeLogEntries[number]): boolean => {
      const d = (e.Details || '').toLowerCase();
      if (d.indexOf('"asactor":"self"') >= 0) return true;
      // Fallback: bei Participant-Aktionen ohne expliziten Marker prüfen wir,
      // ob Actor und Ziel dieselbe Person sind (Target trägt den Namen des
      // Participants, ActorName ist "Nachname, Vorname").
      const action = (e.Action || '').toLowerCase();
      if (action.indexOf('participant') < 0) return false;
      const tgt = (e.TargetName || '').toLowerCase().trim();
      const actorName = (e.ActorName || '').toLowerCase().trim();
      if (!tgt || !actorName) return false;
      // ActorName-Format "Nachname, Vorname" → in "Vorname Nachname" umdrehen
      const parts = actorName.split(',').map(s => s.trim());
      const flipped = parts.length === 2 ? `${parts[1]} ${parts[0]}` : actorName;
      return tgt === flipped || tgt === actorName;
    };
    const filtered = changeLogEntries.filter(e =>
      (!fa || (e.Action || '').toLowerCase().indexOf(fa) >= 0) &&
      (!fe || ((e.EventTitle || '').toLowerCase().indexOf(fe) >= 0 || (e.TargetName || '').toLowerCase().indexOf(fe) >= 0)) &&
      (!fac || (e.ActorName || '').toLowerCase().indexOf(fac) >= 0 || (e.ActorEmail || '').toLowerCase().indexOf(fac) >= 0) &&
      (!changeLogHideSelf || !isSelfAction(e))
    );
    const fmtDate = (iso: string): string => {
      if (!iso) return '';
      try { return new Date(iso).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); }
      catch { return iso; }
    };
    const actionColor = (a: string): string => {
      if (a.indexOf('Deleted') >= 0) return 'var(--dex-red, #c00)';
      if (a.indexOf('Created') >= 0) return 'var(--dex-green-dark)';
      if (a.indexOf('Cancelled') >= 0) return 'var(--dex-orange)';
      return 'var(--dex-gray-700)';
    };
    // v19.30 (Feature D): Details lesbar rendern. Bei ParticipantUpdated &
    // ähnlichen Aktionen steht im Details-JSON `{ changes: { Feld: { old, new } } }`.
    // Wir zeigen pro Feld eine „Feld: alt → neu"-Zeile statt rohes JSON. Bei
    // anderen/unstrukturierten Details fallen wir auf den Klartext zurück.
    const fmtVal = (v: unknown): string => {
      if (v === undefined || v === null || v === '') return '—';
      return String(v);
    };
    const renderDetails = (raw: string): React.ReactNode => {
      if (!raw) return <span style={{ color: 'var(--dex-gray-400)' }}>—</span>;
      let parsed: unknown;
      try { parsed = JSON.parse(raw); } catch { return <span>{raw}</span>; }
      if (!parsed || typeof parsed !== 'object') return <span>{raw}</span>;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const obj = parsed as any;
      const changes = obj.changes;
      if (changes && typeof changes === 'object' && Object.keys(changes).length > 0) {
        return (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            {Object.keys(changes).map(field => {
              const c = changes[field] || {};
              return (
                <div key={field} style={{ fontFamily: 'inherit', fontSize: '0.78rem', lineHeight: 1.4 }}>
                  <strong style={{ color: 'var(--dex-gray-800)' }}>{field}:</strong>{' '}
                  <span style={{ color: 'var(--dex-gray-500)', textDecoration: 'line-through' }}>{fmtVal(c.old)}</span>
                  {' → '}
                  <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>{fmtVal(c.new)}</span>
                </div>
              );
            })}
          </div>
        );
      }
      // Kein changes-Block: übrige aussagekräftige Schlüssel kompakt zeigen
      // (z.B. asActor / via / scope), sonst das rohe JSON.
      const keys = Object.keys(obj).filter(k => k !== 'asActor');
      if (keys.length === 0) {
        return <span style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{isDe ? '(keine Detailänderungen)' : '(no detail changes)'}</span>;
      }
      return (
        <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>
          {keys.map(k => `${k}: ${fmtVal(obj[k])}`).join(' · ')}
        </span>
      );
    };
    return (
      <div
        style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}
        onClick={() => setShowChangeLogModal(false)}
      >
        <div
          className="card"
          style={{ width: '100%', maxWidth: 1200, maxHeight: '90vh', overflow: 'auto', padding: 24, borderRadius: 16, background: '#fff' }}
          onClick={e => e.stopPropagation()}
        >
          <div className="flex-between mb-16">
            <h3 style={{ margin: 0 }}>
              <FileText size={18} /> {isDe ? 'Audit-Log (DEX_ChangeLog)' : 'Audit log (DEX_ChangeLog)'}
            </h3>
            <button onClick={() => setShowChangeLogModal(false)} style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}>
              <X size={20} />
            </button>
          </div>
          <p style={{ margin: '0 0 8px', fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
            {isDe
              ? <>Letzte <strong>{changeLogEntries.length}</strong> Einträge ({filtered.length} sichtbar). Schreibrechte: alle authentifizierten User; Leserechte: Organizer + Admin.</>
              : <>Last <strong>{changeLogEntries.length}</strong> entries ({filtered.length} visible). Write access: all authenticated users; read access: organizer + admin.</>}
          </p>
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.82rem', color: 'var(--dex-gray-700)', marginBottom: 12, cursor: 'pointer' }}>
            <input
              type="checkbox"
              checked={changeLogHideSelf}
              onChange={e => setChangeLogHideSelf(e.target.checked)}
            />
            {isDe
              ? 'Eigenaktionen der User ausblenden (nur Aktionen von Organizer/Admin anzeigen)'
              : 'Hide user self-actions (show only actions performed by organizer/admin)'}
          </label>
          {changeLogLoading && (
            <p style={{ textAlign: 'center', padding: 16, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
              {isDe ? 'Lade Einträge…' : 'Loading entries…'}
            </p>
          )}
          {!changeLogLoading && (
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem' }}>
                <thead style={{ position: 'sticky', top: 0, background: 'var(--dex-gray-50)' }}>
                  <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Datum' : 'Date'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Aktion' : 'Action'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Ziel' : 'Target'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Event' : 'Event'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Wer' : 'Actor'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Details' : 'Details'}</th>
                  </tr>
                  <tr style={{ borderBottom: '1px solid var(--dex-gray-200)', background: '#fff' }}>
                    <th style={{ padding: 4 }} />
                    <th style={{ padding: 4 }}>
                      <input value={changeLogFilterAction} onChange={e => setChangeLogFilterAction(e.target.value)} placeholder={isDe ? 'z.B. Deleted' : 'e.g. Deleted'} style={{ width: '100%', padding: '4px 6px', border: '1px solid var(--dex-gray-200)', borderRadius: 4, fontSize: '0.75rem' }} />
                    </th>
                    <th style={{ padding: 4 }} />
                    <th style={{ padding: 4 }}>
                      <input value={changeLogFilterEvent} onChange={e => setChangeLogFilterEvent(e.target.value)} placeholder={isDe ? 'Event-/Ziel-Name' : 'event/target'} style={{ width: '100%', padding: '4px 6px', border: '1px solid var(--dex-gray-200)', borderRadius: 4, fontSize: '0.75rem' }} />
                    </th>
                    <th style={{ padding: 4 }}>
                      <input value={changeLogFilterActor} onChange={e => setChangeLogFilterActor(e.target.value)} placeholder={isDe ? 'Name/E-Mail' : 'name/email'} style={{ width: '100%', padding: '4px 6px', border: '1px solid var(--dex-gray-200)', borderRadius: 4, fontSize: '0.75rem' }} />
                    </th>
                    <th style={{ padding: 4 }} />
                  </tr>
                </thead>
                <tbody>
                  {filtered.map(e => (
                    <tr key={e.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                      <td style={{ padding: 6, color: 'var(--dex-gray-600)', whiteSpace: 'nowrap' }}>{fmtDate(e.Created)}</td>
                      <td style={{ padding: 6, color: actionColor(e.Action), fontWeight: 600 }}>{e.Action}</td>
                      <td style={{ padding: 6 }}>{e.TargetName || e.TargetId || '-'}</td>
                      <td style={{ padding: 6, color: 'var(--dex-gray-700)' }}>{e.EventTitle || '-'}</td>
                      <td style={{ padding: 6 }}>
                        {e.ActorName || e.ActorEmail || '-'}
                        {e.ActorEmail && e.ActorName && (
                          <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{e.ActorEmail}</div>
                        )}
                      </td>
                      <td style={{ padding: 6, color: 'var(--dex-gray-600)', fontSize: '0.75rem', maxWidth: 360, wordBreak: 'break-word' }}>{renderDetails(e.Details)}</td>
                    </tr>
                  ))}
                  {filtered.length === 0 && (
                    <tr><td colSpan={6} style={{ padding: 16, textAlign: 'center', color: 'var(--dex-gray-500)' }}>
                      {isDe ? 'Keine Einträge passen zum Filter.' : 'No entries match the filter.'}
                    </td></tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    );
  })() : null;

  // v6.20: Access-Gate — wer weder Admin noch Organizer eines Events noch QR-Scanner
  // eines Events ist, darf die Admin-Seite gar nicht erst sehen. Zeigt eine klare
  // "Kein Zugriff"-Meldung statt einer leeren Event-Liste.
  if (!selectedEvent && !isAdmin && adminEvents.length === 0) {
    return (
      <div className="page-container" role="main">
        <h2 className="mb-16">{t('admin.title')}</h2>
        <div className="card" style={{ padding: 48, textAlign: 'center' }}>
          <p style={{ color: 'var(--dex-gray-700)', marginBottom: 8, fontWeight: 600 }}>
            {t('admin.noaccess.title') || 'Kein Zugriff'}
          </p>
          <p style={{ color: 'var(--dex-gray-500)', fontSize: '0.88rem', maxWidth: 520, margin: '0 auto' }}>
            {t('admin.noaccess.msg')}
          </p>
          <div style={{ marginTop: 24, display: 'flex', gap: 12, justifyContent: 'center', flexWrap: 'wrap' }}>
            {/* v7.2: Organizer ohne eigenes Event sehen hier einen direkten
                Shortcut zum Event-Erstellen — sonst sitzen sie in dieser
                Sackgasse ohne sichtbaren nächsten Schritt. Admins sehen den
                Button nicht, weil sie bereits alle Events in adminEvents haben. */}
            {currentUserRole !== 'User' && (
              <button className="btn btn-primary" onClick={() => navigate('create-event')}>
                + {t('admin.newevent')}
              </button>
            )}
            <button className="btn btn-secondary" onClick={() => navigate('landing')}>
              {t('reg.backtoevents') || 'Zurück'}
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (!selectedEvent) {
    // Event-Auswahl
    return (
      <div className="page-container" role="main" style={{ maxWidth: 1200, marginLeft: 'auto', marginRight: 'auto' }}>
        <style>{`@keyframes dex-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
        <h2 style={{ margin: '0 0 16px' }}>{t('admin.title')}</h2>
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 16 }}>
          {isAdmin && (
            <button className="btn btn-secondary" onClick={() => navigate('participants')} style={{ fontSize: '0.85rem' }}>
              <Users size={16} /> {t('admin.participants')}
            </button>
          )}
          {isAdmin && (
            <button className="btn btn-secondary" onClick={() => navigate('flowcharts')} style={{ fontSize: '0.85rem' }}>
              ↻ {t('admin.processes')}
            </button>
          )}
          {isAdmin && (
            <button className="btn btn-secondary" onClick={openChangeLog} style={{ fontSize: '0.85rem' }}>
              <FileText size={16} /> {isDe ? 'Audit-Log' : 'Audit log'}
            </button>
          )}
          <a
            href={`${siteUrl}/Lists/DEX_Events/AllItems.aspx`}
            target="_blank"
            rel="noopener noreferrer"
            className="btn btn-secondary"
            style={{ fontSize: '0.85rem', textDecoration: 'none' }}
          >
            <FileText size={16} /> {t('admin.splist')}
          </a>
          <button className="btn btn-primary" onClick={() => navigate('create-event')} style={{ fontSize: '0.85rem' }}>
            <Plus size={16} /> {t('admin.newevent')}
          </button>
        </div>

        {isEventsLoading ? (
          <div className="card text-center" style={{ padding: 48 }}>
            <div style={{ display: 'flex', justifyContent: 'center', marginBottom: 16 }}>
              <div style={{
                width: 40, height: 40, border: '4px solid var(--dex-gray-200)',
                borderTop: '4px solid var(--dex-green)', borderRadius: '50%',
                animation: 'dexOrbSpin 1s linear infinite',
              }} />
            </div>
            <p style={{ color: 'var(--dex-gray-400)' }}>Events werden geladen...</p>
          </div>
        ) : adminEvents.length === 0 ? (
          <div className="card text-center" style={{ padding: 48 }}>
            <p style={{ color: 'var(--dex-gray-400)' }}>{t('admin.noevents')}</p>
            <button className="btn btn-primary mt-24" onClick={() => navigate('create-event')}>
              {t('create.submit')}
            </button>
          </div>
        ) : (
          <>
          {(() => {
            const renderEventCard = (event: DeloitteEvent, opts?: { muted?: boolean }): React.ReactElement => (
              <div
                key={event.id}
                className="card card-clickable"
                style={{ padding: '20px 24px', cursor: 'pointer', opacity: opts?.muted ? 0.85 : 1 }}
              >
                <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: 12 }}>
                  <div onClick={() => handleSelectEvent(event)} style={{ flex: '1 1 200px', display: 'flex', alignItems: 'center', gap: 16, cursor: 'pointer' }}>
                    {/* v9.11: Thumbnail-Container immer rendern (auch wenn kein Bild
                        gesetzt ist) — sonst rutscht der Text nach links und die
                        Reihen wirken inkonsistent neben Reihen mit Bild. */}
                    <div style={{
                      width: 60, height: 40, borderRadius: 'var(--dex-radius)', flexShrink: 0,
                      background: event.imageUrl
                        ? `url(${event.imageUrl}) center/cover no-repeat`
                        : 'linear-gradient(135deg, var(--dex-gray-200), var(--dex-gray-100))',
                      filter: opts?.muted ? 'grayscale(0.4)' : 'none',
                      display: 'flex', alignItems: 'center', justifyContent: 'center',
                      color: 'var(--dex-gray-400)', fontSize: '0.7rem',
                    }}>
                      {!event.imageUrl && '—'}
                    </div>
                    <div>
                    <h3 style={{ marginBottom: 4 }}>{event.title}</h3>
                    <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', margin: 0 }}>
                      {formatDate(event.startDate)} - {formatDate(event.endDate)}
                      {event.location ? ` · ${event.location}` : ''}
                    </p>
                    <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)', margin: '2px 0 0' }}>
                      Organizer: {event.organizers.map(o => { const p = o.split(',').map(s => s.trim()); return p.length === 2 ? `${p[1]} ${p[0]}` : o; }).join(', ')}
                    </p>
                    </div>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                    <span style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                      {/* v9.10: B2Run-Events haben maxParticipants=0 weil die Kapazität
                          auf durchstarter+funstarter aufgeteilt ist — Summe als
                          effektive Kapazität anzeigen statt "∞". */}
                      {(() => {
                        const split = (event.durchstarterCapacity || 0) + (event.funstarterCapacity || 0);
                        const isSplitEv = (event.durchstarterCapacity || 0) > 0 && (event.funstarterCapacity || 0) > 0;
                        const eff = event.maxParticipants && event.maxParticipants > 0 ? event.maxParticipants : split;
                        // v22.5: Bei Events mit zwei Teilnehmergruppen keine Überbuchung
                        // in der Listenzeile zeigen — die belegte Zahl auf die Kapazität
                        // deckeln (die echte Überbuchungszahl bleibt dem Werkzeug
                        // „Überbuchung prüfen" im Event-Detail vorbehalten).
                        const shown = (isSplitEv && eff > 0) ? Math.min(event.currentParticipants, eff) : event.currentParticipants;
                        return `${shown}/${eff || '∞'} Teilnehmer`;
                      })()}
                    </span>
                    {/* v9.20: Status-Badge mit Entwurfs-Override.
                        Wenn das Event als Entwurf markiert ist, wird "ENTWURF"
                        statt des EventStatus angezeigt — für den Organizer
                        ist dieser Hinweis wichtiger als der technische Status. */}
                    <span className="badge" style={{
                      background: event.isFictive ? 'rgba(237,139,0,0.15)' : getStatusColor(event.status) + '22',
                      color: event.isFictive ? 'var(--dex-orange-dark, #b35a00)' : getStatusColor(event.status),
                      fontWeight: 600,
                    }}>
                      {event.isFictive ? 'ENTWURF' : (isDe ? localizeStatus(event.status) : event.status)}
                    </span>
                    {/* v10.20 / v11.9: Migrations-Button für Legacy-B2Run-Events.
                        Erkennt das Event als 'altes B2Run' wenn entweder
                        type === 'B2Run' (alte EventType-Spalte) ODER mind.
                        ein b2run_*-Custom-Field in den eventSpecificFields
                        steht. Damit erscheint der Knopf auch wenn die alte
                        EventType-Spalte aus DEX_Events bereits gelöscht
                        wurde — entscheidend ist die b2run_*-Spur in der
                        Felder-Konfiguration. Klick: entfernt b2run_*-Fields
                        aus customFields, persistiert 'Durchstarter' /
                        'Funstarter' als Gruppen-Labels, setzt EventType
                        best-effort auf 'Other'. Bestehende Anmeldungen,
                        Wartelisten und Sub-Events bleiben unverändert. */}
                    {isAdmin && (event.type === 'B2Run' || (event.eventSpecificFields || []).some(f => (f.id || '').toLowerCase().startsWith('b2run_'))) && (
                      <button
                        className="btn btn-secondary"
                        style={{ fontSize: '0.8rem', padding: '6px 12px', color: 'var(--dex-green-dark, #4a7c1f)' }}
                        title={isDe
                          ? 'Auf neues Standard-Event-Schema migrieren (Type entfernen, Labels Durchstarter/Funstarter explizit speichern). Bestehende Anmeldungen bleiben unverändert.'
                          : 'Migrate to the new standard event schema (drop type, persist Durchstarter/Funstarter labels). Existing registrations remain unchanged.'}
                        onClick={async (e) => {
                          e.stopPropagation();
                          if (!eventServiceRef) return;
                          // v11.9: Migration nimmt jetzt auch Legacy-B2Run-
                          // Sub-Events mit. Wir scannen alle Child-Events
                          // (childEventsOf) und migrieren die mit, die
                          // entweder type='B2Run' oder mind. ein b2run_*-
                          // Custom-Field haben.
                          const kids = childEventsOf(event.id);
                          const kidsToMigrate = kids.filter(k => k.type === 'B2Run' || (k.eventSpecificFields || []).some(f => (f.id || '').toLowerCase().startsWith('b2run_')));
                          const kidsHint = kidsToMigrate.length > 0
                            ? (isDe
                                ? `\n\nEs werden zusätzlich ${kidsToMigrate.length} Sub-Event(s) mitmigriert: ${kidsToMigrate.map(k => '„' + (k.title || '?') + '"').join(', ')}.`
                                : `\n\nAdditionally ${kidsToMigrate.length} sub-event(s) will be migrated: ${kidsToMigrate.map(k => '"' + (k.title || '?') + '"').join(', ')}.`)
                            : '';
                          const msg = isDe
                            ? `Event "${event.title}" auf Standard-Schema migrieren?\n\n• Type "B2Run" wird entfernt — Event sieht aus wie ein normales Deloitte-Event.\n• Bezeichnungen "Durchstarter" / "Funstarter" werden als Gruppen-Labels gespeichert (kannst du im Wizard frei ändern).\n• Falls Leistungsnachweis-Pflicht aktiv war: wird in ein reguläres Custom-Field „Leistungsnachweis vorhanden" (Checkbox, Pflicht, nur für Gruppe A) umgewandelt — bleibt also als richtige Frage erhalten.\n• Hardcoded Startblock-Mapping pro Gruppe wird ersatzlos entfernt. Bei Bedarf als Custom-Field mit Gruppen-Bindung wieder anlegen.\n• b2run_*-Custom-Fields (Altersgruppe, T-Shirt-Größe, Mobilnummer etc.) BLEIBEN als generische Custom-Fields erhalten — du kannst sie danach im Wizard umbenennen oder löschen, wenn nicht mehr gebraucht.\n• Anmeldungen, Wartelisten und Sub-Events bleiben inhaltlich unverändert.${kidsHint}`
                            : `Migrate event "${event.title}" to the standard schema?\n\n• Type "B2Run" is removed — the event will look like a standard Deloitte event.\n• Labels "Durchstarter" / "Funstarter" are persisted as group labels (editable later in the wizard).\n• If performance-proof requirement was active: it is converted into a regular custom field „Leistungsnachweis vorhanden" (checkbox, required, only for group A) — stays as a proper prompt.\n• Hardcoded per-group start-block mapping is removed. If needed, add it as a custom field bound to a group.\n• b2run_* custom fields (age group, t-shirt size, mobile etc.) are KEPT as generic custom fields — you can rename or remove them later in the wizard if no longer needed.\n• Registrations, waitlists and sub-events stay unchanged content-wise.${kidsHint}`;
                          if (!(await confirmDialog(msg, { title: isDe ? 'B2Run migrieren' : 'Migrate B2Run', confirmLabel: isDe ? 'Migrieren' : 'Migrate' }))) return;
                          const errors: string[] = [];
                          const migrateOne = async (ev: DeloitteEvent): Promise<void> => {
                            try {
                              // v11.11: KEINE Custom-Fields mehr löschen.
                              // Die b2run_*-Felder bleiben als generische
                              // Custom-Fields erhalten — der Organizer kann
                              // sie im Wizard danach selbst umbenennen oder
                              // entfernen. Vorher (v11.9) hat die Migration
                              // sie aggressiv aus customFields entfernt, was
                              // zu Datenverlust geführt hat (Altersgruppe,
                              // T-Shirt-Größe etc. waren weg, obwohl nur
                              // die Type-Spalte und Labels umgestellt
                              // werden sollten).
                              // eslint-disable-next-line @typescript-eslint/no-explicit-any
                              const keptFields: any[] = (ev.eventSpecificFields || []).map(f => ({ ...f }));
                              const splitActive = (ev.durchstarterCapacity || 0) > 0 && (ev.funstarterCapacity || 0) > 0;
                              const baseUpdates: Record<string, unknown> = {
                                'SplitLabelA': (ev.splitLabelA || 'Durchstarter'),
                                'SplitLabelB': (ev.splitLabelB || 'Funstarter'),
                              };
                              // v11.13: B2Run-Extras aus
                              // EmailTemplateOverrides._b2run nicht mehr nur
                              // löschen, sondern in echte Custom-Fields mit
                              // onlyForGroup-Bindung übersetzen:
                              // - durchstarterRequiresProof → Custom-Field
                              //   „Leistungsnachweis vorhanden" (Checkbox,
                              //   Pflicht, onlyForGroup='A').
                              // - durchstarterStartblock / funstarterStart-
                              //   block (Auto-Mapping) waren reine UI-
                              //   Convenience und werden ersatzlos entfernt.
                              //   Wenn der Organizer pro Gruppe einen
                              //   Startblock vorgeben will, lege er das
                              //   manuell als Custom-Field mit
                              //   onlyForGroup A bzw. B an.
                              try {
                                const overridesRaw = (ev.emailTemplateOverrides || '').toString();
                                if (overridesRaw.trim()) {
                                  const parsed = JSON.parse(overridesRaw);
                                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                  const b2 = parsed && typeof parsed === 'object' ? (parsed as any)._b2run : null;
                                  if (b2 && typeof b2 === 'object') {
                                    if (b2.durchstarterRequiresProof) {
                                      const PROOF_ID = 'b2run_leistungsnachweis';
                                      const existing = keptFields.find(f => String(f.id || '').toLowerCase() === PROOF_ID);
                                      if (existing) {
                                        existing.onlyForGroup = 'A';
                                        existing.required = true;
                                        if (!existing.label) existing.label = 'Leistungsnachweis vorhanden';
                                        if (!existing.type) existing.type = 'checkbox';
                                        if (!existing.helpText) existing.helpText = 'Ich bestätige, dass ein entsprechender Leistungsnachweis (z.B. Wettkampfergebnis, Trainingsnachweis) vorliegt.';
                                      } else {
                                        keptFields.push({
                                          id: PROOF_ID,
                                          label: 'Leistungsnachweis vorhanden',
                                          type: 'checkbox',
                                          required: true,
                                          options: [],
                                          visible: true,
                                          onlyForGroup: 'A',
                                          helpText: 'Ich bestätige, dass ein entsprechender Leistungsnachweis (z.B. Wettkampfergebnis, Trainingsnachweis) vorliegt.',
                                        });
                                      }
                                      baseUpdates['CustomFields'] = JSON.stringify(keptFields);
                                    }
                                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                    delete (parsed as any)._b2run;
                                    baseUpdates['EmailTemplateOverrides'] = JSON.stringify(parsed);
                                  }
                                }
                              } catch { /* invalid JSON → einfach ignorieren */ }
                              // v11.14: hardcoded B2Run-Field-Specials in
                              // echte Field-Properties migrieren (showIf
                              // für Mobilnummer, externalLinks für
                              // Datenschutz, required für Laufshirt).
                              const fieldExtras = migrateB2RunFieldExtras(keptFields);
                              if (fieldExtras.changed) {
                                baseUpdates['CustomFields'] = JSON.stringify(keptFields);
                              }
                              const ok = await updateEvent(ev.id, baseUpdates);
                              try { await updateEvent(ev.id, { 'EventType': 'Other' }); } catch { /* SP-Spalte evtl. nicht vorhanden — ignoriert */ }
                              if (!ok) { errors.push(`„${ev.title}"`); return; }
                              // v11.11: Subsite-Spalten syncen — fehlende
                              // Spalten werden angelegt. Die b2run_*-Spalten
                              // bleiben drin, weil sie auch in customFields
                              // bleiben.
                              if (ev.subsiteUrl && eventServiceRef) {
                                try {
                                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                  const cfForFix: any[] = keptFields.map(f => ({
                                    id: f.id,
                                    label: f.label,
                                    type: f.type,
                                    required: !!f.required,
                                    visible: true,
                                    options: f.options || [],
                                    /* eslint-disable @typescript-eslint/no-explicit-any */
                                    spInternalName: (f as any).spInternalName || '',
                                    ...((f as any).helpText ? { helpText: (f as any).helpText } : {}),
                                    ...((f as any).multi ? { multi: true } : {}),
                                    ...((f as any).showIf ? { showIf: (f as any).showIf } : {}),
                                    /* eslint-enable @typescript-eslint/no-explicit-any */
                                  }));
                                  await eventServiceRef.fixRegistrationListColumns(ev.subsiteUrl, {
                                    isB2Run: splitActive,
                                    hasQuiz: (ev.quiz || []).length > 0,
                                    customFields: cfForFix,
                                  });
                                } catch (err) { console.warn('[DEX] fixRegistrationListColumns failed for', ev.id, err); }
                              }
                            } catch (err) {
                              console.warn('[DEX] migrate event failed:', ev.id, err);
                              errors.push(`„${ev.title}"`);
                            }
                          };
                          try {
                            await migrateOne(event);
                            for (const k of kidsToMigrate) {
                              await migrateOne(k);
                            }
                            await refreshEvents();
                            const total = 1 + kidsToMigrate.length;
                            if (errors.length === 0) {
                              showAlert(isDe
                                ? `Migration abgeschlossen — ${total} Event(s) auf das Standard-Schema umgestellt.`
                                : `Migration completed — ${total} event(s) migrated to the standard schema.`);
                            } else {
                              showAlert(isDe
                                ? `Migration teilweise fehlgeschlagen bei: ${errors.join(', ')}. Siehe Browser-Console.`
                                : `Migration partially failed for: ${errors.join(', ')}. See browser console.`);
                            }
                          } catch (err) {
                            console.warn('[DEX] migrate B2Run event failed:', err);
                            showAlert(isDe ? 'Migration fehlgeschlagen — siehe Browser-Console.' : 'Migration failed — see browser console.');
                          }
                        }}
                      >
                        {isDe ? 'B2Run migrieren' : 'Migrate B2Run'}
                      </button>
                    )}
                    {/* v18.3: Demo-Event hat keinen Löschen-Button (kein Backend). */}
                    {!event.isDemoShowcase && (
                      <button
                        className="btn btn-secondary"
                        style={{ fontSize: '0.8rem', padding: '6px 12px', color: 'var(--dex-red, #c00)' }}
                        onClick={(e) => {
                          e.stopPropagation();
                          setConfirmDeleteEvent(event);
                          setConfirmDeleteText('');
                        }}
                        disabled={isDeleting}
                      >
                        <Trash2 size={14} /> {isDeleting && deletingId === event.id ? (isDe ? 'Wird gelöscht...' : 'Deleting...') : (isDe ? 'Löschen' : 'Delete')}
                      </button>
                    )}
                  </div>
                </div>
              </div>
            );
            return (
              <>
                {/* v18.2: Sortier- + Entwurf-Filter-Leiste ueber der Event-Liste. */}
                <div style={{
                  display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap',
                  marginBottom: 16, padding: '10px 14px',
                  background: 'var(--dex-gray-50, #fafafa)', borderRadius: 'var(--dex-radius, 12px)',
                  border: '1px solid var(--dex-gray-200)',
                }}>
                  <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                    <span style={{ fontWeight: 600 }}>{isDe ? 'Sortierung:' : 'Sort:'}</span>
                    <select
                      className="form-select"
                      value={eventSortMode}
                      onChange={e => setEventSortMode(e.target.value as 'alpha' | 'date')}
                      style={{ fontSize: '0.85rem', padding: '4px 34px 4px 10px', minWidth: 210, width: 'auto' }}
                    >
                      <option value="alpha">{isDe ? 'Alphabetisch (A–Z)' : 'Alphabetical (A–Z)'}</option>
                      <option value="date">{isDe ? 'Datum aufsteigend' : 'Date ascending'}</option>
                    </select>
                  </label>
                  {draftCount > 0 && (
                    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', color: 'var(--dex-gray-700)', cursor: 'pointer' }}>
                      <input
                        type="checkbox"
                        checked={hideDrafts}
                        onChange={e => setHideDrafts(e.target.checked)}
                        style={{ accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
                      />
                      {isDe ? `Entwürfe ausblenden (${draftCount})` : `Hide drafts (${draftCount})`}
                    </label>
                  )}
                </div>
                <div className="my-events-list">
                  {currentEvents.map(ev => renderEventCard(ev))}
                </div>
                {isAdmin && pastEvents.length > 0 && (
                  <div style={{ marginTop: 24 }}>
                    <button
                      type="button"
                      onClick={() => setShowPastEvents(!showPastEvents)}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        width: '100%', padding: '12px 20px',
                        background: 'var(--dex-gray-50, #f8f9fa)',
                        border: '1px dashed var(--dex-gray-300)',
                        borderRadius: 'var(--dex-radius, 12px)',
                        cursor: 'pointer',
                        fontSize: '0.9rem', fontWeight: 600,
                        color: 'var(--dex-gray-700)',
                        textAlign: 'left',
                      }}
                    >
                      <span style={{ fontSize: '1rem' }}>{showPastEvents ? '▾' : '▸'}</span>
                      <span style={{ flex: 1 }}>
                        Vergangene Events ({pastEvents.length})
                      </span>
                      <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                        Klicken zum {showPastEvents ? 'Einklappen' : 'Ausklappen'}
                      </span>
                    </button>
                    {showPastEvents && (
                      <div className="my-events-list" style={{ marginTop: 12 }}>
                        {pastEvents.map(ev => renderEventCard(ev, { muted: true }))}
                      </div>
                    )}
                  </div>
                )}
              </>
            );
          })()}
          </>
        )}
        {dangerZoneModal}
        {changeLogModal}
      </div>
    );
  }

  // Event ausgewählt - Detail-Ansicht
  const query = searchQuery.toLowerCase().trim();
  const matchesSearch = (reg: SPRegistration): boolean => {
    if (!query) return true;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || '');
    return name.toLowerCase().includes(query)
      || (reg.ParticipantEmail || '').toLowerCase().includes(query)
      || String(reg.TeilnehmerID || '').includes(query);
  };

  const sortRegs = (a: SPRegistration, b: SPRegistration): number => {
    let cmp = 0;
    switch (sortColumn) {
      case 'id': cmp = (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0); break;
      case 'anrede': cmp = (a.Anrede || '').localeCompare(b.Anrede || '', 'de'); break;
      // v11.26: getrennte Vorname / Nachname Sortierung.
      case 'vorname': {
        const na = (a.Vorname || (a.ParticipantName || '').split(' ')[0] || '');
        const nb = (b.Vorname || (b.ParticipantName || '').split(' ')[0] || '');
        cmp = na.localeCompare(nb, 'de');
        break;
      }
      case 'nachname': {
        // Fallback für Alt-Daten ohne separates Vorname/Nachname:
        // Letztes Wort aus ParticipantName als Nachname.
        const lastWord = (s: string): string => {
          const parts = s.trim().split(/\s+/);
          return parts.length > 0 ? parts[parts.length - 1] : '';
        };
        const na = a.Nachname || lastWord(a.ParticipantName || '');
        const nb = b.Nachname || lastWord(b.ParticipantName || '');
        cmp = na.localeCompare(nb, 'de');
        break;
      }
      case 'email': cmp = (a.ParticipantEmail || '').localeCompare(b.ParticipantEmail || ''); break;
      case 'status': cmp = (a.Status || '').localeCompare(b.Status || ''); break;
      case 'date': cmp = new Date(a.RegistrationDate || 0).getTime() - new Date(b.RegistrationDate || 0).getTime(); break;
    }
    return sortAsc ? cmp : -cmp;
  };

  const handleSort = (col: 'id' | 'anrede' | 'vorname' | 'nachname' | 'email' | 'status' | 'date'): void => {
    if (sortColumn === col) { setSortAsc(!sortAsc); }
    else { setSortColumn(col); setSortAsc(true); }
  };

  const sortIcon = (col: string): string => col === sortColumn ? (sortAsc ? ' \u25B2' : ' \u25BC') : '';

  const activeRegs = registrations
    .filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt')
    .filter(matchesSearch)
    .sort(sortRegs);
  const waitlistRegs = registrations.filter(r => r.Status === 'Warteliste').filter(matchesSearch)
    // v12.10: Warteliste nach TeilnehmerID asc sortieren statt
    // RegistrationDate. Damit ist die UI-Reihenfolge konsistent mit
    // der Nachrück-Logik in promoteFirstWaitlistItem (siehe EventService).
    .sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0));
  // v14.11: konsolidierte Matrix-Zeilen für den „Nur Sub-Events"-Modus
  // (subEventsOnlyMode). Aggregation per ParticipantEmail (lowercase).
  // Standard-Felder werden aus der ersten gefundenen Sub-Event-
  // Registrierung kopiert. Pro Sub-Event wird ein Map-Eintrag mit der
  // jeweiligen aktiven Registrierung gehalten (oder undefined).
  const isConsolidatedMode = !!(selectedEvent
    && selectedEvent.subEventsOnlyMode
    && childEventsOf(selectedEvent.id).length > 0);
  const consolidatedChildren: DeloitteEvent[] = isConsolidatedMode
    ? childEventsOf(selectedEvent!.id)
    : [];
  // v22.59: Abmeldungen-Liste im Klammer-Modus über ALLE Sub-Events
  // konsolidieren (vorher nur die Klammer-Subsite → KPI „Abgemeldet" und Liste
  // klafften auseinander). Jede Zeile trägt ihre Subsite + Section-Titel mit,
  // damit das Löschen die richtige Liste trifft.
  const cancelledRegs: Array<SPRegistration & { _subsiteUrl?: string; _sectionTitle?: string; _sectionId?: string }> = isConsolidatedMode
    ? [
        // v22.63: Klammer-eigene Abmeldungen (z.B. „Ich nehme nicht teil"-Absagen
        // — declineEvent schreibt auf die Klammer-Subsite) MIT aufnehmen, sonst
        // verschwinden sie aus der Liste.
        ...registrations
          .filter(r => r.Status === 'Abgemeldet')
          .map(r => ({ ...r, _subsiteUrl: selectedEvent!.subsiteUrl, _sectionTitle: isDe ? 'Gesamt-Event' : 'Overall event', _sectionId: '__parent' })),
        ...consolidatedChildren.reduce<Array<SPRegistration & { _subsiteUrl?: string; _sectionTitle?: string; _sectionId?: string }>>((acc, ch) => {
          for (const r of (subEventRegsByEventId[ch.id] || [])) {
            if (r.Status !== 'Abgemeldet') continue;
            acc.push({ ...r, _subsiteUrl: ch.subsiteUrl, _sectionTitle: shortSubEventTitle(ch.title, selectedEvent!.title), _sectionId: ch.id });
          }
          return acc;
        }, []),
      ].filter(matchesSearch)
    : registrations.filter(r => r.Status === 'Abgemeldet').filter(matchesSearch);
  // v14.11: wenn ein Sub-Event direkt selektiert ist, das Parent-Event
  // ermitteln — der Sub-Event-Detail-View blendet dessen Custom-Fields
  // (Pastel A) zusätzlich neben den eigenen (Pastel B) ein.
  const parentEventForSelected: DeloitteEvent | null = (selectedEvent && selectedEvent.parentEventId)
    ? (allEvents.find(e => e.id === selectedEvent.parentEventId) || null)
    : null;
  // v15.2 HOTFIX: React.useMemo entfernt — der Hook stand NACH dem early
  // return `if (!selectedEvent)` weiter oben (~Zeile 1940) und feuerte
  // damit nur, wenn ein Event selektiert war. Das verletzte die Rules of
  // Hooks (React error #310 „Rendered more hooks than during the previous
  // render") und crashte die App, sobald der User vom Event-Picker auf
  // eine Detail-Ansicht wechselte. Berechnung läuft jetzt pro Render —
  // ist günstig genug, weil consolidatedFiltered weiter unten ohnehin
  // pro Render neu rechnet.
  const consolidatedRows: ConsolidatedRow[] = (() => {
    if (!isConsolidatedMode || !selectedEvent) return [];
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    const byEmail: Record<string, ConsolidatedRow> = {};
    for (const ch of consolidatedChildren) {
      const regs = subEventRegsByEventId[ch.id] || [];
      for (const r of regs) {
        if (ACTIVE.indexOf(r.Status) < 0) continue;
        const emailKey = (r.ParticipantEmail || '').toLowerCase().trim();
        if (!emailKey) continue;
        let row = byEmail[emailKey];
        if (!row) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const anyR = r as any;
          row = {
            emailKey,
            email: r.ParticipantEmail || '',
            vorname: r.Vorname || ((r.ParticipantName || '').split(' ')[0] || ''),
            nachname: r.Nachname || (() => {
              const parts = (r.ParticipantName || '').trim().split(/\s+/);
              return parts.length > 1 ? parts.slice(1).join(' ') : '';
            })(),
            jobTitle: anyR.JobTitle || '',
            location: anyR.Location || '',
            teilnehmerId: r.TeilnehmerID || null,
            earliestRegistrationTs: r.RegistrationDate ? new Date(r.RegistrationDate).getTime() : Number.POSITIVE_INFINITY,
            perChild: {},
            activeCount: 0,
          };
          byEmail[emailKey] = row;
        } else {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const anyR = r as any;
          if (!row.jobTitle && anyR.JobTitle) row.jobTitle = anyR.JobTitle;
          if (!row.location && anyR.Location) row.location = anyR.Location;
          if (!row.vorname && r.Vorname) row.vorname = r.Vorname;
          if (!row.nachname && r.Nachname) row.nachname = r.Nachname;
          // Früheste RegistrationDate uebernehmen (min).
          const ts = r.RegistrationDate ? new Date(r.RegistrationDate).getTime() : Number.POSITIVE_INFINITY;
          if (ts < row.earliestRegistrationTs) row.earliestRegistrationTs = ts;
        }
        row.perChild[ch.id] = r;
        row.activeCount += 1;
      }
    }
    return Object.values(byEmail);
  })();
  // v14.11: Such-Filter + Sort für die konsolidierten Zeilen.
  const consolidatedFiltered: ConsolidatedRow[] = (() => {
    const q = (searchQuery || '').toLowerCase().trim();
    const matches = (row: ConsolidatedRow): boolean => {
      if (!q) return true;
      return row.vorname.toLowerCase().indexOf(q) >= 0
        || row.nachname.toLowerCase().indexOf(q) >= 0
        || row.email.toLowerCase().indexOf(q) >= 0
        || String(row.teilnehmerId || '').indexOf(q) >= 0;
    };
    const filtered = consolidatedRows.filter(matches);
    const cs = consolidatedSort;
    const dir = consolidatedSortAsc ? 1 : -1;
    const cmp = (a: ConsolidatedRow, b: ConsolidatedRow): number => {
      // v15.23: #-Spalte sortiert nach Reihenfolge der ersten Anmeldung
      // (RegistrationDate min), nicht mehr nach TeilnehmerID — die TID ist
      // pro Sub-Event und bei konsolidierten Personen mehrdeutig.
      if (cs === 'id') return (a.earliestRegistrationTs - b.earliestRegistrationTs) * dir;
      if (cs === 'vorname') return a.vorname.localeCompare(b.vorname, 'de') * dir;
      if (cs === 'nachname') return a.nachname.localeCompare(b.nachname, 'de') * dir;
      if (cs === 'email') return a.email.localeCompare(b.email) * dir;
      if (cs === 'jobTitle') return a.jobTitle.localeCompare(b.jobTitle, 'de') * dir;
      if (cs === 'location') return a.location.localeCompare(b.location, 'de') * dir;
      if (cs && cs.indexOf('child:') === 0) {
        const cid = cs.substring(6);
        const ra = a.perChild[cid] ? 1 : 0;
        const rb = b.perChild[cid] ? 1 : 0;
        return (rb - ra) * dir;
      }
      return 0;
    };
    return filtered.sort(cmp);
  })();
  // v14.11: konsolidierter Matrix-View. Wird nur gerendert, wenn das
  // selektierte Hauptevent `subEventsOnlyMode === true` ist und mind.
  // ein Sub-Event hat. Standard-Spalten neutral, parent-event-level
  // Custom-Fields in Pastel A (hellblau), pro Sub-Event eine X-Spalte,
  // sub-event-spezifische Custom-Fields in Pastel B (hellgelb)
  // gruppiert pro Sub-Event-Header.
  const renderConsolidatedView = (): React.ReactNode => {
    if (!selectedEvent) return null;
    if (isLoadingSubEventRegs) {
      return <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{isDe ? 'Lade Sub-Event-Teilnehmer...' : 'Loading sub-event participants...'}</p>;
    }
    if (consolidatedRows.length === 0) {
      return <p style={{ color: 'var(--dex-gray-400)' }}>{isDe ? 'Noch keine Anmeldungen in den Sub-Events.' : 'No registrations in the sub-events yet.'}</p>;
    }
    // v14.11: pastel A = event-level (parent) fields, pastel B = sub-event-specific fields
    const PASTEL_A_HEADER: React.CSSProperties = { background: 'rgba(0, 118, 168, 0.15)' };
    const PASTEL_A_CELL: React.CSSProperties = { background: 'rgba(0, 118, 168, 0.08)' };
    const PASTEL_B_HEADER: React.CSSProperties = { background: 'rgba(255, 191, 0, 0.18)' };
    const PASTEL_B_CELL: React.CSSProperties = { background: 'rgba(255, 191, 0, 0.10)' };
    const parentCustomFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.label && f.label.trim());
    const parentIds = new Set(parentCustomFields.map(f => f.id));
    const childCustomFieldsByChild: Array<{ child: DeloitteEvent; fields: typeof parentCustomFields }> = consolidatedChildren.map(c => {
      const own = (c.eventSpecificFields || []).filter(f => f.type !== 'user' && f.label && f.label.trim() && !parentIds.has(f.id));
      return { child: c, fields: own };
    });
    const handleSortConsolidated = (key: string): void => {
      if (consolidatedSort === key) setConsolidatedSortAsc(!consolidatedSortAsc);
      else { setConsolidatedSort(key); setConsolidatedSortAsc(true); }
    };
    const sortArrow = (key: string): string => key === consolidatedSort ? (consolidatedSortAsc ? ' ▲' : ' ▼') : '';
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    const abbreviate = (s: string, max: number): string => s.length > max ? s.substring(0, max - 1) + '…' : s;
    const totalColSpan = 6 + parentCustomFields.length + childCustomFieldsByChild.reduce((sum, x) => sum + 1 + x.fields.length, 0) + 1;
    // v19.30: Aktionen (Hauptevent-Felder bearbeiten / abmelden) nur für
    // berechtigte Rollen (Admin oder Organizer dieses Events).
    const canManage = isAdmin || isOrganizerFor(selectedEvent);
    // v19.30 (Feature A): Anzahl der bearbeitbaren Hauptevent-Felder (ohne
    // People-Picker und Dokument-Uploads, die keinen editierbaren Textwert
    // haben). Nur wenn > 0 erscheint der „Felder"-Button.
    const editableParentFieldCount = parentCustomFields.filter(f => f.type !== 'document').length;
    // v19.30 (Feature A): Hat die Person eine Registrierung auf der
    // Hauptevent-Teilnehmerliste? Nur dann gibt es Hauptevent-Antworten zum
    // Bearbeiten. (Im subEventsOnlyMode kann jemand nur in Sub-Events
    // angemeldet sein.)
    const hasParentReg = (emailKey: string): boolean =>
      registrations.some(r => (r.ParticipantEmail || '').toLowerCase().trim() === emailKey);
    return (
      <div style={{ overflowX: 'auto' }}>
        {/* v15.3.1: Legende für die Pastell-Spalten — sonst rät der Organizer,
            was die zwei Hintergrundfarben bedeuten. */}
        {(parentCustomFields.length > 0 || childCustomFieldsByChild.some(x => x.fields.length > 0)) && (
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, marginBottom: 10, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
            {parentCustomFields.length > 0 && (
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                <span style={{ width: 14, height: 14, borderRadius: 3, ...PASTEL_A_HEADER, border: '1px solid rgba(0, 118, 168, 0.3)' }} />
                {isDe ? 'Felder des Hauptevents' : 'Main-event fields'}
              </span>
            )}
            {childCustomFieldsByChild.some(x => x.fields.length > 0) && (
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                <span style={{ width: 14, height: 14, borderRadius: 3, ...PASTEL_B_HEADER, border: '1px solid rgba(255, 191, 0, 0.4)' }} />
                {isDe ? 'Felder eines Sub-Events' : 'Sub-event fields'}
              </span>
            )}
          </div>
        )}
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
          <thead>
            <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
              <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSortConsolidated('id')}>#{sortArrow('id')}</th>
              <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSortConsolidated('vorname')}>{isDe ? 'Vorname' : 'First name'}{sortArrow('vorname')}</th>
              <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSortConsolidated('nachname')}>{isDe ? 'Nachname' : 'Last name'}{sortArrow('nachname')}</th>
              <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSortConsolidated('email')}>Email{sortArrow('email')}</th>
              <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSortConsolidated('jobTitle')}>Job Title{sortArrow('jobTitle')}</th>
              <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSortConsolidated('location')}>{isDe ? 'Standort' : 'Location'}{sortArrow('location')}</th>
              {parentCustomFields.map(f => (
                <th key={`pf-${f.id}`} style={{ textAlign: 'left', padding: 8, fontSize: '0.78rem', whiteSpace: 'normal', overflowWrap: 'break-word', maxWidth: 150, verticalAlign: 'top', lineHeight: 1.25, ...PASTEL_A_HEADER }} title={`${f.label} — ${isDe ? 'Hauptevent-Feld' : 'main-event field'}`}>
                  {f.label}
                </th>
              ))}
              {childCustomFieldsByChild.map(({ child, fields }) => (
                <React.Fragment key={`sub-${child.id}`}>
                  <th
                    style={{ textAlign: 'center', padding: 8, cursor: 'pointer', userSelect: 'none', borderLeft: '1px solid var(--dex-gray-200)' }}
                    onClick={() => handleSortConsolidated(`child:${child.id}`)}
                    title={child.title}
                  >
                    <div style={{ fontSize: '0.78rem', fontWeight: 600 }}>{abbreviate(shortSubEventTitle(child.title, selectedEvent?.title) || '?', 16)}</div>
                    <div style={{ fontSize: '0.68rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>{isDe ? 'angemeldet?' : 'registered?'}{sortArrow(`child:${child.id}`)}</div>
                  </th>
                  {fields.map(f => (
                    <th key={`scf-${child.id}-${f.id}`} style={{ textAlign: 'left', padding: 8, fontSize: '0.78rem', whiteSpace: 'normal', overflowWrap: 'break-word', maxWidth: 150, verticalAlign: 'top', lineHeight: 1.25, ...PASTEL_B_HEADER }} title={`${f.label} — ${child.title}`}>
                      <div style={{ color: 'var(--dex-gray-500)', fontWeight: 400, fontSize: '0.68rem' }}>{abbreviate(shortSubEventTitle(child.title, selectedEvent?.title) || '?', 18)}</div>
                      <div style={{ fontWeight: 600 }}>{f.label}</div>
                    </th>
                  ))}
                </React.Fragment>
              ))}
              <th style={{ textAlign: 'left', padding: 8 }}>{isDe ? 'Details' : 'Details'}</th>
            </tr>
          </thead>
          <tbody>
            {consolidatedFiltered.map((row, idx) => {
              const isExpanded = expandedConsolidatedEmail === row.emailKey;
              return (
                <React.Fragment key={row.emailKey}>
                  <tr style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                    {/* v15.20: Im konsolidierten View einfach fortlaufend
                        durchnummerieren (idx+1). Die Sub-Event-TeilnehmerID
                        macht hier keinen Sinn, weil jede Person eine eigene
                        TID pro Sub-Event hat — sortbar bleibt es ueber
                        Vorname/Nachname/Email-Spalten. */}
                    <td style={{ padding: 8, color: 'var(--dex-gray-400)' }}>{idx + 1}</td>
                    <td style={{ padding: 8, fontWeight: 500 }}>{row.vorname || '-'}</td>
                    <td style={{ padding: 8, fontWeight: 500 }}>{row.nachname || '-'}</td>
                    <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{row.email}</td>
                    <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{row.jobTitle || '-'}</td>
                    <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{row.location || '-'}</td>
                    {parentCustomFields.map(f => {
                      let val = '';
                      // v15.3.1: Parent-Level-Custom-Fields zuerst aus der
                      // PARENT-Teilnehmerliste auflösen (registrations =
                      // selectedEvent.id-Regs), erst danach Fallback auf
                      // Sub-Event-CustomData. Vorher wurde nur Sub-Event-
                      // CustomData gelesen — Parent-Felder waren immer leer.
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const parentReg = registrations.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === row.emailKey) as any;
                      if (parentReg) {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const spName = (f as any).spInternalName || '';
                        let v: unknown = spName ? parentReg[spName] : undefined;
                        if ((v === undefined || v === null || v === '') && parentReg.CustomData) {
                          try { v = JSON.parse(parentReg.CustomData)[f.id]; } catch { /* */ }
                        }
                        if (v !== undefined && v !== null && v !== '') val = String(v);
                      }
                      // Fallback: Sub-Event-CustomData durchsuchen (Legacy-Events,
                      // bei denen Parent-Felder in Sub-Event-CustomData kopiert
                      // wurden — z.B. bei Wizard-„Vom Hauptevent kopieren").
                      if (!val) {
                        for (const ch of consolidatedChildren) {
                          const r = row.perChild[ch.id];
                          if (!r) continue;
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const spName = (f as any).spInternalName || '';
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          let v: any = spName ? (r as any)[spName] : undefined;
                          if ((v === undefined || v === null || v === '') && r.CustomData) {
                            try { v = JSON.parse(r.CustomData)[f.id]; } catch { /* */ }
                          }
                          if (v !== undefined && v !== null && v !== '') { val = String(v); break; }
                        }
                      }
                      return (
                        <td key={`pcv-${f.id}`} style={{ padding: 8, fontSize: '0.8rem', color: 'var(--dex-gray-700)', whiteSpace: 'nowrap', maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', ...PASTEL_A_CELL }} title={val}>
                          {val || '-'}
                        </td>
                      );
                    })}
                    {childCustomFieldsByChild.map(({ child, fields }) => {
                      const r = row.perChild[child.id];
                      const isReg = !!r && ACTIVE.indexOf(r.Status) >= 0;
                      return (
                        <React.Fragment key={`scv-${child.id}`}>
                          <td style={{ padding: 8, textAlign: 'center', borderLeft: '1px solid var(--dex-gray-200)' }}
                              title={r ? `${translateStatus(r.Status, isDe)} — TID ${r.TeilnehmerID || '?'}` : (isDe ? 'Nicht angemeldet' : 'Not registered')}>
                            {isReg ? (
                              r.Status === 'Warteliste'
                                ? <span style={{ color: 'var(--dex-orange, #ed8b00)', fontSize: '0.78rem' }} title={translateStatus(r.Status, isDe)}>W</span>
                                : <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', display: 'inline-flex' }}><Check size={16} /></span>
                            ) : (
                              <span style={{ color: 'var(--dex-gray-300)' }}>—</span>
                            )}
                          </td>
                          {fields.map(f => {
                            let val = '';
                            if (r) {
                              // eslint-disable-next-line @typescript-eslint/no-explicit-any
                              const spName = (f as any).spInternalName || '';
                              // eslint-disable-next-line @typescript-eslint/no-explicit-any
                              let v: any = spName ? (r as any)[spName] : undefined;
                              if ((v === undefined || v === null || v === '') && r.CustomData) {
                                try { v = JSON.parse(r.CustomData)[f.id]; } catch { /* */ }
                              }
                              if (v !== undefined && v !== null && v !== '') val = String(v);
                            }
                            return (
                              <td key={`scv-${child.id}-${f.id}`} style={{ padding: 8, fontSize: '0.8rem', color: 'var(--dex-gray-700)', whiteSpace: 'nowrap', maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', ...PASTEL_B_CELL }} title={val}>
                                {val || (r ? '-' : '')}
                              </td>
                            );
                          })}
                        </React.Fragment>
                      );
                    })}
                    <td style={{ padding: 8 }}>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                        <button
                          type="button"
                          className="btn btn-secondary"
                          style={{ fontSize: '0.75rem', padding: '4px 10px', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                          onClick={() => setExpandedConsolidatedEmail(isExpanded ? null : row.emailKey)}
                        >
                          {isExpanded ? (isDe ? 'Schließen' : 'Close') : (isDe ? 'Details' : 'Details')}
                          {isExpanded ? <ChevronUp size={12} /> : <ChevronDown size={12} />}
                        </button>
                        {/* v19.30 (Feature A): Hauptevent-Felder bearbeiten —
                            nur wenn es Hauptevent-Custom-Felder gibt UND die
                            Person eine Hauptevent-Registrierung hat. */}
                        {canManage && editableParentFieldCount > 0 && hasParentReg(row.emailKey) && (
                          <button
                            type="button"
                            className="btn btn-secondary"
                            style={{ fontSize: '0.75rem', padding: '4px 10px', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                            title={isDe ? 'Felder des Hauptevents dieser Person bearbeiten' : 'Edit this person’s main-event fields'}
                            onClick={() => openMainFieldsEdit(row.emailKey, `${row.vorname} ${row.nachname}`.trim() || row.email)}
                          >
                            <Pencil size={12} /> {isDe ? 'Felder' : 'Fields'}
                          </button>
                        )}
                        {/* v19.30 (Feature B): Abmelden mit Sub-Event-Auswahl. */}
                        {canManage && (
                          <button
                            type="button"
                            className="btn btn-secondary"
                            style={{ fontSize: '0.75rem', padding: '4px 10px', display: 'inline-flex', alignItems: 'center', gap: 4, color: 'var(--dex-red, #c00)' }}
                            title={isDe ? 'Aus einzelnen oder allen Sub-Events abmelden' : 'Deregister from selected or all sub-events'}
                            onClick={() => openDeregModal(row)}
                          >
                            <Trash2 size={12} /> {isDe ? 'Abmelden' : 'Cancel'}
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                  {isExpanded && (
                    <tr style={{ background: 'var(--dex-gray-50, #f7f7f7)' }}>
                      <td colSpan={totalColSpan} style={{ padding: '12px 16px' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                          <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
                            {isDe ? 'Anmeldungen von' : 'Registrations of'} {row.vorname} {row.nachname}
                          </div>
                          {consolidatedChildren.map(ch => {
                            const r = row.perChild[ch.id];
                            if (!r) {
                              return (
                                <div key={`exp-${ch.id}`} style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)' }}>
                                  {shortSubEventTitle(ch.title, selectedEvent?.title)} — {isDe ? 'nicht angemeldet' : 'not registered'}
                                </div>
                              );
                            }
                            return (
                              <div key={`exp-${ch.id}`} style={{ display: 'flex', alignItems: 'center', gap: 12, fontSize: '0.82rem' }}>
                                <span style={{ fontWeight: 500, minWidth: 200 }}>{shortSubEventTitle(ch.title, selectedEvent?.title)}</span>
                                <span className={`badge ${r.Status === 'Eingecheckt' ? 'badge-green' : 'badge-gray'}`}>{translateStatus(r.Status, isDe)}</span>
                                <span style={{ color: 'var(--dex-gray-500)' }}>TID {r.TeilnehmerID || '?'}</span>
                                <span style={{ color: 'var(--dex-gray-400)', fontSize: '0.75rem' }}>{formatDate(r.RegistrationDate)}</span>
                                <button
                                  className="btn btn-secondary"
                                  style={{ marginLeft: 'auto', fontSize: '0.72rem', padding: '3px 8px', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                                  onClick={() => setSelectedEvent(ch)}
                                >
                                  {isDe ? 'In Sub-Event öffnen' : 'Open in sub-event'} <ExternalLink size={12} />
                                </button>
                              </div>
                            );
                          })}
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };
  // Seit v6.5: getrennte Wartelisten bei B2Run-Split-Kapazitäten (Durchstarter/Funstarter).
  // Die Split-Aktivierung erkennen wir daran, dass beide Kapazitäts-Felder gesetzt und > 0 sind.
  const isSplitCapacity = !!selectedEvent
    && typeof selectedEvent.durchstarterCapacity === 'number'
    && typeof selectedEvent.funstarterCapacity === 'number'
    && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
  const waitlistDurch = isSplitCapacity
    ? waitlistRegs.filter(r => r.PreferredStarterType === 'Durchstarter')
    : [];
  const waitlistFun = isSplitCapacity
    ? waitlistRegs.filter(r => r.PreferredStarterType === 'Funstarter')
    : [];
  const waitlistUnassigned = isSplitCapacity
    ? waitlistRegs.filter(r => !r.PreferredStarterType || (r.PreferredStarterType !== 'Durchstarter' && r.PreferredStarterType !== 'Funstarter'))
    : [];

  // Roommate-Matching: durchsucht CustomData nach roommate-Type Feldern, extrahiert
  // Email aus "Name <email>"-Format, baut Map Email -> Partner-Email. Match-Badge,
  // wenn beide sich gegenseitig ausgewählt haben.
  // v11.65: ausschliesslich `roommate`-Felder, nicht mehr `user`. Bei Assistant-
  // /generischen User-Pickern macht ein „Match"-Badge semantisch keinen Sinn —
  // der wurde fälschlich auch dort gezeigt, wenn Person A und B sich
  // gegenseitig als Assistant eingetragen haben.
  const userFieldIds = (selectedEvent?.eventSpecificFields || [])
    .filter(f => f.type === 'roommate')
    .map(f => f.id);

  // Render-Funktionen pro Spalte — als eine Map, damit der Header + die Body-Zeilen
  // die gleiche stabile ID benutzen. Wird bei jedem Registration-Render neu aufgebaut,
  // weil die renderCell-Lambdas auf aktuelle Closures (handleSort, sortIcon, …) angewiesen sind.
  // Das ist günstig, weil die Funktion nur Pointer speichert.
  const roommateChoice: Record<string, { partnerEmail: string; partnerName: string }> = {};
  if (userFieldIds.length > 0) {
    const allActiveAndWaitlist = registrations.filter(r =>
      r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt' || r.Status === 'Warteliste'
    );
    for (const r of allActiveAndWaitlist) {
      const email = (r.ParticipantEmail || '').toLowerCase();
      if (!email) continue;
      let cd: Record<string, string> = {};
      try { cd = JSON.parse(r.CustomData || '{}'); } catch { /* */ }
      for (const fid of userFieldIds) {
        const v = cd[fid];
        if (!v) continue;
        const m = v.match(/<([^>]+@[^>]+)>/);
        if (!m) continue;
        const pEmail = m[1].trim().toLowerCase();
        const pName = v.replace(/<[^>]*>/, '').trim();
        if (pEmail && pEmail !== email) {
          roommateChoice[email] = { partnerEmail: pEmail, partnerName: pName };
          break; // nur erstes user-Feld
        }
      }
    }
  }
  const getRoommateInfo = (reg: { ParticipantEmail?: string }): { partnerName: string; partnerEmail: string; mutual: boolean } | null => {
    const email = (reg.ParticipantEmail || '').toLowerCase();
    if (!email) return null;
    const choice = roommateChoice[email];
    if (!choice) return null;
    const reverse = roommateChoice[choice.partnerEmail];
    const mutual = !!reverse && reverse.partnerEmail === email;
    return { partnerName: choice.partnerName || choice.partnerEmail, partnerEmail: choice.partnerEmail, mutual };
  };

  return (
    <div className="page-container" role="main">
      {/* Admin-Toast: drei Phasen beim Abmelden (seit v6.8).
          1. cancelling — orange, Spinner, läuft während der Abmeldung+Promote-Suche
          2. promoted   — grün, zeigt den Nachrücker
          3. no-promote — grau, Abmeldung ok, keiner auf der Warteliste */}
      {adminToast && (() => {
        const accent = adminToast.kind === 'cancelling'
          ? 'var(--dex-orange, #ed8b00)'
          : adminToast.kind === 'promoted'
            ? 'var(--dex-green, #86bc25)'
            : 'var(--dex-gray-400)';
        const accentDark = adminToast.kind === 'cancelling'
          ? 'var(--dex-orange, #ed8b00)'
          : adminToast.kind === 'promoted'
            ? 'var(--dex-green-dark, #6b9a1e)'
            : 'var(--dex-gray-600)';
        const closable = adminToast.kind !== 'cancelling';
        return (
          <div style={{
            position: 'fixed', top: 80, right: 20, zIndex: 1000, maxWidth: 460,
            padding: '14px 18px', borderRadius: 'var(--dex-radius, 12px)',
            background: '#fff',
            border: `1px solid ${accent}`,
            borderLeft: `4px solid ${accent}`,
            boxShadow: '0 6px 20px rgba(0,0,0,0.15)',
            display: 'flex', alignItems: 'flex-start', gap: 12,
          }}>
            {adminToast.kind === 'cancelling' && (
              <div style={{
                width: 20, height: 20, marginTop: 2, flexShrink: 0,
                border: `3px solid var(--dex-gray-200)`,
                borderTopColor: accent,
                borderRadius: '50%',
                animation: 'dex-spin 0.8s linear infinite',
              }} />
            )}
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: 700, color: accentDark, marginBottom: 4 }}>
                {adminToast.kind === 'cancelling' && `Abmeldung von ${adminToast.name} wird verarbeitet…`}
                {adminToast.kind === 'promoted' && `Nachgerückt: ${adminToast.name}`}
                {adminToast.kind === 'no-promote' && `Abmeldung von ${adminToast.name} verarbeitet`}
              </div>
              <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                {adminToast.kind === 'cancelling' && 'Teilnehmer wird abgemeldet, Warteliste wird geprüft und ggf. ein Nachrücker informiert.'}
                {adminToast.kind === 'promoted' && (
                  <>
                    <strong>{adminToast.email}</strong> wurde automatisch aus der Warteliste{adminToast.type ? ` (${adminToast.type})` : ''} nachgerückt. Nachrück-Mail + Outlook-Einladung wurden versendet.
                  </>
                )}
                {adminToast.kind === 'no-promote' && 'Aktuell ist niemand auf der Warteliste (bzw. kein passender Starter-Typ). Der Platz bleibt frei.'}
              </div>
            </div>
            {closable && (
              <button
                onClick={() => setAdminToast(null)}
                aria-label="Schließen"
                style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.2rem', color: 'var(--dex-gray-500)', lineHeight: 1, padding: 0 }}
              >×</button>
            )}
          </div>
        );
      })()}
      {/* Keyframes für Spinner */}
      <style>{`@keyframes dex-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
      {/* v9.29: Inline Zurück + Aktualisieren entfernt — beides liegt jetzt im Header.
          Eventauswahl-Reset („zurück zur Event-Liste") triggern wir über den Header-Back —
          siehe Listener weiter oben, der bei navigate-Wechsel selectedEvent zurücksetzt. */}

      {/* v22.7: Hinweisbox, wenn Teilnehmer-Konten nicht mehr aktiv sind
          (Person hat womöglich Deloitte verlassen). Hintergrund-Check beim
          Öffnen, max. 1×/Tag pro Event. */}
      {inactiveAccounts.length > 0 && (() => {
        const items = inactiveAccounts.map(email => {
          const reg = registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === email);
          const name = reg ? ((reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || email)) : email;
          return { email, name };
        });
        return (
          <div style={{
            display: 'flex', alignItems: 'flex-start', gap: 12,
            padding: '14px 16px', marginBottom: 20,
            background: '#fff3e0', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 12,
            color: 'var(--dex-gray-800)',
          }}>
            <AlertCircle size={20} />
            <div style={{ minWidth: 0 }}>
              <div style={{ fontWeight: 700, marginBottom: 4, color: 'var(--dex-orange-dark, #b35a00)' }}>
                {isDe
                  ? `${items.length === 1 ? 'Eine Person' : `${items.length} Personen`} hat womöglich Deloitte verlassen`
                  : `${items.length === 1 ? 'One person' : `${items.length} people`} may have left Deloitte`}
              </div>
              <div style={{ fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, marginBottom: 8 }}>
                {isDe
                  ? 'Zu folgenden Teilnehmern wurde kein aktives Deloitte-Konto mehr gefunden — das Konto ist deaktiviert oder existiert nicht mehr. Mails/Outlook-Termine an diese Adressen kommen ggf. nicht an. Bitte prüfen und ggf. abmelden.'
                  : 'No active Deloitte account was found for the following participants — the account is disabled or no longer exists. Emails/Outlook invites to these addresses may not arrive. Please review and deregister if needed.'}
              </div>
              <ul style={{ margin: 0, paddingLeft: 18, fontSize: '0.85rem', color: 'var(--dex-gray-800)' }}>
                {items.map(it => (
                  <li key={it.email}><strong>{it.name}</strong> <span style={{ color: 'var(--dex-gray-500)' }}>({it.email})</span></li>
                ))}
              </ul>
            </div>
          </div>
        );
      })()}

      {/* v12.7: Aktionen-Card aufgelöst — alle ActionTiles registrieren
          sich jetzt im ActionsRegistryProvider. Die Dropdown-Liste sitzt
          unten in der linken Event-Detail-Card. Daher 1-Spalten-Grid
          statt vorher 2-Spalten. */}
      <ActionsRegistryProvider>
      <div className="admin-event-info-grid" style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 24, marginBottom: 24 }}>
        {/* v22.5: Detail-Card + „Nächste Schritte"-Box rechts daneben (Desktop;
            stapelt auf Mobile via flex-wrap). Die Box erscheint nur für Entwürfe
            und nur für Admin/Organizer. */}
        <div style={{ display: 'flex', gap: 24, alignItems: 'flex-start', flexWrap: 'wrap' }}>
        <div ref={detailCardRef} className="card" style={{ padding: 24, minHeight: reservedDetailHeight, flex: '1 1 420px', minWidth: 0 }}>
          {/* Header: Event-Titel + Status-Badge + Schnellaktionen (v13.11) */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap', marginBottom: 16 }}>
            <h2 style={{ margin: 0, fontSize: '1.2rem', lineHeight: 1.2 }}>{selectedEvent.title}</h2>
            {/* v20.3: Status-Badge ist klickbar — Klick auf „Aktiv" setzt das
                Event auf Entwurf, Klick auf „Entwurf" schaltet es live
                (jeweils mit Sicherheitsabfrage). v22.15: auch Abgeschlossen/
                Abgesagt sind für Admin/Organizer klickbar und lassen sich
                wieder auf Aktiv setzen (vorher Sackgasse — z.B. wenn der
                Auto-Cleanup ein Event mit altem Testdatum auf „Abgeschlossen"
                gesetzt hatte und das Datum später korrigiert wurde). */}
            {(() => {
              const isDraft = !!selectedEvent.isFictive;
              const badgeBg = isDraft ? 'rgba(237,139,0,0.15)' : getStatusColor(selectedEvent.status) + '22';
              const badgeFg = isDraft ? 'var(--dex-orange-dark, #b35a00)' : getStatusColor(selectedEvent.status);
              const label = isDraft ? 'ENTWURF' : (isDe ? localizeStatus(selectedEvent.status) : selectedEvent.status);
              const isFinalState = !isDraft && (selectedEvent.status === 'Completed' || selectedEvent.status === 'Cancelled');
              const canToggleStatus = (isAdmin || isOrganizerFor(selectedEvent))
                && !(isImpersonating && selectedEvent.isDemoShowcase)
                && (isDraft || selectedEvent.status === 'Active' || isFinalState);
              if (!canToggleStatus) {
                return (
                  <span className="badge" style={{ background: badgeBg, color: badgeFg }}>{label}</span>
                );
              }
              return (
                <button
                  type="button"
                  className="badge"
                  onClick={() => { toggleDraftStatus().catch(() => { /* */ }); }}
                  title={isDraft
                    ? (isDe ? 'Klicken: Event live schalten (Aktiv). Alle Berechtigten sehen das Event danach und können sich anmelden.' : 'Click: publish event (Active). All eligible users will see the event and can register.')
                    : isFinalState
                      ? (isDe ? 'Klicken: Event wieder auf Aktiv setzen. Danach ist es für die Berechtigten wieder sichtbar und buchbar.' : 'Click: set event back to Active. It will be visible and bookable for eligible users again.')
                      : (isDe ? 'Klicken: Event auf Entwurf setzen. Reguläre User sehen das Event danach nicht mehr; Anmeldungen bleiben erhalten.' : 'Click: set event to draft. Regular users will no longer see the event; registrations are kept.')}
                  style={{
                    background: badgeBg, color: badgeFg,
                    border: `1px solid ${badgeFg}`,
                    cursor: 'pointer',
                    display: 'inline-flex', alignItems: 'center', gap: 5,
                  }}
                >
                  {label}
                  <span style={{ fontSize: '0.75em', opacity: 0.85 }}>⇄</span>
                </button>
              );
            })()}
            {/* v13.11: Event bearbeiten + Check-In starten als Schnell-
                Buttons direkt neben dem Status-Badge — die häufigsten
                Aktionen aus dem Aktionen-Dropdown nach oben gezogen,
                damit Organizer am Eventtag nicht erst scrollen müssen. */}
            <div style={{ display: 'flex', gap: 8, marginLeft: 'auto', flexWrap: 'wrap' }}>
              {/* v18.3: Im Demo-Modus ist das Demo-Event read-only — Edit /
                  Check-In / Aktionen sind ausgeblendet (kein SharePoint-
                  Backend), stattdessen ein Demo-Hinweis. */}
              {selectedEvent.isDemoShowcase ? (
                <span style={{
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                  fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-blue, #0076a8)',
                  background: 'rgba(0,118,168,0.08)', border: '1px solid var(--dex-blue, #0076a8)',
                  borderRadius: 999, padding: '4px 12px',
                }}>
                  {isDe ? 'Demo — nur Ansicht (keine Aktionen)' : 'Demo — view only (no actions)'}
                </span>
              ) : (
                <>
                  {(isAdmin || isOrganizerFor(selectedEvent)) && (
                    <button
                      type="button"
                      className="btn btn-secondary"
                      onClick={() => navigate('edit-event', selectedEvent.id)}
                      style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', padding: '6px 12px' }}
                      title={t('admin.editbutton') || (isDe ? 'Event bearbeiten' : 'Edit event')}
                    >
                      <Pencil size={14} />
                      {isDe ? 'Event bearbeiten' : 'Edit event'}
                    </button>
                  )}
                  <button
                    type="button"
                    className="btn btn-primary"
                    onClick={() => navigate('check-in', selectedEvent.id)}
                    style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', padding: '6px 12px' }}
                    title={t('admin.checkin') || (isDe ? 'Check-In starten' : 'Start check-in')}
                  >
                    <Hash size={14} />
                    {isDe ? 'Check-In starten' : 'Start check-in'}
                  </button>
                </>
              )}
            </div>
          </div>
          {/* Foto immer als Kreis links, Detail-Rows rechts. Layout
              unabhängig vom Bildformat (cover-Crop sorgt für den Kreis). */}
          <div style={{ display: 'flex', gap: 20, alignItems: 'flex-start' }}>
            {/* v12.6: Event-Bild jetzt prominent als großes Rechteck-
                Format (wie auf der Registrierungs-Seite) statt kleinem
                Avatar-Kreis. Hintergrund weiß für saubere Darstellung
                transparenter PNG-Logos.
                v20.2: darunter die Self-Check-in-QR-Kachel — sichtbar ab
                5 Tagen vor Event-Start ODER sobald QR-Codes versendet
                wurden (und solange das Event nicht länger als 1 Tag vorbei
                ist). Klick öffnet das Erklär-/Einstell-Modal. */}
            {(() => {
              const canManageSci = isAdmin || isOrganizerFor(selectedEvent);
              const dayMs = 24 * 60 * 60 * 1000;
              const startTs = selectedEvent.startDate ? new Date(selectedEvent.startDate).getTime() : 0;
              const endTs = selectedEvent.endDate ? new Date(selectedEvent.endDate).getTime() : startTs;
              const nowTs = Date.now();
              const qrPhase = registrations.some(r => r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
              const within5Days = startTs > 0 && nowTs >= startTs - 5 * dayMs;
              const notLongPast = (endTs || startTs) === 0 || nowTs <= (endTs || startTs) + dayMs;
              const showSciTile = canManageSci && notLongPast && (qrPhase || within5Days);
              if (!selectedEvent.imageUrl && !showSciTile) return null;
              return (
                <div style={{ flex: '0 0 auto', width: 260, maxWidth: '38%', display: 'flex', flexDirection: 'column', gap: 12 }}>
                  {selectedEvent.imageUrl && (
                    <div
                      style={{
                        background: '#fff',
                        borderRadius: 'var(--dex-radius, 12px)',
                        overflow: 'hidden',
                        border: '1px solid var(--dex-gray-200, #e5e7eb)',
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                      }}
                    >
                      <img
                        src={selectedEvent.imageUrl}
                        alt={selectedEvent.title}
                        style={{
                          display: 'block',
                          width: '100%',
                          height: 'auto',
                          maxHeight: 240,
                          objectFit: 'contain',
                        }}
                      />
                    </div>
                  )}
                  {showSciTile && (
                    <button
                      type="button"
                      onClick={openSelfCheckInModal}
                      disabled={sciBusy}
                      title={isDe ? 'Self-Check-in-QR anzeigen, drucken und Zeitfenster einstellen' : 'Show/print the self check-in QR and set the time window'}
                      style={{
                        background: '#fff',
                        border: '1px solid var(--dex-green, #86bc25)',
                        borderRadius: 'var(--dex-radius, 12px)',
                        padding: 12, cursor: sciBusy ? 'wait' : 'pointer',
                        display: 'flex', alignItems: 'center', gap: 12, textAlign: 'left',
                      }}
                    >
                      {sciMiniQr ? (
                        <img src={sciMiniQr} alt="Self-Check-in QR" style={{ width: 64, height: 64, flexShrink: 0 }} />
                      ) : (
                        <span style={{ width: 64, height: 64, flexShrink: 0, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(134,188,37,0.10)', borderRadius: 8, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                          <QrCode size={32} />
                        </span>
                      )}
                      <span style={{ minWidth: 0 }}>
                        <span style={{ display: 'block', fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
                          Self-Check-in QR
                        </span>
                        <span style={{ display: 'block', fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                          {isDe ? 'Anklicken: anzeigen, drucken + Zeitfenster festlegen' : 'Click: show, print + set time window'}
                        </span>
                      </span>
                    </button>
                  )}
                </div>
              );
            })()}
            <div style={{ flex: 1, minWidth: 0 }}>
              <h3 className="mb-16">{isDe ? 'Event-Details' : 'Event details'}</h3>
                {/* v11.28: Bookmark-Tabs statt Dropdown für schnelles Umschalten
                    zwischen Hauptevent und Sub-Events. Pro Tab wird die aktuelle
                    Teilnehmerzahl (currentParticipants aus EventContext) als
                    kleiner Badge angezeigt. */}
                {selectedEvent && (() => {
                  const isChild = !!selectedEvent.parentEventId;
                  const siblings = isChild
                    ? childEventsOf(selectedEvent.parentEventId || '')
                    : childEventsOf(selectedEvent.id);
                  if (!isChild && siblings.length === 0) return null;
                  const parent = isChild ? events.find(e => e.id === selectedEvent.parentEventId) : selectedEvent;
                  // v22.75: Der aktuell GEWÄHLTE Tab zeigt die LIVE-Zahl aus den
                  // gerade geladenen Registrierungen (registrations) — die
                  // Tab-Badges stammen sonst aus dem zwischengespeicherten
                  // Event-Zustand (letzter Listen-Load) und hinken neuen
                  // Anmeldungen hinterher (Badge 112 vs. Tabelle 126).
                  const liveSelectedActive = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length;
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const tabs: Array<{ id: string; label: string; count: number; isParent: boolean; ev: any }> = [];
                  if (parent) {
                    // v22.64: Im Klammer-Modus zeigt der HAUPT-Badge die ECHTE
                    // Zahl eindeutiger aktiver Personen über alle Sub-Events
                    // (live), nicht den gespeicherten Counter `currentParticipants`
                    // der Klammer — der zählt nicht buchbare Klammern unzuverlässig
                    // und kann verrutschen.
                    let parentCount = parent.currentParticipants || 0;
                    const pKids = childEventsOf(parent.id);
                    const haveSubData = parent.subEventsOnlyMode && pKids.length > 0 && pKids.every(c => subEventRegsByEventId[c.id] !== undefined);
                    if (haveSubData) {
                      const activeSet = new Set<string>();
                      for (const c of pKids) {
                        for (const r of (subEventRegsByEventId[c.id] || [])) {
                          if (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt') {
                            const k = (r.ParticipantEmail || '').toLowerCase().trim();
                            if (k) activeSet.add(k);
                          }
                        }
                      }
                      parentCount = activeSet.size;
                    } else if (parent.id === selectedEvent.id) {
                      // Normales Hauptevent ist selbst gewählt → Live-Zahl.
                      parentCount = liveSelectedActive;
                    }
                    tabs.push({ id: parent.id, label: parent.title || (isDe ? 'Hauptevent' : 'Main event'), count: parentCount, isParent: true, ev: parent });
                  }
                  for (const c of siblings) {
                    // v23.2: Nicht-gewählte Sub-Tabs zeigen — sofern die Liste
                    // bereits geladen ist — die LIVE-Zeilenzahl aus
                    // subEventRegsByEventId statt des veralteten Counters
                    // `currentParticipants`. Sonst „springt" der Badge je nach
                    // gewähltem Tab (gewählt = live, sonst = Cache), siehe der
                    // 188-vs-190-Effekt. Gewählter Tab bleibt die Live-Zahl der
                    // aktuell geladenen Tabelle.
                    const subRegs = subEventRegsByEventId[c.id];
                    const subLiveCount = subRegs
                      ? subRegs.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length
                      : (c.currentParticipants || 0);
                    tabs.push({ id: c.id, label: shortSubEventTitle(c.title, parent?.title) || (isDe ? 'ohne Titel' : 'untitled'), count: c.id === selectedEvent.id ? liveSelectedActive : subLiveCount, isParent: false, ev: c });
                  }
                  // v22.70: Einzelnen Tab-Button rendern (für flaches Layout
                  // UND die Sub-Event-Reihe im Klammer-Layout wiederverwendet).
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const renderTab = (t: { id: string; label: string; count: number; isParent: boolean; ev: any }): React.ReactElement => {
                    const active = t.id === selectedEvent.id;
                    return (
                      <button
                        key={t.id}
                        type="button"
                        role="tab"
                        aria-selected={active}
                        onClick={() => handleSelectEvent(t.ev).catch(() => { /* */ })}
                        style={{
                          display: 'inline-flex', alignItems: 'center', gap: 8,
                          padding: '8px 14px',
                          border: '1px solid var(--dex-gray-200)',
                          borderBottom: active ? '2px solid var(--dex-green, #86bc25)' : '1px solid var(--dex-gray-200)',
                          borderRadius: '8px 8px 0 0',
                          background: active ? '#fff' : 'var(--dex-gray-50, #fafafa)',
                          color: active ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-700)',
                          fontWeight: active ? 700 : 500,
                          fontSize: '0.85rem',
                          cursor: 'pointer',
                          marginBottom: -1,
                          whiteSpace: 'nowrap',
                          maxWidth: 280,
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          transition: 'background 0.15s, color 0.15s, border-color 0.15s',
                        }}
                        title={t.label}
                      >
                        {t.isParent && (
                          <span style={{ fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: 0.4, color: active ? 'var(--dex-green-dark)' : 'var(--dex-gray-400)' }}>
                            {isDe ? 'Haupt' : 'Main'}
                          </span>
                        )}
                        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>{t.label}</span>
                        <span
                          style={{
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            minWidth: 24, height: 20, padding: '0 6px',
                            borderRadius: 999,
                            background: active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)',
                            color: active ? '#fff' : 'var(--dex-gray-700)',
                            fontSize: '0.72rem', fontWeight: 700,
                          }}
                        >
                          {t.count}
                        </span>
                      </button>
                    );
                  };
                  const parentTab = tabs.find(tb => tb.isParent);
                  const childTabs = tabs.filter(tb => !tb.isParent);
                  // v22.70: Im Klammer-Modus die Klammer als ECHTE Klammer ÜBER
                  // den Sub-Event-Tabs darstellen (oben volle Breite, darunter die
                  // eingerückten Sub-Events). Normales Hauptevent bleibt flach.
                  const klammerLayout = !!(parentTab && parentTab.ev && parentTab.ev.subEventsOnlyMode && childTabs.length > 0);
                  if (klammerLayout && parentTab) {
                    const pActive = parentTab.id === selectedEvent.id;
                    return (
                      <div role="tablist" aria-label={isDe ? 'Event wechseln' : 'Switch event'} style={{ marginBottom: 16 }}>
                        {/* Klammer-Ebene oben — volle Breite, gefüllter Kopf. */}
                        <button
                          type="button"
                          role="tab"
                          aria-selected={pActive}
                          onClick={() => handleSelectEvent(parentTab.ev).catch(() => { /* */ })}
                          title={parentTab.label}
                          style={{
                            display: 'flex', alignItems: 'center', gap: 10, width: '100%',
                            padding: '10px 16px', cursor: 'pointer', textAlign: 'left',
                            border: `1.5px solid ${pActive ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                            borderRadius: '10px 10px 0 0',
                            background: pActive ? 'var(--dex-green, #86bc25)' : 'rgba(134,188,37,0.10)',
                            color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)',
                            fontWeight: 700, fontSize: '0.9rem',
                          }}
                        >
                          {/* v22.73: Zahl LINKS, dann Event-Name, dann „(Klammer)"
                              + Info-Icon mit Erklärung. */}
                          <span style={{
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            minWidth: 26, height: 22, padding: '0 8px', borderRadius: 999,
                            background: pActive ? 'rgba(255,255,255,0.25)' : 'var(--dex-green, #86bc25)',
                            color: '#fff', fontSize: '0.74rem', fontWeight: 700, flexShrink: 0,
                          }}>{parentTab.count}</span>
                          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1, display: 'inline-flex', alignItems: 'center', gap: 6, minWidth: 0, color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)' }}>
                            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)' }}>{parentTab.label}</span>
                            <span style={{ fontWeight: 600, opacity: 0.9, flexShrink: 0, color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)' }}>({isDe ? 'Klammer' : 'Bracket'})</span>
                            <span style={{ flexShrink: 0, display: 'inline-flex', color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)' }} onClick={e => e.stopPropagation()}>
                              <InfoTooltip placement="bottom" text={isDe
                                ? <>Das <strong>Klammer-Event selbst wird nicht gebucht</strong> — Teilnehmer melden sich nur für die einzelnen <strong>Sub-Events</strong> an. Die Klammer fasst die Sub-Events nur zusammen. Die Zahl links zeigt, <strong>wie viele Personen sich insgesamt (kumuliert) für die Sub-Events angemeldet haben</strong>.</>
                                : <>The <strong>bracket event itself is not booked</strong> — attendees only register for the individual <strong>sub-events</strong>. The bracket just groups them. The number on the left shows <strong>how many people registered for the sub-events in total (cumulative)</strong>.</>} />
                            </span>
                          </span>
                        </button>
                        {/* Sub-Events darunter — eingerückt unter einer Klammer-Linie. */}
                        <div style={{
                          display: 'flex', flexWrap: 'wrap', gap: 6, alignItems: 'flex-end',
                          marginLeft: 18, paddingLeft: 16, paddingTop: 10,
                          borderLeft: '2px solid var(--dex-green, #86bc25)',
                          borderBottom: '1px solid var(--dex-gray-200)',
                        }}>
                          {childTabs.map(t => renderTab(t))}
                        </div>
                      </div>
                    );
                  }
                  return (
                    <div
                      role="tablist"
                      aria-label={isDe ? 'Event wechseln' : 'Switch event'}
                      style={{
                        display: 'flex', flexWrap: 'wrap', gap: 6,
                        marginBottom: 16,
                        borderBottom: '1px solid var(--dex-gray-200)',
                        paddingBottom: 0,
                      }}
                    >
                      {tabs.map(t => renderTab(t))}
                    </div>
                  );
                })()}
              {/* Eigenes Row-Layout (zwei Spalten: Label fett, Wert links-
                  bündig). Das globale .settings-info SCSS macht stattdessen
                  space-between (also Wert rechts-bündig) — hier wollen wir
                  beide Spalten links ausgerichtet. */}
              {(() => {
                const rowStyle: React.CSSProperties = {
                  display: 'grid',
                  gridTemplateColumns: '160px 1fr',
                  gap: 12,
                  padding: '10px 0',
                  borderBottom: '1px solid var(--dex-gray-200)',
                  fontSize: '0.9rem',
                };
                const labelStyle: React.CSSProperties = { fontWeight: 700, color: 'var(--dex-gray-700)' };
                const valueStyle: React.CSSProperties = { fontWeight: 400, color: 'var(--dex-gray-800)' };
                return (
                  <>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Zeitraum' : 'Time period'}</span>
                      <span style={valueStyle}>{formatDate(selectedEvent.startDate)} - {formatDate(selectedEvent.endDate)}</span>
                    </div>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Organizer' : 'Organizer'}</span>
                      <span style={valueStyle}>{selectedEvent.organizers.map(o => {
                        const parts = o.split(',').map(s => s.trim());
                        return parts.length === 2 ? `${parts[1]} ${parts[0]}` : o;
                      }).join(', ')}</span>
                    </div>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Ort' : 'Location'}</span>
                      <span style={valueStyle}>{selectedEvent.location || '-'}</span>
                    </div>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Max. Teilnehmer' : 'Max. attendees'}</span>
                      <span style={valueStyle}>{(() => {
                        // v9.11: B2Run-Events nutzen Split-Kapazität statt maxParticipants —
                        // hier die Summe anzeigen statt "Unbegrenzt".
                        const split = (selectedEvent.durchstarterCapacity || 0) + (selectedEvent.funstarterCapacity || 0);
                        const eff = selectedEvent.maxParticipants && selectedEvent.maxParticipants > 0
                          ? selectedEvent.maxParticipants
                          : split;
                        return eff || (isDe ? 'Unbegrenzt' : 'Unlimited');
                      })()}</span>
                    </div>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Aktuell registriert' : 'Currently registered'}</span>
                      <span style={valueStyle}>{isConsolidatedMode ? consolidatedFiltered.length : activeRegs.length}</span>
                    </div>
                    {waitlistRegs.length > 0 && (
                      <div style={rowStyle}>
                        <span style={labelStyle}>{isDe ? 'Warteliste' : 'Waitlist'}</span>
                        <span style={valueStyle}>{waitlistRegs.length}</span>
                      </div>
                    )}
                    {/* v12.7: Aktionen-Dropdown direkt unter „Aktuell
                        registriert" — alphabetisch sortiert, mit Hover-
                        Tooltip pro Action (desc-Text rechts daneben).
                        Ersetzt die separate Aktionen-Card. */}
                    <div style={{ marginTop: 14 }}>
                      <div style={{ ...labelStyle, marginBottom: 6 }}>
                        {isDe ? 'Aktionen' : 'Actions'}
                      </div>
                      <ActionsDropdown isDe={isDe} />
                    </div>
                    {/* v12.2: 'Abgefragte Felder'-Zeile entfernt — die
                        Custom-Field-Pills hier waren redundant; sie tauchen
                        ohnehin als Spalten in der Teilnehmer-Tabelle auf. */}
                  </>
                );
              })()}
            </div>
          </div>
        </div>
        {/* v22.5: „Nächste Schritte"-Box rechts neben der Detail-Card — nur für
            Entwürfe (Admin/Organizer). Erklärt, was nach dem Anlegen noch zu tun
            ist: finalisieren, Test-An-/Abmeldung, live schalten (+ wer es sieht),
            Einladungsmail verschicken, Anmeldungen verfolgen. */}
        {(isAdmin || isOrganizerFor(selectedEvent)) && !!selectedEvent.isFictive && !selectedEvent.isDemoShowcase && (
          <aside style={{ flex: '0 1 340px', minWidth: 290 }}>
            <div className="card" style={{ padding: 20, background: 'rgba(134,188,37,0.05)', border: '1px solid var(--dex-green, #86bc25)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', display: 'inline-flex' }}><Info size={18} /></span>
                <h3 style={{ margin: 0, fontSize: '1rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'Nächste Schritte' : 'Next steps'}</h3>
              </div>
              <p style={{ margin: '0 0 14px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {isDe ? 'Dein Event ist angelegt — so machst du es startklar:' : 'Your event is created — here is how to get it ready:'}
              </p>
              <ol style={{ margin: 0, padding: 0, listStyle: 'none', display: 'flex', flexDirection: 'column', gap: 13 }}>
                {(() => {
                  // v22.6: Sichtbarkeit eines Events als Klartext (für Haupt-
                  // event UND je Sub-Section).
                  const visText = (lc: string[], au: string[]): string => {
                    if (lc.length === 0 && au.length === 0) {
                      return isDe ? 'alle Mitarbeiter von Deloitte Deutschland' : 'all Deloitte Germany employees';
                    }
                    const parts: string[] = [];
                    if (lc.length) parts.push((isDe ? 'Standorte: ' : 'Locations: ') + lc.join(', '));
                    if (au.length) parts.push(isDe ? `${au.length} Verteiler/Personen` : `${au.length} distributions/people`);
                    return parts.join(isDe ? ' und ' : ' and ');
                  };
                  const locs = (selectedEvent.locationAudience || []).filter(Boolean);
                  const auds = (selectedEvent.audienceFilter || []).filter(Boolean);
                  const children = childEventsOf(selectedEvent.id);
                  const hasChildren = children.length > 0;
                  const parentVisText = visText(locs, auds);
                  const visSummary = (isDe ? 'Sichtbar für ' : 'Visible to ') + parentVisText + '.';
                  // Pro Sub-Section die Sichtbarkeit; wenn alle gleich → nur einmal.
                  // v22.22: Eine Sub-Section OHNE eigene Filter ist zur Laufzeit
                  // NICHT für „alle Mitarbeiter" sichtbar — der Zugang läuft immer
                  // über das Gesamt-Event (dessen Sichtbarkeit gilt dann auch für
                  // die Sub-Section). Das auch so benennen, statt irreführend
                  // „alle Mitarbeiter von Deloitte Deutschland" anzuzeigen.
                  const parentRestricted = locs.length > 0 || auds.length > 0;
                  const childVis = children.map(c => {
                    const cl = (c.locationAudience || []).filter(Boolean);
                    const ca = (c.audienceFilter || []).filter(Boolean);
                    const inherits = cl.length === 0 && ca.length === 0 && parentRestricted;
                    return {
                      title: shortSubEventTitle(c.title, selectedEvent.title) || c.title,
                      inherits,
                      text: inherits
                        ? (isDe ? 'wie das Gesamt-Event (keine eigene Einschränkung)' : 'same as the overall event (no own restriction)')
                        : visText(cl, ca),
                    };
                  });
                  const allChildrenSame = childVis.length > 0 && childVis.every(c => c.text === childVis[0].text);
                  // v22.8: Wenn Gesamt-Event UND alle Sub-Sections dieselbe
                  // Sichtbarkeit haben, ist die Unterscheidung überflüssig — dann
                  // nur EINE Aussage zeigen. (v22.22: gilt auch, wenn alle
                  // Sub-Sections die Sichtbarkeit des Gesamt-Events erben.)
                  const everythingSame = hasChildren && allChildrenSame && (childVis[0].text === parentVisText || childVis[0].inherits);
                  const steps: Array<{ title: string; body: React.ReactNode }> = [
                    {
                      title: isDe ? 'Event finalisieren' : 'Finalize the event',
                      body: isDe
                        ? 'Über „Event bearbeiten" Felder, Bild und Texte vervollständigen.'
                        : 'Use “Edit event” to complete fields, image and texts.',
                    },
                    {
                      title: isDe ? 'Test-An- und Abmeldung' : 'Test registration & cancellation',
                      body: isDe
                        ? 'Melde dich einmal selbst an und wieder ab, um zu prüfen, ob die automatische Kommunikation (Bestätigungs-Mail, Outlook-Termin, Abmelde-Mail) richtig ankommt.'
                        : 'Register and cancel yourself once to check that the automatic communication (confirmation email, Outlook invite, cancellation email) works correctly.',
                    },
                    {
                      title: isDe ? 'Event live schalten' : 'Publish the event',
                      body: (
                        <>
                          {isDe
                            ? 'Oben über das Status-Häkchen „Entwurf → Aktiv" schalten. Danach ist es für die berechtigten Gruppen sichtbar.'
                            : 'Switch the status badge above from “Draft → Active”. It is then visible to the eligible groups.'}
                          <span style={{ display: 'block', marginTop: 5, padding: '6px 8px', borderRadius: 6, background: '#fff', border: '1px solid var(--dex-gray-200)', color: 'var(--dex-gray-600)', fontSize: '0.74rem', lineHeight: 1.45 }}>
                            {everythingSame ? (
                              // Gesamt-Event und alle Sub-Sections gleich → eine Aussage.
                              <>{visSummary} {isDe ? `(Gesamt-Event und alle ${childVis.length} Sub-Section${childVis.length === 1 ? '' : 's'}.)` : `(Overall event and all ${childVis.length} sub-section${childVis.length === 1 ? '' : 's'}.)`}</>
                            ) : !hasChildren ? (
                              <>{visSummary}</>
                            ) : (
                              <>
                                <strong style={{ color: 'var(--dex-gray-700)' }}>{isDe ? 'Gesamt-Event: ' : 'Overall event: '}</strong>{visSummary}
                                <span style={{ display: 'block', marginTop: 4 }}>
                                  {allChildrenSame ? (
                                    <>
                                      <strong style={{ color: 'var(--dex-gray-700)' }}>
                                        {isDe ? `Für alle ${childVis.length} Sub-Sections gilt: ` : `For all ${childVis.length} sub-sections: `}
                                      </strong>
                                      {childVis[0].text}.
                                    </>
                                  ) : (
                                    <>
                                      <strong style={{ color: 'var(--dex-gray-700)' }}>{isDe ? 'Sub-Sections:' : 'Sub-sections:'}</strong>
                                      {childVis.map((c, ci) => (
                                        <span key={ci} style={{ display: 'block', paddingLeft: 8 }}>• <strong>{c.title}:</strong> {c.text}</span>
                                      ))}
                                    </>
                                  )}
                                </span>
                              </>
                            )}
                          </span>
                        </>
                      ),
                    },
                    {
                      title: isDe ? 'Einladungsmail verschicken' : 'Send the invitation email',
                      body: (
                        <>
                          {isDe
                            ? 'Optional kannst du die Einladung mit Anmelde-Link direkt über DEX verschicken — an dich zum Weiterleiten oder an den Mailverteiler.'
                            : 'Optionally send the invitation with the registration link directly via DEX — to yourself for forwarding or to the mail distribution.'}
                          {' '}
                          <button
                            type="button"
                            onClick={openInviteModal}
                            style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, fontSize: '0.78rem', textDecoration: 'underline' }}
                          >
                            {isDe ? 'Einladungsmail öffnen' : 'Open invitation email'}
                          </button>
                        </>
                      ),
                    },
                    {
                      title: isDe ? 'Anmeldungen verfolgen' : 'Track registrations',
                      body: isDe
                        ? 'Sobald sich Teilnehmer anmelden, siehst du hier im Admin-Panel alle Infos — Anzahl, Status und die komplette Teilnehmerliste.'
                        : 'As soon as participants register, you see everything here in the admin panel — count, status and the full participant list.',
                    },
                  ];
                  return steps.map((s, i) => (
                    <li key={i} style={{ display: 'flex', gap: 10, alignItems: 'flex-start' }}>
                      <span style={{ flexShrink: 0, width: 24, height: 24, borderRadius: '50%', background: 'var(--dex-green, #86bc25)', color: '#fff', fontWeight: 700, fontSize: '0.82rem', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}>{i + 1}</span>
                      <span style={{ fontSize: '0.82rem', lineHeight: 1.5 }}>
                        <strong style={{ color: 'var(--dex-gray-800)' }}>{s.title}.</strong>{' '}
                        <span style={{ color: 'var(--dex-gray-600)' }}>{s.body}</span>
                      </span>
                    </li>
                  ));
                })()}
              </ol>
            </div>
          </aside>
        )}
        {/* v22.16: „Hinweise"-Box für AKTIVE Events — Pendant zur „Nächste
            Schritte"-Box bei Entwürfen. Zeigt smarte Empfehlungen (z.B.
            englischer Inhalt → Anmeldesprache fest auf Englisch stellen).
            Erscheint nur, wenn mindestens ein Hinweis zutrifft; jeder Hinweis
            ist pro Event ausblendbar (localStorage). */}
        {(isAdmin || isOrganizerFor(selectedEvent)) && !selectedEvent.isFictive && !selectedEvent.isDemoShowcase && (() => {
          void hintsDismissTick; // erzwingt Re-Render nach „Ausblenden"
          const dismissKey = (id: string): string => `dex_hint_dismiss_${selectedEvent.id}_${id}`;
          const isDismissed = (id: string): boolean => {
            try { return window.localStorage.getItem(dismissKey(id)) === '1'; } catch { return false; }
          };
          const hints: Array<{ id: string; title: string; body: React.ReactNode; action?: React.ReactNode }> = [];
          // 1) Englischer Inhalt, aber Anmeldesprache nicht fest auf Englisch.
          const fieldsText = (selectedEvent.eventSpecificFields || [])
            .map(f => [f.label, f.helpText, (f.options || []).join(' ')].filter(Boolean).join(' '))
            .join(' ');
          const contentText = `${stripHtmlToText(selectedEvent.description || '')} ${fieldsText}`;
          if ((selectedEvent.registrationLanguage || '') !== 'en' && looksEnglishText(contentText)) {
            hints.push({
              id: 'lang-en',
              title: isDe ? 'Anmeldesprache auf Englisch festlegen?' : 'Fix registration language to English?',
              body: isDe
                ? 'Beschreibung und Felder dieses Events sind offenbar auf Englisch — die Anmeldeseite folgt aber der App-Sprache des Teilnehmers. Bei deutscher App-Einstellung mischt das Formular dann Deutsch (Buttons, Hinweise, Datenschutz) und Englisch (Inhalte). Empfehlung: die Anmeldesprache fest auf Englisch stellen. (Auch im Wizard änderbar: Schritt 5 „Felder" → „Sprache des Anmeldeformulars".)'
                : 'The description and fields of this event appear to be in English — but the registration page follows each participant\'s app language. With a German app setting the form then mixes German (buttons, hints, privacy note) and English (content). Recommendation: fix the registration language to English. (Also changeable in the wizard: step 5 “Fields” → “Registration form language”.)',
              action: (
                <button
                  type="button"
                  className="btn btn-primary"
                  disabled={hintLangBusy}
                  style={{ fontSize: '0.78rem', padding: '6px 12px' }}
                  onClick={() => {
                    (async () => {
                      setHintLangBusy(true);
                      const ok = await updateEvent(selectedEvent.id, { 'RegistrationLanguage': 'en' });
                      setHintLangBusy(false);
                      if (ok) {
                        setSelectedEvent(prev => prev ? { ...prev, registrationLanguage: 'en' } : prev);
                        await refreshEvents();
                        showAlert(isDe
                          ? 'Anmeldesprache auf Englisch festgelegt — das Anmeldeformular erscheint jetzt für alle Teilnehmer durchgängig auf Englisch.'
                          : 'Registration language fixed to English — the registration form now appears consistently in English for everyone.', { variant: 'success' });
                      } else {
                        showAlert(isDe ? 'Anmeldesprache konnte nicht gespeichert werden.' : 'Could not save the registration language.', { variant: 'error' });
                      }
                    })().catch(() => { /* */ });
                  }}
                >
                  {hintLangBusy ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Auf Englisch festlegen' : 'Fix to English')}
                </button>
              ),
            });
          }
          // 2) Beschreibung fehlt oder ist sehr kurz.
          if (stripHtmlToText(selectedEvent.description || '').length < 20) {
            hints.push({
              id: 'no-desc',
              title: isDe ? 'Beschreibung ergänzen' : 'Add a description',
              body: isDe
                ? 'Das Event hat (fast) keine Beschreibung — Teilnehmer sehen auf der Anmeldeseite dann kaum, worum es geht. Über „Event bearbeiten" → Schritt 1 (Grundlagen) ergänzen.'
                : 'The event has (almost) no description — participants see very little about it on the registration page. Add one via “Edit event” → step 1 (Basics).',
            });
          }
          // 3) Event-Bild fehlt.
          if (!selectedEvent.imageUrl) {
            hints.push({
              id: 'no-image',
              title: isDe ? 'Event-Bild hochladen' : 'Upload an event image',
              body: isDe
                ? 'Ohne Bild wirkt die Event-Karte in der Übersicht und der Mail-Kopf deutlich weniger einladend. Über „Event bearbeiten" → Schritt 1 (Grundlagen) hochladen.'
                : 'Without an image the event card in the list and the email header look much less inviting. Upload one via “Edit event” → step 1 (Basics).',
            });
          }
          // 3b) v22.63: Der frühere allgemeine „Sichtbarkeit gilt fürs ganze
          // Event"-Hinweis ist entfallen — er erschien immer und war reines
          // Erklär-Rauschen. Hinweise kommen jetzt nur noch bei echten
          // möglichen Inkonsistenzen (3c/3d/3e).
          // 3c) v22.59: Sehr kleine Zielgruppe — nur wenige Einzeladressen, kein
          // Standortfilter und kein „alle Mitarbeiter". Oft hat der Organizer
          // dann nur sich/die Organizer eingetragen und den eigentlichen Kreis
          // vergessen.
          {
            const aud = selectedEvent.audienceFilter || [];
            const loc = selectedEvent.locationAudience || [];
            const hasAllPattern = aud.some(a => { const f = (a || '').toLowerCase(); return f === 'all' || f === 'deall'; });
            const resolved = (selectedEvent.audienceResolvedEmails || []).map(s => (s || '').trim()).filter(Boolean);
            const atCount = aud.filter(a => a.indexOf('@') >= 0).length;
            const effCount = resolved.length > 0 ? resolved.length : atCount;
            if (aud.length > 0 && loc.length === 0 && !hasAllPattern && effCount > 0 && effCount < 10) {
              hints.push({
                id: 'tiny-audience',
                title: isDe ? 'Sehr kleine Zielgruppe — Absicht?' : 'Very small audience — intended?',
                body: isDe
                  ? <>Die Sichtbarkeit umfasst aktuell nur <strong>{effCount} {effCount === 1 ? 'Person' : 'Personen'}</strong> (einzelne Adressen — kein Standortfilter, kein Mailverteiler, nicht „alle Mitarbeiter“). Falls du eigentlich einen größeren Kreis erreichen willst, ergänze einen Standort oder Verteiler über „Event bearbeiten“ → Schritt 4. Ist das Event bewusst nur für diese wenigen Personen, kannst du den Hinweis ausblenden.</>
                  : <>The visibility currently covers only <strong>{effCount} {effCount === 1 ? 'person' : 'people'}</strong> (individual addresses — no location filter, no mailing list, not “all employees”). If you actually want to reach a larger group, add a location or distribution list via “Edit event” → step 4. If the event is intentionally for these few people only, you can dismiss this hint.</>,
              });
            }
          }
          // 3d) v22.60: Sub-Events mit EIGENER kleiner Sichtbarkeit. Die Klammer
          // kann korrekt sein, aber ein Sub-Event hat eine abweichende, sehr
          // enge Zielgruppe (z.B. nur ein paar Einzeladressen) → eigener Hinweis
          // pro betroffenem Sub-Event. Nur wenn die Sub-Sichtbarkeit sich von
          // der Klammer unterscheidet (sonst ist es nur geerbt).
          {
            const parentAudKey = (selectedEvent.audienceFilter || []).join('|');
            const parentLocKey = (selectedEvent.locationAudience || []).join('|');
            const smallSubs: string[] = [];
            for (const ch of childEventsOf(selectedEvent.id)) {
              const cAud = ch.audienceFilter || [];
              const cLoc = ch.locationAudience || [];
              if (cAud.length === 0 && cLoc.length === 0) continue; // keine eigene Sichtbarkeit
              if (cAud.join('|') === parentAudKey && cLoc.join('|') === parentLocKey) continue; // nur geerbt
              const cHasAll = cAud.some(a => { const f = (a || '').toLowerCase(); return f === 'all' || f === 'deall'; });
              const cResolved = (ch.audienceResolvedEmails || []).map(s => (s || '').trim()).filter(Boolean);
              const cAt = cAud.filter(a => a.indexOf('@') >= 0).length;
              const cEff = cResolved.length > 0 ? cResolved.length : cAt;
              if (cAud.length > 0 && cLoc.length === 0 && !cHasAll && cEff > 0 && cEff < 10) {
                smallSubs.push(`${shortSubEventTitle(ch.title, selectedEvent.title)} (${cEff})`);
              }
            }
            if (smallSubs.length > 0) {
              hints.push({
                id: 'tiny-sub-audience',
                title: isDe ? 'Sub-Event mit sehr kleiner Zielgruppe' : 'Sub-event with very small audience',
                body: isDe
                  ? <>Diese Sub-Events haben eine <strong>eigene, sehr kleine Sichtbarkeit</strong> (nur wenige Einzeladressen, kein Standort/Verteiler): <strong>{smallSubs.join(', ')}</strong>. Die Klammer-Sichtbarkeit kann passen, aber wer in der jeweiligen Sub-Event-Liste nicht steht, kann sich für dieses Sub-Event nicht anmelden. Prüfen/anpassen über „Event bearbeiten“ → Schritt 4 (Tab des Sub-Events) → „Sichtbarkeit prüfen“.</>
                  : <>These sub-events have their <strong>own, very small visibility</strong> (only a few individual addresses, no location/distribution list): <strong>{smallSubs.join(', ')}</strong>. The bracket visibility may be fine, but anyone not in the respective sub-event list cannot register for that sub-event. Check/adjust via “Edit event” → step 4 (sub-event tab) → “Check visibility”.</>,
              });
            }
          }
          // 3e) v22.61: Person nur im Sub-Event, aber nicht in der Klammer →
          // kann das Event gar nicht öffnen (Zugang läuft zur Laufzeit über die
          // Klammer-Sichtbarkeit; Sub-Events werden erst auf der Anmeldeseite
          // der Klammer gefiltert). Risiko besteht, wenn die Klammer eingeschränkt
          // ist UND ein Sub-Event eine abweichende eigene Sichtbarkeit hat.
          {
            const pAud = selectedEvent.audienceFilter || [];
            const pLoc = selectedEvent.locationAudience || [];
            const parentShowsAll = (pAud.length === 0 && pLoc.length === 0)
              || pAud.some(a => { const f = (a || '').toLowerCase(); return f === 'all' || f === 'deall'; });
            const pAudKey = pAud.join('|');
            const pLocKey = pLoc.join('|');
            const riskySubs: string[] = [];
            if (!parentShowsAll) {
              for (const ch of childEventsOf(selectedEvent.id)) {
                const cAud = ch.audienceFilter || [];
                const cLoc = ch.locationAudience || [];
                if (cAud.length === 0 && cLoc.length === 0) continue; // erbt die Klammer → kein zusätzliches Publikum
                if (cAud.join('|') === pAudKey && cLoc.join('|') === pLocKey) continue; // identisch zur Klammer
                riskySubs.push(shortSubEventTitle(ch.title, selectedEvent.title));
              }
            }
            if (riskySubs.length > 0) {
              hints.push({
                id: 'sub-not-in-parent',
                title: isDe ? 'Sub-Event für mehr Leute geöffnet als das Event?' : 'Sub-event open to more people than the event?',
                body: isDe
                  ? <>Bei diesen Sub-Events hast du <strong>andere Leute ausgewählt als beim Event selbst</strong>: <strong>{riskySubs.join(', ')}</strong>. Das kann ein Problem sein: Wer nur beim Sub-Event ausgewählt ist, aber nicht beim Event, <strong>kann das Event gar nicht öffnen</strong> — und sieht das Sub-Event deshalb nie. Damit das passt, sollten beim Event mindestens alle dabei sein, die irgendein Sub-Event sehen sollen. Zum Vergleichen: „Sichtbarkeit prüfen“ (dort pro Sub-Event).</>
                  : <>For these sub-events you picked <strong>different people than for the event itself</strong>: <strong>{riskySubs.join(', ')}</strong>. That can be a problem: anyone picked only for the sub-event but not for the event <strong>can’t open the event at all</strong> — and therefore never sees the sub-event. To make it work, the event should include at least everyone who should see any sub-event. To compare: “Check visibility” (per sub-event).</>,
              });
            }
          }
          // 4) v22.34: End-Datum fehlt (Hauptevent oder Sub-Event) — ohne Ende
          // kann der Outlook-Termin nicht angelegt werden (der Kalendereintrag
          // braucht Start UND Ende; das Sub-Event bekommt dann nie eine
          // OutlookEventId). Praxisfall: Organizerin vergaß beim Anlegen das
          // End-Datum eines Sub-Events → kein Outlook-Termin für die Teilnehmer.
          {
            const noEndChildren = childEventsOf(selectedEvent.id)
              .filter(c => !!(c.startDate || '').trim() && !(c.endDate || '').trim());
            const mainNoEnd = !!(selectedEvent.startDate || '').trim() && !(selectedEvent.endDate || '').trim();
            if (mainNoEnd || noEndChildren.length > 0) {
              const names: string[] = [];
              if (mainNoEnd) names.push(isDe ? 'das Hauptevent' : 'the main event');
              for (const c of noEndChildren) {
                names.push(`„${shortSubEventTitle(c.title, selectedEvent.title) || c.title}"`);
              }
              hints.push({
                id: 'no-enddate',
                title: isDe
                  ? 'End-Datum fehlt — Outlook-Termin kann nicht erstellt werden'
                  : 'End date missing — Outlook invite cannot be created',
                body: isDe
                  ? <>Ohne End-Datum kann für die Teilnehmer <strong>kein Outlook-Termin</strong> angelegt werden (ein Kalendereintrag braucht Start UND Ende) — betroffen: <strong>{names.join(', ')}</strong>. Bitte über „Event bearbeiten“ das End-Datum nachtragen (Hauptevent: Schritt 1 „Grundlagen“, Sub-Events: Schritt 2 „Sub-Events“). Beim Speichern fragt die App dann, ob der Outlook-Termin angelegt bzw. aktualisiert werden soll.</>
                  : <>Without an end date <strong>no Outlook invite</strong> can be created for attendees (a calendar entry needs a start AND an end) — affected: <strong>{names.join(', ')}</strong>. Please add the end date via “Edit event” (main event: step 1 “Basics”, sub-events: step 2 “Sub-events”). When saving, the app then asks whether the Outlook invite should be created or updated.</>,
              });
            }
          }
          // 5) v22.69: Hauptevent/Klammer ist live, aber ein Sub-Event steht
          // noch auf Entwurf — Entwurf-Sub-Events sind für reguläre Teilnehmer
          // NICHT buchbar (seit v22.68). Der Organizer denkt sonst, alles sei
          // buchbar.
          if (!selectedEvent.isFictive) {
            const draftKids = childEventsOf(selectedEvent.id).filter(c => c.isFictive);
            if (draftKids.length > 0) {
              const draftNames = draftKids.map(c => shortSubEventTitle(c.title, selectedEvent.title)).join(', ');
              hints.push({
                id: 'draft-subevent-live-parent',
                title: isDe ? 'Sub-Event noch im Entwurf — nicht buchbar' : 'Sub-event still a draft — not bookable',
                body: isDe
                  ? <>Das Event ist live, aber diese Sub-Events stehen noch auf <strong>Entwurf</strong>: <strong>{draftNames}</strong>. Entwurf-Sub-Events sind für reguläre Teilnehmer <strong>nicht sichtbar und nicht buchbar</strong>. Wenn sie buchbar sein sollen, schalte sie über den Status-Badge oben (Entwurf ⇄ Aktiv) auf den jeweiligen Sub-Event-Tab live.</>
                  : <>The event is live, but these sub-events are still in <strong>draft</strong>: <strong>{draftNames}</strong>. Draft sub-events are <strong>not visible and not bookable</strong> for regular attendees. If they should be bookable, publish them via the status badge (draft ⇄ active) on the respective sub-event tab.</>,
              });
            }
          }
          const visible = hints.filter(h => !isDismissed(h.id));
          if (visible.length === 0) return null;
          return (
            <aside style={{ flex: '0 1 340px', minWidth: 290 }}>
              <div className="card" style={{ padding: 20, background: 'rgba(0,118,168,0.04)', border: '1px solid var(--dex-blue, #0076a8)' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                  <span style={{ color: 'var(--dex-blue, #0076a8)', display: 'inline-flex' }}><Info size={18} /></span>
                  <h3 style={{ margin: 0, fontSize: '1rem', color: 'var(--dex-blue, #0076a8)' }}>{isDe ? 'Hinweise zu diesem Event' : 'Hints for this event'}</h3>
                </div>
                <p style={{ margin: '0 0 14px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe ? 'Der App sind ein paar Dinge aufgefallen, die du dir kurz anschauen solltest:' : 'The app noticed a few things worth a quick look:'}
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
                  {visible.map(h => (
                    <div key={h.id} style={{ borderTop: '1px solid rgba(0,118,168,0.15)', paddingTop: 12 }}>
                      <div style={{ fontWeight: 700, fontSize: '0.85rem', color: 'var(--dex-gray-800)', marginBottom: 4 }}>{h.title}</div>
                      <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>{h.body}</div>
                      <div style={{ marginTop: 8, display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                        {h.action}
                        <button
                          type="button"
                          onClick={() => {
                            try { window.localStorage.setItem(dismissKey(h.id), '1'); } catch { /* */ }
                            setHintsDismissTick(t => t + 1);
                          }}
                          style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--dex-gray-500)', fontSize: '0.74rem', textDecoration: 'underline' }}
                        >
                          {isDe ? 'Hinweis ausblenden' : 'Dismiss hint'}
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </aside>
          );
        })()}
        </div>

        {/* v7.6: Aktionen-Bereich als Kachel-Grid (auto-fit ab 220px, max 4
            pro Zeile auf Desktop). Default Grau, beim Hover Deloitte-Grün mit
            leichtem Schatten. Jede Kachel zeigt SVG-Icon + Titel + ausführliche
            Beschreibung + Rollen-Badge ("Organizer" oder "Nur Admin"). Die
            ehemals in der TN-Toolbar versteckten Wartungs-Aktionen (IDs neu
            vergeben, Spalten fixen, Felder reparieren, Profile neu laden) sind
            seit v7.6 hier integriert — der Organizer/Admin findet alle Event-
            relevanten Aktionen an einem Ort. QR-Scanner sehen den ganzen Block
            nicht. */}
        {!isQRScannerOnlyForSelected && !selectedEvent.isDemoShowcase && (
        <ActionsCollapsibleCard isDe={isDe}>
          <div className="admin-actions-grid" style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
            gap: 12,
          }}>
            {/* v9.20: Check-In starten — prominent als erster Tile.
                Sowohl Organizer als auch Check-In-Team-Mitglieder dürfen
                diese Aktion auslösen (siehe Header.canCheckIn-Logik). */}
            <ActionTile
              icon={<Hash size={18} />}
              category="checkin"
              title={t('admin.checkin')}
              desc={isDe
                ? 'Öffnet das Check-In-Tool: QR-Codes scannen, manuell ein-/auschecken, Live-KPIs (wie viele angemeldet / eingecheckt / ausstehend) sehen. Am Eventtag das wichtigste Werkzeug.'
                : 'Opens the check-in tool: scan QR codes, check in/out manually, see live KPIs (how many registered / checked in / pending). The most important tool on event day.'}
              badge="organizer"
              onClick={() => navigate('check-in', selectedEvent.id)}
            />

            {/* v9.20: QR-Codes versenden als ActionTile (Modal-Trigger). */}
            <ActionTile
              icon={<Send size={18} />}
              category="checkin"
              title={isSendingQR ? (isDe ? `QR-Codes werden versendet... (${qrSentCount})` : `Sending QR codes... (${qrSentCount})`) : (isDe ? 'QR-Codes versenden' : 'Send QR codes')}
              desc={isDe
                ? 'Öffnet ein Modal mit zwei Optionen: Test (nur an dich) oder Versand an alle ohne Code. Nach dem ERSTEN Versand bekommt jede weitere Anmeldung an diesem Event ihren QR-Code automatisch per Mail — auch nach der Anmeldefrist. Teilnehmer finden ihren QR-Code zusätzlich jederzeit unter „Meine Events".'
                : 'Opens a modal with two options: test (only to you) or send to everyone without a code. After the FIRST send, every further registration on this event receives its QR code automatically by email — even after the registration deadline. Participants can also find their QR code anytime under "My events".'}
              badge="organizer"
              busy={isSendingQR}
              onClick={() => {
                setQrSendResult(null);
                setQrSendModalOpen(true);
              }}
            />

            {/* v11.89/v20.3: Der Event-Live/Entwurf-Toggle ist aus dem
                Aktionen-Menü ausgezogen — der Status-Badge neben dem
                Event-Titel ist jetzt selbst der klickbare Umschalter. */}

            {/* 1. Event bearbeiten */}
            <ActionTile
              icon={<Pencil size={18} />}
              category="event"
              title={t('admin.editbutton') || 'Event bearbeiten'}
              desc={isDe
                ? 'Öffnet das Event im Schritt-für-Schritt-Wizard. Titel, Datum, Ort, Kapazität, Custom-Fields, E-Mail-Templates und Quiz nachträglich anpassen.'
                : 'Opens the event in the step-by-step wizard. Adjust title, date, location, capacity, custom fields, email templates and quiz afterwards.'}
              badge="organizer"
              onClick={() => navigate('edit-event', selectedEvent.id)}
            />

            {/* 2. Teilnehmerliste in SharePoint öffnen */}
            <ActionTile
              icon={<ExternalLink size={18} />}
              category="event"
              title={t('admin.opensp') || 'In SharePoint öffnen'}
              desc={isDe
                ? 'Öffnet die SharePoint-Teilnehmerliste der Subsite in einem neuen Tab — für tiefere Bearbeitung jenseits dieser App (z.B. Massen-Edit per Spreadsheet-View).'
                : 'Opens the SharePoint participant list of the subsite in a new tab — for deeper editing beyond this app (e.g. bulk edit via spreadsheet view).'}
              badge="organizer"
              href={selectedEvent.subsiteUrl ? `${selectedEvent.subsiteUrl}/Lists/Teilnehmer/AllItems.aspx` : `${siteUrl}/Lists`}
            />

            {/* v18.33/v20.1: Self-Check-in — QR-PDF + rotierende Live-Anzeige.
                Seit v20.1 IMMER sichtbar für Admin/(Co-)Organizer: hat das
                Event noch keinen aktiven Token, wird Self-Check-in beim Klick
                automatisch aktiviert (Token erzeugen + am Event speichern). */}
            {(isAdmin || isOrganizerFor(selectedEvent)) && (
              <ActionTile
                icon={<QrCode size={18} />}
                category="checkin"
                subCategory={isDe ? 'Self-Check-in' : 'Self check-in'}
                title={isDe ? 'Self-Check-in einstellen' : 'Set up self check-in'}
                desc={isDe
                  ? 'Öffnet die Self-Check-in-Übersicht dieses Events: großer QR-Code, PDF-Download, Live-Anzeige und das Check-in-Zeitfenster (Von/Bis) — so kannst du das Zeitfenster auch schon Wochen vor dem Event festlegen. Standard: 2 Stunden vor Event-Start bis Event-Ende.'
                  : 'Opens the self check-in overview of this event: large QR code, PDF download, live display and the check-in time window (from/until) — so you can set the window weeks before the event. Default: 2 hours before event start until event end.'}
                badge="organizer"
                busy={sciBusy}
                onClick={() => { openSelfCheckInModal().catch(() => { /* best-effort */ }); }}
              />
            )}
            {(isAdmin || isOrganizerFor(selectedEvent)) && (
              <ActionTile
                icon={<Download size={18} />}
                category="checkin"
                subCategory={isDe ? 'Self-Check-in' : 'Self check-in'}
                title={isDe ? 'Self-Check-in: QR-PDF' : 'Self check-in: QR PDF'}
                desc={isDe
                  ? 'Lädt ein druckbares PDF mit dem QR-Code und einer kurzen Anleitung herunter. Zum Aushängen am Eingang — Teilnehmer scannen mit der Handy-Kamera und checken sich selbst ein. Das Check-in-Zeitfenster (Standard: 2 Stunden vor Start bis Event-Ende) begrenzt, wann der Code funktioniert — einstellbar über die QR-Kachel unter dem Event-Bild.'
                  : 'Downloads a printable PDF with the QR code and short instructions. For posting at the entrance — attendees scan with their phone camera and check themselves in. The check-in time window (default: 2 hours before start until event end) limits when the code works — adjustable via the QR tile below the event image.'}
                badge="organizer"
                onClick={() => {
                  (async () => {
                    const token = await ensureSelfCheckInReady(selectedEvent);
                    if (!token) return;
                    await downloadSelfCheckInPdf({
                      eventTitle: selectedEvent.title || 'Event',
                      eventDateLabel: selectedEvent.startDate ? new Date(selectedEvent.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '',
                      locationLabel: selectedEvent.location || '',
                      token,
                    });
                  })().catch(() => { /* best-effort */ });
                }}
              />
            )}
            {(isAdmin || isOrganizerFor(selectedEvent)) && (
              <ActionTile
                icon={<QrCode size={18} />}
                category="checkin"
                subCategory={isDe ? 'Self-Check-in' : 'Self check-in'}
                title={isDe ? 'Self-Check-in: Live-Anzeige' : 'Self check-in: live display'}
                desc={isDe
                  ? 'Öffnet eine rotierende QR-Anzeige für einen Bildschirm am Eingang (Laptop, Tablet, Beamer). Der Code wechselt automatisch — ein abfotografierter Code verfällt sofort. Die foto-sichere Variante.'
                  : 'Opens a rotating QR display for a screen at the entrance (laptop, tablet, projector). The code changes automatically — a photographed code expires instantly. The photo-safe option.'}
                badge="organizer"
                onClick={() => {
                  (async () => {
                    const token = await ensureSelfCheckInReady(selectedEvent);
                    if (token) navigate('self-checkin-display', selectedEvent.id);
                  })().catch(() => { /* best-effort */ });
                }}
              />
            )}

            {/* v10.19: Deep-Link kopieren — Organizer/Admin können den Link
                des aktuell offenen Events in die Zwischenablage legen und z.B.
                an Co-Organizer / Helfer weitergeben. Zielseite ist exakt
                dieses Admin-Center-Detail (?action=admin&event=<SP-ID>). Beim
                Aufruf landet der Empfänger nach Login direkt auf der gleichen
                Detail-Ansicht statt in der Event-Auswahl-Liste. */}
            <ActionTile
              icon={<Link2 size={18} />}
              category="event"
              title={copiedDeepLink ? (t('admin.copied') || 'Kopiert') : (isDe ? 'Deep-Link kopieren' : 'Copy deep link')}
              desc={isDe
                ? 'Legt den direkten Link auf dieses Event-Admin in die Zwischenablage. Per Mail / Teams an Co-Organizer schicken — sie landen nach Login direkt hier, ohne sich erst durch die Event-Liste klicken zu müssen.'
                : 'Copies the direct link to this event admin to the clipboard. Send it via email / Teams to co-organizers — after login they land directly here without clicking through the event list first.'}
              badge="organizer"
              onClick={() => {
                const base = (typeof window !== 'undefined' && window.location)
                  ? `${window.location.origin}${window.location.pathname}`
                  : `${siteUrl}/SitePages/DEX.aspx`;
                const url = `${base}?env=WebView&action=admin&event=${selectedEvent.id}`;
                if (typeof navigator !== 'undefined' && navigator.clipboard) {
                  navigator.clipboard.writeText(url).then(() => {
                    setCopiedDeepLink(true);
                    setTimeout(() => setCopiedDeepLink(false), 2000);
                  }).catch(() => { showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{url}</span>, { title: isDe ? 'Deep-Link manuell kopieren' : 'Copy deep link manually' }); });
                } else {
                  showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{url}</span>, { title: isDe ? 'Deep-Link manuell kopieren' : 'Copy deep link manually' });
                }
              }}
            />

            {/* 3. E-Mail-Adressen kopieren */}
            <ActionTile
              icon={<Copy size={18} />}
              category="mails"
              title={copiedEmails ? (t('admin.copied') || 'Kopiert') : (t('admin.copyemails') || 'E-Mails kopieren')}
              desc={isDe
                ? 'Legt alle aktiven Teilnehmer-Mails (Semikolon-getrennt) in die Zwischenablage. Direkt in Outlook-Empfänger oder externe Tools einfügbar.'
                : 'Copies all active participant emails (semicolon-separated) to the clipboard. Can be pasted directly into Outlook recipients or external tools.'}
              badge="organizer"
              onClick={() => {
                const emails = registrations
                  .filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt')
                  .map(r => r.ParticipantEmail)
                  .join('; ');
                if (emails) {
                  navigator.clipboard.writeText(emails).then(() => {
                    setCopiedEmails(true);
                    setTimeout(() => setCopiedEmails(false), 2000);
                  }).catch(() => { showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{emails}</span>, { title: isDe ? 'E-Mail-Adressen manuell kopieren' : 'Copy email addresses manually' }); });
                }
              }}
            />

            {/* 4. Massenmail an alle aktiven Teilnehmer */}
            <ActionTile
              icon={<Mail size={18} />}
              category="mails"
              title={isDe ? 'E-Mail versenden an Teilnehmergruppen' : 'Send email to participant groups'}
              desc={isDe
                ? 'Öffnet einen RichText-Editor mit Deloitte-Mail-Template. Geht an alle aktiven Teilnehmer (nicht Wartelistler / Abgemeldete).'
                : 'Opens a rich-text editor with the Deloitte mail template. Goes to all active participants (not waitlisted / cancelled).'}
              badge="organizer"
              onClick={openMassmailPicker}
            />

            {/* v11.40: 4b. Einladungsmail — Mail mit Anmelde-Link an dich
                (zum Weiterleiten an Kollegen / Teams / externe Adressen)
                oder direkt an den auf dem Event hinterlegten Mailverteiler.
                Default-Text + Link werden vorbefüllt, sind aber im RichText-
                Editor frei editierbar. */}
            <ActionTile
              icon={<Send size={18} />}
              category="mails"
              title={isDe ? 'Einladungsmail' : 'Invitation email'}
              desc={isDe
                ? 'Versendet eine Einladungs-Mail mit Anmelde-Link — an dich zum Weiterleiten oder direkt an den hinterlegten Mailverteiler des Events.'
                : 'Sends an invitation email with the registration link — to yourself for forwarding or directly to the configured mail distribution list of the event.'}
              badge="organizer"
              onClick={openInviteModal}
            />

            {/* 5. Excel-Download (mit Dropdown Deloitte/B2Run-View)
                Wrapper braucht display:flex, damit der innere Button auf die
                volle Grid-Zellen-Höhe gestreckt wird — sonst sieht die Kachel
                niedriger aus als ihre Nachbarn, die zwei Zeilen Titel haben. */}
            <div style={{ position: 'relative', display: 'flex' }}>
              <ActionTile
                icon={<Download size={18} />}
                category="participants"
                title={isDe ? 'Excel-Export' : 'Excel export'}
                desc={selectedEvent && selectedEvent.type === 'B2Run'
                  ? (isDe
                    ? "Lädt die Teilnehmerliste als Excel. Wahl zwischen 'Deloitte Felder' (alle internen Spalten + Custom-Fields) oder 'B2Run View' (importierbar in b2run.com)."
                    : "Downloads the participant list as Excel. Choose between 'Deloitte fields' (all internal columns + custom fields) or 'B2Run view' (importable into b2run.com).")
                  : (isDe
                    ? 'Lädt die Teilnehmerliste als Excel mit allen internen Spalten + Custom-Fields des Events.'
                    : 'Downloads the participant list as Excel with all internal columns + custom fields of the event.')}
                badge="organizer"
                onClick={() => {
                  // v17.12: Erst Zielgruppe abfragen, dann erst exportieren.
                  // Bei B2Run zusätzlich noch View-Auswahl im Dropdown.
                  if (selectedEvent && selectedEvent.type === 'B2Run') {
                    setShowExportMenu(!showExportMenu);
                  } else {
                    setExcelAudience('active');
                    setExcelTargetModal({ mode: 'deloitte' });
                  }
                }}
              />
              {showExportMenu && (
                <div style={{
                  position: 'absolute', top: '100%', left: 0, right: 0,
                  background: '#fff', border: '1px solid var(--dex-gray-200)',
                  borderRadius: 'var(--dex-radius, 8px)',
                  boxShadow: '0 8px 24px rgba(0,0,0,0.15)',
                  marginTop: 4, padding: 6, zIndex: 100,
                }}>
                  <button
                    type="button"
                    onClick={() => { setShowExportMenu(false); setExcelAudience('active'); setExcelTargetModal({ mode: 'deloitte' }); }}
                    style={{
                      display: 'block', width: '100%', textAlign: 'left',
                      padding: '10px 12px', border: 'none', background: 'transparent',
                      cursor: 'pointer', borderRadius: 6,
                    }}
                    onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.background = 'var(--dex-gray-50)'; }}
                    onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.background = 'transparent'; }}
                  >
                    <div style={{ fontWeight: 600, fontSize: '0.85rem', color: 'var(--dex-gray-800)' }}>{isDe ? 'Deloitte Felder' : 'Deloitte fields'}</div>
                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)', lineHeight: 1.4, marginTop: 2 }}>
                      {isDe
                        ? 'Alle internen Felder: Name, E-Mail, Department, Standort, Position, Status, Registrierungsdatum + alle Custom-Fields des Events.'
                        : 'All internal fields: name, email, department, location, position, status, registration date + all custom fields of the event.'}
                    </div>
                  </button>
                  {selectedEvent && selectedEvent.type === 'B2Run' && (
                    <button
                      type="button"
                      onClick={() => { setShowExportMenu(false); setExcelAudience('active'); setExcelTargetModal({ mode: 'b2run' }); }}
                      style={{
                        display: 'block', width: '100%', textAlign: 'left',
                        padding: '10px 12px', border: 'none', background: 'transparent',
                        cursor: 'pointer', borderRadius: 6,
                        borderTop: '1px solid var(--dex-gray-100)',
                      }}
                      onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.background = 'var(--dex-gray-50)'; }}
                      onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.background = 'transparent'; }}
                    >
                      <div style={{ fontWeight: 600, fontSize: '0.85rem', color: 'var(--dex-gray-800)' }}>B2Run View</div>
                      <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)', lineHeight: 1.4, marginTop: 2 }}>
                        Spaltenformat exakt wie das B2Run-Excel-Template (Nr, Anrede, Name, E-Mail, Startblock, AGB, Gruppe, Mobilnummer, Altersklasse, …) — direkt importierbar in b2run.com.
                      </div>
                    </button>
                  )}
                </div>
              )}
            </div>

            {/* 6. Outlook-Absagen prüfen — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<AlertCircle size={18} />}
                category="participants"
                title={isCheckingDeclines ? (isDe ? 'Outlook wird geprüft…' : 'Checking Outlook…') : (isDe ? 'Outlook-Absagen prüfen' : 'Check Outlook declines')}
                desc={isDe
                  ? 'Zeigt dir, wer den Outlook-Termin abgesagt hat, aber noch als Teilnehmer angemeldet ist — damit du diese Personen gezielt ansprechen oder abmelden kannst.'
                  : 'Reads the Outlook declines from the no_reply.events mailbox and matches them against active participants. Shows who declined the appointment but is still on the list.'}
                badge="admin"
                busy={isCheckingDeclines}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent) return;
                  setIsCheckingDeclines(true);
                  setDeclineResult(null);
                  setDeclineCopied(false);
                  try {
                    const result = await eventServiceRef.getDeclinedAttendees(selectedEvent.id);
                    if (result.ok) {
                      const activeByEmail = new Map<string, SPRegistration>();
                      for (const r of registrations) {
                        if (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt') {
                          activeByEmail.set(String(r.ParticipantEmail || '').toLowerCase(), r);
                        }
                      }
                      const hits: Array<{ email: string; name: string; reg: SPRegistration }> = [];
                      for (const d of result.attendees) {
                        const reg = activeByEmail.get(d.email);
                        if (reg) hits.push({ email: d.email, name: d.name, reg });
                      }
                      setDeclineResult({
                        declinedAndRegistered: hits,
                        declinedTotal: result.attendees.length,
                        error: null,
                      });
                      setShowDeclineModal(true);
                    } else {
                      let msg = result.message || 'Unbekannter Fehler beim Lesen des Outlook-Termins.';
                      if (!result.message) {
                        if (result.reason === 'no-pointer') {
                          msg = 'Für dieses Event ist kein Outlook-Termin verknüpft (OutlookEventId / CalendarLink fehlen).';
                        } else if (result.reason === 'not-found') {
                          msg = 'Outlook-Termin wurde im Postfach no_reply.events@deloitte.de nicht gefunden.';
                        } else if (result.reason === 'forbidden') {
                          msg = 'Graph-API-Zugriff abgelehnt (HTTP 403). Tenant-Admin muss "Calendars.Read.Shared" genehmigen, und der User braucht Reviewer-Rechte auf dem Postfach-Kalender.';
                        }
                      }
                      setDeclineResult({ declinedAndRegistered: [], declinedTotal: 0, error: msg });
                      setShowDeclineModal(true);
                    }
                  } catch (err) {
                    setDeclineResult({
                      declinedAndRegistered: [],
                      declinedTotal: 0,
                      error: err instanceof Error ? err.message : String(err),
                    });
                    setShowDeclineModal(true);
                  }
                  setIsCheckingDeclines(false);
                }}
              />
            )}

            {/* 7. TeilnehmerIDs neu vergeben — Admin ODER Organizer des Events (v11.36) */}
            {(isAdmin || (!!selectedEvent && isOrganizerFor(selectedEvent))) && (
              <ActionTile
                icon={<Hash size={18} />}
                category="participants"
                title={isReorderingIDs ? (isDe ? 'IDs werden vergeben…' : 'Assigning IDs…') : (isDe ? 'IDs neu vergeben' : 'Reassign IDs')}
                desc={isDe
                  ? 'Vergibt die TeilnehmerIDs sequentiell (1, 2, 3, …) nach Erstellungsreihenfolge. Schließt Lücken nach Stornos und sortiert die Liste sauber durch. Hinweis: nicht ausführen während gerade viele Anmeldungen laufen — erst wenn die Anmeldewelle vorbei ist.'
                  : 'Assigns the participant IDs sequentially (1, 2, 3, …) by creation order. Closes gaps after cancellations and sorts the list cleanly. Note: do not run while many registrations are coming in — wait until the registration wave is over.'}
                badge="organizer"
                busy={isReorderingIDs}
                disabled={!selectedEvent?.subsiteUrl}
                result={reorderResult}
                resultIsError={!!reorderResult && (reorderResult.indexOf('Fehler') >= 0 || reorderResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!(await confirmDialog(isDe
                    ? 'TeilnehmerIDs neu vergeben (1, 2, 3, …)? Sortierung nach Erstellungsreihenfolge.\n\nNICHT ausführen, während gerade viele Anmeldungen laufen — bitte erst wenn die Anmeldewelle vorbei ist.'
                    : 'Reassign participant IDs (1, 2, 3, …)? Sorted by creation order.\n\nDo NOT run while many registrations are coming in — please wait until the registration wave is over.'))) return;
                  await runIdReorder();
                }}
              />
            )}

            {/* 7a2. Nachrücken — Admin ODER Organizer des Events (v18.70).
                Füllt einen freien Platz mit dem ersten Wartelistler (nach
                TeilnehmerID), inkl. Nachrück-Mail + Outlook-Einladung, danach
                IDs neu vergeben + Seat-Sync. */}
            {(isAdmin || (!!selectedEvent && isOrganizerFor(selectedEvent))) && (
              <ActionTile
                icon={<Users size={18} />}
                category="participants"
                title={isPromoting ? (isDe ? 'Rückt nach…' : 'Promoting…') : (isDe ? 'Von Warteliste nachrücken' : 'Promote from waitlist')}
                desc={isDe
                  ? 'Rückt den ersten Teilnehmer von der Warteliste (nach TeilnehmerID) auf einen freien Platz nach. Die Person bekommt Status „Angemeldet", eine Nachrück-Mail und eine Outlook-Einladung; danach werden die IDs neu vergeben. Promotet nur, wenn tatsächlich ein Platz frei ist.'
                  : 'Promotes the first person on the waitlist (by participant ID) into a free seat. The person gets status “Registered”, a promotion email and an Outlook invite; IDs are then reassigned. Only promotes if a seat is actually free.'}
                badge="organizer"
                busy={isPromoting}
                disabled={!selectedEvent?.subsiteUrl || isPromoting || isReorderingIDs}
                result={promoteResult}
                resultIsError={!!promoteResult && (promoteResult.indexOf('Fehler') >= 0 || promoteResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!(await confirmDialog(isDe
                    ? 'Den ersten Teilnehmer von der Warteliste auf einen freien Platz nachrücken?'
                    : 'Promote the first waitlist participant into a free seat?', { confirmLabel: isDe ? 'Nachrücken' : 'Promote' }))) return;
                  await runManualPromote();
                }}
              />
            )}

            {/* 7b. Counter zurücksetzen — Admin only (v9.13 → v11.27).
                Recovery-Button um den DEX_TeilnehmerCounter EXAKT auf
                max(TeilnehmerID) der Subsite zu setzen. Bidirektional:
                Counter wird hochgezogen wenn er drunter steht (gegen
                Doppel-IDs), oder runtergesetzt wenn er drüber steht
                (z.B. nach vielen Abmeldungen, die TIDs gefressen
                haben). Vorher (vor v11.27) lief es nur monotonic-up,
                weshalb ein zu hoher Counter (Counter=11, Max-TID=4)
                nicht zurückgesetzt wurde — Klick auf den Button
                hatte dann keinen sichtbaren Effekt. */}
            {isAdmin && (
              <ActionTile
                icon={<Hash size={18} />}
                category="maintenance"
                title={isResettingCounter ? (isDe ? 'Counter wird zurückgesetzt…' : 'Resetting counter…') : (isDe ? 'Counter zurücksetzen' : 'Reset counter')}
                desc={isDe
                  ? 'Repariert die automatische Nummern-Vergabe: Neue Anmeldungen bekommen danach wieder die nächste passende Teilnehmer-Nummer. Nutzen, wenn neue Anmeldungen mit offensichtlich falschen Nummern starten (viel zu hoch oder wieder bei 1).'
                  : 'Sets the participant ID counter exactly to the current max ID of the participant list. Helps when new registrations start with IDs that are too high (gaps from earlier cancellations) or when they would accidentally start at IDs that are too low (e.g. back at 1). Bidirectional — regardless of whether the counter is too high or too low.'}
                badge="admin"
                busy={isResettingCounter}
                disabled={!selectedEvent?.subsiteUrl}
                result={resetCounterResult}
                resultIsError={!!resetCounterResult && (resetCounterResult.indexOf('Fehler') >= 0 || resetCounterResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!(await confirmDialog(isDe ? 'Counter auf aktuellen Max-Wert zurücksetzen?' : 'Reset counter to the current max value?'))) return;
                  setIsResettingCounter(true);
                  setResetCounterResult(null);
                  try {
                    const result = await eventServiceRef.resetCounterToMax(selectedEvent.subsiteUrl);
                    setResetCounterResult(isDe
                      ? `Counter steht jetzt auf ${result.counter} (Max-TID: ${result.max})`
                      : `Counter is now at ${result.counter} (max ID: ${result.max})`);
                  } catch {
                    setResetCounterResult(isDe ? 'Fehler beim Zurücksetzen des Counters' : 'Error resetting the counter');
                  }
                  setIsResettingCounter(false);
                }}
              />
            )}

            {/* 7c. Überbuchung prüfen — Admin ODER Organizer des Events (v11.36).
                Markiert pro Gruppe (bzw. gesamt) die zuletzt über Kapazität
                Angemeldeten mit OverbookReview='Pending'. Ändert KEINEN
                Status — Admin/Organizer entscheidet danach pro Person über
                die Buttons in der „Überbuchung – zu prüfen"-Box oben in der
                Teilnehmerliste. Organizer dürfen das für eigene Events, weil
                es Teilnehmerverwaltung ist (analog Abmelden/QR/Massenmail). */}
            {(isAdmin || (!!selectedEvent && isOrganizerFor(selectedEvent))) && (
              <ActionTile
                icon={<Users size={18} />}
                category="participants"
                title={isDetectingOverbook ? (isDe ? 'Wird geprüft…' : 'Checking…') : (isDe ? 'Überbuchung prüfen' : 'Check overbooking')}
                desc={isDe
                  ? 'Findet pro Gruppe (Durchstarter/Funstarter, bzw. gesamt) die zuletzt angemeldeten Personen ÜBER der Kapazität und markiert sie zur Prüfung. Es wird nichts automatisch geändert — danach entscheidest du pro Person (auf Warteliste / Platz behalten) über die Buttons oben in der Teilnehmerliste.'
                  : 'Finds, per group (Durchstarter/Funstarter, or overall), the most recently registered people OVER capacity and marks them for review. Nothing is changed automatically — afterwards you decide per person (move to waitlist / keep seat) via the buttons at the top of the participant list.'}
                badge="organizer"
                busy={isDetectingOverbook}
                disabled={!selectedEvent?.subsiteUrl}
                result={detectOverbookResult}
                resultIsError={!!detectOverbookResult && (detectOverbookResult.indexOf('Fehler') >= 0 || detectOverbookResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!(await confirmDialog(isDe ? 'Überbuchung prüfen und betroffene Personen markieren? (ändert keinen Status)' : 'Check overbooking and mark affected people? (does not change any status)', { confirmLabel: isDe ? 'Prüfen' : 'Check' }))) return;
                  setIsDetectingOverbook(true);
                  setDetectOverbookResult(null);
                  try {
                    const res = await eventServiceRef.detectOverbooking(selectedEvent.subsiteUrl, {
                      isSplit: isSplitCapacity,
                      maxParticipants: selectedEvent.maxParticipants || 0,
                      durchstarterCapacity: selectedEvent.durchstarterCapacity || 0,
                      funstarterCapacity: selectedEvent.funstarterCapacity || 0,
                    });
                    // Counter mit echtem Bestand abgleichen (best-effort).
                    try { await eventServiceRef.syncSeatsToActiveCount(selectedEvent.subsiteUrl, { isSplit: isSplitCapacity }); } catch { /* */ }
                    const parts = res.groups
                      .map(g => `${g.group}: ${g.activeBefore}/${g.cap || '∞'} → ${g.marked} ${isDe ? 'markiert' : 'marked'}`)
                      .join(' · ');
                    setDetectOverbookResult(res.total > 0
                      ? (isDe
                        ? `${res.total} markiert (${parts})${res.errors ? ` — ${res.errors} Fehler` : ''}`
                        : `${res.total} marked (${parts})${res.errors ? ` — ${res.errors} errors` : ''}`)
                      : (isDe ? `Keine Überbuchung gefunden (${parts})` : `No overbooking found (${parts})`));
                    const regs = await getAllRegistrations(selectedEvent.id);
                    setRegistrations(regs);
                  } catch {
                    setDetectOverbookResult(isDe ? 'Fehler beim Prüfen der Überbuchung' : 'Error checking overbooking');
                  }
                  setIsDetectingOverbook(false);
                }}
              />
            )}

            {/* 8. Spalten fixen — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<Columns size={18} />}
                category="maintenance"
                title={isFixingColumns ? (isDe ? 'Spalten werden gefixt…' : 'Fixing columns…') : (isDe ? 'Spalten fixen' : 'Fix columns')}
                desc={isDe
                  ? 'Bringt die Teilnehmerliste auf den aktuellen Stand: legt fehlende Spalten an, räumt überflüssige weg und sortiert die Spalten-Reihenfolge richtig. Nutzen, wenn in der Liste Spalten fehlen oder Antworten nicht ankommen.'
                  : 'Creates missing columns in the participant list, removes superfluous ones (e.g. StarterType for non-B2Run events) and fixes the default view order.'}
                badge="admin"
                busy={isFixingColumns}
                disabled={!selectedEvent?.subsiteUrl}
                result={fixColumnsResult}
                resultIsError={!!fixColumnsResult && (fixColumnsResult.indexOf('Fehler') >= 0 || fixColumnsResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  setIsFixingColumns(true);
                  setFixColumnsResult(null);
                  try {
                    const isB2Run = !!(selectedEvent.durchstarterCapacity || selectedEvent.funstarterCapacity);
                    const hasQuiz = !!(selectedEvent.quiz && selectedEvent.quiz.length > 0);
                    const customFields = (selectedEvent.eventSpecificFields || []).map(f => ({
                      id: f.id, label: f.label, type: f.type, required: f.required, options: f.options,
                      visible: true,
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      spInternalName: (f as any).spInternalName || '',
                    }));
                    const result = await eventServiceRef.fixRegistrationListColumns(
                      selectedEvent.subsiteUrl,
                      { isB2Run, hasQuiz, customFields },
                      (count, titles) => {
                        const preview = titles.slice(0, 8).map(t => `„${t}"`).join(', ');
                        const more = titles.length > 8 ? (isDe ? ` …und ${titles.length - 8} weitere` : ` …and ${titles.length - 8} more`) : '';
                        // v20.4: App-Modal statt window.confirm — der Service
                        // akzeptiert boolean | Promise<boolean> und awaitet.
                        return confirmDialog(isDe
                          ? `${count} überflüssige (leere) Duplikat-Spalten in der Teilnehmerliste gefunden ` +
                            `(${titles.length} Titel betroffen: ${preview}${more}).\n\n` +
                            `Diese werden jetzt gelöscht (irreversibel). Spalten mit Daten bleiben erhalten ` +
                            `und werden zur manuellen Prüfung gemeldet.\n\nFortfahren?`
                          : `Found ${count} redundant (empty) duplicate columns in the participant list ` +
                            `(${titles.length} titles affected: ${preview}${more}).\n\n` +
                            `These will now be deleted (irreversible). Columns with data are kept ` +
                            `and reported for manual review.\n\nProceed?`
                        );
                      }
                    );
                    const msgs: string[] = [];
                    if (result.added.length > 0) msgs.push(isDe ? `Spalten hinzugefügt: ${result.added.join(', ')}` : `Columns added: ${result.added.join(', ')}`);
                    if (result.removed.length > 0) msgs.push(isDe ? `Spalten entfernt: ${result.removed.join(', ')}` : `Columns removed: ${result.removed.join(', ')}`);
                    if (result.duplicatesRemoved && result.duplicatesRemoved.length > 0) {
                      msgs.push(isDe ? `${result.duplicatesRemoved.length} leere Duplikate gelöscht` : `${result.duplicatesRemoved.length} empty duplicates deleted`);
                    }
                    if (result.duplicatesWithData && result.duplicatesWithData.length > 0) {
                      const list = result.duplicatesWithData.map(t => `„${t}"`).join(', ');
                      msgs.push(isDe ? `${result.duplicatesWithData.length} Duplikate mit Daten — bitte manuell prüfen: ${list}` : `${result.duplicatesWithData.length} duplicates with data — please review manually: ${list}`);
                    }
                    if (result.viewFixed) msgs.push(isDe ? 'View-Reihenfolge korrigiert' : 'View order fixed');
                    if (result.customFieldMap && Object.keys(result.customFieldMap).length > 0) {
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const updatedCf = (customFields as any[]).map(f => {
                        const sp = result.customFieldMap![f.id];
                        return sp ? { ...f, spInternalName: sp } : f;
                      });
                      try {
                        await updateEvent(selectedEvent.id, { 'CustomFields': JSON.stringify(updatedCf) });
                        msgs.push(isDe ? `Custom-Field-Zuordnung aktualisiert (${Object.keys(result.customFieldMap).length})` : `Custom field mapping updated (${Object.keys(result.customFieldMap).length})`);
                      } catch {
                        msgs.push(isDe ? 'WARN: Custom-Field-Mapping konnte nicht am Event gespeichert werden' : 'WARN: custom field mapping could not be saved on the event');
                      }
                    }
                    const finalMsg = msgs.length > 0 ? msgs.join(' | ') : (isDe ? 'Alles OK, keine Änderungen nötig.' : 'All OK, no changes needed.');
                    setFixColumnsResult(finalMsg);
                    // v19.10: Ergebnis zusätzlich als Dialog zeigen. Im Aktionen-
                    // Dropdown rendert die ActionTile (und damit ihr `result`-Text)
                    // `null` — vorher kam nach „Spalten fixen" daher GAR KEINE
                    // sichtbare Rückmeldung. Ein window.alert ist garantiert sichtbar.
                    showAlert((isDe ? '„Spalten fixen" — Ergebnis:\n\n' : 'Fix columns — result:\n\n') + finalMsg);
                  } catch {
                    const errMsg = isDe ? 'Fehler beim Fixen der Spalten.' : 'Error fixing columns.';
                    setFixColumnsResult(errMsg);
                    showAlert(errMsg);
                  }
                  setIsFixingColumns(false);
                }}
              />
            )}

            {/* v20.6: Fremd-Anmeldungen: Zugriff reparieren (alle aktiven
                Events) — Admin only. Geht alle Teilnehmerlisten durch, stellt
                die "nur eigene Elemente"-Sicherheit sicher und setzt bei
                Anmeldungen durch Dritte den Zeilen-Autor auf den Teilnehmer,
                damit die angemeldete Person ihre Anmeldung in "Meine Events"
                sieht und sich selbst abmelden kann (v20.5 rückwirkend). */}
            {isAdmin && (
              <ActionTile
                icon={<Wrench size={18} />}
                category="maintenance"
                title={isRepairingAccess ? (isDe ? 'Prüfung läuft…' : 'Check running…') : (isDe ? 'Fremd-Anmeldungen: Zugriff reparieren (alle aktiven Events)' : 'Proxy registrations: repair access (all active events)')}
                desc={isDe
                  ? 'Geht alle Teilnehmerlisten aller aktiven Events (inkl. Sub-Events) durch und prüft zwei Dinge: (1) dass jede Liste auf „nur eigene Elemente" steht — also niemand fremde Anmeldedaten sehen kann; (2) dass bei Anmeldungen, die jemand FÜR eine andere Person gemacht hat, die angemeldete Person ihre eigene Anmeldung sehen und sich selbst abmelden kann. Gefundene Probleme werden direkt repariert. Externe Personen ohne Deloitte-Login können dabei nicht berücksichtigt werden.'
                  : 'Walks the participant lists of all active events (incl. sub-events) and checks two things: (1) that every list is set to „own items only" — so nobody can see other people’s registration data; (2) that for registrations someone made FOR another person, the registered person can see their own registration and cancel it themselves. Found issues are repaired directly. External people without a Deloitte login cannot be covered.'}
                badge="admin"
                busy={isRepairingAccess}
                result={repairAccessResult}
                resultIsError={!!repairAccessResult && (repairAccessResult.indexOf('Fehler') >= 0 || repairAccessResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef) return;
                  // Alle aktiven Events (inkl. Sub-Events) mit Teilnehmerliste,
                  // dedupliziert nach Subsite (Reuse-Pfade teilen sich eine).
                  const seen = new Set<string>();
                  const targets = allEvents.filter(ev => {
                    if (ev.status !== 'Active' || !ev.subsiteUrl) return false;
                    const key = ev.subsiteUrl.toLowerCase();
                    if (seen.has(key)) return false;
                    seen.add(key);
                    return true;
                  });
                  if (targets.length === 0) {
                    setRepairAccessResult(isDe ? 'Keine aktiven Events mit Teilnehmerliste gefunden.' : 'No active events with a participant list found.');
                    return;
                  }
                  if (!(await confirmDialog(isDe
                    ? `Zugriffs-Prüfung über ${targets.length} aktive Event-Teilnehmerlisten starten?\n\nGeprüft wird pro Liste die „nur eigene Elemente"-Sicherheit, und bei Anmeldungen durch Dritte bekommt die angemeldete Person Zugriff auf ihre eigene Anmeldung. Je nach Teilnehmerzahl kann das einige Minuten dauern.`
                    : `Start the access check across ${targets.length} active event participant lists?\n\nEach list is checked for „own items only" security, and for registrations made by a third party the registered person gets access to their own registration. Depending on participant counts this can take a few minutes.`, { confirmLabel: isDe ? 'Prüfen & reparieren' : 'Check & repair' }))) return;
                  setIsRepairingAccess(true);
                  setRepairAccessResult(null);
                  // v20.7: Fortschritts-Modal öffnen.
                  setAccessFixModal({ running: true, evIdx: 0, evTotal: targets.length, evTitle: '', itemDone: 0, itemTotal: 0, summary: null });
                  // v22: Globale Queue-Listen (DEX_Outlook/DEX_IDReorder) auf
                  // „nur eigene Elemente" härten — einmal pro Lauf, idempotent.
                  let queueIlsLine = '';
                  try {
                    const q = await eventServiceRef.hardenQueueListsIls();
                    queueIlsLine = q.failed.length === 0
                      ? (isDe ? 'Globale Queue-Listen (Outlook/IDReorder) stehen auf „nur eigene Elemente".' : 'Global queue lists (Outlook/IDReorder) set to „own items only".')
                      : (isDe ? `Queue-Listen-Härtung fehlgeschlagen bei: ${q.failed.join(', ')}.` : `Queue list hardening failed for: ${q.failed.join(', ')}.`);
                  } catch { /* best-effort */ }
                  let listsChecked = 0;
                  let ilsWrong = 0;
                  let ilsFixed = 0;
                  let proxyFound = 0;
                  let authorFixed = 0;
                  let authorFailed = 0;
                  const errorEvents: string[] = [];
                  try {
                    for (let i = 0; i < targets.length; i++) {
                      const ev = targets[i];
                      const shortTitle = (ev.title || '?').slice(0, 40);
                      setAccessFixModal(prev => prev ? { ...prev, evIdx: i + 1, evTitle: shortTitle, itemDone: 0, itemTotal: 0 } : prev);
                      try {
                        const r = await eventServiceRef.repairProxyRegistrationAccess(ev.subsiteUrl as string, (done, total) => {
                          if (done % 10 === 0 || done === total) {
                            setAccessFixModal(prev => prev ? { ...prev, itemDone: done, itemTotal: total } : prev);
                          }
                        });
                        listsChecked++;
                        if (r.ilsWasWrong) {
                          ilsWrong++;
                          if (r.ilsFixed) ilsFixed++;
                        }
                        proxyFound += r.proxyFound;
                        authorFixed += r.authorFixed;
                        authorFailed += r.authorFailed;
                      } catch {
                        errorEvents.push(shortTitle);
                      }
                    }
                    const parts: string[] = [];
                    if (isDe) {
                      parts.push(`${listsChecked} Listen geprüft.`);
                      parts.push(ilsWrong === 0
                        ? 'Listen-Sicherheit überall korrekt („nur eigene Elemente").'
                        : `Listen-Sicherheit bei ${ilsWrong} Liste(n) falsch — ${ilsFixed} davon repariert${ilsFixed < ilsWrong ? ', Rest bitte manuell prüfen!' : '.'}`);
                      parts.push(`${proxyFound} Anmeldungen durch Dritte gefunden, ${authorFixed} Zugriffe repariert${authorFailed > 0 ? `, ${authorFailed} nicht möglich (z.B. externe Personen)` : ''}.`);
                      if (errorEvents.length > 0) parts.push(`Fehler bei: ${errorEvents.join(', ')}`);
                    } else {
                      parts.push(`${listsChecked} lists checked.`);
                      parts.push(ilsWrong === 0
                        ? 'List security correct everywhere („own items only").'
                        : `List security wrong on ${ilsWrong} list(s) — ${ilsFixed} repaired${ilsFixed < ilsWrong ? ', please check the rest manually!' : '.'}`);
                      parts.push(`${proxyFound} third-party registrations found, ${authorFixed} access repaired${authorFailed > 0 ? `, ${authorFailed} not possible (e.g. external people)` : ''}.`);
                      if (errorEvents.length > 0) parts.push(`Errors on: ${errorEvents.join(', ')}`);
                    }
                    if (queueIlsLine) parts.push(queueIlsLine);
                    setRepairAccessResult(parts.join(' '));
                    // v20.7: Summary im Fortschritts-Modal anzeigen.
                    setAccessFixModal(prev => prev ? { ...prev, running: false, summary: parts } : prev);
                  } catch {
                    const err = isDe ? 'Fehler bei der Zugriffs-Prüfung.' : 'Error during the access check.';
                    setRepairAccessResult(err);
                    setAccessFixModal(prev => prev ? { ...prev, running: false, summary: [err] } : prev);
                  } finally {
                    setIsRepairingAccess(false);
                  }
                }}
              />
            )}

            {/* 8b. Organizer-Mails reparieren (alle Events) — Admin only.
                Findet Events mit Längen-Mismatch zwischen organizers (Names) und
                organizerEmails — typisch nach Legacy-Korruption aus v10.0–v10.2-
                Closure-Bug. Versucht via Graph-Search die fehlenden Emails per
                Lastname-Match nachzufüllen. Persistiert das gefixte Pair-Mapping
                via updateEvent. Bricht NICHT bei einzelnen Fehlern ab — ein Event
                mit unauflösbarem Namen wird übersprungen, der Rest läuft weiter.
                Operiert über ALLE adminEvents (nicht nur das gerade ausgewählte). */}
            {isAdmin && (
              <ActionTile
                icon={<Wrench size={18} />}
                category="maintenance"
                title={isRepairingOrganizers ? (isDe ? 'Reparatur läuft…' : 'Repair running…') : (isDe ? 'Organizer-Mails reparieren (alle Events)' : 'Repair organizer emails (all events)')}
                desc={isDe
                  ? 'Scannt alle Events nach Mismatches zwischen Organizer-Namen und Organizer-Emails (Legacy-Korruption aus früheren App-Versionen). Sucht fehlende Emails per Tenant-Suche über den Nachnamen und persistiert die gefixten Paare. Manuell nicht auflösbare Personen bleiben mit leerem Email-Slot — User muss diese im Wizard nachziehen.'
                  : 'Scans all events for mismatches between organizer names and organizer emails (legacy corruption from earlier app versions). Looks up missing emails via tenant search by last name and persists the fixed pairs. People that cannot be resolved automatically keep an empty email slot — the user must add them in the wizard.'}
                badge="admin"
                busy={isRepairingOrganizers}
                result={repairOrganizersResult}
                resultIsError={!!repairOrganizersResult && (repairOrganizersResult.indexOf('Fehler') >= 0 || repairOrganizersResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef) return;
                  if (!(await confirmDialog(isDe
                    ? `Organizer-Mails über ALLE ${adminEvents.length} Events reparieren? Dauert je nach Anzahl ca. 1–2 Minuten.`
                    : `Repair organizer emails across ALL ${adminEvents.length} events? Depending on the count this takes about 1–2 minutes.`, { confirmLabel: isDe ? 'Reparieren' : 'Repair' }))) return;
                  setIsRepairingOrganizers(true);
                  setRepairOrganizersResult(null);
                  let scanned = 0;
                  let mismatched = 0;
                  let eventsUpdated = 0;
                  let orgsRecovered = 0;
                  let orgsUnresolved = 0;
                  const unresolvedNames: string[] = [];
                  try {
                    for (const ev of adminEvents) {
                      scanned++;
                      const names = (ev.organizers || []).slice();
                      const emails = (ev.organizerEmails || []).slice();
                      // Pad das kürzere Array auf max() — nichts geht verloren
                      const max = Math.max(names.length, emails.length);
                      while (names.length < max) names.push('');
                      while (emails.length < max) emails.push('');
                      // Mismatch erkannt? Mindestens ein Name ohne Email oder eine Email ohne Name
                      const hasMismatch = names.some((n, i) => (n || '').trim() && !((emails[i] || '').trim()))
                        || emails.some((e, i) => (e || '').trim() && !((names[i] || '').trim()));
                      if (!hasMismatch) continue;
                      mismatched++;
                      // Für jeden Slot mit Name aber ohne Email: Graph-Search nach Lastname
                      // pro Slot, EINS-zu-EINS-Match wenn Local-Part den Lastname enthält.
                      let recoveredHere = 0;
                      const fixedNames = names.slice();
                      const fixedEmails = emails.slice();
                      for (let i = 0; i < max; i++) {
                        const name = (fixedNames[i] || '').trim();
                        const email = (fixedEmails[i] || '').trim();
                        if (!name || email) continue;
                        // Lastname extrahieren — egal ob "Lastname, Firstname" oder
                        // "Firstname Lastname", als Suchquery für Graph nehmen wir
                        // den ganzen Namen (Graph ist tolerant).
                        try {
                          // Lastname als Suchterm — Graph-Search ist tolerant für
                          // 'Lastname' als Query und liefert eindeutigere Ergebnisse
                          // als die kombinierte Form 'Lastname, Firstname'.
                          const queryRaw = name.indexOf(',') >= 0 ? name.split(',')[0].trim() : name;
                          const hits = await searchUsers(queryRaw);
                          // Lastname-Substring-Match: filtere die Hits auf Personen,
                          // deren Email-Local-Part den Lastname enthält. Damit greifen
                          // wir den richtigen Eintrag auch bei Häufigkeitsnamen.
                          const lastname = queryRaw.toLowerCase().split(/\s+/).filter(t => t.length >= 3).pop() || '';
                          const matched = lastname
                            ? hits.filter(h => ((h.email || '').toLowerCase().split('@')[0]).indexOf(lastname) >= 0)
                            : hits;
                          if (matched.length === 1 && matched[0].email) {
                            fixedEmails[i] = matched[0].email;
                            recoveredHere++;
                            orgsRecovered++;
                          } else if (matched.length === 0 && hits.length === 1 && hits[0].email) {
                            // Kein Lastname-Match aber genau 1 Treffer überhaupt — übernehmen
                            fixedEmails[i] = hits[0].email;
                            recoveredHere++;
                            orgsRecovered++;
                          } else {
                            // Mehrdeutig oder nichts gefunden — leer lassen
                            unresolvedNames.push(`${name} (Event ${ev.eventNumber})`);
                            orgsUnresolved++;
                          }
                        } catch {
                          unresolvedNames.push(`${name} (Event ${ev.eventNumber})`);
                          orgsUnresolved++;
                        }
                      }
                      // Nichts wiederhergestellt? Skip Update — Storage ist eh schon
                      // im aktuellen Zustand, Pad allein bringt keinen Mehrwert
                      // (bei Save aus Wizard heilt sich das ohnehin).
                      if (recoveredHere === 0) continue;
                      // Alle vollständig leeren Slots aussortieren bevor wir schreiben
                      const finalPairs = fixedNames.map((n, i) => ({ n: (n || '').trim(), e: (fixedEmails[i] || '').trim() }))
                        .filter(p => p.n || p.e);
                      const finalNames = finalPairs.map(p => p.n).join('; ');
                      const finalEmails = finalPairs.map(p => p.e).join(';');
                      try {
                        const ok = await updateEvent(ev.id, { 'Organizer': finalNames, 'OrganizerEmail': finalEmails });
                        if (ok) eventsUpdated++;
                      } catch {
                        // Update fehlgeschlagen — counts trotzdem belassen, einfach
                        // skip dieses Event.
                      }
                    }
                    const lines = isDe
                      ? [`Gescannt: ${scanned}`, `Mit Mismatch: ${mismatched}`, `Aktualisiert: ${eventsUpdated}`, `Emails wiederhergestellt: ${orgsRecovered}`]
                      : [`Scanned: ${scanned}`, `With mismatch: ${mismatched}`, `Updated: ${eventsUpdated}`, `Emails recovered: ${orgsRecovered}`];
                    if (orgsUnresolved > 0) {
                      lines.push(isDe
                        ? `Manuell nachziehen (${orgsUnresolved}): ${unresolvedNames.slice(0, 5).join(', ')}${unresolvedNames.length > 5 ? '…' : ''}`
                        : `Add manually (${orgsUnresolved}): ${unresolvedNames.slice(0, 5).join(', ')}${unresolvedNames.length > 5 ? '…' : ''}`);
                    }
                    setRepairOrganizersResult(lines.join(' · '));
                  } catch (err) {
                    setRepairOrganizersResult(isDe ? `Fehler: ${err instanceof Error ? err.message : String(err)}` : `Error: ${err instanceof Error ? err.message : String(err)}`);
                  }
                  setIsRepairingOrganizers(false);
                }}
              />
            )}

            {/* v11.9: B2Run-Migration als Action-Tile im Admin-Event-Detail.
                Erkennt Legacy-Events (type='B2Run' oder b2run_*-Custom-
                Fields vorhanden) und bietet die gleiche Migration an wie
                der „B2Run migrieren"-Button in der Event-Liste. Damit
                findet der Admin den Knopf auch wenn er das Event bereits
                ausgewählt hat. */}
            {isAdmin && selectedEvent && (selectedEvent.type === 'B2Run' || (selectedEvent.eventSpecificFields || []).some(f => (f.id || '').toLowerCase().startsWith('b2run_'))) && (
              <ActionTile
                icon={<RefreshCw size={18} />}
                category="maintenance"
                title={isDe ? 'Legacy-B2Run migrieren' : 'Migrate legacy B2Run'}
                desc={isDe
                  ? "Stellt ein altes B2Run-Event auf das normale Eventschema um. Die Gruppen heißen danach 'Durchstarter' / 'Funstarter' (im Wizard frei umbenennbar), alle Anmeldefelder, Anmeldungen, Wartelisten und Sub-Events bleiben erhalten."
                  : "Removes the B2Run type and persists 'Durchstarter' / 'Funstarter' as regular group labels (you can rename them freely in the wizard afterwards). b2run_* custom fields (age group, t-shirt size etc.) are KEPT as generic custom fields. Registrations, waitlists and sub-events remain unchanged."}
                badge="admin"
                onClick={async () => {
                  if (!eventServiceRef) return;
                  const kids = childEventsOf(selectedEvent.id);
                  const kidsToMigrate = kids.filter(k => k.type === 'B2Run' || (k.eventSpecificFields || []).some(f => (f.id || '').toLowerCase().startsWith('b2run_')));
                  const msg = isDe
                    ? `Event "${selectedEvent.title}" auf Standard-Schema migrieren?\n\n` +
                      `• B2Run-Type wird entfernt — Event sieht aus wie ein normales Deloitte-Event.\n` +
                      `• Bezeichnungen "Durchstarter" / "Funstarter" werden als Gruppen-Labels gespeichert (frei umbenennbar im Wizard).\n` +
                      `• Falls Leistungsnachweis-Pflicht aktiv war: wird in ein reguläres Custom-Field „Leistungsnachweis vorhanden" (Checkbox, Pflicht, nur für Gruppe A) umgewandelt.\n` +
                      `• Hardcoded Startblock-Mapping pro Gruppe wird ersatzlos entfernt.\n` +
                      `• b2run_*-Custom-Fields (Altersgruppe, T-Shirt-Größe, Mobilnummer etc.) BLEIBEN als generische Custom-Fields erhalten.\n` +
                      `• Anmeldungen, Wartelisten und Sub-Events bleiben inhaltlich unverändert.\n\n` +
                      (kidsToMigrate.length > 0
                        ? `Es werden zusätzlich ${kidsToMigrate.length} Sub-Event(s) mitmigriert: ${kidsToMigrate.map(k => '„' + (k.title || '?') + '"').join(', ')}.`
                        : `Keine Sub-Events mit Legacy-B2Run-Spuren gefunden — nur das Hauptevent wird migriert.`)
                    : `Migrate event "${selectedEvent.title}" to the standard schema?\n\n` +
                      `• The B2Run type is removed — the event will look like a normal Deloitte event.\n` +
                      `• Labels "Durchstarter" / "Funstarter" are stored as group labels (freely renamable in the wizard).\n` +
                      `• If a performance-proof requirement was active: it is converted into a regular custom field „Leistungsnachweis vorhanden" (checkbox, required, only for group A).\n` +
                      `• The hardcoded per-group start-block mapping is removed.\n` +
                      `• b2run_* custom fields (age group, t-shirt size, mobile etc.) are KEPT as generic custom fields.\n` +
                      `• Registrations, waitlists and sub-events stay unchanged content-wise.\n\n` +
                      (kidsToMigrate.length > 0
                        ? `Additionally, ${kidsToMigrate.length} sub-event(s) will be migrated: ${kidsToMigrate.map(k => '„' + (k.title || '?') + '"').join(', ')}.`
                        : `No sub-events with legacy B2Run traces found — only the main event will be migrated.`);
                  if (!(await confirmDialog(msg, { title: isDe ? 'B2Run migrieren' : 'Migrate B2Run', confirmLabel: isDe ? 'Migrieren' : 'Migrate' }))) return;
                  const errors: string[] = [];
                  const migrateOne = async (ev: DeloitteEvent): Promise<void> => {
                    try {
                      // v11.11: Custom-Fields werden NICHT mehr gelöscht.
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const keptFields: any[] = (ev.eventSpecificFields || []).map(f => ({ ...f }));
                      const baseUpdates: Record<string, unknown> = {
                        'SplitLabelA': (ev.splitLabelA || 'Durchstarter'),
                        'SplitLabelB': (ev.splitLabelB || 'Funstarter'),
                      };
                      // v11.13: B2Run-Extras aus EmailTemplateOverrides._b2run
                      // in echte Custom-Fields mit onlyForGroup übersetzen
                      // (siehe ausführlicher Kommentar im Card-Button-Pfad).
                      try {
                        const overridesRaw = (ev.emailTemplateOverrides || '').toString();
                        if (overridesRaw.trim()) {
                          const parsed = JSON.parse(overridesRaw);
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const b2 = parsed && typeof parsed === 'object' ? (parsed as any)._b2run : null;
                          if (b2 && typeof b2 === 'object') {
                            if (b2.durchstarterRequiresProof) {
                              const PROOF_ID = 'b2run_leistungsnachweis';
                              const existing = keptFields.find(f => String(f.id || '').toLowerCase() === PROOF_ID);
                              if (existing) {
                                existing.onlyForGroup = 'A';
                                existing.required = true;
                                if (!existing.label) existing.label = 'Leistungsnachweis vorhanden';
                                if (!existing.type) existing.type = 'checkbox';
                                if (!existing.helpText) existing.helpText = 'Ich bestätige, dass ein entsprechender Leistungsnachweis (z.B. Wettkampfergebnis, Trainingsnachweis) vorliegt.';
                              } else {
                                keptFields.push({
                                  id: PROOF_ID,
                                  label: 'Leistungsnachweis vorhanden',
                                  type: 'checkbox',
                                  required: true,
                                  options: [],
                                  visible: true,
                                  onlyForGroup: 'A',
                                  helpText: 'Ich bestätige, dass ein entsprechender Leistungsnachweis (z.B. Wettkampfergebnis, Trainingsnachweis) vorliegt.',
                                });
                              }
                              baseUpdates['CustomFields'] = JSON.stringify(keptFields);
                            }
                            // eslint-disable-next-line @typescript-eslint/no-explicit-any
                            delete (parsed as any)._b2run;
                            baseUpdates['EmailTemplateOverrides'] = JSON.stringify(parsed);
                          }
                        }
                      } catch { /* invalid JSON → einfach ignorieren */ }
                      // v11.14: hardcoded B2Run-Field-Specials in echte
                      // Field-Properties migrieren.
                      const fieldExtras = migrateB2RunFieldExtras(keptFields);
                      if (fieldExtras.changed) {
                        baseUpdates['CustomFields'] = JSON.stringify(keptFields);
                      }
                      const ok = await updateEvent(ev.id, baseUpdates);
                      try { await updateEvent(ev.id, { 'EventType': 'Other' }); } catch { /* SP-Spalte evtl. nicht vorhanden — ignoriert */ }
                      if (!ok) errors.push(`„${ev.title}"`);
                    } catch (err) {
                      console.warn('[DEX] migrate event failed:', ev.id, err);
                      errors.push(`„${ev.title}"`);
                    }
                  };
                  try {
                    // Hauptevent zuerst, dann alle b2run-Sub-Events.
                    await migrateOne(selectedEvent);
                    for (const k of kidsToMigrate) {
                      await migrateOne(k);
                    }
                    await refreshEvents();
                    if (errors.length === 0) {
                      const total = 1 + kidsToMigrate.length;
                      showAlert(isDe
                        ? `Migration abgeschlossen — ${total} Event(s) auf das Standard-Schema umgestellt.`
                        : `Migration complete — ${total} event(s) converted to the standard schema.`);
                    } else {
                      showAlert(isDe
                        ? `Migration teilweise fehlgeschlagen bei: ${errors.join(', ')}. Siehe Browser-Console für Details.`
                        : `Migration partially failed for: ${errors.join(', ')}. See browser console for details.`);
                    }
                  } catch (err) {
                    console.warn('[DEX] migrate B2Run event failed:', err);
                    showAlert(isDe ? 'Migration fehlgeschlagen — siehe Browser-Console.' : 'Migration failed — see browser console.');
                  }
                }}
              />
            )}

            {/* v11.11: Custom-Fields aus Versionsverlauf zurückholen.
                Hilft den Admins, denen die v11.9-Migration die b2run_*-
                Felder (Altersgruppe, T-Shirt-Größe etc.) versehentlich
                aus customFields entfernt hat. Liest die SP-Versionen des
                Event-Items, sucht die jüngste Version mit b2run_*-
                Feldern und mergt diese zurück in das aktuelle
                CustomFields-Array. Bestehende Felder bleiben unverändert
                — es werden NUR fehlende b2run_*-Felder ergänzt. */}

            {isAdmin && selectedEvent && (
              <ActionTile
                icon={<RefreshCw size={18} />}
                category="maintenance"
                title={isDe ? 'Custom-Fields aus Versionsverlauf zurückholen' : 'Restore custom fields from version history'}
                desc={isDe
                  ? 'Holt versehentlich verloren gegangene Anmeldefelder (z.B. Altersgruppe, T-Shirt-Größe, Startblock, Mobilnummer) aus einer früheren Version des Events zurück. Bestehende Felder bleiben unangetastet — es wird nur Fehlendes ergänzt.'
                  : 'Reads the SharePoint version history of the event and restores lost b2run_* custom fields (age group, t-shirt size, start block, mobile number etc.). Useful after the v11.9 migration which deleted these fields by accident. Existing fields are NOT overwritten — only missing fields are added.'}
                badge="admin"
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent) return;
                  try {
                    const history = await eventServiceRef.getEventCustomFieldsHistory(parseInt(selectedEvent.id, 10));
                    if (history.length === 0) {
                      showAlert(isDe ? 'Kein Versionsverlauf gefunden — entweder hat das Event keine Versionen oder der Zugriff wurde verweigert.' : 'No version history found — the event has no versions or access was denied.');
                      return;
                    }
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const currentFields: any[] = (selectedEvent.eventSpecificFields || []).map(f => ({ ...f }));
                    const currentIds = new Set(currentFields.map(f => String(f.id || '').toLowerCase()));
                    // Jüngste Version mit b2run_*-Feldern finden, die noch
                    // NICHT in currentFields stecken.
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    let foundFields: any[] = [];
                    let foundVersion = '';
                    let foundModified = '';
                    for (const v of history) {
                      const missingB2run = v.customFields.filter(f => {
                        const id = String(f.id || '').toLowerCase();
                        return id.indexOf('b2run_') === 0 && !currentIds.has(id);
                      });
                      if (missingB2run.length > 0) {
                        foundFields = missingB2run;
                        foundVersion = v.versionLabel;
                        foundModified = v.modified;
                        break;
                      }
                    }
                    if (foundFields.length === 0) {
                      showAlert(isDe ? 'Keine fehlenden b2run_*-Felder im Versionsverlauf gefunden — entweder sind alle Felder schon vorhanden oder es gab nie welche.' : 'No missing b2run_* fields found in the version history — either all fields already exist or there never were any.');
                      return;
                    }
                    const fieldList = foundFields.map(f => `• ${f.label || f.id}`).join('\n');
                    const modifiedDate = foundModified ? new Date(foundModified).toLocaleString(isDe ? 'de-DE' : 'en-GB') : '?';
                    if (!(await confirmDialog(isDe
                      ? `Folgende ${foundFields.length} Custom-Field(s) aus Version ${foundVersion} (${modifiedDate}) zurückholen?\n\n${fieldList}\n\nDie Felder werden ans Ende deiner aktuellen Felder-Liste angehängt. Du kannst sie danach im Wizard frei umbenennen, neu sortieren oder löschen.`
                      : `Restore the following ${foundFields.length} custom field(s) from version ${foundVersion} (${modifiedDate})?\n\n${fieldList}\n\nThe fields are appended to the end of your current field list. You can rename, reorder or delete them afterwards in the wizard.`, { confirmLabel: isDe ? 'Zurückholen' : 'Restore' }))) {
                      return;
                    }
                    const merged = [...currentFields, ...foundFields];
                    const ok = await updateEvent(selectedEvent.id, { 'CustomFields': JSON.stringify(merged) });
                    if (!ok) {
                      showAlert(isDe ? 'Update fehlgeschlagen — siehe Browser-Console.' : 'Update failed — see browser console.');
                      return;
                    }
                    // Subsite-Spalten gleich mit-syncen, damit die b2run_*-
                    // Spalten in der Teilnehmerliste wieder existieren.
                    if (selectedEvent.subsiteUrl) {
                      try {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const cfForFix: any[] = merged.map((f: any) => ({
                          id: f.id,
                          label: f.label,
                          type: f.type,
                          required: !!f.required,
                          visible: true,
                          options: f.options || [],
                          spInternalName: f.spInternalName || '',
                          ...(f.helpText ? { helpText: f.helpText } : {}),
                          ...(f.multi ? { multi: true } : {}),
                          ...(f.showIf ? { showIf: f.showIf } : {}),
                        }));
                        const splitActive = (selectedEvent.durchstarterCapacity || 0) > 0 && (selectedEvent.funstarterCapacity || 0) > 0;
                        await eventServiceRef.fixRegistrationListColumns(selectedEvent.subsiteUrl, {
                          isB2Run: splitActive,
                          hasQuiz: (selectedEvent.quiz || []).length > 0,
                          customFields: cfForFix,
                        });
                      } catch (err) { console.warn('[DEX] fixRegistrationListColumns nach Restore fehlgeschlagen:', err); }
                    }
                    await refreshEvents();
                    showAlert(isDe
                      ? `${foundFields.length} Custom-Field(s) erfolgreich aus Version ${foundVersion} zurückgeholt.`
                      : `${foundFields.length} custom field(s) successfully restored from version ${foundVersion}.`);
                  } catch (err) {
                    console.warn('[DEX] restore custom fields from history failed:', err);
                    showAlert(isDe ? 'Zurückholen fehlgeschlagen — siehe Browser-Console.' : 'Restore failed — see browser console.');
                  }
                }}
              />
            )}

            {/* 9. Felder reparieren — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<Wrench size={18} />}
                category="maintenance"
                title={isFixingFields ? (isDe ? 'Felder werden repariert…' : 'Repairing fields…') : (isDe ? 'Felder reparieren' : 'Repair fields')}
                desc={isDe
                  ? "Räumt die Anmeldefelder dieses Events automatisch auf: AGB/Datenschutz wird eine richtige Checkbox, T-Shirt-Auswahl bekommt eine 'Kein T-Shirt'-Option, doppelte '(Pflicht)'-Zusätze verschwinden."
                  : "Normalizes custom fields: terms/privacy → checkbox, t-shirt → 'no t-shirt' option, add B2Run special fields, remove redundant '(required)' suffixes."}
                badge="admin"
                busy={isFixingFields}
                disabled={!selectedEvent}
                result={fixFieldsResult}
                resultIsError={!!fixFieldsResult && (fixFieldsResult.startsWith('Fehler') || fixFieldsResult.startsWith('Update fehl') || fixFieldsResult.startsWith('Error') || fixFieldsResult.startsWith('Update failed'))}
                onClick={async () => {
                  if (!selectedEvent) return;
                  setIsFixingFields(true);
                  setFixFieldsResult(null);
                  try {
                    const changes: string[] = [];
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const raw: any[] = (selectedEvent.eventSpecificFields || []).map((f: any) => ({ ...f }));
                    const hasField = (id: string): boolean => raw.some(f => f.id === id);
                    const isB2Run = raw.some(f => String(f.id || '').indexOf('b2run_') === 0);
                    if (isB2Run) {
                      if (!hasField('b2run_infoservice')) {
                        raw.push({ id: 'b2run_infoservice', label: 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)', type: 'checkbox', required: false, options: [], visible: true });
                        changes.push("Feld ergänzt: 'Infoservice'");
                      }
                      if (!hasField('b2run_anonym')) {
                        raw.push({ id: 'b2run_anonym', label: 'Anonym teilnehmen', type: 'checkbox', required: false, options: [], visible: true });
                        changes.push("Feld ergänzt: 'Anonym teilnehmen'");
                      }
                      const hasLaufshirt = raw.some(f => f.id === 'b2run_laufshirt' || /laufshirt/i.test(String(f.label || '')));
                      if (!hasLaufshirt) {
                        raw.push({ id: 'b2run_laufshirt', label: 'Deloitte-Laufshirt', type: 'select', required: true, options: ['Habe bereits ein Laufshirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true });
                        changes.push("Feld ergänzt: 'Deloitte-Laufshirt' (Pflicht)");
                      }
                    }
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const fixed = raw.map((f: any) => {
                      const nf = { ...f };
                      const label = String(nf.label || '');
                      const lowLabel = label.toLowerCase();
                      const isConsent = lowLabel.indexOf('zustimmung') >= 0
                        || lowLabel.indexOf('agb') >= 0
                        || lowLabel.indexOf('datenschutz') >= 0;
                      const isB2RunCheckbox = ['b2run_infoservice', 'b2run_anonym', 'b2run_datenschutz'].indexOf(nf.id) >= 0;
                      if ((isConsent || isB2RunCheckbox) && nf.type !== 'checkbox') {
                        nf.type = 'checkbox';
                        nf.options = [];
                        changes.push(`${label} -> Checkbox`);
                      }
                      const isShirt = lowLabel.indexOf('t-shirt') >= 0 || lowLabel.indexOf('tshirt') >= 0 || lowLabel.indexOf('shirt') >= 0;
                      if (isShirt && nf.type === 'select') {
                        const opts: string[] = Array.isArray(nf.options) ? nf.options.slice() : [];
                        const hasNo = opts.some((o: string) => o.toLowerCase().indexOf('kein') >= 0);
                        if (!hasNo) {
                          opts.unshift('Ohne T-Shirt');
                          nf.options = opts;
                          changes.push(`${label} -> 'Ohne T-Shirt'-Option`);
                        }
                        if (nf.required) {
                          nf.required = false;
                          changes.push(`${label} -> optional`);
                        }
                      }
                      const stripped = label.replace(/\s*\((?:pflicht|mandatory|required)\)\s*$/i, '').trim();
                      if (stripped && stripped !== label) {
                        nf.label = stripped;
                        changes.push(`Label "${label}" -> "${stripped}"`);
                      }
                      if (nf.id === 'b2run_mobilnummer') {
                        if (nf.required) { nf.required = false; changes.push('Mobilnummer -> optional'); }
                        if (nf.label === 'Mobilnummer') {
                          nf.label = 'Mobilnummer (nur bei aktiviertem Infoservice)';
                          changes.push("Mobilnummer-Label präzisiert");
                        }
                      }
                      if (nf.id === 'b2run_infoservice' && nf.label && nf.label.indexOf('benötigt') >= 0) {
                        nf.label = 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)';
                        changes.push('Infoservice-Label modernisiert');
                      }
                      if (nf.id === 'b2run_datenschutz') {
                        const needLinks = !Array.isArray(nf.externalLinks) || nf.externalLinks.length === 0;
                        if (needLinks) {
                          nf.externalLinks = [
                            { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
                            { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
                          ];
                          changes.push('B2Run-Datenschutz: AGB + Datenschutz Links ergänzt');
                        }
                      }
                      if (nf.id === 'b2run_laufshirt' || /laufshirt/i.test(label)) {
                        if (!nf.required) {
                          nf.required = true;
                          changes.push(`${label || nf.id}: als Pflichtfeld markiert`);
                        }
                        if (nf.type === 'select') {
                          const opts: string[] = Array.isArray(nf.options) ? nf.options.slice() : [];
                          const hasNo = opts.some((o: string) => o.toLowerCase().indexOf('kein') >= 0);
                          if (!hasNo) {
                            opts.unshift('Habe bereits ein Laufshirt');
                            nf.options = opts;
                            changes.push(`${label || nf.id}: 'Habe bereits ein Laufshirt'-Option hinzugefügt`);
                          }
                        }
                      }
                      return nf;
                    });
                    const dsIdx = fixed.findIndex((f: { id: string }) => f.id === 'b2run_datenschutz');
                    if (dsIdx >= 0 && dsIdx !== fixed.length - 1) {
                      const [ds] = fixed.splice(dsIdx, 1);
                      fixed.push(ds);
                      changes.push('Zustimmung-Checkbox ans Ende verschoben');
                    }
                    const ok = await updateEvent(selectedEvent.id, { CustomFields: JSON.stringify(fixed) });
                    if (ok) {
                      setFixFieldsResult(changes.length > 0
                        ? (isDe ? `Geändert: ${changes.join(' | ')}` : `Changed: ${changes.join(' | ')}`)
                        : (isDe ? 'Keine Änderungen nötig.' : 'No changes needed.'));
                    } else {
                      setFixFieldsResult(isDe ? 'Update fehlgeschlagen.' : 'Update failed.');
                    }
                  } catch (err) {
                    setFixFieldsResult((isDe ? 'Fehler: ' : 'Error: ') + (err instanceof Error ? err.message : String(err)));
                  }
                  setIsFixingFields(false);
                }}
              />
            )}

            {/* 10. Profile neu laden — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<RefreshCw size={18} />}
                category="maintenance"
                title={isRefreshingProfiles ? (isDe ? 'Teilnehmer werden nachgeladen…' : 'Reloading attendees…') : (isDe ? 'Teilnehmer nachladen (Daten reparieren)' : 'Reload attendees (repair data)')}
                desc={isDe
                  ? 'Lädt Name, JobTitle, Standort, Department und Telefonnummer der letzten N Teilnehmer frisch aus dem Microsoft-365-Benutzerprofil. Repariert auch kaputte Namen — z.B. wenn statt des Vornamens ein technisches Anmelde-Kürzel in der Liste steht.'
                  : 'Reloads name, job title, location, department and phone of the last N attendees from the Microsoft 365 user profile. Also repairs broken names — e.g. when a technical login token appears instead of the first name.'}
                badge="admin"
                busy={isRefreshingProfiles}
                disabled={!selectedEvent?.subsiteUrl}
                result={refreshProfilesResult}
                resultIsError={!!refreshProfilesResult && (refreshProfilesResult.indexOf('Fehler') >= 0 || refreshProfilesResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  const ans = prompt(isDe
                    ? 'Wie viele der letzten Teilnehmer sollen aus dem Benutzerprofil neu geladen werden? (JobTitle, Standort, Department, Phone)'
                    : 'How many of the most recent participants should be reloaded from the user profile? (job title, location, department, phone)', '20');
                  if (!ans) return;
                  const n = parseInt(ans, 10);
                  if (isNaN(n) || n <= 0) { showAlert(isDe ? 'Bitte eine positive Zahl eingeben.' : 'Please enter a positive number.'); return; }
                  setIsRefreshingProfiles(true);
                  setRefreshProfilesResult(null);
                  try {
                    const result = await eventServiceRef.fixEventParticipantsProfileData(selectedEvent.subsiteUrl, n);
                    setRefreshProfilesResult(isDe
                      ? `${result.scanned} geprüft, ${result.updated} aktualisiert, ${result.failedLookups} Profil-Lookups fehlgeschlagen`
                      : `${result.scanned} checked, ${result.updated} updated, ${result.failedLookups} profile lookups failed`);
                    const regs = await getAllRegistrations(selectedEvent.id);
                    setRegistrations(regs);
                  } catch {
                    setRefreshProfilesResult(isDe ? 'Fehler beim Auffrischen der Profile' : 'Error refreshing profiles');
                  }
                  setIsRefreshingProfiles(false);
                }}
              />
            )}

            {/* v19.30 (Feature D): Audit-Log / Änderungsprotokoll dieses
                Events öffnen — vorgefiltert auf den Event-Titel. Sichtbar für
                Admin oder Organizer dieses Events. Zeigt pro Eintrag Zeitpunkt,
                Akteur, Aktion, Ziel-Teilnehmer und bei Daten-Änderungen das
                Vorher → Nachher je Feld. */}
            {(isAdmin || isOrganizerFor(selectedEvent)) && (
              <ActionTile
                icon={<FileText size={18} />}
                category="event"
                title={isDe ? 'Audit-Log / Änderungsprotokoll' : 'Audit log / change history'}
                desc={isDe
                  ? 'Öffnet das Änderungsprotokoll vorgefiltert auf dieses Event. Du siehst pro Eintrag: wann, wer, welche Aktion (z.B. bearbeitet, abgemeldet, gelöscht), welcher Teilnehmer betroffen war und bei Daten-Änderungen den genauen Vorher → Nachher-Vergleich je Feld.'
                  : 'Opens the change history pre-filtered to this event. Each entry shows: when, who, which action (e.g. edited, deregistered, deleted), which participant was affected and — for data changes — the exact before → after comparison per field.'}
                badge="organizer"
                onClick={openChangeLogForEvent}
              />
            )}
          </div>
        </ActionsCollapsibleCard>
        )}
      </div>
      </ActionsRegistryProvider>

      {/* Zähler + QR/Check-in Aktionen.
          v9.14: Warteliste-KPI wird nur gerendert wenn Event eine Warteliste hat.
          Sonst Grid auf 4 Spalten.
          v11.32: Bei Split-Capacity wird die separate Kapazitäts-Karten-Reihe
          unten in die „Angemeldet"-Kachel hochgezogen. Die Kachel bekommt
          dann doppelte Breite (2fr) damit Group-A/B-Breakdown sauber drin
          Platz hat — keine zwei breiten Vollbreite-Karten mehr. */}
      {(() => {
        const hasWaitlistKPI = !!(selectedEvent?.waitlistEnabled && (selectedEvent?.maxParticipants || 0) > 0);
        // Fraktionen pro Spalte — Angemeldet bekommt 2fr wenn Split aktiv ist.
        const angeFr = isSplitCapacity ? '2fr' : '1fr';
        const tail = `1fr 1fr${hasWaitlistKPI ? ' 1fr' : ''} 1fr`; // QR / Eingecheckt / [Warteliste] / Abgemeldet
        const gridCols = `${angeFr} ${tail}`;
        // v15.14: Im subEventsOnlyMode (Hauptevent ohne eigene Anmeldungen)
        // beziehen sich die Stat-Cards auf die konsolidierten Teilnehmer
        // über alle Sub-Events. Die Hauptevent-Liste selbst hat hier nur
        // Alt-Daten und würde das echte Bild verfälschen.
        const consolidatedRegs: SPRegistration[] = isConsolidatedMode
          // v22.63: Klammer-eigene Abmeldungen (Absagen auf der Klammer-Subsite)
          // in die KPI „Abgemeldet" mitzählen, damit KPI und Abmeldungs-Liste
          // übereinstimmen. Nur Abgemeldet-Zeilen, um die Aktiv-Zahlen nicht zu
          // verfälschen.
          ? ([] as SPRegistration[]).concat(
              registrations.filter(r => r.Status === 'Abgemeldet'),
              ...Object.values(subEventRegsByEventId),
            )
          : [];
        const consolidatedActiveByEmail = new Set<string>();
        const consolidatedQRByEmail = new Set<string>();
        const consolidatedCheckedByEmail = new Set<string>();
        const consolidatedWaitlistByEmail = new Set<string>();
        const consolidatedCancelledByEmail = new Set<string>();
        const consolidatedAnyByEmail = new Set<string>();
        for (const r of consolidatedRegs) {
          const key = (r.ParticipantEmail || '').toLowerCase().trim();
          if (!key) continue;
          consolidatedAnyByEmail.add(key);
          if (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt') consolidatedActiveByEmail.add(key);
          if (r.Status === 'QR versendet') consolidatedQRByEmail.add(key);
          if (r.Status === 'Eingecheckt') consolidatedCheckedByEmail.add(key);
          if (r.Status === 'Warteliste') consolidatedWaitlistByEmail.add(key);
          if (r.Status === 'Abgemeldet') consolidatedCancelledByEmail.add(key);
        }
        const active = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
        const totalActive = isConsolidatedMode ? consolidatedActiveByEmail.size : active.length;
        // v19.12: nach EFFEKTIVER Gruppe zählen (StarterType ODER, falls leer,
        // PreferredStarterType). Sonst fehlen angemeldete Nachrücker, deren
        // StarterType der Flow nicht gesetzt hat, in der Gruppen-Zahl — dann ist
        // „Angemeldet" gesamt ≠ Summe der beiden Gruppen (Beobachtung: 142 vs.
        // 130+11). Eine angemeldete Person mit Wunsch „Funstarter" belegt real
        // einen Funstarter-Platz und muss dort mitgezählt werden.
        const durchActive = active.filter(r => (r.StarterType || r.PreferredStarterType) === 'Durchstarter').length;
        const funActive = active.filter(r => (r.StarterType || r.PreferredStarterType) === 'Funstarter').length;
        const durchCap = selectedEvent?.durchstarterCapacity || 0;
        const funCap = selectedEvent?.funstarterCapacity || 0;
        const labelA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
        const labelB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
        const reversed = !!selectedEvent?.splitDisplayOrderReversed;
        const grpA = (
          <div key="grpA" style={{ display: 'flex', alignItems: 'baseline', gap: 8 }}>
            <span style={{ color: 'var(--dex-green-dark, #6b9a1e)', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={labelA}>● {labelA}</span>
            <strong style={{ whiteSpace: 'nowrap' }}>{durchActive}<span style={{ color: 'var(--dex-gray-400)' }}>/{durchCap}</span></strong>
          </div>
        );
        const grpB = (
          <div key="grpB" style={{ display: 'flex', alignItems: 'baseline', gap: 8 }}>
            <span style={{ color: 'var(--dex-orange, #ff8c00)', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={labelB}>● {labelB}</span>
            <strong style={{ whiteSpace: 'nowrap' }}>{funActive}<span style={{ color: 'var(--dex-gray-400)' }}>/{funCap}</span></strong>
          </div>
        );
        return (
          <div className="admin-counters" style={{ display: 'grid', gridTemplateColumns: gridCols, gap: 12, marginBottom: 24 }}>
            <div className="card" style={{ padding: 16, textAlign: 'center' }}>
              <div style={{ fontSize: '1.8rem', fontWeight: 700, color: '#1565c0' }}>{totalActive}</div>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.registered')}</div>
              {isSplitCapacity && (
                <div style={{
                  marginTop: 10, paddingTop: 10,
                  borderTop: '1px solid var(--dex-gray-200)',
                  fontSize: '0.82rem', textAlign: 'left',
                  display: 'flex', flexDirection: 'column', gap: 4,
                }}>
                  {reversed ? <>{grpB}{grpA}</> : <>{grpA}{grpB}</>}
                </div>
              )}
            </div>
            <div className="card" style={{ padding: 16, textAlign: 'center' }}>
              <div style={{ fontSize: '1.8rem', fontWeight: 700, color: '#6a1b9a' }}>
                {isConsolidatedMode ? consolidatedQRByEmail.size : registrations.filter(r => r.Status === 'QR versendet').length}
              </div>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.qrsent')}</div>
            </div>
            <div className="card" style={{ padding: 16, textAlign: 'center' }}>
              <div style={{ fontSize: '1.8rem', fontWeight: 700, color: 'var(--dex-green)' }}>
                {isConsolidatedMode ? consolidatedCheckedByEmail.size : registrations.filter(r => r.Status === 'Eingecheckt').length}
              </div>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.checkedin')}</div>
            </div>
            {hasWaitlistKPI && (
              <div className="card" style={{ padding: 16, textAlign: 'center' }}>
                <div style={{ fontSize: '1.8rem', fontWeight: 700, color: 'var(--dex-orange)' }}>
                  {isConsolidatedMode ? consolidatedWaitlistByEmail.size : registrations.filter(r => r.Status === 'Warteliste').length}
                </div>
                <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.waitlist')}</div>
              </div>
            )}
            <div className="card" style={{ padding: 16, textAlign: 'center' }}>
              <div style={{ fontSize: '1.8rem', fontWeight: 700, color: 'var(--dex-gray-400)' }}>
                {isConsolidatedMode ? consolidatedCancelledByEmail.size : registrations.filter(r => r.Status === 'Abgemeldet').length}
              </div>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.cancelled')}</div>
            </div>
          </div>
        );
      })()}

      {/* v9.20: Check-In starten + QR-Codes versenden sind jetzt im Aktionen-Grid
          unten als ActionTile gerendert (nicht mehr als eigene Button-Reihe).
          Damit sind alle Quick-Actions an EINEM Ort zusammengefasst. Auch für
          Check-In-only-User (qrScanner-Mode) — die sehen weiterhin nur den
          Check-In-Tile, da das Aktionen-Grid für sie unten gefiltert ist. */}
      {!isQRScannerOnlyForSelected && (<>

      {/* ===== QUIZ-STATISTIK (collapsible, oberhalb Teilnehmerliste) ===== */}
      {selectedEvent && selectedEvent.quiz && selectedEvent.quiz.length > 0 && (() => {
        // Teilnehmer mit mindestens einer beantworteten Frage (nicht nur "komplett durchgeführt").
        // Dadurch erscheinen auch Teilnehmer, die mittendrin aufgehört haben.
        const regsWithQuiz = registrations.filter(r => {
          if (!r.QuizAnswers) return false;
          try {
            const parsed = JSON.parse(r.QuizAnswers);
            return Array.isArray(parsed) && parsed.some((a: number[]) => Array.isArray(a) && a.length > 0);
          } catch { return false; }
        });
        const regsCompleted = regsWithQuiz.filter(r => typeof r.QuizCompletedAt === 'string' && r.QuizCompletedAt);
        const totalQuizzes = regsWithQuiz.length;
        const totalCompleted = regsCompleted.length;

        // Pro Frage: wie viele haben sie überhaupt beantwortet, wie viele richtig
        const perQuestion = selectedEvent.quiz.map((q, qIdx) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const correct = (q as any).correctIndices || [(q as any).correctIndex || 0];
          let correctCount = 0;
          let answeredCount = 0;
          for (const reg of regsWithQuiz) {
            try {
              const answers = JSON.parse(reg.QuizAnswers || '[]');
              const given: number[] = Array.isArray(answers[qIdx]) ? answers[qIdx] : [];
              if (given.length === 0) continue;
              answeredCount++;
              const isRight = correct.length === given.length && correct.every((c: number) => given.indexOf(c) >= 0);
              if (isRight) correctCount++;
            } catch { /* skip */ }
          }
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const imageBase64 = (q as any).imageBase64 as string | undefined;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const section = (q as any).section as string | undefined;
          return { question: q.question, imageBase64, section, correctCount, answeredCount, total: totalQuizzes };
        });

        // Top 10 nach Score, bei Gleichstand: abgeschlossene vor nicht-abgeschlossenen, dann Zeitpunkt
        const top10 = regsWithQuiz.slice().sort((a, b) => {
          const sa = a.QuizScore || 0;
          const sb = b.QuizScore || 0;
          if (sb !== sa) return sb - sa;
          const aDone = !!a.QuizCompletedAt;
          const bDone = !!b.QuizCompletedAt;
          if (aDone !== bDone) return aDone ? -1 : 1;
          const ta = new Date(a.QuizCompletedAt || 0).getTime();
          const tb = new Date(b.QuizCompletedAt || 0).getTime();
          return ta - tb;
        }).slice(0, 10);

        return (
          <details className="card" style={{ padding: 0, marginBottom: 16 }}>
            <summary style={{
              padding: '16px 24px', cursor: 'pointer', listStyle: 'none',
              display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12,
              fontSize: '1rem', fontWeight: 600,
            }}>
              <span style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <FileText size={18} /> Quiz-Statistik
              </span>
              <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                {totalQuizzes === 0
                  ? 'Keine Daten'
                  : `${totalCompleted} abgeschlossen, ${totalQuizzes - totalCompleted} teilweise (Klick zum Ausklappen)`}
              </span>
            </summary>
            <div style={{ padding: '0 24px 24px 24px' }}>
              {totalQuizzes === 0 ? (
                <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', margin: 0 }}>
                  Noch kein Teilnehmer hat das Quiz gestartet.
                </p>
              ) : (
                <>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 24 }}>
                    <div style={{ padding: 16, background: 'var(--dex-green-light, #f0fdf4)', borderRadius: 12, textAlign: 'center' }}>
                      <div style={{ fontSize: '1.6rem', fontWeight: 700, color: 'var(--dex-green-dark, #6b9a1e)' }}>{totalCompleted}</div>
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Abgeschlossen</div>
                    </div>
                    <div style={{ padding: 16, background: 'var(--dex-orange-light, #fff7ed)', borderRadius: 12, textAlign: 'center' }}>
                      <div style={{ fontSize: '1.6rem', fontWeight: 700, color: 'var(--dex-orange, #ed8b00)' }}>{totalQuizzes - totalCompleted}</div>
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Teilweise</div>
                    </div>
                    <div style={{ padding: 16, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, textAlign: 'center' }}>
                      <div style={{ fontSize: '1.6rem', fontWeight: 700 }}>{selectedEvent.quiz.length}</div>
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Fragen</div>
                    </div>
                    <div style={{ padding: 16, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, textAlign: 'center' }}>
                      <div style={{ fontSize: '1.6rem', fontWeight: 700 }}>
                        {totalQuizzes > 0
                          ? (regsWithQuiz.reduce((sum, r) => sum + (r.QuizScore || 0), 0) / totalQuizzes).toFixed(1)
                          : '0'}
                      </div>
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Ø Score</div>
                    </div>
                  </div>

                  {/* Pro Frage - gruppiert nach Bereich falls vorhanden */}
                  <h4 style={{ marginTop: 0, marginBottom: 12 }}>Pro Frage</h4>
                  {(() => {
                    const hasSections = perQuestion.some(pq => !!pq.section);
                    if (!hasSections) return null;
                    // Gruppen in Reihenfolge der ersten Erwähnung
                    const sectionsInOrder: string[] = [];
                    for (const pq of perQuestion) {
                      if (pq.section && sectionsInOrder.indexOf(pq.section) < 0) sectionsInOrder.push(pq.section);
                    }
                    return (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 16, marginBottom: 24 }}>
                        {sectionsInOrder.map(sec => (
                          <div key={`stat-sec-${sec}`}>
                            <h5 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.92rem' }}>
                              Bereich: {sec}
                            </h5>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              {perQuestion.map((pq, idx) => pq.section === sec ? (() => {
                                const pct = pq.answeredCount > 0 ? Math.round((pq.correctCount / pq.answeredCount) * 100) : 0;
                                return (
                                  <div key={idx} style={{ padding: 10, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 8 }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 4, gap: 12 }}>
                                      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flex: 1 }}>
                                        {pq.imageBase64 && (
                                          <img src={pq.imageBase64} alt="" style={{ width: 60, height: 44, objectFit: 'cover', borderRadius: 6, flexShrink: 0, border: '1px solid var(--dex-gray-200)' }} />
                                        )}
                                        <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>{idx + 1}. {pq.question}</span>
                                      </div>
                                      <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>
                                        {pq.correctCount} / {pq.answeredCount} richtig ({pct}%)
                                      </span>
                                    </div>
                                    <div style={{ height: 6, background: 'var(--dex-gray-200)', borderRadius: 3, overflow: 'hidden' }}>
                                      <div style={{ height: '100%', width: `${pct}%`, background: pct >= 70 ? 'var(--dex-green, #86bc25)' : pct >= 40 ? 'var(--dex-orange, #ff8c00)' : 'var(--dex-red, #c00)', transition: 'width 0.3s' }} />
                                    </div>
                                  </div>
                                );
                              })() : null)}
                            </div>
                          </div>
                        ))}
                        {/* Fragen ohne Bereich */}
                        {perQuestion.some(pq => !pq.section) && (
                          <div>
                            <h5 style={{ margin: '0 0 6px', color: 'var(--dex-gray-600)', fontSize: '0.92rem' }}>Ohne Bereich</h5>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              {perQuestion.map((pq, idx) => !pq.section ? (() => {
                                const pct = pq.answeredCount > 0 ? Math.round((pq.correctCount / pq.answeredCount) * 100) : 0;
                                return (
                                  <div key={idx} style={{ padding: 10, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 8 }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 4, gap: 12 }}>
                                      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flex: 1 }}>
                                        {pq.imageBase64 && (
                                          <img src={pq.imageBase64} alt="" style={{ width: 60, height: 44, objectFit: 'cover', borderRadius: 6, flexShrink: 0, border: '1px solid var(--dex-gray-200)' }} />
                                        )}
                                        <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>{idx + 1}. {pq.question}</span>
                                      </div>
                                      <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>
                                        {pq.correctCount} / {pq.answeredCount} richtig ({pct}%)
                                      </span>
                                    </div>
                                    <div style={{ height: 6, background: 'var(--dex-gray-200)', borderRadius: 3, overflow: 'hidden' }}>
                                      <div style={{ height: '100%', width: `${pct}%`, background: pct >= 70 ? 'var(--dex-green, #86bc25)' : pct >= 40 ? 'var(--dex-orange, #ff8c00)' : 'var(--dex-red, #c00)', transition: 'width 0.3s' }} />
                                    </div>
                                  </div>
                                );
                              })() : null)}
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })()}
                  {!perQuestion.some(pq => !!pq.section) && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 24 }}>
                    {perQuestion.map((pq, idx) => {
                      const pct = pq.answeredCount > 0 ? Math.round((pq.correctCount / pq.answeredCount) * 100) : 0;
                      return (
                        <div key={idx} style={{ padding: 10, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 8 }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 4, gap: 12 }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 10, flex: 1 }}>
                              {pq.imageBase64 && (
                                <img
                                  src={pq.imageBase64}
                                  alt=""
                                  style={{ width: 60, height: 44, objectFit: 'cover', borderRadius: 6, flexShrink: 0, border: '1px solid var(--dex-gray-200)' }}
                                />
                              )}
                              <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>
                                {idx + 1}. {pq.question}
                              </span>
                            </div>
                            <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>
                              {pq.correctCount} / {pq.answeredCount} richtig ({pct}%)
                            </span>
                          </div>
                          <div style={{ height: 6, background: 'var(--dex-gray-200)', borderRadius: 3, overflow: 'hidden' }}>
                            <div style={{
                              height: '100%',
                              width: `${pct}%`,
                              background: pct >= 70 ? 'var(--dex-green, #86bc25)' : pct >= 40 ? 'var(--dex-orange, #ff8c00)' : 'var(--dex-red, #c00)',
                              transition: 'width 0.3s',
                            }} />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                  )}

                  {/* Top 10 */}
                  <h4 style={{ marginTop: 0, marginBottom: 12 }}>Top 10 Teilnehmer</h4>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                          <th style={{ textAlign: 'left', padding: 8, width: 40 }}>#</th>
                          <th style={{ textAlign: 'left', padding: 8 }}>Name</th>
                          <th style={{ textAlign: 'left', padding: 8 }}>E-Mail</th>
                          <th style={{ textAlign: 'left', padding: 8, width: 80 }}>Score</th>
                          <th style={{ textAlign: 'left', padding: 8 }}>Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {top10.map((reg, i) => {
                          const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                          const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `${i + 1}.`;
                          const done = !!reg.QuizCompletedAt;
                          // Beantwortete Fragen zählen (für Partial)
                          let answeredN = 0;
                          try {
                            const parsed = JSON.parse(reg.QuizAnswers || '[]');
                            if (Array.isArray(parsed)) answeredN = parsed.filter((a: number[]) => Array.isArray(a) && a.length > 0).length;
                          } catch { /* */ }
                          return (
                            <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                              <td style={{ padding: 8, fontWeight: 700 }}>{medal}</td>
                              <td style={{ padding: 8, fontWeight: 500 }}>{name}</td>
                              <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                              <td style={{ padding: 8, fontWeight: 700, color: 'var(--dex-green-dark, #6b9a1e)' }}>
                                {reg.QuizScore ?? 0} / {selectedEvent.quiz.length}
                              </td>
                              <td style={{ padding: 8, color: done ? 'var(--dex-gray-500)' : 'var(--dex-orange, #ed8b00)' }}>
                                {done
                                  ? `Abgeschlossen ${formatDate(reg.QuizCompletedAt || '')}`
                                  : `Teilweise (${answeredN}/${selectedEvent.quiz.length})`}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          </details>
        );
      })()}

      {/* Teilnehmerliste */}
      <div className="card" style={{ padding: 24 }}>
        {/* v11.28: Suchfeld direkt neben dem „Teilnehmer (N)"-Header
            statt rechtsbündig — flüssiger Lese-Flow von links nach
            rechts, kein Sprung ueber die ganze Card-Breite mehr. */}
        <div className="mb-16" style={{ display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap' }}>
          <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
            <Users size={18} /> Teilnehmer ({isConsolidatedMode ? consolidatedFiltered.length : activeRegs.length})
            {isConsolidatedMode && (() => {
              const term = (selectedEvent && selectedEvent.childEventTermPlural) || (isDe ? 'Sub-Events' : 'sub-events');
              return (
                <span style={{ marginLeft: 8, fontSize: '0.72rem', fontWeight: 500, color: 'var(--dex-gray-500)' }}>
                  — {isDe ? 'konsolidiert über' : 'consolidated across'} {consolidatedChildren.length} {term}
                </span>
              );
            })()}
          </h3>
          <input
            type="text"
            className="form-input"
            placeholder="Teilnehmer suchen..."
            value={searchQuery}
            onChange={e => setSearchQuery(e.target.value)}
            style={{ maxWidth: 280, padding: '6px 12px', fontSize: '0.85rem' }}
          />
        </div>
        {/* v15.14: Legende für die Pastel-Hintergründe — sowohl in der
            Sub-Event-Detail-Ansicht (Parent-CFs + eigene CFs) als auch im
            konsolidierten Hauptevent-View. Vorher war die Legende NUR im
            konsolidierten View sichtbar und der Organizer hat im Sub-
            Event-Tab nicht gewusst, was Pastel A vs Pastel B bedeutet. */}
        {parentEventForSelected && (
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, marginBottom: 10, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
            {((parentEventForSelected.eventSpecificFields || []).filter(f => f.type !== 'user' && f.label && f.label.trim()).length > 0) && (
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                <span style={{ width: 14, height: 14, borderRadius: 3, background: 'rgba(0, 118, 168, 0.16)', border: '1px solid rgba(0, 118, 168, 0.3)' }} />
                {isDe ? 'Felder des Hauptevents' : 'Main-event fields'}
              </span>
            )}
            {((selectedEvent?.eventSpecificFields || []).filter(f => f.type !== 'user' && f.label && f.label.trim()).length > 0) && (
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                <span style={{ width: 14, height: 14, borderRadius: 3, background: 'rgba(255, 191, 0, 0.22)', border: '1px solid rgba(255, 191, 0, 0.4)' }} />
                {isDe
                  ? `Felder dieses ${(selectedEvent && (selectedEvent as DeloitteEvent & { childEventTermSingular?: string }).childEventTermSingular) || 'Sub-Events'}`
                  : `Fields of this ${(selectedEvent && (selectedEvent as DeloitteEvent & { childEventTermSingular?: string }).childEventTermSingular) || 'sub-event'}`}
              </span>
            )}
          </div>
        )}

        {/* v11.70: Inline-Hinweisbox statt Modal — bei einer kürzlich
            erfolgten Abmeldung läuft die automatische Korrektur evtl. noch
            (Nachrücken + ID-Neuvergabe per Power-Automate-Batch). Solange
            sich die IDs evtl. noch verschieben, soll der Organizer nicht
            parallel manuell „IDs neu vergeben" anstoßen. */}
        {(() => {
          if (!selectedEvent) return null;
          // v22.67: Im Klammer-Modus („Nur Sub-Events") greift die
          // TeilnehmerID-Durchgängigkeits-Prüfung NICHT — die geprüfte Liste
          // sind die Schatten-Zeilen der Klammer (ohne fortlaufende Nummern);
          // die echten TeilnehmerIDs leben pro Sub-Event. Die Warnung war hier
          // ein Fehlalarm.
          if (selectedEvent.subEventsOnlyMode) return null;
          const info = recentCancellation(registrations);
          if (!info.recent) return null;
          const whenStr = info.whenIso ? formatDate(info.whenIso) : '';
          // v22.12: zweiphasig — innerhalb von ~10 Min nach der letzten
          // Abmeldung läuft die automatische Korrektur evtl. noch (warten);
          // danach ist die Lücke ECHT stehengeblieben (typisch: die höchste
          // Nummer wurde abgemeldet, während gleichzeitig neue Anmeldungen
          // bereits höhere Nummern gezogen haben — ein späterer automatischer
          // Lauf kommt nicht, weil nur Abmeldungen die Korrektur anstoßen).
          const minutesSinceCancel = info.whenIso
            ? Math.floor((Date.now() - new Date(info.whenIso).getTime()) / 60000)
            : 999;
          const probablyStillRunning = minutesSinceCancel >= 0 && minutesSinceCancel < 10;
          return (
            <div style={{
              margin: '0 0 16px',
              padding: '14px 16px',
              borderRadius: 8,
              background: 'rgba(237,139,0,0.10)',
              border: '1px solid var(--dex-orange, #ed8b00)',
              color: 'var(--dex-orange-dark, #b35a00)',
            }}>
              <div style={{ fontWeight: 700, marginBottom: 6, fontSize: '0.9rem' }}>
                {probablyStillRunning
                  ? 'TeilnehmerIDs sind gerade nicht durchgängig — automatische Korrektur läuft vermutlich noch'
                  : 'TeilnehmerIDs sind nicht durchgängig — bitte einmal korrigieren'}
              </div>
              <div style={{ fontSize: '0.82rem', lineHeight: 1.5 }}>
                <strong>Geprüft an der geladenen Teilnehmerliste:</strong> {info.detail}.{whenStr ? <> Letzte Abmeldung: <strong>{whenStr}</strong>.</> : ''}{' '}
                {probablyStillRunning ? (
                  <>Die automatische Korrektur — <strong>Nachrücken von der Warteliste</strong> und <strong>Neu-Nummerierung</strong> — braucht nach einer Abmeldung typischerweise 1–5 Minuten. Die Liste wird hier <strong>automatisch alle 30 Sekunden neu geladen</strong>; diese Box verschwindet von selbst, sobald alles stimmt. Bitte in dieser Phase NICHT manuell korrigieren (sonst laufen zwei Korrekturen ineinander).</>
                ) : (
                  <>Die letzte Abmeldung liegt länger zurück — die automatische Korrektur ist also bereits durchgelaufen, die Lücke ist trotzdem geblieben. Das passiert, wenn genau die <strong>höchste Nummer abgemeldet</strong> wurde, während <strong>gleichzeitig neue Anmeldungen</strong> schon höhere Nummern bekommen haben — ein weiterer automatischer Lauf kommt erst bei der nächsten Abmeldung. Das ist rein kosmetisch (Nachrücken/Check-in funktionieren trotzdem) und jetzt <strong>gefahrlos per Klick zu beheben</strong>:</>
                )}
              </div>
              {/* v22.12: manueller Sofort-Check (lädt die Liste neu). */}
              <button
                type="button"
                className="btn btn-secondary"
                disabled={idRecheckBusy}
                style={{ marginTop: 12, marginRight: 10, fontSize: '0.82rem', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                onClick={() => { reloadRegistrationsForIdCheck().catch(() => { /* */ }); }}
              >
                <RefreshCw size={14} /> {idRecheckBusy ? 'Prüft…' : 'Jetzt neu prüfen'}
              </button>
              {/* v18.60: Direkter Korrektur-Button in der Box — Admin ODER
                  Organizer des Events. Use-Case: die automatische Batch-
                  Korrektur ist offensichtlich NICHT gelaufen (IDs seit längerem
                  falsch). Vorher musste der Organizer den Button im Aktionen-
                  Dropdown suchen. */}
              {(isAdmin || isOrganizerFor(selectedEvent)) && !!selectedEvent.subsiteUrl && (
                <button
                  type="button"
                  className="btn btn-primary"
                  disabled={isReorderingIDs}
                  style={{ marginTop: 12, fontSize: '0.82rem', opacity: isReorderingIDs ? 0.6 : 1, display: 'inline-flex', alignItems: 'center', gap: 6 }}
                  onClick={async () => {
                    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                    if (!(await confirmDialog(isDe
                      ? 'TeilnehmerIDs jetzt neu vergeben (1, 2, 3, …)?\n\nNur klicken, wenn die automatische Korrektur offensichtlich nicht gelaufen ist (IDs schon länger falsch) — NICHT mitten in einer Anmeldewelle.'
                      : 'Reassign participant IDs now (1, 2, 3, …)?\n\nOnly click if the automatic correction clearly did not run (IDs wrong for a while) — NOT in the middle of a registration wave.'))) return;
                    await runIdReorder();
                  }}
                >
                  <Hash size={14} /> {isReorderingIDs ? (isDe ? 'IDs werden korrigiert…' : 'Fixing IDs…') : (isDe ? 'IDs jetzt korrigieren' : 'Fix IDs now')}
                </button>
              )}
            </div>
          );
        })()}

        {(() => {
          // v23.2: Doppel-Anmelde-Hinweis. Listet jede Person, die mit
          // derselben E-Mail ≥2 nicht-abgemeldete Zeilen hat (z.B. dieselbe
          // Person in zwei Teams). Die betroffenen Zeilen sind in der Tabelle
          // zusätzlich rot markiert; pro Person kann der Organizer über den
          // „Abmelden"-Button das Duplikat still entfernen.
          if (duplicateEmails.size === 0 || !selectedEvent) return null;
          // Pro betroffener E-Mail die aktiven Zeilen sammeln (Name + Teams).
          const dupGroups: Array<{ email: string; rows: SPRegistration[] }> = [];
          duplicateEmails.forEach(em => {
            const rows = registrations.filter(r => (r.Status || '') !== 'Abgemeldet' && (r.ParticipantEmail || '').trim().toLowerCase() === em);
            if (rows.length > 1) dupGroups.push({ email: em, rows });
          });
          if (dupGroups.length === 0) return null;
          return (
            <div style={{ marginBottom: 20, padding: 16, borderRadius: 12, border: '1px solid var(--dex-red, #c00)', background: 'rgba(200,0,0,0.06)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 10 }}>
                <Icon iconName="Warning" style={{ fontSize: 18, color: 'var(--dex-red, #c00)' }} />
                <strong style={{ color: 'var(--dex-red, #c00)', fontSize: '0.95rem' }}>
                  {isDe ? `Doppel-Anmeldungen erkannt (${dupGroups.length})` : `Duplicate registrations detected (${dupGroups.length})`}
                </strong>
                <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Dieselbe Person ist mehrfach angemeldet. Die betroffenen Zeilen sind unten rot markiert — über „Abmelden" kannst du die doppelte Zeile still entfernen.'
                    : 'The same person is registered more than once. The affected rows are marked red below — use „Cancel" to silently remove the duplicate row.'}
                </span>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {dupGroups.map(g => {
                  const first = g.rows[0];
                  const dispName = (first.Vorname && first.Nachname) ? `${first.Vorname} ${first.Nachname}` : (first.ParticipantName || g.email);
                  const teamList = g.rows
                    .map(r => r.TeamName ? `„${r.TeamName}"` : (r.TeamId ? (isDe ? '(Team ohne Namen)' : '(unnamed team)') : (isDe ? '(Einzel-Anmeldung)' : '(individual)')))
                    .join(', ');
                  return (
                    <div key={g.email} style={{ fontSize: '0.84rem', color: 'var(--dex-gray-800)', display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'baseline' }}>
                      <strong>{dispName}</strong>
                      <span style={{ color: 'var(--dex-gray-500)' }}>{g.email}</span>
                      <span style={{ padding: '1px 8px', borderRadius: 999, background: 'var(--dex-red, #c00)', color: '#fff', fontSize: '0.72rem', fontWeight: 700 }}>
                        {isDe ? `${g.rows.length}× angemeldet` : `${g.rows.length}× registered`}
                      </span>
                      <span style={{ color: 'var(--dex-gray-600)' }}>— {teamList}</span>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })()}

        {(() => {
          // v11.36: Überbuchungs-Review-Box. Zeigt alle per „Überbuchung
          // prüfen" markierten Personen (OverbookReview='Pending') mit
          // Fairness-Kontext + Aktions-Buttons. Erst durch eine Aktion
          // ändert sich der Status.
          const flaggedRaw = registrations.filter(r => r.OverbookReview === 'Pending');
          if (flaggedRaw.length === 0 || !selectedEvent) return null;
          const groupOf = (r: SPRegistration): string => r.StarterType || r.PreferredStarterType || '';
          const ACTIVE_ST = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
          // Gruppen-Key: bei Split die Gruppe, sonst ein gemeinsamer Topf.
          const keyOf = (r: SPRegistration): string => isSplitCapacity ? (groupOf(r) || '?') : 'all';
          const capOf = (key: string): number => {
            if (!isSplitCapacity) return selectedEvent.maxParticipants || 0;
            if (key === 'Durchstarter') return selectedEvent.durchstarterCapacity || 0;
            if (key === 'Funstarter') return selectedEvent.funstarterCapacity || 0;
            return 0;
          };
          // Pro Gruppe: aktive Anmeldungen in Anmeldereihenfolge (Id asc =
          // Reihenfolge der Registrierung — identisch zur Detect-Logik).
          const activeByGroup: Record<string, SPRegistration[]> = {};
          registrations
            .filter(r => ACTIVE_ST.indexOf(r.Status) >= 0)
            .slice()
            .sort((a, b) => a.Id - b.Id)
            .forEach(r => { const k = keyOf(r); (activeByGroup[k] = activeByGroup[k] || []).push(r); });
          // v22.40: Nur Personen anzeigen, die WIRKLICH noch über Kapazität
          // sind. Hat sich zwischenzeitlich jemand abgemeldet, passt eine
          // markierte Person ggf. wieder regulär in die Liste (oder ist selbst
          // nicht mehr aktiv) — solche stale Marker hier ausblenden (der
          // Auto-Heal-Effekt entfernt sie zusätzlich dauerhaft).
          const flagged = flaggedRaw.filter(r => {
            const k = keyOf(r); const cap = capOf(k); const bucket = activeByGroup[k] || [];
            const idx = bucket.findIndex(x => x.Id === r.Id);
            if (idx < 0) return false;
            return cap > 0 && (idx + 1) > cap;
          });
          if (flagged.length === 0) return null;
          // Faire Wartelisten-Reihenfolge je Gruppe: die über Kapazität
          // Aktiven + bereits vorhandene Warteliste, nach RegistrationDate.
          const fairWaitByGroup: Record<string, SPRegistration[]> = {};
          Object.keys(activeByGroup).forEach(k => {
            const cap = capOf(k);
            const overCap = cap > 0 ? activeByGroup[k].slice(cap) : [];
            const existingWl = registrations.filter(r => r.Status === 'Warteliste' && keyOf(r) === k);
            fairWaitByGroup[k] = [...overCap, ...existingWl].sort((a, b) =>
              new Date(a.RegistrationDate).getTime() - new Date(b.RegistrationDate).getTime());
          });
          // Faire Aktiv-Gesamtzahl (bei sauberer Liste) für die faire ID.
          let totalFairActive = 0;
          Object.keys(activeByGroup).forEach(k => {
            const cap = capOf(k);
            totalFairActive += cap > 0 ? Math.min(activeByGroup[k].length, cap) : activeByGroup[k].length;
          });
          const fmtGap = (ms: number): string => {
            if (!isFinite(ms) || ms < 0) return '—';
            const s = Math.round(ms / 1000);
            if (s < 90) return `${s} ${isDe ? 'Sek' : 'sec'}`;
            const m = Math.round(s / 60);
            if (m < 90) return `${m} ${isDe ? 'Min' : 'min'}`;
            const h = Math.round(m / 60);
            if (h < 48) return `${h} ${isDe ? 'Std' : 'h'}`;
            return `${Math.round(h / 24)} ${isDe ? 'Tage' : 'days'}`;
          };
          return (
            <div style={{ marginBottom: 20, padding: 16, borderRadius: 12, border: '1px solid var(--dex-orange, #ed8b00)', background: 'rgba(237,139,0,0.07)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 12 }}>
                <strong style={{ color: 'var(--dex-orange-dark, #b35a00)', fontSize: '0.95rem' }}>
                  {isDe ? `Überbuchung – zu prüfen (${flagged.length})` : `Overbooking – to review (${flagged.length})`}
                </strong>
                <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Über Kapazität angemeldet. Pro Person entscheiden — danach werden IDs automatisch neu vergeben.'
                    : 'Registered over capacity. Decide per person — afterwards the IDs are reassigned automatically.'}
                </span>
                <div style={{ marginLeft: 'auto', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  <button
                    className="btn btn-secondary"
                    style={{ fontSize: '0.78rem', padding: '5px 12px', color: 'var(--dex-red, #c00)' }}
                    onClick={() => { setOverbookModal({ mode: 'confirm', targets: flagged }); setObWithMail(true); setObRemoveCalendar(true); }}
                  >
                    {isDe ? `Alle bestätigen (${flagged.length})` : `Confirm all (${flagged.length})`}
                  </button>
                  <InfoTooltip placement="left" text={isDe
                    ? (
                    <>
                      <strong>Sammel-Aktion:</strong> setzt <strong>alle</strong> markierten Personen auf die <strong>Warteliste</strong> (gruppentreu).<br /><br />
                      Die Optionen <strong>mit/ohne Mail</strong>, <strong>Kalender-Abmeldung</strong> und <strong>Sprache</strong> gelten <strong>für alle gleich</strong> — eine gemeinsame Entscheidung.<br /><br />
                      Der Mailtext ist trotzdem <strong>pro Person personalisiert</strong> (Name + individuelle neue Warteliste-Position).<br /><br />
                      Sollen einzelne Personen <strong>anders</strong> behandelt werden (z.B. &bdquo;Platz behalten&ldquo;), nutze stattdessen die <strong>Einzel-Buttons</strong> pro Zeile.
                    </>
                    ) : (
                    <>
                      <strong>Bulk action:</strong> moves <strong>all</strong> marked people to the <strong>waitlist</strong> (group-faithful).<br /><br />
                      The options <strong>with/without email</strong>, <strong>calendar removal</strong> and <strong>language</strong> apply <strong>to all alike</strong> — a single shared decision.<br /><br />
                      The email text is still <strong>personalized per person</strong> (name + individual new waitlist position).<br /><br />
                      If individual people should be treated <strong>differently</strong> (e.g. &bdquo;keep seat&ldquo;), use the <strong>per-row buttons</strong> instead.
                    </>
                    )
                  } />
                </div>
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '1px solid rgba(237,139,0,0.4)', textAlign: 'left', color: 'var(--dex-gray-600)' }}>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Aktuell' : 'Current'}</th>
                      <th style={{ padding: '4px 8px' }}>Name</th>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Gruppe' : 'Group'}</th>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Angemeldet' : 'Registered'}</th>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Über Kapazität' : 'Over capacity'}</th>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Abstand zum letzten fairen Platz' : 'Gap to last fair seat'}</th>
                      <th style={{ padding: '4px 8px' }}>{isDe ? 'Fairer Platz' : 'Fair seat'}</th>
                      <th style={{ padding: '4px 8px', textAlign: 'right' }}>{isDe ? 'Aktion' : 'Action'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {flagged.map(reg => {
                      const nm = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                      const k = keyOf(reg);
                      const grpLabel = isSplitCapacity ? (groupOf(reg) || '—') : '—';
                      const cap = capOf(k);
                      const bucket = activeByGroup[k] || [];
                      const idx = bucket.findIndex(x => x.Id === reg.Id); // 0-basiert
                      const position = idx >= 0 ? idx + 1 : null;
                      const overBy = (position !== null && cap > 0) ? position - cap : null;
                      const cutoff = (cap > 0 && cap - 1 < bucket.length) ? bucket[cap - 1] : null;
                      const cutoffNm = cutoff ? ((cutoff.Vorname && cutoff.Nachname) ? `${cutoff.Vorname} ${cutoff.Nachname}` : cutoff.ParticipantName) : '';
                      const gapMs = cutoff ? (new Date(reg.RegistrationDate).getTime() - new Date(cutoff.RegistrationDate).getTime()) : NaN;
                      const wl = fairWaitByGroup[k] || [];
                      const wlRank = wl.findIndex(x => x.Id === reg.Id) + 1; // 1-basiert; 0 = nicht gefunden
                      const fairId = totalFairActive + (wlRank > 0 ? wlRank : (overBy || 0));
                      return (
                        <tr key={reg.Id} style={{ borderBottom: '1px solid rgba(237,139,0,0.25)' }}>
                          <td style={{ padding: '6px 8px', fontWeight: 600 }}>#{reg.TeilnehmerID ?? '—'}</td>
                          <td style={{ padding: '6px 8px' }}>
                            {nm}
                            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{reg.ParticipantEmail}</div>
                          </td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>{grpLabel}</td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>
                            {position !== null && cap > 0
                              ? (isDe
                                ? <>Platz <strong>{position}</strong> bei Kap. {cap} <span style={{ color: 'var(--dex-red, #c00)' }}>(+{overBy})</span></>
                                : <>Seat <strong>{position}</strong> at cap. {cap} <span style={{ color: 'var(--dex-red, #c00)' }}>(+{overBy})</span></>)
                              : '—'}
                          </td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>
                            {cutoff
                              ? <><strong>+{fmtGap(gapMs)}</strong><div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{isDe ? 'nach' : 'after'} {cutoffNm} ({formatDate(cutoff.RegistrationDate)})</div></>
                              : '—'}
                          </td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>
                            {wlRank > 0
                              ? (isDe
                                ? <>Warteliste-Platz <strong>{wlRank}</strong>{isSplitCapacity ? ` (${grpLabel})` : ''}<div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>= TeilnehmerID ~#{fairId} bei sauberer Liste</div></>
                                : <>Waitlist position <strong>{wlRank}</strong>{isSplitCapacity ? ` (${grpLabel})` : ''}<div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>= participant ID ~#{fairId} with a clean list</div></>)
                              : '—'}
                          </td>
                          <td style={{ padding: '6px 8px', textAlign: 'right', whiteSpace: 'nowrap' }}>
                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, marginRight: 10 }}>
                              <button
                                className="btn btn-secondary"
                                style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #c00)' }}
                                onClick={() => { setOverbookModal({ mode: 'confirm', targets: [reg] }); setObWithMail(true); setObRemoveCalendar(true); }}
                              >
                                {isDe ? 'Auf Warteliste' : 'To waitlist'}
                              </button>
                              <InfoTooltip placement="left" text={isDe
                                ? (
                                <>
                                  <strong>&bdquo;Auf Warteliste&ldquo;</strong> — die Person wird (gruppentreu) auf die <strong>Warteliste</strong> gesetzt; sie hatte fälschlich einen Platz.<br /><br />
                                  Im nächsten Dialog wählst du: <strong>mit oder ohne Entschuldigungs-Mail</strong> (Deloitte-Layout, geht in die Mail-Queue — nicht direkt versendet) und ob sie <strong>vom Kalendereintrag abgemeldet</strong> wird.<br /><br />
                                  Es wird ein <strong>Audit-Eintrag</strong> geschrieben (war fälschlich angemeldet, Original-Registrierung). Danach werden die <strong>TeilnehmerIDs automatisch neu vergeben</strong>.
                                </>
                                ) : (
                                <>
                                  <strong>&bdquo;To waitlist&ldquo;</strong> — the person is moved (group-faithful) to the <strong>waitlist</strong>; they had a seat by mistake.<br /><br />
                                  In the next dialog you choose: <strong>with or without an apology email</strong> (Deloitte layout, goes into the mail queue — not sent directly) and whether they are <strong>removed from the calendar entry</strong>.<br /><br />
                                  An <strong>audit entry</strong> is written (was registered by mistake, original registration). Afterwards the <strong>participant IDs are reassigned automatically</strong>.
                                </>
                                )
                              } />
                            </span>
                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                              <button
                                className="btn btn-secondary"
                                style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                onClick={() => { setOverbookModal({ mode: 'keep', targets: [reg] }); setObKeepVariant('firstWaitlist'); }}
                              >
                                {isDe ? 'Platz behalten' : 'Keep seat'}
                              </button>
                              <InfoTooltip placement="left" text={isDe
                                ? (
                                <>
                                  <strong>&bdquo;Platz behalten&ldquo;</strong> — die Person verliert den Platz <strong>nicht</strong>. Im nächsten Dialog wählst du:<br /><br />
                                  <strong>(a) Erste(r) auf der Warteliste</strong> der Gruppe — rückt beim nächsten frei werdenden Platz garantiert als Erste(r) nach.<br /><br />
                                  <strong>(b) Bleibt angemeldet</strong> — die Gruppe ist dann <strong>+1</strong> über Kapazität; der nächste frei werdende Platz wird <strong>einmal nicht</strong> nachgerückt, bis die Überzahl absorbiert ist.<br /><br />
                                  Beide Varianten mit <strong>Audit-Eintrag</strong>, danach IDs neu.
                                </>
                                ) : (
                                <>
                                  <strong>&bdquo;Keep seat&ldquo;</strong> — the person does <strong>not</strong> lose the seat. In the next dialog you choose:<br /><br />
                                  <strong>(a) First on the waitlist</strong> of the group — guaranteed to move up first when the next seat becomes free.<br /><br />
                                  <strong>(b) Stays registered</strong> — the group is then <strong>+1</strong> over capacity; the next freed seat is <strong>skipped once</strong> until the surplus is absorbed.<br /><br />
                                  Both variants with an <strong>audit entry</strong>, then IDs reassigned.
                                </>
                                )
                              } />
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          );
        })()}

        {(() => {
          // v11.84: Teams-Section — Admin-Center-Team-Management.
          // Sichtbar nur für Events mit aktivierter Team-Anmeldung. Listet
          // alle Teams (gruppiert per TeamId, abgemeldete Mitglieder
          // ausgeblendet), mit Lead-Badge und Buttons für „+ Person
          // hinzufügen" und „Lead-Rolle uebergeben". Reagiert live auf
          // `registrations` — kein zusätzlicher Roundtrip.
          if (!selectedEvent || !selectedEvent.teamRegistrationEnabled) return null;
          if (isLoadingRegs) return null;

          // groupBy TeamId, abgemeldete Personen NICHT eingehen lassen.
          const teamsByid: Record<string, SPRegistration[]> = {};
          // v16.2: Teilnehmer ohne Team in eine eigene Liste — werden
          // unten als „Teilnehmer ohne Team"-Sektion gerendert, damit
          // der Organizer sie sieht und ggf. einem (neuen) Team zuordnen
          // kann.
          const teamlessActive: SPRegistration[] = [];
          for (const r of registrations) {
            if (r.Status === 'Abgemeldet') continue;
            const tid = r.TeamId || '';
            if (!tid) {
              teamlessActive.push(r);
              continue;
            }
            (teamsByid[tid] = teamsByid[tid] || []).push(r);
          }
          // Sortierung: aelteste Lead-RegistrationDate zuerst.
          const teamEntries = Object.entries(teamsByid)
            .map(([tid, members]) => {
              // Lead oben, dann TeilnehmerID aufsteigend.
              members.sort((a, b) => {
                if (!!a.TeamLead !== !!b.TeamLead) return a.TeamLead ? -1 : 1;
                const aT = (a.TeilnehmerID ?? 9_999_999) as number;
                const bT = (b.TeilnehmerID ?? 9_999_999) as number;
                return aT - bT;
              });
              const lead = members.find(m => !!m.TeamLead) || members[0];
              const leadDate = lead?.RegistrationDate ? new Date(lead.RegistrationDate).getTime() : Number.MAX_SAFE_INTEGER;
              return { tid, members, lead, leadDate };
            })
            .sort((a, b) => a.leadDate - b.leadDate);

          const teamSizeCfg = selectedEvent.teamSize || 0;
          const count = teamEntries.length;
          const canManage = isAdmin || isOrganizerFor(selectedEvent);

          const statusBadge = (st: string): React.ReactElement | null => {
            if (!st || st === 'Angemeldet') return null;
            const colorMap: Record<string, string> = {
              'Warteliste': '#b35a00',
              'QR versendet': '#3a7dbf',
              'Eingecheckt': '#4a7c1f',
            };
            const color = colorMap[st] || 'var(--dex-gray-500)';
            return (
              <span style={{
                display: 'inline-block', padding: '1px 8px', borderRadius: 10,
                background: `${color}15`, color, fontSize: '0.7rem', fontWeight: 600, marginLeft: 6,
              }}>{st}</span>
            );
          };

          return (
            <div style={{ marginBottom: 20, padding: 16, borderRadius: 12, border: '1px solid var(--dex-gray-200)', background: '#fff' }}>
              <div
                role="button"
                tabIndex={0}
                onClick={() => setTeamsCollapsed(v => !v)}
                onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setTeamsCollapsed(v => !v); } }}
                style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer', userSelect: 'none' }}
              >
                <Users size={20} />
                <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1rem' }}>
                  {(selectedEvent?.teamTermPlural || 'Teams')} ({count})
                </strong>
                <span style={{ marginLeft: 'auto', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  {teamsCollapsed ? <ChevronDown size={18} /> : <ChevronUp size={18} />}
                </span>
              </div>
              {!teamsCollapsed && (
                <div style={{ marginTop: 14, display: 'flex', flexDirection: 'column', gap: 12 }}>
                  {/* v23.0: Drag&Drop-Hinweis. */}
                  {canManage && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.8rem', color: 'var(--dex-gray-600)', background: 'rgba(134,188,37,0.08)', border: '1px solid var(--dex-green, #86bc25)', borderRadius: 8, padding: '7px 12px' }}>
                      <Icon iconName="DragObject" style={{ fontSize: 15, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                      {isDe
                        ? `Tipp: Personen per Drag & Drop zwischen ${(selectedEvent?.teamTermPlural || 'Teams')} und „ohne ${(selectedEvent?.teamTermSingular || 'Team')}" verschieben.`
                        : `Tip: drag & drop people between ${(selectedEvent?.teamTermPlural || 'teams')} and “no ${(selectedEvent?.teamTermSingular || 'team')}”.`}
                    </div>
                  )}
                  {teamEntries.length === 0 && (
                    <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.88rem', fontStyle: 'italic' }}>
                      Keine Team-Anmeldungen bisher.
                    </div>
                  )}
                  {/* v16.2: „Neues Team anlegen"-Button + Teamless-Sektion.
                      v23.0: zusätzlich „Mail an <Teams>"-Button (Per-Team-Info-Mail). */}
                  {canManage && (
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.85rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                        onClick={() => {
                          // Neue lokale TeamID generieren und Add-Member-Dialog
                          // direkt damit oeffnen. Sobald die erste Person hinzu-
                          // gefügt wird, wird die TeamId im SP-Item gespeichert.
                          const newTid = (typeof crypto !== 'undefined' && crypto.randomUUID)
                            ? crypto.randomUUID()
                            : `team-${Date.now()}-${Math.floor(Math.random() * 1e9)}`;
                          setAdminAddMemberDialog({ teamId: newTid, teamName: '', freeSlots: teamSizeCfg || 99, isNewTeam: true });
                          setAdminAddMemberPick(null);
                          setAdminAddMemberQuery('');
                          setAdminAddMemberResults([]);
                          setAdminAddMemberConsent(false);
                          setAdminAddMemberError('');
                          setAdminAddTeamlessPicks(new Set());
                          setAdminAddLeadRegId(null);
                          setAdminAddSendMail(false);
                          setAdminAddCcOrganizer(false);
                          setAdminAddNotifyOthers(false);
                          setAdminAddNotifyScope('all');
                          setAdminAddNewPersonMail(true);
                        }}
                      >
                        <Plus size={14} /> {isDe ? `Neue ${selectedEvent?.teamTermSingular || 'Team'} anlegen` : `Create new ${selectedEvent?.teamTermSingular || 'team'}`}
                      </button>
                      {getActiveTeams().length > 0 && (
                        <button
                          type="button"
                          className="btn btn-secondary"
                          style={{ fontSize: '0.85rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                          onClick={openTeamMailDialog}
                          title={isDe ? 'Jedem Mitglied eine eigene Mail mit team-spezifischer Info senden (z.B. Teams-Einwahllink).' : 'Send each member an individual mail with team-specific info (e.g. a Teams join link).'}
                        >
                          <Icon iconName="Mail" style={{ fontSize: 14 }} /> {isDe ? `Mail an ${selectedEvent?.teamTermPlural || 'Teams'}` : `Mail to ${selectedEvent?.teamTermPlural || 'teams'}`}
                        </button>
                      )}
                    </div>
                  )}
                  {teamlessActive.length > 0 && (
                    <div
                      onDragOver={canManage ? (e => { e.preventDefault(); setDragOverTid(''); }) : undefined}
                      onDragLeave={canManage ? (() => setDragOverTid(prev => (prev === '' ? null : prev))) : undefined}
                      onDrop={canManage ? (() => onTeamDrop('', undefined)) : undefined}
                      style={{ padding: 14, border: dragOverTid === '' ? '2px dashed var(--dex-green, #86bc25)' : '1px dashed var(--dex-orange, #ed8b00)', borderRadius: 10, background: dragOverTid === '' ? 'rgba(134,188,37,0.10)' : 'rgba(237,139,0,0.04)' }}
                    >
                      <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, flexWrap: 'wrap', marginBottom: 10 }}>
                        <strong style={{ fontSize: '0.95rem', color: 'var(--dex-orange-dark, #b35a00)' }}>
                          {isDe ? `Teilnehmer ohne ${selectedEvent?.teamTermSingular || 'Team'}` : `Attendees without ${selectedEvent?.teamTermSingular || 'team'}`} ({teamlessActive.length})
                        </strong>
                        <span style={{ color: 'var(--dex-gray-600)', fontSize: '0.82rem' }}>
                          — Einzel-Anmeldungen ohne Team-Zuordnung
                        </span>
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                        {teamlessActive.map(m => {
                          const name = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const dept = (m as any).Department || '';
                          return (
                            <div
                              key={m.Id}
                              draggable={canManage}
                              onDragStart={canManage ? (() => setDragRegId(m.Id)) : undefined}
                              onDragEnd={canManage ? (() => { setDragRegId(null); setDragOverTid(null); }) : undefined}
                              title={canManage ? (isDe ? 'Ziehen, um zuzuordnen' : 'Drag to assign') : undefined}
                              style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '4px 6px', borderRadius: 6, cursor: canManage ? 'grab' : 'default', opacity: dragRegId === m.Id ? 0.4 : 1, background: dragRegId === m.Id ? 'var(--dex-gray-100)' : 'transparent' }}
                            >
                              <img
                                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=L`}
                                alt={name}
                                onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                              />
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontSize: '0.88rem', fontWeight: 500 }}>{name}{statusBadge(m.Status)}</div>
                                <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)' }}>{m.ParticipantEmail}</div>
                                {dept && <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', marginTop: 1 }}>{dept}</div>}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                  {/* v19.0: Teams in einem responsiven 3-Spalten-Raster +
                      durchnummeriert — spart vertikalen Platz im Organizer-Center. */}
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))', gap: 12, alignItems: 'stretch' }}>
                  {teamEntries.map(({ tid, members, lead }, teamIdx) => {
                    const teamName = members.find(m => !!m.TeamName)?.TeamName || '';
                    const total = members.length;
                    const free = teamSizeCfg > 0 ? Math.max(0, teamSizeCfg - total) : 0;
                    const canAdd = canManage && (teamSizeCfg === 0 || total < teamSizeCfg);
                    const leadEmail = lead?.ParticipantEmail || '';
                    const otherMembers = members.filter(m => m.Id !== lead?.Id);
                    // v19.19: Teams mit freien Plätzen farblich (orange) hervorheben,
                    // damit der Organizer auf einen Blick sieht, welche Teams noch
                    // nicht voll belegt sind.
                    const hasFreeSlots = free > 0;
                    const isDropTarget = dragOverTid === tid;
                    return (
                      <div
                        key={tid}
                        onDragOver={canManage ? (e => { e.preventDefault(); setDragOverTid(tid); }) : undefined}
                        onDragLeave={canManage ? (() => setDragOverTid(prev => (prev === tid ? null : prev))) : undefined}
                        onDrop={canManage ? (() => onTeamDrop(tid, teamName || undefined)) : undefined}
                        style={{
                          padding: 14,
                          border: isDropTarget ? '2px solid var(--dex-green, #86bc25)' : (hasFreeSlots ? '1px solid var(--dex-orange, #ed8b00)' : '1px solid var(--dex-gray-200)'),
                          borderRadius: 10,
                          background: isDropTarget ? 'rgba(134,188,37,0.12)' : (hasFreeSlots ? 'rgba(237,139,0,0.06)' : 'var(--dex-gray-50, #f7f7f7)'),
                          // v19.19: Flex-Spalte, damit der Aktions-Block (u.a.
                          // „Lead-Rolle übergeben") per marginTop:auto immer am
                          // unteren Kartenrand sitzt → alle Karten gleich hoch.
                          display: 'flex',
                          flexDirection: 'column',
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, flexWrap: 'wrap', marginBottom: 10 }}>
                          <strong style={{ fontSize: '0.95rem', color: 'var(--dex-gray-800)' }}>
                            <span style={{ color: 'var(--dex-gray-400)', marginRight: 4 }}>{teamIdx + 1}.</span>{teamName ? `Team „${teamName}"` : 'Team (ohne Namen)'}
                          </strong>
                          <span style={{ color: hasFreeSlots ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-600)', fontSize: '0.85rem', fontWeight: hasFreeSlots ? 600 : 400 }}>
                            {teamSizeCfg > 0 ? `${total}/${teamSizeCfg} belegt` : `${total} Mitglieder`}
                          </span>
                          {hasFreeSlots && (
                            <span style={{
                              display: 'inline-block', padding: '1px 8px', borderRadius: 10,
                              background: 'var(--dex-orange, #ed8b00)', color: '#fff',
                              fontSize: '0.7rem', fontWeight: 700,
                            }}>{free} frei</span>
                          )}
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                          {members.map(m => {
                            const name = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
                            const isLead = !!m.TeamLead;
                            return (
                              <div
                                key={m.Id}
                                draggable={canManage}
                                onDragStart={canManage ? (() => setDragRegId(m.Id)) : undefined}
                                onDragEnd={canManage ? (() => { setDragRegId(null); setDragOverTid(null); }) : undefined}
                                title={canManage ? (isDe ? 'Ziehen, um in ein anderes Team / „ohne Team" zu verschieben' : 'Drag to move to another team / “no team”') : undefined}
                                style={{
                                  display: 'flex', alignItems: 'center', gap: 10,
                                  padding: '4px 6px', borderRadius: 6,
                                  cursor: canManage ? 'grab' : 'default',
                                  opacity: dragRegId === m.Id ? 0.4 : 1,
                                  background: dragRegId === m.Id ? 'var(--dex-gray-100)' : 'transparent',
                                }}
                              >
                                <div style={{ position: 'relative', width: 32, height: 32, flexShrink: 0 }}>
                                  <img
                                    src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=L`}
                                    alt={name}
                                    onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                    style={{
                                      width: 32, height: 32, borderRadius: '50%', objectFit: 'cover',
                                      background: 'var(--dex-gray-100)',
                                      transition: 'transform 0.18s ease',
                                      transformOrigin: 'left center',
                                      cursor: 'pointer',
                                    }}
                                    onMouseEnter={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(2.4)'; (e.currentTarget as HTMLImageElement).style.zIndex = '10'; (e.currentTarget as HTMLImageElement).style.position = 'relative'; (e.currentTarget as HTMLImageElement).style.boxShadow = '0 8px 24px rgba(0,0,0,0.25)'; }}
                                    onMouseLeave={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(1)'; (e.currentTarget as HTMLImageElement).style.zIndex = ''; (e.currentTarget as HTMLImageElement).style.position = ''; (e.currentTarget as HTMLImageElement).style.boxShadow = ''; }}
                                  />
                                </div>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontSize: '0.88rem', fontWeight: 500 }}>
                                    {name}
                                    {statusBadge(m.Status)}
                                  </div>
                                  <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)' }}>{m.ParticipantEmail}</div>
                                  {/* v16.1: Business Area / Department aus
                                      der SP-Registrierung mit anzeigen,
                                      damit der Organizer auf einen Blick
                                      sieht, aus welcher Practice die
                                      Mitglieder kommen. */}
                                  {(() => {
                                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                    const dept = (m as any).Department || '';
                                    if (!dept) return null;
                                    return (
                                      <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', marginTop: 1 }}>{dept}</div>
                                    );
                                  })()}
                                </div>
                                {isLead && (
                                  <span style={{
                                    display: 'inline-block', padding: '2px 10px', borderRadius: 12,
                                    background: 'var(--dex-green, #86bc25)', color: '#fff',
                                    fontSize: '0.72rem', fontWeight: 700,
                                  }}>Lead</span>
                                )}
                                {/* v22.41/v22.45: „Aus Team entfernen" — löst NUR
                                    die Team-Zuordnung (TeamId/Lead/Name leeren),
                                    die Anmeldung inkl. Status (z.B. Warteliste)
                                    bleibt bestehen. Erscheint erst im „Anpassen"-
                                    Modus des Teams (nicht dauerhaft an jedem Namen). */}
                                {canManage && teamEditOpenFor === tid && eventServiceRef && selectedEvent.subsiteUrl && (
                                  <button
                                    type="button"
                                    title="Aus dem Team entfernen (Anmeldung bleibt bestehen)"
                                    onClick={async () => {
                                      const sub = selectedEvent.subsiteUrl;
                                      if (!sub) return;
                                      const stHint = m.Status && m.Status !== 'Angemeldet' ? ` (Status: ${m.Status})` : '';
                                      const ok = await confirmDialog(
                                        `${name} aus dem Team „${teamName || ''}" entfernen?\n\nDie Anmeldung${stHint} bleibt bestehen — die Person steht danach ohne Team da und kann einem anderen Team zugeordnet werden.`,
                                        { danger: true, confirmLabel: 'Aus Team entfernen' }
                                      );
                                      if (!ok) return;
                                      try {
                                        await eventServiceRef.assignRegistrationToTeam(sub, m.Id, '', '', false);
                                        // Lead entfernt + andere bleiben → frühestes Mitglied nachziehen.
                                        if (isLead) {
                                          const rest = members
                                            .filter(x => x.Id !== m.Id && x.Status !== 'Abgemeldet')
                                            .sort((a, b) => ((a.TeilnehmerID ?? 9_999_999) as number) - ((b.TeilnehmerID ?? 9_999_999) as number));
                                          if (rest.length > 0) {
                                            await eventServiceRef.assignRegistrationToTeam(sub, rest[0].Id, tid, teamName || undefined, true);
                                          }
                                        }
                                        await eventServiceRef.writeChangeLog({
                                          action: 'TeamMemberRemoved',
                                          targetType: 'Participant',
                                          targetId: m.ParticipantEmail,
                                          targetName: name,
                                          eventId: selectedEvent.id,
                                          eventTitle: selectedEvent.title,
                                          details: { teamId: tid, removedBy: currentUser.email, keptStatus: m.Status },
                                        }).catch(() => { /* */ });
                                        setTeamsToast(`${name} wurde aus dem Team entfernt — Anmeldung bleibt bestehen.`);
                                        window.setTimeout(() => setTeamsToast(''), 4500);
                                        const regs = await getAllRegistrations(selectedEvent.id);
                                        setRegistrations(regs);
                                      } catch (err) {
                                        console.warn('[DEX] removeFromTeam failed:', err);
                                        showAlert('Entfernen aus dem Team fehlgeschlagen.', { variant: 'error' });
                                      }
                                    }}
                                    style={{
                                      background: 'none', border: 'none', cursor: 'pointer',
                                      color: 'var(--dex-red, #c00)', fontSize: '0.72rem',
                                      textDecoration: 'underline', padding: '2px 4px', flexShrink: 0,
                                    }}
                                  >
                                    Entfernen
                                  </button>
                                )}
                              </div>
                            );
                          })}
                        </div>
                        {canManage && (
                          <div style={{ marginTop: 'auto', paddingTop: 12, display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 8, position: 'relative' }}>
                            {canAdd && (
                              <button
                                type="button"
                                className="btn btn-secondary"
                                style={{ fontSize: '0.8rem', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                                onClick={() => {
                                  setAdminAddMemberDialog({ teamId: tid, teamName, freeSlots: free });
                                  setAdminAddMemberPick(null);
                                  setAdminAddMemberQuery('');
                                  setAdminAddMemberResults([]);
                                  setAdminAddMemberConsent(false);
                                  setAdminAddMemberError('');
                                  setAdminAddTeamlessPicks(new Set());
                                  setAdminAddLeadRegId(null);
                                  setAdminAddSendMail(false);
                                  setAdminAddCcOrganizer(false);
                                  setAdminAddNotifyOthers(false);
                                  setAdminAddNotifyScope('all');
                                  setAdminAddNewPersonMail(true);
                                }}
                              >
                                <Plus size={14} /> Person hinzufügen
                                {teamSizeCfg > 0 && ` (${free} Slot${free === 1 ? '' : 's'} frei)`}
                              </button>
                            )}
                            {/* v22.45: „Anpassen" schaltet den Bearbeiten-Modus
                                des Teams ein/aus — erst dann erscheinen die
                                „Entfernen"-Buttons pro Mitglied. */}
                            {canManage && (
                              <button
                                type="button"
                                className={teamEditOpenFor === tid ? 'btn btn-primary' : 'btn btn-secondary'}
                                style={{ fontSize: '0.8rem', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                                onClick={() => setTeamEditOpenFor(teamEditOpenFor === tid ? null : tid)}
                              >
                                <Pencil size={14} /> {teamEditOpenFor === tid ? 'Fertig' : 'Anpassen'}
                              </button>
                            )}
                            {otherMembers.length > 0 && (
                              <>
                                <button
                                  type="button"
                                  className="btn btn-secondary"
                                  style={{ fontSize: '0.8rem', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                                  onClick={() => setLeadTransferOpenFor(leadTransferOpenFor === tid ? null : tid)}
                                >
                                  <RefreshCw size={14} /> Lead-Rolle übergeben
                                </button>
                                {leadTransferOpenFor === tid && (
                                  <div style={{
                                    position: 'absolute', top: '100%', left: 0, marginTop: 6,
                                    background: '#fff', border: '1px solid var(--dex-gray-300)',
                                    borderRadius: 8, boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
                                    zIndex: 20, minWidth: 280, maxWidth: 360, padding: 6,
                                  }}>
                                    <div style={{ padding: '6px 10px', fontSize: '0.78rem', color: 'var(--dex-gray-600)', borderBottom: '1px solid var(--dex-gray-100)' }}>
                                      Neue Lead-Rolle übertragen an:
                                    </div>
                                    {otherMembers.map(m => {
                                      const nm = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
                                      return (
                                        <button
                                          key={m.Id}
                                          type="button"
                                          disabled={leadTransferBusy}
                                          onClick={async () => {
                                            if (leadTransferBusy) return;
                                            setLeadTransferBusy(true);
                                            try {
                                              const res = await transferTeamLead(selectedEvent.id, tid, m.ParticipantEmail);
                                              if (res.ok) {
                                                setTeamsToast(`Lead-Rolle wurde an ${nm} übergeben.`);
                                                const regs = await getAllRegistrations(selectedEvent.id);
                                                setRegistrations(regs);
                                                window.setTimeout(() => setTeamsToast(''), 4500);
                                              } else {
                                                setTeamsToast(`Lead-Übergabe fehlgeschlagen: ${res.reason || 'Unbekannter Fehler'}.`);
                                                window.setTimeout(() => setTeamsToast(''), 4500);
                                              }
                                            } finally {
                                              setLeadTransferBusy(false);
                                              setLeadTransferOpenFor(null);
                                            }
                                          }}
                                          style={{
                                            display: 'flex', alignItems: 'center', gap: 8,
                                            width: '100%', padding: '8px 10px', border: 'none',
                                            background: 'transparent', cursor: 'pointer',
                                            textAlign: 'left', borderRadius: 6,
                                          }}
                                          onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.background = 'var(--dex-gray-100)'; }}
                                          onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.background = 'transparent'; }}
                                        >
                                          <img
                                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=S`}
                                            alt={nm}
                                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                            style={{ width: 24, height: 24, borderRadius: '50%' }}
                                          />
                                          <div style={{ minWidth: 0 }}>
                                            <div style={{ fontSize: '0.85rem', fontWeight: 500 }}>{nm}</div>
                                            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{m.ParticipantEmail}</div>
                                          </div>
                                        </button>
                                      );
                                    })}
                                    <button
                                      type="button"
                                      onClick={() => setLeadTransferOpenFor(null)}
                                      style={{
                                        width: '100%', padding: '6px 10px',
                                        border: 'none', borderTop: '1px solid var(--dex-gray-100)',
                                        background: 'transparent', cursor: 'pointer',
                                        fontSize: '0.78rem', color: 'var(--dex-gray-500)',
                                      }}
                                    >Abbrechen</button>
                                  </div>
                                )}
                              </>
                            )}
                            {/* leadEmail nur als Referenz für den Lead-Lookup behalten — nicht für's TS-Linting wegwerfen. */}
                            <span style={{ display: 'none' }}>{leadEmail}</span>
                          </div>
                        )}
                      </div>
                    );
                  })}
                  </div>
                </div>
              )}
            </div>
          );
        })()}

        {teamsToast && (
          <div style={{
            marginBottom: 14, padding: '10px 14px', borderRadius: 8,
            background: 'rgba(134,188,37,0.12)', border: '1px solid var(--dex-green, #86bc25)',
            color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.88rem',
          }}>
            {teamsToast}
          </div>
        )}

        {/* v23.0: Per-Team-Info-Mail. Jedes aktive Mitglied bekommt eine eigene
            Mail; pro Team trägt der Organizer eine team-spezifische Info ein
            (z.B. einen eigenen Teams-Einwahllink). */}
        {teamMailOpen && selectedEvent && (
          <Modal
            open={true}
            onClose={() => { if (!teamMailSending) setTeamMailOpen(false); }}
            dismissable={!teamMailSending}
            maxWidth={720}
            padding={24}
            ariaLabel={isDe ? 'Mail an Teams' : 'Mail to teams'}
          >
            <h3 style={{ marginTop: 0, marginBottom: 6, color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {isDe ? `Mail an ${selectedEvent.teamTermPlural || 'Teams'}` : `Mail to ${selectedEvent.teamTermPlural || 'teams'}`}
            </h3>
            <p style={{ marginTop: 0, fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
              {isDe
                ? 'Jedes aktive Mitglied bekommt eine eigene Mail. Pro Gruppe trägst du unten eine eigene Info ein (z.B. einen Teams-Einwahllink) — sie ersetzt den Platzhalter {{TeamInfo}}. Verfügbare Platzhalter: {{Vorname}}, {{Name}}, {{TeamName}}, {{EventTitle}}, {{TeamInfo}}.'
                : 'Each active member receives an individual mail. Per group you enter its own info below (e.g. a Teams join link) — it replaces the {{TeamInfo}} placeholder. Available placeholders: {{Vorname}}, {{Name}}, {{TeamName}}, {{EventTitle}}, {{TeamInfo}}.'}
            </p>

            <label style={{ display: 'block', fontWeight: 600, fontSize: '0.85rem', marginTop: 10 }}>{isDe ? 'Betreff' : 'Subject'}</label>
            <input
              type="text"
              value={teamMailSubject}
              onChange={e => setTeamMailSubject(e.target.value)}
              disabled={teamMailSending}
              style={{ width: '100%', padding: '8px 10px', borderRadius: 8, border: '1px solid var(--dex-gray-300)', fontSize: '0.9rem', boxSizing: 'border-box' }}
            />

            <label style={{ display: 'block', fontWeight: 600, fontSize: '0.85rem', marginTop: 12 }}>{isDe ? 'Mail-Text (HTML erlaubt)' : 'Mail body (HTML allowed)'}</label>
            <textarea
              value={teamMailBody}
              onChange={e => setTeamMailBody(e.target.value)}
              disabled={teamMailSending}
              rows={8}
              style={{ width: '100%', padding: '8px 10px', borderRadius: 8, border: '1px solid var(--dex-gray-300)', fontSize: '0.85rem', fontFamily: 'monospace', boxSizing: 'border-box', resize: 'vertical' }}
            />

            <div style={{ marginTop: 14, fontWeight: 600, fontSize: '0.85rem' }}>
              {isDe ? 'Info pro Gruppe' : 'Info per group'}
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 8, maxHeight: 280, overflowY: 'auto' }}>
              {getActiveTeams().map(t => {
                const tName = t.teamName || (selectedEvent.teamTermSingular || 'Team');
                return (
                  <div key={t.tid} style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: '10px 12px', background: 'var(--dex-gray-50, #f7f7f7)' }}>
                    <div style={{ fontWeight: 600, fontSize: '0.86rem', marginBottom: 6 }}>
                      {tName} <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}>· {t.members.length} {isDe ? 'Mitglieder' : 'members'}</span>
                    </div>
                    <textarea
                      value={teamMailInfoByTid[t.tid] || ''}
                      onChange={e => setTeamMailInfoByTid(prev => ({ ...prev, [t.tid]: e.target.value }))}
                      disabled={teamMailSending}
                      rows={3}
                      placeholder={isDe ? 'z.B. https://teams.microsoft.com/l/meetup-join/…' : 'e.g. https://teams.microsoft.com/l/meetup-join/…'}
                      style={{ width: '100%', padding: '7px 9px', borderRadius: 6, border: '1px solid var(--dex-gray-300)', fontSize: '0.83rem', boxSizing: 'border-box', resize: 'vertical' }}
                    />
                  </div>
                );
              })}
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 18 }}>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => setTeamMailOpen(false)}
                disabled={teamMailSending}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => { sendTeamMails().catch(() => { /* */ }); }}
                disabled={teamMailSending || !teamMailSubject.trim() || !teamMailBody.trim()}
              >
                {teamMailSending
                  ? (isDe ? 'Wird gesendet…' : 'Sending…')
                  : (isDe ? 'Mails senden' : 'Send mails')}
              </button>
            </div>
          </Modal>
        )}

        {regLoadError ? (
          <p style={{ color: 'var(--dex-red)', fontStyle: 'italic' }}>{regLoadError}</p>
        ) : isLoadingRegs ? (
          <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{isDe ? 'Lade Teilnehmer...' : 'Loading participants...'}</p>
        ) : isConsolidatedMode ? (
          // v14.11: konsolidierter Matrix-View für Events im
          // „Nur Sub-Events"-Modus. Eine Zeile pro eindeutigem Teilnehmer,
          // X-Spalten pro Sub-Event, plus Event-Level- (Pastel A) und
          // Sub-Event-Level- (Pastel B) Custom-Field-Spalten gruppiert.
          renderConsolidatedView()
        ) : activeRegs.length === 0 ? (
          <p style={{ color: 'var(--dex-gray-400)' }}>{isDe ? 'Noch keine Teilnehmer registriert.' : 'No participants registered yet.'}</p>
        ) : (
          /* v17.13: overflowX: 'auto' entfernt — der scrollbare Wrapper
             hat die sticky-thead-Berechnung gebrochen (sticky relative zum
             Scroll-Container statt zum Window). Tabelle lässt die Karte
             jetzt horizontal ueberlaufen, was bei vielen Spalten zu einer
             Scrollbar AM AUSSEREN Container (SP-Page) führt — Sticky-
             thead funktioniert dort einwandfrei. */
          <div style={{ overflowX: 'visible' }}>
            {(() => {
              // v6.17: Spaltenkonfiguration — Header und Body-Zellen werden dynamisch
              // anhand `columnOrder` (+ `hiddenColumns`) gerendert. So kann der User
              // Spalten ein-/ausblenden und per Pfeilen umsortieren. Die Render-Logik
              // selbst (Sort-Buttons, Badges, Custom-Field-Anzeige etc.) bleibt gleich,
              // nur die Iteration ist umgebaut.
              const visibleColumnIds = columnOrder.filter(id => hiddenColumns.indexOf(id) < 0);

              const sortableCols: Record<string, 'id' | 'anrede' | 'vorname' | 'nachname' | 'email' | 'status' | 'date'> = {
                id: 'id', anrede: 'anrede', vorname: 'vorname', nachname: 'nachname', email: 'email', status: 'status', date: 'date',
              };

              const hideButton = (id: string): React.ReactNode => {
                const col = availableColumns.find(c => c.id === id);
                if (!col || col.alwaysVisible) return null;
                return (
                  <button
                    type="button"
                    onClick={(e) => { e.stopPropagation(); hideColumn(id); }}
                    aria-label={isDe ? `Spalte ${col.label} ausblenden` : `Hide column ${col.label}`}
                    title={isDe ? 'Spalte ausblenden' : 'Hide column'}
                    style={{
                      marginLeft: 6, padding: 0, width: 16, height: 16, lineHeight: '14px',
                      border: 'none', background: 'transparent', cursor: 'pointer',
                      color: 'var(--dex-gray-400)', fontSize: '0.8rem', borderRadius: 3,
                    }}
                    onMouseEnter={(e) => { e.currentTarget.style.color = 'var(--dex-red, #c00)'; }}
                    onMouseLeave={(e) => { e.currentTarget.style.color = 'var(--dex-gray-400)'; }}
                  >
                    ✕
                  </button>
                );
              };

              const renderHeader = (id: string): React.ReactNode => {
                // v15.3: lange Spalten-Überschriften (Custom-Field-Labels wie
                // „Please check if you have marked all parts you want to attend
                // and confirm") brechen jetzt um statt mit Ellipsis abgeschnitten
                // zu werden. Begrenzte maxWidth + Wortumbruch — der Header bleibt
                // lesbar ohne dass der User hovern muss.
                // v15.4.1: wordBreak:'break-word' war zu aggressiv (Edge
                // brach kurze Wörter wie „Vorname" → „Vorna\nme"). Jetzt
                // overflowWrap:'break-word' — Umbruch nur an Wort-Grenzen
                // oder wenn ein einzelnes Wort breiter als die Spalte ist.
                const baseStyle: React.CSSProperties = {
                  textAlign: 'left', padding: 8,
                  whiteSpace: 'normal',
                  overflowWrap: 'break-word',
                  hyphens: 'auto',
                  maxWidth: 180,
                  verticalAlign: 'top',
                  lineHeight: 1.3,
                  // v17.10: Sticky-Header — beim Scrollen bleiben die
                  // Spaltenüberschriften der Teilnehmer-Tabelle sichtbar.
                  position: 'sticky',
                  top: 0,
                  background: '#fff',
                  zIndex: 5,
                  borderBottom: '2px solid var(--dex-gray-200)',
                };
                const sortable = sortableCols[id];
                if (sortable) {
                  return (
                    <th
                      key={id}
                      style={{ ...baseStyle, cursor: 'pointer', userSelect: 'none' }}
                      onClick={() => handleSort(sortable)}
                    >
                      {id === 'id' ? '#' : id === 'anrede' ? (isDe ? 'Anrede' : 'Salutation') : id === 'vorname' ? (isDe ? 'Vorname' : 'First name') : id === 'nachname' ? (isDe ? 'Nachname' : 'Last name') : id === 'email' ? 'Email' : id === 'status' ? 'Status' : (isDe ? 'Registriert am' : 'Registered on')}
                      {sortIcon(sortable)}
                      {hideButton(id)}
                    </th>
                  );
                }
                if (id === 'jobTitle') return <th key={id} style={baseStyle}>Job Title{hideButton(id)}</th>;
                if (id === 'location') return <th key={id} style={baseStyle}>{isDe ? 'Standort' : 'Location'}{hideButton(id)}</th>;
                if (id === 'starterType') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? "Starter-Typ: Durchstarter oder Funstarter. Wird bei der Anmeldung gewählt und steuert die Split-Kapazität + Warteliste. Der eigentliche Startblock steht in der Custom-Field-Spalte 'Start block'." : "Starter type: Durchstarter or Funstarter. Chosen at registration and controls the split capacity + waitlist. The actual start block is in the custom field column 'Start block'."}>
                      {isDe ? 'Starter-Typ' : 'Starter type'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'promotedDate') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Zeitpunkt des Nachrückens — gesetzt sobald der Teilnehmer von der Warteliste in den Aktiv-Bereich promotet wurde. Leer für Personen die sich direkt angemeldet haben.' : 'Time of promotion — set as soon as the participant was promoted from the waitlist into the active area. Empty for people who registered directly.'}>
                      {isDe ? 'Nachgerückt am' : 'Promoted on'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'replaced') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Die abgemeldete Person, deren Platz diese Person übernommen hat. Nur gesetzt für nachgerückte Personen.' : 'The cancelled person whose seat this person took. Only set for promoted people.'}>
                      {isDe ? 'Hat ersetzt' : 'Replaced'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'replacedBy') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Wer nach der Abmeldung dieses Teilnehmers den Platz übernommen hat. Nur gesetzt für abgemeldete Personen, deren Cancel einen Promote ausgelöst hat.' : 'Who took the seat after this participant cancelled. Only set for cancelled people whose cancellation triggered a promotion.'}>
                      {isDe ? 'Ersetzt durch' : 'Replaced by'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'registeredBy') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Selbst = der Teilnehmer hat sich selbst registriert. Ansonsten Name des Users, der die Registrierung durchgeführt hat.' : 'Self = the participant registered themselves. Otherwise the name of the user who performed the registration.'}>
                      {isDe ? 'Registriert von' : 'Registered by'}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'team') {
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Team-Name des Teilnehmers (falls Team-Anmeldung aktiv).' : 'Team name of the participant (if team registration is active).'}>
                      Team{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'roommate') {
                  // v11.56: Label dynamisch aus availableColumns nehmen (entstammt dem
                  // ersten roommate-/user-Feld der Custom-Field-Definition) statt
                  // hartcodiertem „Zimmerpartner".
                  const roommateCol = availableColumns.find(c => c.id === 'roommate');
                  const roommateLabel = roommateCol?.label || 'Zimmerpartner';
                  return (
                    <th key={id} style={baseStyle} title={isDe ? 'Ausgewählter User-Picker-Wert aus diesem Feld. Match = beide haben sich gegenseitig ausgewählt.' : 'Selected user-picker value from this field. Match = both selected each other.'}>
                      {roommateLabel}{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'action') {
                  return <th key={id} style={{ textAlign: 'left', padding: 8 }}>Aktion</th>;
                }
                // v14.11: pastel A = event-level (parent) fields, pastel B = sub-event-specific fields.
                // Pastel-Hintergrund nur im Sub-Event-Detail-View (parentEventForSelected gesetzt),
                // sonst neutraler Hintergrund wie bisher.
                const inSubEventDetail = !!parentEventForSelected;
                const pastelAHeader: React.CSSProperties = inSubEventDetail ? { background: 'rgba(0, 118, 168, 0.15)' } : {};
                const pastelBHeader: React.CSSProperties = inSubEventDetail ? { background: 'rgba(255, 191, 0, 0.18)' } : {};
                if (id.indexOf('cfp-') === 0) {
                  const cfId = id.substring(4);
                  const field = (parentEventForSelected?.eventSpecificFields || []).find(f => f.id === cfId);
                  if (!field) return null;
                  const label = field.label || '';
                  return (
                    <th key={id} style={{ ...baseStyle, fontSize: '0.78rem', ...pastelAHeader }} title={`${label} — ${isDe ? 'Hauptevent-Feld' : 'main-event field'}`}>
                      {label}
                      {hideButton(id)}
                    </th>
                  );
                }
                if (id.indexOf('cf-') === 0) {
                  const cfId = id.substring(3);
                  const field = (selectedEvent?.eventSpecificFields || []).find(f => f.id === cfId);
                  if (!field) return null;
                  const label = field.label || '';
                  return (
                    <th key={id} style={{ ...baseStyle, fontSize: '0.78rem', ...pastelBHeader }} title={inSubEventDetail ? `${label} — ${isDe ? 'Sub-Event-Feld' : 'sub-event field'}` : label}>
                      {label}
                      {hideButton(id)}
                    </th>
                  );
                }
                return null;
              };

              const renderCell = (id: string, reg: SPRegistration, i: number): React.ReactNode => {
                if (id === 'id') {
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-400)' }}>{reg.TeilnehmerID || (i + 1)}</td>;
                }
                if (id === 'anrede') {
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{reg.Anrede || '-'}</td>;
                }
                if (id === 'vorname') {
                  // Fallback für Alt-Daten: erstes Wort aus ParticipantName.
                  const v = reg.Vorname || ((reg.ParticipantName || '').split(' ')[0] || '');
                  return <td key={id} style={{ padding: 8, fontWeight: 500 }}>{v || '-'}</td>;
                }
                if (id === 'nachname') {
                  // Fallback für Alt-Daten: alles ausser dem ersten Wort als Nachname.
                  let n = reg.Nachname || '';
                  if (!n && reg.ParticipantName) {
                    const parts = reg.ParticipantName.trim().split(/\s+/);
                    if (parts.length > 1) n = parts.slice(1).join(' ');
                  }
                  return <td key={id} style={{ padding: 8, fontWeight: 500 }}>{n || '-'}</td>;
                }
                if (id === 'email') {
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>;
                }
                if (id === 'jobTitle') {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{(reg as any).JobTitle || '-'}</td>;
                }
                if (id === 'location') {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{(reg as any).Location || '-'}</td>;
                }
                if (id === 'starterType') {
                  return (
                    <td key={id} style={{ padding: 8, fontSize: '0.8rem' }}>
                      {(() => {
                        // Tatsächlicher Startblock (StarterType) + Wunsch (PreferredStarterType).
                        // Wenn beide identisch: nur einen anzeigen. Wenn unterschiedlich (z.B. per
                        // Fallback-Dialog auf anderen Typ umgestiegen): Wunsch in Klammern daneben.
                        const actual = reg.StarterType || '';
                        const pref = reg.PreferredStarterType || '';
                        if (!actual && !pref) return <span style={{ color: 'var(--dex-gray-400)' }}>—</span>;
                        if (actual && pref && actual !== pref) {
                          return <span>{actual} <span style={{ color: 'var(--dex-gray-500)' }}>(Wunsch: {pref})</span></span>;
                        }
                        if (actual) return <span>{actual}</span>;
                        // v19.12: StarterType ist leer. Bei AKTIVEN (angemeldeten/
                        // eingecheckten) Personen ist die effektive Gruppe der Wunsch
                        // — der Nachrück-Flow hat den StarterType beim Promoten nur
                        // nicht gesetzt. Solche Personen NEHMEN ihren Wunsch-Platz ein,
                        // also plain die Gruppe zeigen (NICHT „Wunsch:"). „Wunsch:"
                        // bleibt den Warteliste-Personen vorbehalten (dort ist die
                        // Gruppe wirklich noch nicht zugewiesen).
                        const isWaitlist = reg.Status === 'Warteliste';
                        return <span>{isWaitlist ? `Wunsch: ${pref}` : pref}</span>;
                      })()}
                    </td>
                  );
                }
                if (id === 'status') {
                  return (
                    <td key={id} style={{ padding: 8 }}>
                      <span className={`badge ${reg.Status === 'Eingecheckt' ? 'badge-green' : 'badge-gray'}`}>
                        {translateStatus(reg.Status, isDe)}
                      </span>
                    </td>
                  );
                }
                if (id === 'date') {
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>;
                }
                if (id === 'promotedDate') {
                  // v17.15: „Nachgerückt am" — gesetzt beim Promote
                  // von Warteliste → Angemeldet. Leer für Personen die
                  // sich direkt in den Aktiv-Bereich angemeldet haben.
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const v = (reg as any).PromotedDate as string | undefined;
                  return <td key={id} style={{ padding: 8, color: v ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-300)', fontSize: '0.8rem' }}>{v ? formatDate(v) : '—'}</td>;
                }
                if (id === 'replaced') {
                  // v17.15: „Ersetzt" — die Person, deren Cancel diesen
                  // Promote ausgelöst hat. Wenn die Person in den
                  // aktuellen registrations gefunden wird, zeigen wir den
                  // Namen — sonst fallback auf die rohe E-Mail.
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const email = ((reg as any).ReplacedParticipantEmail as string | undefined) || '';
                  if (!email) return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-300)' }}>—</td>;
                  const other = registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === email.toLowerCase());
                  const label = other ? ((other.Vorname || '') + ' ' + (other.Nachname || '')).trim() || other.ParticipantName || email : email;
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-700)', fontSize: '0.8rem' }} title={email}>{label}</td>;
                }
                if (id === 'replacedBy') {
                  // v17.15: „Ersetzt durch" — die Person die nach Cancel
                  // dieses Eintrags den Platz uebernommen hat. Spiegelbild
                  // von „Ersetzt".
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const email = ((reg as any).ReplacedByParticipantEmail as string | undefined) || '';
                  if (!email) return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-300)' }}>—</td>;
                  const other = registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === email.toLowerCase());
                  const label = other ? ((other.Vorname || '') + ' ' + (other.Nachname || '')).trim() || other.ParticipantName || email : email;
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.8rem' }} title={email}>{label}</td>;
                }
                if (id === 'joinOrder') {
                  // v17.9 (deprecated): joinOrder-Spalte seit v17.10 entfernt.
                  return <td key={id} style={{ padding: 8 }}>—</td>;
                }
                if (id === 'registeredBy') {
                  return (
                    <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>
                      {(() => {
                        const actorEmail = (reg.RegisteredByEmail || '').toLowerCase();
                        const participantEmail = (reg.ParticipantEmail || '').toLowerCase();
                        if (!actorEmail) return <span style={{ color: 'var(--dex-gray-400)' }}>-</span>;
                        if (actorEmail === participantEmail) {
                          return <span style={{ color: 'var(--dex-green-dark)' }}>Selbst</span>;
                        }
                        return (
                          <span title={reg.RegisteredByEmail || ''} style={{ color: 'var(--dex-orange)' }}>
                            {reg.RegisteredByName || reg.RegisteredByEmail}
                          </span>
                        );
                      })()}
                    </td>
                  );
                }
                if (id === 'team') {
                  // v16.1: Team-Name + Lead-Markierung. Wenn der TN in
                  // keinem Team ist, „—" anzeigen.
                  const tName = (reg.TeamName || '').trim();
                  const inTeam = !!reg.TeamId;
                  if (!inTeam) return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-400)' }}>—</td>;
                  return (
                    <td key={id} style={{ padding: 8, color: 'var(--dex-gray-700)', fontSize: '0.82rem' }}>
                      {tName ? `„${tName}"` : <span style={{ color: 'var(--dex-gray-500)' }}>ohne Namen</span>}
                      {reg.TeamLead && (
                        <span style={{ marginLeft: 6, padding: '1px 7px', background: 'var(--dex-green, #86bc25)', color: '#fff', borderRadius: 8, fontSize: '0.66rem', fontWeight: 700 }}>Lead</span>
                      )}
                    </td>
                  );
                }
                if (id === 'roommate') {
                  return (
                    <td key={id} style={{ padding: 8, fontSize: '0.8rem' }}>
                      {(() => {
                        const info = getRoommateInfo(reg);
                        if (!info) return <span style={{ color: 'var(--dex-gray-300)' }}>-</span>;
                        const photoEmail = (info.partnerEmail || '').trim();
                        return (
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                            {photoEmail && (
                              <img
                                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(photoEmail)}&size=S`}
                                alt={info.partnerName}
                                onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                style={{ width: 24, height: 24, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                              />
                            )}
                            <span>{info.partnerName}</span>
                            {info.mutual && (
                              <span
                                className="badge"
                                style={{ marginLeft: 2, background: 'var(--dex-green)', color: '#fff', padding: '1px 6px', borderRadius: 4, fontSize: '0.7rem' }}
                                title={isDe ? 'Beide haben sich gegenseitig als Zimmerpartner ausgewählt' : 'Both selected each other as roommates'}
                              >
                                Match
                              </span>
                            )}
                          </span>
                        );
                      })()}
                    </td>
                  );
                }
                // v14.11: cfp-* sind Parent-Event-Custom-Fields (Pastel A) im
                // Sub-Event-Detail-View. Wert kommt entweder aus reg.CustomData
                // (Sub-Events erben i.d.R. die Parent-Felder via Wizard-Copy)
                // oder, falls leer, aus dem SP-Internal-Name-Property.
                const inSubEventDetailCell = !!parentEventForSelected;
                const pastelACell: React.CSSProperties = inSubEventDetailCell ? { background: 'rgba(0, 118, 168, 0.08)' } : {};
                const pastelBCell: React.CSSProperties = inSubEventDetailCell ? { background: 'rgba(255, 191, 0, 0.10)' } : {};
                if (id.indexOf('cfp-') === 0) {
                  const cfId = id.substring(4);
                  const field = (parentEventForSelected?.eventSpecificFields || []).find(f => f.id === cfId);
                  if (!field) return null;
                  // v15.14: Werte für Parent-Custom-Fields kommen primär aus
                  // der Parent-Event-Registrierung der Person (lookup per
                  // ParticipantEmail in parentRegsByEmail) — die Sub-Event-
                  // Registrierung enthält diese Antworten i.d.R. nicht. Nur
                  // wenn keine Parent-Reg existiert, fallen wir auf die Sub-
                  // Event-Daten zurück.
                  const emailKey = (reg.ParticipantEmail || '').toLowerCase().trim();
                  const parentReg = emailKey ? parentRegsByEmail[emailKey] : undefined;
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const spName = (field as any).spInternalName || '';
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  let val: any = undefined;
                  if (parentReg) {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    val = spName ? (parentReg as any)[spName] : undefined;
                    if ((val === undefined || val === null || val === '') && parentReg.CustomData) {
                      try {
                        const cd = JSON.parse(parentReg.CustomData);
                        val = cd[field.id];
                      } catch { /* no-op */ }
                    }
                  }
                  if (val === undefined || val === null || val === '') {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    val = spName ? (reg as any)[spName] : undefined;
                    if ((val === undefined || val === null || val === '') && reg.CustomData) {
                      try {
                        const cd = JSON.parse(reg.CustomData);
                        val = cd[field.id];
                      } catch { /* no-op */ }
                    }
                  }
                  let display: React.ReactNode = '-';
                  if (val !== undefined && val !== null && val !== '') {
                    if (field.type === 'checkbox') {
                      const truthy = val === true || val === 'true' || val === 1 || val === '1';
                      display = <span style={{ color: truthy ? 'var(--dex-green-dark)' : 'var(--dex-gray-400)' }}>{truthy ? '✓' : '–'}</span>;
                    } else if (field.type === 'select' && field.multi) {
                      display = String(val).split(' | ').map(s => s.trim()).filter(Boolean).join(', ');
                    } else {
                      display = String(val);
                    }
                  }
                  return (
                    <td key={id} style={{ padding: 8, color: 'var(--dex-gray-700)', fontSize: '0.8rem', whiteSpace: 'nowrap', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', ...pastelACell }} title={String(val || '')}>
                      {display}
                    </td>
                  );
                }
                if (id.indexOf('cf-') === 0) {
                  const cfId = id.substring(3);
                  const field = (selectedEvent?.eventSpecificFields || []).find(f => f.id === cfId);
                  if (!field) return null;
                  // v19.2: Dokument-Felder haben keinen Spaltenwert — die Datei
                  // liegt als Attachment. In der Spalte einen Download-Link (oder
                  // mehrere) zeigen, statt „-".
                  if (field.type === 'document') {
                    const att = attachmentsByReg[reg.Id] || [];
                    const prefix = `dxf-${(field.id || '').replace(/[^a-zA-Z0-9]/g, '')}--`;
                    const docs = att.filter(a => a.fileName.startsWith(prefix));
                    const pretty = (fn: string): string => fn
                      .replace(/^dxf-[a-zA-Z0-9]+--\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}_/, '')
                      .replace(/^dxf-[a-zA-Z0-9]+--/, '');
                    return (
                      <td key={id} style={{ padding: 8, fontSize: '0.8rem', whiteSpace: 'nowrap', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', ...pastelBCell }}>
                        {docs.length === 0 ? (
                          <span style={{ color: 'var(--dex-gray-400)' }}>–</span>
                        ) : (
                          docs.map((d, i) => (
                            <a
                              key={d.fileName}
                              href={d.serverRelativeUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              title={pretty(d.fileName)}
                              style={{ color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'underline', marginRight: i < docs.length - 1 ? 8 : 0, display: 'inline-flex', alignItems: 'center', gap: 3 }}
                            >
                              <FileText size={12} />{docs.length > 1 ? `${isDe ? 'Datei' : 'File'} ${i + 1}` : (isDe ? 'Datei' : 'File')}
                            </a>
                          ))
                        )}
                      </td>
                    );
                  }
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const spName = (field as any).spInternalName || '';
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  let val: any = spName ? (reg as any)[spName] : undefined;
                  if ((val === undefined || val === null || val === '') && reg.CustomData) {
                    try {
                      const cd = JSON.parse(reg.CustomData);
                      val = cd[field.id];
                    } catch { /* no-op */ }
                  }
                  let display: React.ReactNode = '-';
                  if (val !== undefined && val !== null && val !== '') {
                    if (field.type === 'checkbox') {
                      const truthy = val === true || val === 'true' || val === 1 || val === '1';
                      display = <span style={{ color: truthy ? 'var(--dex-green-dark)' : 'var(--dex-gray-400)' }}>{truthy ? '✓' : '–'}</span>;
                    } else if (field.type === 'select' && field.multi) {
                      // v7.11: Mehrfachauswahl wird " | "-getrennt gespeichert.
                      // In der Admin-Tabelle als Komma-Liste anzeigen, damit
                      // der Spalten-Inhalt sauberer scanbar ist.
                      display = String(val).split(' | ').map(s => s.trim()).filter(Boolean).join(', ');
                    } else {
                      display = String(val);
                    }
                  }
                  return (
                    <td key={id} style={{ padding: 8, color: 'var(--dex-gray-700)', fontSize: '0.8rem', whiteSpace: 'nowrap', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', ...pastelBCell }} title={String(val || '')}>
                      {display}
                    </td>
                  );
                }
                if (id === 'action') {
                  const att = attachmentsByReg[reg.Id] || [];
                  return (
                    <td key={id} style={{ padding: 8, display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                      <button
                        className="btn btn-secondary"
                        style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                        title={isDe ? 'Teilnehmer-Daten bearbeiten' : 'Edit attendee data'}
                        onClick={() => openEditModal(reg)}
                      >
                        <Pencil size={12} /> {isDe ? 'Bearbeiten' : 'Edit'}
                      </button>
                      {/* v11.0: Anhang-Button — wenn das Event den Teilnehmer-
                          Upload erlaubt ODER ein Dokument-Custom-Feld hat (v19.0).
                          Zeigt Counter wenn mind. eine Datei hochgeladen wurde. */}
                      {(selectedEvent?.allowAttendeeUpload || (selectedEvent?.eventSpecificFields || []).some(f => f.type === 'document')) && (
                        <button
                          className="btn btn-secondary"
                          style={{ fontSize: '0.75rem', padding: '4px 10px', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                          title={isDe ? 'Hochgeladene Dateien anzeigen' : 'Show uploaded files'}
                          onClick={() => setAttachmentsModalReg(reg)}
                        >
                          <FileText size={12} />
                          {att.length > 0 ? `${isDe ? 'Datei' : 'File'} (${att.length})` : (isDe ? 'Datei' : 'File')}
                        </button>
                      )}
                      {reg.Status === 'Eingecheckt' ? (
                        <button
                          className="btn btn-secondary"
                          style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                          onClick={async () => {
                            if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                            await eventServiceRef.checkOutParticipant(selectedEvent.subsiteUrl, reg.Id);
                            const regs = await getAllRegistrations(selectedEvent.id);
                            setRegistrations(regs);
                          }}
                        >
                          {isDe ? 'Auschecken' : 'Check out'}
                        </button>
                      ) : (
                        <button
                          className="btn btn-primary"
                          style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                          onClick={async () => {
                            if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                            await eventServiceRef.checkInParticipant(selectedEvent.subsiteUrl, reg.Id);
                            const regs = await getAllRegistrations(selectedEvent.id);
                            setRegistrations(regs);
                          }}
                        >
                          {isDe ? 'Einchecken' : 'Check in'}
                        </button>
                      )}
                      <button
                        className="btn btn-secondary"
                        style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #c00)' }}
                        onClick={async () => {
                          if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                          // v23.2: Doppel-Anmeldung? Statt direkt abzumelden das
                          // Duplikat-Modal öffnen (still löschen vs. normal abmelden).
                          if (duplicateEmails.has((reg.ParticipantEmail || '').trim().toLowerCase())) { setDupCancelReg(reg); return; }
                          const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                          // v22.22: Vergangenes Event → stille Abmeldung (keine
                          // Abmelde-Mail, keine Outlook-Absage, kein Nachrücken,
                          // kein ID-Reorder). Der Confirm sagt das explizit.
                          const eventWasOver = isEventOver(selectedEvent);
                          const confirmMsg = eventWasOver
                            ? (isDe
                              ? `${name} (${reg.ParticipantEmail}) wirklich abmelden?\n\nDas Event liegt in der Vergangenheit — die Abmeldung läuft still: Es gehen keine Abmelde-Mail und keine Outlook-Absage raus, und es rückt niemand von der Warteliste nach.`
                              : `Really cancel ${name} (${reg.ParticipantEmail})?\n\nThe event is in the past — the cancellation runs silently: no cancellation email, no Outlook removal, and nobody is promoted from the waitlist.`)
                            : (isDe ? `${name} (${reg.ParticipantEmail}) wirklich abmelden?` : `Really cancel ${name} (${reg.ParticipantEmail})?`);
                          if (!(await confirmDialog(confirmMsg, { danger: true, confirmLabel: isDe ? 'Abmelden' : 'Cancel registration' }))) return;
                          await performStandardCancel(reg);
                        }}
                      >
                        {isDe ? 'Abmelden' : 'Cancel'}
                      </button>
                    </td>
                  );
                }
                return null;
              };

              return (
                <>
                  {/* v6.17: Kontrollzeile mit Column-Picker-Button. Der Popover
                      zeigt alle verfügbaren Spalten inkl. Checkbox zum Ein-/
                      Ausblenden und Pfeilen zum Umsortieren. Die Config wird
                      pro Event in localStorage persistiert (s. useEffect oben). */}
                  <div style={{ display: 'flex', justifyContent: 'flex-end', alignItems: 'center', marginBottom: 8, gap: 8, position: 'relative' }}>
                    <button
                      className="btn btn-secondary"
                      style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                      onClick={() => setShowColumnPicker(!showColumnPicker)}
                    >
                      {isDe ? 'Spalten anpassen' : 'Customize columns'}
                    </button>
                    {showColumnPicker && (
                      <div
                        style={{
                          position: 'absolute', right: 0, top: '100%', marginTop: 4,
                          background: '#fff', border: '1px solid var(--dex-gray-200)',
                          borderRadius: 8, padding: 12,
                          boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                          width: 280, zIndex: 100, maxHeight: 400, overflowY: 'auto',
                        }}
                      >
                        <div style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 8 }}>
                          {isDe ? 'Spalten verwalten' : 'Manage columns'}
                        </div>
                        {columnOrder.map((id, idx) => {
                          const col = availableColumns.find(c => c.id === id);
                          if (!col) return null;
                          const isHidden = hiddenColumns.indexOf(id) >= 0;
                          const isVisible = !isHidden;
                          const canMoveUp = isVisible && idx > 0 && columnOrder[idx - 1] !== undefined;
                          // "action" bleibt immer letzte → niemand darf unter "action" wandern
                          // und "action" selbst darf nicht verschoben werden.
                          const nextId = columnOrder[idx + 1];
                          const canMoveDown = isVisible && idx < columnOrder.length - 1 && id !== 'action' && nextId !== 'action';
                          return (
                            <div
                              key={id}
                              style={{
                                display: 'flex', alignItems: 'center', gap: 6,
                                padding: '4px 2px', fontSize: '0.82rem',
                                opacity: isVisible ? 1 : 0.55,
                              }}
                            >
                              <input
                                type="checkbox"
                                checked={isVisible}
                                disabled={!!col.alwaysVisible}
                                onChange={() => {
                                  if (col.alwaysVisible) return;
                                  if (isHidden) showColumn(id); else hideColumn(id);
                                }}
                                style={{ cursor: col.alwaysVisible ? 'not-allowed' : 'pointer' }}
                                title={col.alwaysVisible ? 'Pflicht-Spalte — kann nicht ausgeblendet werden' : (isHidden ? 'Einblenden' : 'Ausblenden')}
                              />
                              <span style={{ flex: 1, color: 'var(--dex-gray-700)' }}>{col.label}</span>
                              <button
                                type="button"
                                onClick={() => moveColumn(id, -1)}
                                disabled={!canMoveUp}
                                aria-label={isDe ? 'Spalte nach oben' : 'Move column up'}
                                title={isDe ? 'Nach oben' : 'Up'}
                                style={{
                                  border: 'none', background: 'transparent',
                                  cursor: canMoveUp ? 'pointer' : 'not-allowed',
                                  color: canMoveUp ? 'var(--dex-gray-600)' : 'var(--dex-gray-300)',
                                  fontSize: '0.9rem', padding: '0 4px',
                                }}
                              >
                                ↑
                              </button>
                              <button
                                type="button"
                                onClick={() => moveColumn(id, 1)}
                                disabled={!canMoveDown}
                                aria-label={isDe ? 'Spalte nach unten' : 'Move column down'}
                                title={isDe ? 'Nach unten' : 'Down'}
                                style={{
                                  border: 'none', background: 'transparent',
                                  cursor: canMoveDown ? 'pointer' : 'not-allowed',
                                  color: canMoveDown ? 'var(--dex-gray-600)' : 'var(--dex-gray-300)',
                                  fontSize: '0.9rem', padding: '0 4px',
                                }}
                              >
                                ↓
                              </button>
                            </div>
                          );
                        })}
                        <div style={{ marginTop: 8, textAlign: 'right' }}>
                          <button
                            className="btn btn-secondary"
                            style={{ fontSize: '0.72rem', padding: '3px 8px' }}
                            onClick={() => setShowColumnPicker(false)}
                          >
                            {isDe ? 'Schließen' : 'Close'}
                          </button>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* v11.98: Split-/Merged-Toggle bei Split-Kapazität.
                      Default 'split' — getrennte Tabellen pro Gruppe,
                      kleinere zuerst. */}
                  {(() => {
                    const renderTable = (rows: SPRegistration[], indexOffset: number): React.ReactElement => (
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                        <thead>
                          <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                            {visibleColumnIds.map(id => renderHeader(id))}
                          </tr>
                        </thead>
                        <tbody>
                          {rows.map((reg, i) => {
                            const isOverbook = reg.OverbookReview === 'Pending';
                            // v22.44: Inaktive Deloitte-Konten dauerhaft orange
                            // markieren (bis zur Abmeldung) — gleiche Optik wie
                            // die Überbuchungs-Markierung. inactiveAccounts kommt
                            // aus dem Konten-Aktiv-Check (nur @deloitte-Adressen).
                            const isInactiveAcct = inactiveAccounts.indexOf((reg.ParticipantEmail || '').trim().toLowerCase()) >= 0;
                            // v23.2: Doppel-Anmeldung — rote Markierung (hat Vorrang
                            // vor der orangen Überbuchungs-/Inaktiv-Markierung).
                            const isDuplicate = (reg.Status || '') !== 'Abgemeldet'
                              && duplicateEmails.has((reg.ParticipantEmail || '').trim().toLowerCase());
                            const highlight = isOverbook || isInactiveAcct;
                            const rowTitle = isDuplicate
                              ? (isDe ? 'Doppel-Anmeldung — diese Person ist mehrfach angemeldet. Über „Abmelden" lässt sich die doppelte Zeile still entfernen.' : 'Duplicate registration — this person is registered more than once. Use „Cancel" to silently remove the duplicate row.')
                              : isInactiveAcct
                              ? (isDe ? 'Kein aktives Deloitte-Konto gefunden — Person hat womöglich Deloitte verlassen. Mails/Outlook kommen ggf. nicht an.' : 'No active Deloitte account found — person may have left Deloitte. Emails/Outlook may not arrive.')
                              : isOverbook
                                ? (isDe ? 'Über Kapazität angemeldet — siehe Box „Überbuchung – zu prüfen" oben' : 'Registered over capacity — see the „Overbooking – to review" box above')
                                : undefined;
                            return (
                              <tr
                                key={reg.Id}
                                title={rowTitle}
                                style={{
                                  borderBottom: '1px solid var(--dex-gray-100)',
                                  ...(isDuplicate
                                    ? { background: 'rgba(200,0,0,0.10)', boxShadow: 'inset 3px 0 0 var(--dex-red, #c00)' }
                                    : highlight
                                    ? { background: 'rgba(237,139,0,0.13)', boxShadow: 'inset 3px 0 0 var(--dex-orange, #ed8b00)' }
                                    : {}),
                                }}
                              >
                                {visibleColumnIds.map(id => renderCell(id, reg, indexOffset + i))}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    );

                    if (!isSplitCapacity || splitParticipantsView === 'merged') {
                      return (
                        <>
                          {isSplitCapacity && (
                            <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 8 }}>
                              <SplitMergeToggle view={splitParticipantsView} setView={setSplitParticipantsView} isDe={isDe} />
                            </div>
                          )}
                          {renderTable(activeRegs, 0)}
                        </>
                      );
                    }

                    // Split-View: nach Gruppe trennen (StarterType ||
                    // PreferredStarterType), kleinere Gruppe zuerst.
                    const lblA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
                    const lblB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
                    const groupA = activeRegs.filter(r => (r.StarterType || r.PreferredStarterType) === 'Durchstarter');
                    const groupB = activeRegs.filter(r => (r.StarterType || r.PreferredStarterType) === 'Funstarter');
                    const groupNone = activeRegs.filter(r => !(r.StarterType || r.PreferredStarterType));
                    const groups = [
                      { label: lblA, key: 'A', rows: groupA, cap: selectedEvent?.durchstarterCapacity || 0 },
                      { label: lblB, key: 'B', rows: groupB, cap: selectedEvent?.funstarterCapacity || 0 },
                    ].sort((x, y) => x.rows.length - y.rows.length);
                    let runningIdx = 0;
                    return (
                      <>
                        <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 8 }}>
                          <SplitMergeToggle view={splitParticipantsView} setView={setSplitParticipantsView} isDe={isDe} />
                        </div>
                        {groups.map(g => {
                          const offset = runningIdx;
                          runningIdx += g.rows.length;
                          return (
                            <div key={g.key} style={{ marginBottom: 20 }}>
                              <h4 style={{
                                margin: '0 0 8px', color: 'var(--dex-green-dark, #4a7c1f)',
                                fontSize: '0.95rem', fontWeight: 700, display: 'flex',
                                alignItems: 'baseline', gap: 8,
                              }}>
                                <span>{g.label}</span>
                                <span style={{ color: 'var(--dex-gray-500)', fontWeight: 500, fontSize: '0.85rem' }}>
                                  ({g.rows.length}{g.cap > 0 ? ` / ${g.cap}` : ''})
                                </span>
                              </h4>
                              {g.rows.length === 0 ? (
                                <p style={{ margin: '0 0 12px', fontSize: '0.82rem', color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
                                  {isDe ? 'Keine Teilnehmer in dieser Gruppe.' : 'No participants in this group.'}
                                </p>
                              ) : renderTable(g.rows, offset)}
                            </div>
                          );
                        })}
                        {groupNone.length > 0 && (
                          <div style={{ marginBottom: 20 }}>
                            <h4 style={{ margin: '0 0 8px', color: 'var(--dex-gray-500)', fontSize: '0.95rem', fontWeight: 700 }}>
                              {isDe ? 'Ohne Gruppe' : 'No group'} <span style={{ color: 'var(--dex-gray-400)', fontWeight: 500, fontSize: '0.85rem' }}>({groupNone.length})</span>
                            </h4>
                            {renderTable(groupNone, runningIdx)}
                          </div>
                        )}
                      </>
                    );
                  })()}
                </>
              );
            })()}
          </div>
        )}

        {/* v17.8: Anker für Floating-Jump-Button „Zur Warteliste". */}
        <div id="admin-waitlist-anchor" style={{ scrollMarginTop: 80 }} />
        {(() => {
          // Seit v6.5: bei B2Run-Split-Kapazitäten getrennte Wartelisten-Tabellen pro
          // PreferredStarterType. Ohne Split: eine einzige Warteliste wie bisher.
          const renderWaitlistTable = (title: string, regs: SPRegistration[], accentColor: string): React.ReactElement | null => {
            if (regs.length === 0) return null;
            // v17.8: Sortierung pro Spalte. Default 'pos' = TeilnehmerID asc
            // (FIFO-Position der Warteliste — wie vorher).
            const sortedRegs = (() => {
              const arr = regs.slice();
              const dir = waitlistSortAsc ? 1 : -1;
              const safe = (s: string | undefined): string => (s || '').toLowerCase();
              const dateMs = (s: string | undefined): number => s ? new Date(s).getTime() : Number.POSITIVE_INFINITY;
              arr.sort((a, b) => {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const anyA = a as any; const anyB = b as any;
                switch (waitlistSortColumn) {
                  case 'pos': return ((a.TeilnehmerID || 0) - (b.TeilnehmerID || 0)) * dir;
                  case 'vorname': return safe(a.Vorname).localeCompare(safe(b.Vorname), 'de') * dir;
                  case 'nachname': return safe(a.Nachname).localeCompare(safe(b.Nachname), 'de') * dir;
                  case 'email': return safe(a.ParticipantEmail).localeCompare(safe(b.ParticipantEmail)) * dir;
                  case 'jobtitle': return safe(anyA.JobTitle).localeCompare(safe(anyB.JobTitle), 'de') * dir;
                  case 'location': return safe(anyA.Location).localeCompare(safe(anyB.Location), 'de') * dir;
                  case 'date': return (dateMs(a.RegistrationDate) - dateMs(b.RegistrationDate)) * dir;
                }
                return 0;
              });
              return arr;
            })();
            const arrow = (k: typeof waitlistSortColumn): string => k === waitlistSortColumn ? (waitlistSortAsc ? ' ▲' : ' ▼') : '';
            const toggleSort = (k: typeof waitlistSortColumn): void => {
              if (waitlistSortColumn === k) setWaitlistSortAsc(v => !v);
              else { setWaitlistSortColumn(k); setWaitlistSortAsc(true); }
            };
            const thClickable: React.CSSProperties = { textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none', position: 'sticky', top: 0, background: '#fff', zIndex: 5, borderBottom: '2px solid var(--dex-gray-200)' };
            return (
              <React.Fragment key={title}>
                <h4 style={{ marginTop: 24, color: accentColor }}>{title} ({regs.length})</h4>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                    <thead>
                      <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                        <th style={thClickable} onClick={() => toggleSort('pos')}>Platz{arrow('pos')}</th>
                        <th style={thClickable} onClick={() => toggleSort('vorname')}>Vorname{arrow('vorname')}</th>
                        <th style={thClickable} onClick={() => toggleSort('nachname')}>Nachname{arrow('nachname')}</th>
                        <th style={thClickable} onClick={() => toggleSort('email')}>Email{arrow('email')}</th>
                        <th style={thClickable} onClick={() => toggleSort('jobtitle')}>Job Title{arrow('jobtitle')}</th>
                        <th style={thClickable} onClick={() => toggleSort('location')}>Standort{arrow('location')}</th>
                        {isSplitCapacity && <th style={{ textAlign: 'left', padding: 8 }}>Wunsch-Typ</th>}
                        <th style={thClickable} onClick={() => toggleSort('date')}>Registriert am{arrow('date')}</th>
                        <th style={{ textAlign: 'left', padding: 8 }}>Aktion</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sortedRegs.map((reg, i) => {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const anyReg = reg as any;
                        // v17.8: Position bleibt die FIFO-Position basierend auf der ORIGINAL-
                        // Reihenfolge (TeilnehmerID asc), unabhängig von der aktuellen Sortierung.
                        // Wenn der User nach Nachname sortiert, soll trotzdem klar sein, wer
                        // Platz 1 / 2 / 3 ist.
                        const fifoIdx = regs
                          .slice()
                          .sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0))
                          .findIndex(r => r.Id === reg.Id);
                        const pos = fifoIdx >= 0 ? fifoIdx + 1 : i + 1;
                        return (
                        <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                          <td style={{ padding: 8, fontWeight: 600, color: accentColor }}>{pos}</td>
                          <td style={{ padding: 8, fontWeight: 500 }}>{reg.Vorname || '-'}</td>
                          <td style={{ padding: 8, fontWeight: 500 }}>{reg.Nachname || '-'}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{anyReg.JobTitle || '-'}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{anyReg.Location || '-'}</td>
                          {isSplitCapacity && (
                            <td style={{ padding: 8, color: 'var(--dex-gray-700)' }}>
                              {reg.PreferredStarterType || '—'}
                            </td>
                          )}
                          <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                          <td style={{ padding: 8 }}>
                            <button
                              className="btn btn-secondary"
                              style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #c00)' }}
                              onClick={async () => {
                                if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                                const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                                // v22.22: Vergangenes Event → stilles Entfernen
                                // (keine Abmelde-Mail, kein ID-Reorder).
                                const eventWasOver = isEventOver(selectedEvent);
                                if (!(await confirmDialog(`${name} von der Warteliste entfernen?${eventWasOver ? (isDe ? '\n\nDas Event liegt in der Vergangenheit — es geht keine Abmelde-Mail raus.' : '\n\nThe event is in the past — no cancellation email will be sent.') : ''}`, { danger: true, confirmLabel: isDe ? 'Entfernen' : 'Remove' }))) return;
                                await eventServiceRef.cancelRegistration(selectedEvent.subsiteUrl, reg.Id, `${currentUser.firstName} ${currentUser.surname}`.trim(), currentUser.email);
                                if (reg.ParticipantEmail && !selectedEvent.disableEmails && !selectedEvent.disableCancellationEmail && !eventWasOver) {
                                  const emailData = cancellationEmail(name, selectedEvent.title);
                                  eventServiceRef.queueEmail(
                                    emailData.subject, reg.ParticipantEmail, name, emailData.body,
                                    'Abmeldung', selectedEvent.title, selectedEvent.id
                                  ).catch(err => console.warn('[DEX]', err));
                                }
                                if (reg.ParticipantEmail && selectedEvent.eventNumber) {
                                  eventServiceRef.removeParticipantEvent(reg.ParticipantEmail, selectedEvent.eventNumber).catch(err => console.warn('[DEX]', err));
                                }
                                if (selectedEvent.subsiteUrl && !eventWasOver) {
                                  try {
                                    const ok = await eventServiceRef.queueIDReorder(
                                      selectedEvent.id, selectedEvent.eventNumber || 0,
                                      selectedEvent.subsiteUrl, selectedEvent.title,
                                      `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantName || undefined,
                                      reg.ParticipantEmail || undefined
                                    );
                                    if (!ok) {
                                      showAlert(isDe ? 'Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.' : 'Cancellation successful, but the ID reorder entry could not be written to the queue. Please click "Reassign IDs" once.');
                                    }
                                  } catch {
                                    showAlert(isDe ? 'Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.' : 'Cancellation successful, but the ID reorder entry could not be written to the queue. Please click "Reassign IDs" once.');
                                  }
                                }
                                const allRegs = await getAllRegistrations(selectedEvent.id);
                                setRegistrations(allRegs);
                              }}
                            >
                              Entfernen
                            </button>
                          </td>
                        </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </React.Fragment>
            );
          };

          if (isSplitCapacity) {
            // v11.6: Wartelisten-Tabellen mit den frei wählbaren Gruppen-
            // Labels statt hartcodeten 'Durchstarter'/'Funstarter'.
            const wlLabelA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
            const wlLabelB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
            // v11.29: Reihenfolge respektiert splitDisplayOrderReversed
            // (gleicher Toggle wie auf Register-Page + Kapazitäts-Cards).
            const wlA = renderWaitlistTable(`Warteliste ${wlLabelA}`, waitlistDurch, 'var(--dex-green-dark, #6b9a1e)');
            const wlB = renderWaitlistTable(`Warteliste ${wlLabelB}`, waitlistFun, 'var(--dex-orange, #ff8c00)');
            const reversed = !!selectedEvent?.splitDisplayOrderReversed;
            return (
              <>
                {reversed ? <>{wlB}{wlA}</> : <>{wlA}{wlB}</>}
                {renderWaitlistTable('Warteliste ohne Gruppe', waitlistUnassigned, 'var(--dex-gray-500)')}
              </>
            );
          }
          return renderWaitlistTable('Warteliste', waitlistRegs, 'var(--dex-orange)');
        })()}

        {cancelledRegs.length > 0 && (() => {
          // v18.11: Abmeldungs-Liste mit denselben Spalten + Sortierung wie
          // Teilnehmer-/Warteliste. Unterscheidet proaktive Absagen
          // (CustomData _declined = „Ich nehme nicht teil", ohne vorherige
          // Anmeldung) von regulären Abmeldungen.
          const isDeclined = (reg: SPRegistration): boolean => {
            try { return !!(JSON.parse(reg.CustomData || '{}')._declined); } catch { return false; }
          };
          const safe = (s: string | undefined): string => (s || '').toLowerCase();
          const dateMs = (s: string | undefined): number => s ? new Date(s).getTime() : 0;
          const dir = cancelledSortAsc ? 1 : -1;
          const sorted = cancelledRegs.slice().sort((a, b) => {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const anyA = a as any; const anyB = b as any;
            switch (cancelledSortColumn) {
              case 'vorname': return safe(a.Vorname).localeCompare(safe(b.Vorname), 'de') * dir;
              case 'nachname': return safe(a.Nachname).localeCompare(safe(b.Nachname), 'de') * dir;
              case 'email': return safe(a.ParticipantEmail).localeCompare(safe(b.ParticipantEmail)) * dir;
              case 'jobtitle': return safe(anyA.JobTitle).localeCompare(safe(anyB.JobTitle), 'de') * dir;
              case 'location': return safe(anyA.Location).localeCompare(safe(anyB.Location), 'de') * dir;
              case 'type': return ((isDeclined(a) ? 1 : 0) - (isDeclined(b) ? 1 : 0)) * dir;
              case 'date': return (dateMs(a.CancellationDate) - dateMs(b.CancellationDate)) * dir;
            }
            return 0;
          });
          const arrow = (k: typeof cancelledSortColumn): string => k === cancelledSortColumn ? (cancelledSortAsc ? ' ▲' : ' ▼') : '';
          const toggleSort = (k: typeof cancelledSortColumn): void => {
            if (cancelledSortColumn === k) setCancelledSortAsc(v => !v);
            else { setCancelledSortColumn(k); setCancelledSortAsc(true); }
          };
          const thClickable: React.CSSProperties = { textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none', position: 'sticky', top: 0, background: '#fff', zIndex: 5, borderBottom: '2px solid var(--dex-gray-200)' };
          const declineCount = cancelledRegs.filter(isDeclined).length;
          // v19.28: Abgemeldete Registrierungen endgültig löschen (Admin/Organizer)
          // — z.B. um Test-Anmeldungen aus der Übersicht zu entfernen. Hartes
          // DELETE nach Sicherheits-Confirm; Audit-Eintrag im ChangeLog.
          const canDelete = !!selectedEvent && (isAdmin || isOrganizerFor(selectedEvent)) && !!selectedEvent.subsiteUrl;
          const deleteCancelled = async (reg: SPRegistration): Promise<void> => {
            if (!selectedEvent) return;
            // v22.59: im Klammer-Modus die Subsite der jeweiligen Sub-Section
            // nutzen (die Zeile trägt sie mit), sonst die Klammer-Subsite.
            const targetSubsite = (reg as SPRegistration & { _subsiteUrl?: string })._subsiteUrl || selectedEvent.subsiteUrl;
            if (!targetSubsite) return;
            const nm = `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantName || reg.ParticipantEmail;
            const msg = isDe
              ? `Diese abgemeldete Registrierung von „${nm}" ENDGÜLTIG löschen?\n\nDie Zeile wird komplett aus der Teilnehmerliste entfernt und kann NICHT wiederhergestellt werden. (Nützlich z.B. zum Aufräumen von Test-Anmeldungen.)`
              : `Permanently DELETE this cancelled registration of „${nm}"?\n\nThe row is removed entirely from the participant list and CANNOT be restored. (Useful e.g. for cleaning up test registrations.)`;
            if (!(await confirmDialog(msg, { danger: true, title: isDe ? 'Registrierung löschen' : 'Delete registration', confirmLabel: isDe ? 'Endgültig löschen' : 'Delete permanently' }))) return;
            const ok = await eventServiceRef.deleteRegistration(targetSubsite, reg.Id);
            if (ok) {
              try {
                await eventServiceRef.writeChangeLog({
                  action: 'RegistrationDeleted',
                  targetType: 'Participant',
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  targetId: ((reg as any).ParticipantEmail || '') + '#' + reg.Id,
                  targetName: nm,
                  eventId: selectedEvent.id,
                  eventTitle: selectedEvent.title,
                  details: { deletedStatus: reg.Status, cancellationDate: reg.CancellationDate || '' },
                });
              } catch { /* Audit best-effort */ }
              // v22.59/v22.63: im Klammer-Modus sowohl die Sub-Event-Listen
              // (Sub-Section-Abmeldungen) ALS AUCH die Klammer-Registrierungen
              // (z.B. Absagen auf der Klammer) neu laden, sonst die Event-Regs.
              if (isConsolidatedMode) {
                setSubRegReloadTick(t => t + 1);
                const regs = await getAllRegistrations(selectedEvent.id);
                setRegistrations(regs);
              } else {
                const regs = await getAllRegistrations(selectedEvent.id);
                setRegistrations(regs);
              }
            } else {
              // eslint-disable-next-line no-alert
              showAlert(isDe ? 'Löschen fehlgeschlagen.' : 'Delete failed.');
            }
          };
          // v22.63: Konsolidierte Abmelde-Matrix — EINE Zeile pro Person, mit
          // einem ✗ je Section (Gesamt-Event + Sub-Events), in der sich die
          // Person abgemeldet hat. Analog zur konsolidierten Anmelde-Matrix.
          if (isConsolidatedMode) {
            const sectionCols: Array<{ id: string; title: string }> = [
              ...(cancelledRegs.some(r => r._sectionId === '__parent') ? [{ id: '__parent', title: isDe ? 'Gesamt-Event' : 'Overall event' }] : []),
              ...consolidatedChildren
                .filter(ch => cancelledRegs.some(r => r._sectionId === ch.id))
                .map(ch => ({ id: ch.id, title: shortSubEventTitle(ch.title, selectedEvent!.title) })),
            ];
            type CancelRow = SPRegistration & { _subsiteUrl?: string; _sectionId?: string };
            interface CancelPerson { email: string; firstName: string; lastName: string; jobTitle: string; location: string; latest: number; declinedAny: boolean; bySection: Record<string, CancelRow> }
            const peopleMap = new Map<string, CancelPerson>();
            for (const r of cancelledRegs) {
              const key = (r.ParticipantEmail || '').toLowerCase().trim();
              if (!key) continue;
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const anyR = r as any;
              let p = peopleMap.get(key);
              if (!p) { p = { email: r.ParticipantEmail || '', firstName: '', lastName: '', jobTitle: '', location: '', latest: 0, declinedAny: false, bySection: {} }; peopleMap.set(key, p); }
              if (!p.firstName && r.Vorname) p.firstName = r.Vorname;
              if (!p.lastName && r.Nachname) p.lastName = r.Nachname;
              if (!p.jobTitle && anyR.JobTitle) p.jobTitle = anyR.JobTitle;
              if (!p.location && anyR.Location) p.location = anyR.Location;
              p.bySection[r._sectionId || '__parent'] = r;
              const tms = r.CancellationDate ? new Date(r.CancellationDate).getTime() : 0;
              if (tms > p.latest) p.latest = tms;
              if (isDeclined(r)) p.declinedAny = true;
            }
            const pdir = cancelledSortAsc ? 1 : -1;
            const people = Array.from(peopleMap.values()).sort((a, b) => {
              switch (cancelledSortColumn) {
                case 'nachname': return a.lastName.toLowerCase().localeCompare(b.lastName.toLowerCase(), 'de') * pdir;
                case 'email': return a.email.toLowerCase().localeCompare(b.email.toLowerCase()) * pdir;
                case 'date': return (a.latest - b.latest) * pdir;
                default: return a.firstName.toLowerCase().localeCompare(b.firstName.toLowerCase(), 'de') * pdir;
              }
            });
            const declinePeople = people.filter(p => p.declinedAny).length;
            const deletePerson = async (p: CancelPerson): Promise<void> => {
              if (!selectedEvent) return;
              const emailLc = p.email.toLowerCase().trim();
              // v22.66: Person ÜBERALL löschen — nicht nur die Abmelde-Zeilen,
              // sondern ALLE Zeilen dieser E-Mail über die Klammer UND alle
              // Sub-Events (inkl. der aktiven „Schatten"-Zeile auf der Klammer,
              // die beim reinen Abmelden/Löschen sonst verwaist liegen bleibt).
              const targets: Array<{ sub: string; id: number }> = [];
              if (selectedEvent.subsiteUrl) {
                for (const r of registrations) {
                  if ((r.ParticipantEmail || '').toLowerCase().trim() === emailLc) targets.push({ sub: selectedEvent.subsiteUrl, id: r.Id });
                }
              }
              for (const c of consolidatedChildren) {
                if (!c.subsiteUrl) continue;
                for (const r of (subEventRegsByEventId[c.id] || [])) {
                  if ((r.ParticipantEmail || '').toLowerCase().trim() === emailLc) targets.push({ sub: c.subsiteUrl, id: r.Id });
                }
              }
              const nm = `${p.firstName} ${p.lastName}`.trim() || p.email;
              const msg = isDe
                ? `„${nm}" wirklich überall löschen? Alle ${targets.length} Einträge dieser Person (Gesamt-Event + Sub-Events) werden endgültig entfernt und können nicht wiederhergestellt werden.`
                : `Permanently delete „${nm}" everywhere? All ${targets.length} entries of this person (overall event + sub-events) will be removed and cannot be restored.`;
              if (!(await confirmDialog(msg, { danger: true, title: isDe ? 'Person überall löschen' : 'Delete person everywhere', confirmLabel: isDe ? 'Endgültig löschen' : 'Delete permanently' }))) return;
              for (const t of targets) {
                try { await eventServiceRef.deleteRegistration(t.sub, t.id); } catch { /* best-effort */ }
              }
              try {
                await eventServiceRef.writeChangeLog({ action: 'RegistrationDeleted', targetType: 'Participant', targetId: p.email, targetName: nm, eventId: selectedEvent.id, eventTitle: selectedEvent.title, details: { deletedStatus: 'Abgemeldet', count: targets.length, everywhere: true } });
              } catch { /* */ }
              setSubRegReloadTick(t => t + 1);
              const regs = await getAllRegistrations(selectedEvent.id);
              setRegistrations(regs);
            };
            return (
              <>
                <h4 style={{ marginTop: 24, color: 'var(--dex-gray-400)' }}>
                  {isDe ? 'Abmeldungen' : 'Cancellations'} ({people.length})
                  {declinePeople > 0 && (
                    <span style={{ fontSize: '0.8rem', fontWeight: 400, marginLeft: 8, color: 'var(--dex-gray-500)' }}>
                      {isDe ? `davon ${declinePeople} Absage(n) ohne Anmeldung` : `incl. ${declinePeople} decline(s) without registration`}
                    </span>
                  )}
                </h4>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                    <thead>
                      <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                        <th style={thClickable} onClick={() => toggleSort('vorname')}>Vorname{arrow('vorname')}</th>
                        <th style={thClickable} onClick={() => toggleSort('nachname')}>Nachname{arrow('nachname')}</th>
                        <th style={thClickable} onClick={() => toggleSort('email')}>Email{arrow('email')}</th>
                        <th style={{ ...thClickable, cursor: 'default' }}>Job Title</th>
                        <th style={{ ...thClickable, cursor: 'default' }}>Standort</th>
                        {sectionCols.map(sc => (
                          <th key={sc.id} style={{ ...thClickable, cursor: 'default', textAlign: 'center' }} title={sc.title}>{sc.title}</th>
                        ))}
                        <th style={thClickable} onClick={() => toggleSort('date')}>{isDe ? 'Letzte Abmeldung' : 'Last cancellation'}{arrow('date')}</th>
                        {canDelete && (
                          <th style={{ ...thClickable, cursor: 'default', textAlign: 'right' }}>{isDe ? 'Löschen' : 'Delete'}</th>
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {people.map(p => (
                        <tr key={p.email} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                          <td style={{ padding: 8, fontWeight: 500 }}>{p.firstName || '-'}</td>
                          <td style={{ padding: 8, fontWeight: 500 }}>{p.lastName || '-'}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{p.email}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{p.jobTitle || '-'}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{p.location || '-'}</td>
                          {sectionCols.map(sc => {
                            const r = p.bySection[sc.id];
                            // v23.2: In der „Gesamt-Event"-Spalte kein nacktes X,
                            // sondern ein sprechender Badge — die Person hat erklärt,
                            // dass sie nicht (am Gesamt-Event) teilnehmen wird.
                            if (sc.id === '__parent') {
                              return (
                                <td key={sc.id} style={{ padding: 8, textAlign: 'center' }}>
                                  {r
                                    ? <span
                                        title={`${isDeclined(r) ? (isDe ? 'Absage ohne vorherige Anmeldung' : 'Decline without prior registration') : (isDe ? 'Vom Gesamt-Event abgemeldet' : 'Cancelled from the overall event')} — ${formatDate(r.CancellationDate)}`}
                                        style={{ display: 'inline-block', fontSize: '0.72rem', fontWeight: 600, padding: '2px 8px', borderRadius: 999, background: 'rgba(218,41,28,0.10)', color: 'var(--dex-red, #da291c)', whiteSpace: 'nowrap' }}
                                      >
                                        {isDe ? 'Nimmt nicht teil' : 'Will not attend'}
                                      </span>
                                    : <span style={{ color: 'var(--dex-gray-300)' }}>–</span>}
                                </td>
                              );
                            }
                            return (
                              <td key={sc.id} style={{ padding: 8, textAlign: 'center' }}>
                                {r
                                  ? <span title={`${isDeclined(r) ? (isDe ? 'Absage (nicht angemeldet)' : 'Decline (never registered)') : (isDe ? 'Abgemeldet' : 'Cancelled')} — ${formatDate(r.CancellationDate)}`} style={{ color: 'var(--dex-red, #da291c)', fontWeight: 700, fontSize: '1rem' }}>&#10007;</span>
                                  : <span style={{ color: 'var(--dex-gray-300)' }}>–</span>}
                              </td>
                            );
                          })}
                          <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{p.latest ? formatDate(new Date(p.latest).toISOString()) : '-'}</td>
                          {canDelete && (
                            <td style={{ padding: 8, textAlign: 'right' }}>
                              <button
                                type="button"
                                className="btn btn-outline"
                                style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #da291c)', borderColor: 'var(--dex-red, #da291c)' }}
                                onClick={() => { deletePerson(p).catch(() => { /* */ }); }}
                              >
                                {isDe ? 'Löschen' : 'Delete'}
                              </button>
                            </td>
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            );
          }
          return (
            <>
              <h4 style={{ marginTop: 24, color: 'var(--dex-gray-400)' }}>
                {isDe ? 'Abmeldungen' : 'Cancellations'} ({cancelledRegs.length})
                {declineCount > 0 && (
                  <span style={{ fontSize: '0.8rem', fontWeight: 400, marginLeft: 8, color: 'var(--dex-gray-500)' }}>
                    {isDe ? `davon ${declineCount} Absage(n)` : `incl. ${declineCount} decline(s)`}
                  </span>
                )}
              </h4>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                      <th style={thClickable} onClick={() => toggleSort('vorname')}>Vorname{arrow('vorname')}</th>
                      <th style={thClickable} onClick={() => toggleSort('nachname')}>Nachname{arrow('nachname')}</th>
                      <th style={thClickable} onClick={() => toggleSort('email')}>Email{arrow('email')}</th>
                      <th style={thClickable} onClick={() => toggleSort('jobtitle')}>Job Title{arrow('jobtitle')}</th>
                      <th style={thClickable} onClick={() => toggleSort('location')}>Standort{arrow('location')}</th>
                      <th style={thClickable} onClick={() => toggleSort('type')}>{isDe ? 'Art' : 'Type'}{arrow('type')}</th>
                      {isConsolidatedMode && (
                        <th style={{ ...thClickable, cursor: 'default' }}>{isDe ? 'Sub-Event' : 'Sub-event'}</th>
                      )}
                      <th style={thClickable} onClick={() => toggleSort('date')}>{isDe ? 'Abgemeldet am' : 'Cancelled on'}{arrow('date')}</th>
                      {/* v19.4: „Wurde ersetzt durch" — die nachgerückte Person, die
                          den frei gewordenen Platz übernommen hat (vom Flow gesetzt).
                          v19.11: nur bei Events mit echter Warteliste-/Nachrück-
                          Aktivität (sonst durchgehend leer). */}
                      {hasWaitlistActivity && (
                        <th style={{ ...thClickable, cursor: 'default' }}>{isDe ? 'Wurde ersetzt durch' : 'Replaced by'}</th>
                      )}
                      {canDelete && (
                        <th style={{ ...thClickable, cursor: 'default', textAlign: 'right' }}>{isDe ? 'Löschen' : 'Delete'}</th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {sorted.map(reg => {
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const anyReg = reg as any;
                      const declined = isDeclined(reg);
                      return (
                        <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                          <td style={{ padding: 8, fontWeight: 500 }}>{reg.Vorname || '-'}</td>
                          <td style={{ padding: 8, fontWeight: 500 }}>{reg.Nachname || '-'}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{anyReg.JobTitle || '-'}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{anyReg.Location || '-'}</td>
                          <td style={{ padding: 8 }}>
                            {declined
                              ? <span style={{ fontSize: '0.72rem', fontWeight: 600, padding: '2px 8px', borderRadius: 999, background: 'rgba(0,118,168,0.10)', color: 'var(--dex-blue, #0076a8)' }}>{isDe ? 'Absage (nicht angemeldet)' : 'Decline (never registered)'}</span>
                              : <span style={{ fontSize: '0.72rem', fontWeight: 600, padding: '2px 8px', borderRadius: 999, background: 'rgba(218,41,28,0.08)', color: 'var(--dex-red, #da291c)' }}>{isDe ? 'Abgemeldet' : 'Cancelled'}</span>}
                          </td>
                          {isConsolidatedMode && (
                            <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{(reg as SPRegistration & { _sectionTitle?: string })._sectionTitle || '-'}</td>
                          )}
                          <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.CancellationDate)}</td>
                          {hasWaitlistActivity && (
                            <td style={{ padding: 8, color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.8rem' }}>
                              {(() => {
                                const email = (anyReg.ReplacedByParticipantEmail as string | undefined) || '';
                                if (!email) return <span style={{ color: 'var(--dex-gray-300)' }}>—</span>;
                                const other = registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === email.toLowerCase());
                                const label = other ? (((other.Vorname || '') + ' ' + (other.Nachname || '')).trim() || other.ParticipantName || email) : email;
                                return <span title={email}>{label}</span>;
                              })()}
                            </td>
                          )}
                          {canDelete && (
                            <td style={{ padding: 8, textAlign: 'right' }}>
                              <button
                                type="button"
                                title={isDe ? 'Registrierung endgültig löschen' : 'Permanently delete registration'}
                                onClick={() => { deleteCancelled(reg).catch(() => { /* */ }); }}
                                style={{
                                  display: 'inline-flex', alignItems: 'center', gap: 5,
                                  border: '1px solid var(--dex-red, #da291c)', background: 'rgba(218,41,28,0.06)',
                                  color: 'var(--dex-red, #da291c)', borderRadius: 6, padding: '4px 9px',
                                  fontSize: '0.74rem', fontWeight: 600, cursor: 'pointer',
                                }}
                              >
                                <Trash2 size={13} /> {isDe ? 'Löschen' : 'Delete'}
                              </button>
                            </td>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </>
          );
        })()}
      </div>

      {/* ===== TEILNEHMER-EDIT MODAL (v8.0) ===== */}
      {dangerZoneModal}

      {changeLogModal}

      {/* v20.7: Fortschritts-Modal der Zugriffs-Reparatur („Fremd-Anmeldungen:
          Zugriff reparieren") — Event i/N, Eintrag x/y, Balken, Abschluss-
          Summary. Während des Laufs nicht wegklickbar. */}
      {accessFixModal && (
        <Modal
          open={true}
          onClose={() => { if (!accessFixModal.running) setAccessFixModal(null); }}
          dismissable={!accessFixModal.running}
          maxWidth={520}
          padding={24}
          ariaLabel={isDe ? 'Zugriffs-Prüfung' : 'Access check'}
        >
          <h3 style={{ margin: '0 0 10px', fontSize: '1.05rem' }}>
            {isDe ? 'Fremd-Anmeldungen: Zugriff reparieren' : 'Proxy registrations: repair access'}
          </h3>
          {accessFixModal.running ? (
            <>
              <p style={{ margin: '0 0 6px', fontSize: '0.9rem', color: 'var(--dex-gray-700)' }}>
                {isDe ? 'Event' : 'Event'} {accessFixModal.evIdx}/{accessFixModal.evTotal}: <strong>{accessFixModal.evTitle || '…'}</strong>
              </p>
              <p style={{ margin: '0 0 12px', fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>
                {accessFixModal.itemTotal > 0
                  ? (isDe
                    ? `Eintrag ${accessFixModal.itemDone}/${accessFixModal.itemTotal} wird geprüft…`
                    : `Checking item ${accessFixModal.itemDone}/${accessFixModal.itemTotal}…`)
                  : (isDe ? 'Liste wird geladen…' : 'Loading list…')}
              </p>
              {(() => {
                const evBase = Math.max(0, accessFixModal.evIdx - 1);
                const inner = accessFixModal.itemTotal > 0 ? accessFixModal.itemDone / accessFixModal.itemTotal : 0;
                const pct = Math.min(100, Math.round(((evBase + inner) / Math.max(1, accessFixModal.evTotal)) * 100));
                return (
                  <div style={{ background: 'var(--dex-gray-100, #f0f0f0)', borderRadius: 999, height: 10, overflow: 'hidden' }}>
                    <div style={{ width: `${pct}%`, height: '100%', background: 'var(--dex-green, #86bc25)', borderRadius: 999, transition: 'width 0.2s ease' }} />
                  </div>
                );
              })()}
              <p style={{ margin: '10px 0 0', fontSize: '0.78rem', color: 'var(--dex-gray-400)' }}>
                {isDe ? 'Bitte das Fenster geöffnet lassen, bis die Prüfung abgeschlossen ist.' : 'Please keep this window open until the check completes.'}
              </p>
            </>
          ) : (
            <>
              <ul style={{ margin: '0 0 14px', paddingLeft: 18, fontSize: '0.88rem', color: 'var(--dex-gray-700)', lineHeight: 1.6 }}>
                {(accessFixModal.summary || []).map((s, i) => <li key={i}>{s}</li>)}
              </ul>
              <div style={{ textAlign: 'right' }}>
                <button className="btn btn-primary" onClick={() => setAccessFixModal(null)} style={{ fontSize: '0.88rem', padding: '9px 18px' }}>
                  {isDe ? 'Schließen' : 'Close'}
                </button>
              </div>
            </>
          )}
        </Modal>
      )}

      {/* v20.2: Self-Check-in-Modal (von der QR-Kachel unter dem Event-Logo):
          großer QR, Erklärtext, PDF-/Live-Aktionen + editierbares Zeitfenster. */}
      {sciModalOpen && selectedEvent && (
        <Modal
          open={sciModalOpen}
          onClose={() => setSciModalOpen(false)}
          dismissable={!sciBusy}
          maxWidth={560}
          padding={24}
          ariaLabel="Self-Check-in"
        >
          <h3 style={{ margin: '0 0 8px', fontSize: '1.05rem' }}>
            {isDe ? 'Self-Check-in — QR-Code' : 'Self check-in — QR code'}
          </h3>
          <p style={{ margin: '0 0 12px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? 'Diesen QR-Code kannst du am Eingang aushängen oder auf einem Bildschirm zeigen. Teilnehmer scannen ihn mit der Kamera ihres Firmenhandys und checken sich damit selbst ein — ganz ohne Scanner-Team. Jede Person kann nur sich selbst einchecken (Login-gebunden).'
              : 'Post this QR code at the entrance or show it on a screen. Attendees scan it with their company phone camera and check themselves in — no scanner team needed. Each person can only check in themselves (login-bound).'}
          </p>
          {sciModalQr ? (
            <div style={{ textAlign: 'center', margin: '0 0 14px' }}>
              <img src={sciModalQr} alt="Self-Check-in QR" style={{ width: 260, maxWidth: '80%', height: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: 10, background: '#fff' }} />
            </div>
          ) : (
            <p style={{ textAlign: 'center', color: 'var(--dex-gray-400)', fontSize: '0.85rem' }}>
              {isDe ? 'QR-Code konnte nicht erzeugt werden.' : 'QR code could not be generated.'}
            </p>
          )}
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'center', marginBottom: 18 }}>
            <button
              className="btn btn-primary"
              disabled={sciBusy}
              style={{ fontSize: '0.88rem', padding: '10px 18px' }}
              onClick={() => {
                (async () => {
                  await downloadSelfCheckInPdf({
                    eventTitle: selectedEvent.title || 'Event',
                    eventDateLabel: selectedEvent.startDate ? new Date(selectedEvent.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '',
                    locationLabel: selectedEvent.location || '',
                    token: sciToken,
                  });
                })().catch(() => { /* best-effort */ });
              }}
            >
              {isDe ? 'QR-PDF herunterladen (drucken)' : 'Download QR PDF (print)'}
            </button>
            <button
              className="btn btn-secondary"
              disabled={sciBusy}
              style={{ fontSize: '0.88rem', padding: '10px 18px' }}
              onClick={() => { setSciModalOpen(false); navigate('self-checkin-display', selectedEvent.id); }}
            >
              {isDe ? 'Live-QR anzeigen (rotierend)' : 'Show live QR (rotating)'}
            </button>
          </div>
          {/* Zeitfenster: Von/Bis — verhindert verfrühte UND nachträgliche Check-ins. */}
          <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: '12px 14px' }}>
            <div style={{ fontWeight: 700, fontSize: '0.88rem', marginBottom: 4 }}>
              {isDe ? 'Check-in-Zeitfenster' : 'Check-in time window'}
            </div>
            <p style={{ margin: '0 0 10px', fontSize: '0.78rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              {isDe
                ? 'Von wann bis wann der Self-Check-in möglich ist. Vor „Von" und nach „Bis" sind keine Check-ins möglich — also auch keine nachträglichen. Vorbelegt mit dem Standard: 2 Stunden vor Event-Start bis Event-Ende (gilt auch, solange du nichts anderes speicherst).'
                : 'From when until when self check-in is possible. Before "from" and after "until" no check-ins are possible — including late ones. Prefilled with the default: 2 hours before event start until event end (which also applies as long as you do not save anything else).'}
            </p>
            <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', alignItems: 'flex-end' }}>
              <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                {isDe ? 'Von' : 'From'}
                <input
                  type="datetime-local"
                  value={sciFrom}
                  onChange={e => setSciFrom(e.target.value)}
                  className="form-input"
                  style={{ padding: '8px 10px', fontSize: '0.85rem' }}
                />
              </label>
              <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                {isDe ? 'Bis' : 'Until'}
                <input
                  type="datetime-local"
                  value={sciTo}
                  onChange={e => setSciTo(e.target.value)}
                  className="form-input"
                  style={{ padding: '8px 10px', fontSize: '0.85rem' }}
                />
              </label>
              <button
                className="btn btn-secondary"
                disabled={sciBusy}
                style={{ fontSize: '0.85rem', padding: '9px 16px' }}
                onClick={() => { saveSelfCheckInWindow().catch(() => { /* */ }); }}
              >
                {sciBusy ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Zeitfenster speichern' : 'Save time window')}
              </button>
            </div>
            {sciSaveMsg && (
              <p style={{ margin: '8px 0 0', fontSize: '0.78rem', color: sciSaveMsg.indexOf('fehlgeschlagen') >= 0 || sciSaveMsg.indexOf('failed') >= 0 ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)' }}>
                {sciSaveMsg}
              </p>
            )}
          </div>
          <div style={{ textAlign: 'right', marginTop: 14 }}>
            <button className="btn btn-secondary" onClick={() => setSciModalOpen(false)} style={{ fontSize: '0.85rem' }}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          </div>
        </Modal>
      )}

      {/* v9.15: QR-Code-Versand-Modal — Test (nur Organizer) / Volldurchlauf
          (alle Angemeldeten) / Auto-Send-Toggle für zukünftige Anmeldungen. */}
      {qrSendModalOpen && selectedEvent && (
        <Modal
          open={qrSendModalOpen}
          onClose={() => setQrSendModalOpen(false)}
          dismissable={!isSendingQR}
          maxWidth={860}
          padding={24}
          ariaLabel="QR-Codes versenden"
        >
            <h3 style={{ margin: '0 0 4px', fontSize: '1.15rem' }}>{isDe ? 'QR-Code-Versand' : 'QR code sending'}</h3>
            <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              {isDe
                ? 'Persönliche QR-Codes an deine Teilnehmer — fürs schnelle Einchecken am Event-Tag. Jede Mail kommt im Deloitte-Layout und zeigt unter dem Code Name + Event als Klartext.'
                : 'Personal QR codes for your participants — for fast check-in on event day. Each email comes in the Deloitte layout and shows name + event as plain text below the code.'}
            </p>
            {/* v22.6: kompakte Status-Pills statt mehrerer Boxen. */}
            {(() => {
              const without = registrations.filter(r => r.Status === 'Angemeldet').length;
              const withQr = registrations.filter(r => r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length;
              const externalCount = registrations.filter(r => r.Status === 'Angemeldet').filter(r => r.ParticipantEmail && !/@(.*\.)?deloitte\.de$/i.test(r.ParticipantEmail)).length;
              const pill = (bg: string, fg: string, content: React.ReactNode): React.ReactElement => (
                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5, padding: '5px 12px', borderRadius: 999, background: bg, color: fg, fontSize: '0.8rem', fontWeight: 600 }}>{content}</span>
              );
              return (
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 18 }}>
                  {pill(without > 0 ? 'rgba(237,139,0,0.12)' : 'rgba(134,188,37,0.12)', without > 0 ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)', <><strong>{without}</strong> {isDe ? 'ohne Code' : 'without code'}</>)}
                  {pill('var(--dex-gray-100, #eee)', 'var(--dex-gray-600)', <><strong>{withQr}</strong> {isDe ? 'mit Code' : 'with code'}</>)}
                  {pill('rgba(134,188,37,0.10)', 'var(--dex-gray-600)', isDe ? 'Neue Anmeldungen automatisch (ab 1. Versand)' : 'New registrations automatic (after 1st send)')}
                  {externalCount > 0 && pill('#fff3e0', '#7a4a00', isDe ? `${externalCount} extern → QR an Organizer` : `${externalCount} external → QR to organizer`)}
                </div>
              );
            })()}
            {/* v22.6: Querformat — links der Versand-Flow, rechts die
                Self-Check-in-Alternative. Auto-stack auf schmalen Screens. */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: 16, alignItems: 'stretch' }}>
              {/* LINKS: QR-Codes an Teilnehmer senden */}
              <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: 16, display: 'flex', flexDirection: 'column' }}>
                <div style={{ fontWeight: 700, fontSize: '0.92rem', marginBottom: 10, color: 'var(--dex-gray-800)' }}>
                  {isDe ? 'QR-Codes an deine Teilnehmer senden' : 'Send QR codes to your participants'}
                </div>
                <ol style={{ margin: '0 0 14px', padding: 0, listStyle: 'none', display: 'flex', flexDirection: 'column', gap: 8 }}>
                  {[
                    { n: 1, d: isDe ? 'Vorschau ansehen — wie sieht die Mail aus?' : 'Preview — what does the email look like?' },
                    { n: 2, d: isDe ? 'Test an dich — Mail einmal selbst bekommen.' : 'Test to yourself — receive the email once.' },
                    { n: 3, d: isDe ? 'An alle ohne Code senden.' : 'Send to everyone without a code.' },
                  ].map(s => (
                    <li key={s.n} style={{ display: 'flex', gap: 9, alignItems: 'center' }}>
                      <span style={{ flexShrink: 0, width: 20, height: 20, borderRadius: '50%', background: 'var(--dex-green, #86bc25)', color: '#fff', fontWeight: 700, fontSize: '0.72rem', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}>{s.n}</span>
                      <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.4 }}>{s.d}</span>
                    </li>
                  ))}
                </ol>
                {/* v22.6: Aktions-Buttons im linken Block — Reihenfolge wie die Schritte. */}
                <div style={{ marginTop: 'auto', display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                    <button
                      className="btn btn-outline"
                      disabled={isSendingQR || qrPreviewLoading}
                      onClick={() => { qrPreviewAction().catch(() => { /* */ }); }}
                      style={{ fontSize: '0.85rem', flex: 1, minWidth: 130 }}
                      title={isDe ? 'So sieht die Mail aus, die rausgeht — inklusive echtem QR-Code für dich.' : 'How the email looks when sent — including a real QR code for you.'}
                    >
                      {qrPreviewLoading ? (isDe ? 'Lädt…' : 'Loading…') : (isDe ? '1. Vorschau' : '1. Preview')}
                    </button>
                    <button
                      className="btn btn-secondary"
                      onClick={() => { qrTestSendAction().catch(() => { /* */ }); }}
                      disabled={isSendingQR}
                      style={{ fontSize: '0.85rem', flex: 1, minWidth: 130 }}
                    >
                      {isDe ? '2. Test an mich' : '2. Test to me'}
                    </button>
                  </div>
                  <button
                    className="btn btn-primary"
                    onClick={() => { qrFullSendAction().catch(() => { /* */ }); }}
                    disabled={isSendingQR || registrations.filter(r => r.Status === 'Angemeldet').length === 0}
                    style={{ fontSize: '0.9rem', width: '100%', padding: '11px 18px', fontWeight: 700 }}
                  >
                    {(() => {
                      const n = registrations.filter(r => r.Status === 'Angemeldet').length;
                      if (isSendingQR) return `${isDe ? 'Versende' : 'Sending'}… (${qrSentCount})`;
                      if (n === 0) return isDe ? 'Alle haben ihren QR-Code' : 'Everyone has their QR code';
                      if (isDe) return `3. QR-${n === 1 ? 'Code' : 'Codes'} an ${n} Teilnehmer senden`;
                      return `3. Send QR ${n === 1 ? 'code' : 'codes'} to ${n} participant${n === 1 ? '' : 's'}`;
                    })()}
                  </button>
                  {/* v22.18: Mail-Text pro Event anpassbar — QR-Block bleibt fix. */}
                  <button
                    type="button"
                    className="btn btn-outline"
                    disabled={isSendingQR}
                    onClick={() => { openQrMailEditor().catch(() => { /* */ }); }}
                    style={{ fontSize: '0.82rem', width: '100%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 6 }}
                  >
                    <Pencil size={14} />
                    {isDe
                      ? `Mail-Text anpassen${getQrMailOverride(selectedEvent) ? ' (angepasst)' : ''}`
                      : `Customize email text${getQrMailOverride(selectedEvent) ? ' (customized)' : ''}`}
                  </button>
                  {/* Hinweis, falls Organizer selbst nicht angemeldet ist (nur fürs Testen relevant). */}
                  {(() => {
                    const orgEmail = (currentUser.email || '').toLowerCase();
                    const isOrgRegistered = !!orgEmail && registrations.some(r => (r.ParticipantEmail || '').toLowerCase() === orgEmail && (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt'));
                    if (isOrgRegistered) return null;
                    return (
                      <div style={{ fontSize: '0.74rem', color: '#7a5a00', background: '#fff8e1', border: '1px solid #f5b400', borderRadius: 6, padding: '7px 10px', lineHeight: 1.4 }}>
                        {isDe
                          ? 'Test-Hinweis: Du bist selbst nicht angemeldet — die Test-Mail kommt an, aber ein späterer Check-in-Scan findet dich nicht in der Liste.'
                          : 'Test note: you are not registered yourself — the test email arrives, but a later check-in scan will not find you in the list.'}
                      </div>
                    );
                  })()}
                </div>
              </div>
              {/* RECHTS: Self-Check-in als Alternative */}
              <div style={{ border: '1px solid var(--dex-green, #86bc25)', background: 'rgba(134,188,37,0.06)', borderRadius: 12, padding: 16, display: 'flex', flexDirection: 'column' }}>
                <div style={{ fontWeight: 700, fontSize: '0.92rem', marginBottom: 8, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                  {isDe ? 'Alternative: Self-Check-in' : 'Alternative: self check-in'}
                </div>
                <p style={{ margin: '0 0 14px', fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Teilnehmer checken sich am Eingang selbst ein — sie scannen einen Event-QR mit der normalen Handy-Kamera, ganz ohne Scanner-Team.'
                    : 'Participants check in at the entrance themselves — they scan an event QR with their phone camera, no scanner team needed.'}
                </p>
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 'auto' }}>
                  <button
                    className="btn btn-outline"
                    style={{ fontSize: '0.82rem', flex: 1, minWidth: 130 }}
                    onClick={() => { (async () => { const token = await ensureSelfCheckInReady(selectedEvent); if (token) { setQrSendModalOpen(false); navigate('self-checkin-display', selectedEvent.id); } })().catch(() => { /* */ }); }}
                  >
                    {isDe ? 'Live-QR anzeigen' : 'Show live QR'}
                  </button>
                  <button
                    className="btn btn-outline"
                    style={{ fontSize: '0.82rem', flex: 1, minWidth: 130 }}
                    onClick={() => { (async () => { const token = await ensureSelfCheckInReady(selectedEvent); if (!token) return; await downloadSelfCheckInPdf({ eventTitle: selectedEvent.title || 'Event', eventDateLabel: selectedEvent.startDate ? new Date(selectedEvent.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '', locationLabel: selectedEvent.location || '', token }); })().catch(() => { /* */ }); }}
                  >
                    {isDe ? 'QR-PDF herunterladen' : 'Download QR PDF'}
                  </button>
                </div>
                <p style={{ margin: '14px 0 0', fontSize: '0.74rem', color: 'var(--dex-gray-500)', lineHeight: 1.45 }}>
                  {isDe ? 'Ablauf am Event-Tag: ' : 'Process on event day: '}
                  <a href="javascript:void(0)" onClick={(e) => { e.preventDefault(); navigate('manual'); window.location.hash = 'check-in'; }} style={{ color: 'var(--dex-green-dark)', fontWeight: 600 }}>
                    {isDe ? 'Handbuch „Check-In am Event-Tag"' : 'Manual “Check-in on event day”'}
                  </a>
                </p>
              </div>
            </div>

            {/* v22.7: Info-Box „neue Anmeldungen automatisch" — erklärt, dass
                nach dem ersten Versand niemand mehr manuell nachversorgt werden
                muss. */}
            <div style={{ marginTop: 14, display: 'flex', alignItems: 'flex-start', gap: 9, padding: '10px 12px', borderRadius: 8, background: 'var(--dex-green-light, #f0f8e8)', border: '1px solid rgba(134,188,37,0.4)' }}>
              <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', flexShrink: 0, marginTop: 1 }}><Check size={16} /></span>
              <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                {isDe
                  ? <><strong>Neue Anmeldungen bekommen ihren QR-Code automatisch.</strong> Sobald du den Versand einmal gestartet hast (Schritt 3), erhält jede weitere Anmeldung an diesem Event ihren QR-Code direkt zusammen mit der Anmeldebestätigung — auch nach der Anmeldefrist. Du musst dann nichts mehr manuell nachsenden.</>
                  : <><strong>New registrations get their QR code automatically.</strong> Once you have started sending (step 3), every further registration for this event receives its QR code together with the registration confirmation — even after the deadline. You no longer have to resend anything manually.</>}
              </span>
            </div>

            {qrSendResult && (
              <div style={{
                padding: '8px 12px', marginBottom: 12, borderRadius: 6,
                background: qrSendResult.startsWith('Fehler') ? 'var(--dex-red-light, #ffe5e5)' : 'var(--dex-green-light, #f0f8e8)',
                fontSize: '0.85rem', color: qrSendResult.startsWith('Fehler') ? 'var(--dex-red-dark, #b00)' : 'var(--dex-green-dark)',
              }}>{qrSendResult}</div>
            )}

            <div style={{ marginTop: 16, textAlign: 'right' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setQrSendModalOpen(false)}
                disabled={isSendingQR}
                style={{ fontSize: '0.85rem' }}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
            </div>
        </Modal>
      )}

      {/* ===== v22.18: QR-MAIL-TEXT ANPASSEN (HtmlEditorModal mit Live-Vorschau,
          gespeichert im Event — gilt auch für den Auto-Versand) ===== */}
      {qrEditOpen && selectedEvent && (() => {
        const myName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
        const previewVars: Record<string, string> = {
          EventTitle: selectedEvent.title,
          Vorname: currentUser.firstName || myName,
          Name: myName,
        };
        const customLogo = (() => {
          try {
            const o = JSON.parse(selectedEvent.emailTemplateOverrides || '{}');
            return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
          } catch { return ''; }
        })();
        const resolvePlain = (s: string): string => s
          .replace(/\{\{EventTitle\}\}/g, selectedEvent.title)
          .replace(/\{\{Vorname\}\}/g, previewVars.Vorname)
          .replace(/\{\{Name\}\}/g, myName);
        const def = qrEmailDefaults(selectedEvent.emailLanguage || 'EN');
        // Versand-Spalte links: ungespeicherte Änderungen sperren den
        // Massen-Versand (der nutzt den GESPEICHERTEN Text) — der Test an
        // mich nutzt bewusst den aktuellen Editor-Text (Test = Vorschau).
        const savedOv = getQrMailOverride(selectedEvent);
        const savedSubject = (savedOv && savedOv.subject) || def.subject;
        const savedHeading = (savedOv && savedOv.heading) || def.heading;
        const savedSubheading = (savedOv && savedOv.subheading) || def.subheading;
        const savedBody = (savedOv && savedOv.bodyHtml) || def.body;
        const qrEditDirty = qrEditSubject.trim() !== savedSubject.trim()
          || qrEditHeading.trim() !== savedHeading.trim()
          || qrEditSubheading.trim() !== savedSubheading.trim()
          || qrEditBody.trim() !== savedBody.trim();
        const noCodeCount = registrations.filter(r => r.Status === 'Angemeldet').length;
        const withCodeCount = registrations.filter(r => r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length;
        const leftPanel = (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <div style={{ fontWeight: 700, fontSize: '0.92rem' }}>{isDe ? 'QR-Code-Versand' : 'QR code sending'}</div>
            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
              <span style={{ background: '#eef6e3', color: 'var(--dex-green-dark, #4a7c1f)', borderRadius: 12, padding: '3px 10px', fontSize: '0.74rem', fontWeight: 600 }}>
                <strong>{noCodeCount}</strong> {isDe ? 'ohne Code' : 'without code'}
              </span>
              <span style={{ background: 'var(--dex-gray-100, #f0f0f0)', color: 'var(--dex-gray-600)', borderRadius: 12, padding: '3px 10px', fontSize: '0.74rem', fontWeight: 600 }}>
                <strong>{withCodeCount}</strong> {isDe ? 'mit Code' : 'with code'}
              </span>
            </div>
            <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', lineHeight: 1.5 }}>
              {isDe
                ? 'Die Live-Vorschau rechts zeigt deinen aktuellen Text. Der Test an dich nutzt ebenfalls den aktuellen Text — der Versand an die Teilnehmer immer den gespeicherten.'
                : 'The live preview on the right shows your current text. The test to yourself also uses the current text — sending to participants always uses the saved one.'}
            </div>
            <button
              type="button"
              className="btn btn-outline"
              disabled={isSendingQR}
              onClick={() => { qrTestSendAction({ subject: qrEditSubject, heading: qrEditHeading, subheading: qrEditSubheading, bodyHtml: qrEditBody }).catch(() => { /* */ }); }}
              style={{ fontSize: '0.82rem', width: '100%' }}
            >
              {isDe ? 'Test an mich (aktueller Text)' : 'Test to me (current text)'}
            </button>
            <button
              type="button"
              className="btn btn-primary"
              disabled={isSendingQR || qrEditDirty || noCodeCount === 0}
              onClick={() => { qrFullSendAction().catch(() => { /* */ }); }}
              style={{ fontSize: '0.82rem', width: '100%', fontWeight: 700 }}
              title={qrEditDirty ? (isDe ? 'Erst speichern — der Versand nutzt den gespeicherten Text.' : 'Save first — sending uses the saved text.') : undefined}
            >
              {isSendingQR
                ? `${isDe ? 'Versende' : 'Sending'}… (${qrSentCount})`
                : (noCodeCount === 0
                  ? (isDe ? 'Alle haben ihren QR-Code' : 'Everyone has their QR code')
                  : (isDe ? `An ${noCodeCount} Teilnehmer senden` : `Send to ${noCodeCount} participant${noCodeCount === 1 ? '' : 's'}`))}
            </button>
            {qrEditDirty && (
              <div style={{ background: '#fff3e0', border: '1px solid #ed8b00', borderRadius: 'var(--dex-radius)', padding: '8px 10px', fontSize: '0.72rem', color: '#7a4a00', lineHeight: 1.5 }}>
                {isDe
                  ? 'Ungespeicherte Änderungen — erst „Für dieses Event speichern" klicken, dann an die Teilnehmer senden.'
                  : 'Unsaved changes — click "Save for this event" first, then send to participants.'}
              </div>
            )}
            {qrSendResult && (
              <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, borderTop: '1px solid var(--dex-gray-200)', paddingTop: 8 }}>
                {qrSendResult}
              </div>
            )}
            <button
              type="button"
              className="btn btn-secondary"
              disabled={isSendingQR}
              onClick={closeQrMailEditor}
              style={{ fontSize: '0.78rem', width: '100%', marginTop: 4 }}
            >
              {isDe ? 'Zurück zum Versand-Modal' : 'Back to the send dialog'}
            </button>
          </div>
        );
        const headerExtra = (
          <div style={{ padding: 12, background: 'var(--dex-gray-50, #fafafa)', border: '1px solid var(--dex-gray-200)', borderRadius: 'var(--dex-radius)', marginBottom: 4, fontSize: '0.78rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? <><strong>Fester Bestandteil:</strong> Der Platzhalter <code>{'{{QR_BLOCK}}'}</code> steht für den persönlichen QR-Code mit Name + Event als Klartext — er lässt sich verschieben, aber nicht entfernen (fehlt er im Text, wird der Block beim Versand automatisch ans Ende gesetzt). Verfügbare Platzhalter: <code>{'{{Vorname}}'}</code>, <code>{'{{Name}}'}</code>, <code>{'{{EventTitle}}'}</code>. <strong>Der gespeicherte Text gilt für alle QR-Mails dieses Events</strong> — manueller Versand UND automatischer Versand bei neuen Anmeldungen.</>
              : <><strong>Fixed element:</strong> the placeholder <code>{'{{QR_BLOCK}}'}</code> represents the personal QR code with name + event as plain text — it can be moved but not removed (if missing, the block is appended automatically when sending). Available placeholders: <code>{'{{Vorname}}'}</code>, <code>{'{{Name}}'}</code>, <code>{'{{EventTitle}}'}</code>. <strong>The saved text applies to all QR emails of this event</strong> — manual sending AND the automatic send for new registrations.</>}
          </div>
        );
        return (
          <HtmlEditorModal
            open={qrEditOpen}
            onClose={() => { if (!qrEditSaving && !isSendingQR) closeQrMailEditor(); }}
            title={isDe ? `QR-Mail anpassen: ${selectedEvent.title}` : `Customize QR email: ${selectedEvent.title}`}
            leftPanel={leftPanel}
            value={qrEditBody}
            onChange={setQrEditBody}
            previewMode="email"
            emailSubject={qrEditSubject}
            onEmailSubjectChange={setQrEditSubject}
            emailHeading={qrEditHeading}
            onEmailHeadingChange={setQrEditHeading}
            emailSubheading={qrEditSubheading}
            onEmailSubheadingChange={setQrEditSubheading}
            emailHeadingColor="#86bc25"
            previewVars={previewVars}
            previewHtmlVars={{ QR_BLOCK: qrEditSampleBlock }}
            previewToLine={`${currentUser.email} ${isDe ? '(Beispiel — du selbst)' : '(example — yourself)'}`}
            previewSubjectLine={resolvePlain(qrEditSubject)}
            defaultBodyHtml={def.body}
            insertableVars={[
              { key: '{{Vorname}}', label: isDe ? 'Vorname' : 'First name' },
              { key: '{{Name}}', label: isDe ? 'Voller Name' : 'Full name' },
              { key: '{{EventTitle}}', label: isDe ? 'Event-Titel' : 'Event title' },
              { key: '{{QR_BLOCK}}', label: isDe ? 'QR-Code-Block (fix)' : 'QR code block (fixed)' },
            ]}
            imageBase64={customLogo}
            headerExtra={headerExtra}
            extraAction={{
              label: qrEditSaving
                ? (isDe ? 'Speichert…' : 'Saving…')
                : (isDe ? 'Für dieses Event speichern' : 'Save for this event'),
              onClick: saveQrMailOverride,
              disabled: qrEditSaving || !qrEditSubject.trim() || !qrEditBody.trim(),
              icon: <Check size={16} />,
            }}
          />
        );
      })()}

      {editingReg && selectedEvent && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
          onClick={() => { if (!isSavingEdit) closeEditModal(); }}
        >
          <div
            className="card"
            style={{
              width: '100%', maxWidth: 920, maxHeight: '90vh', overflow: 'auto',
              padding: 24, borderRadius: 16, background: '#fff',
              boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
            }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 8 }}>
                <Pencil size={18} />{' '}
                {isDe ? 'Teilnehmer bearbeiten' : 'Edit attendee'}
                {' — '}
                <span style={{ color: 'var(--dex-green-dark)' }}>
                  {editForm.Vorname} {editForm.Nachname}
                </span>
              </h3>
              <button
                onClick={closeEditModal}
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}
                aria-label={isDe ? 'Schließen' : 'Close'}
                disabled={isSavingEdit}
              ><X size={20} /></button>
            </div>
            <p style={{ margin: '0 0 16px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
              {isDe
                ? 'Hier kannst du Vorname, Nachname und E-Mail-Adresse korrigieren (z.B. nach einem Tippfehler bei der manuellen Anlage) sowie die event-spezifischen Felder anpassen. Beim Ändern der E-Mail wird geprüft, ob die Adresse zum Deloitte-Tenant gehört und die Person dort existiert — externe Adressen sind nicht erlaubt. Phone, Department, Standort und Job Title kommen aus dem M365-Profil und sind read-only — sie werden bei einem Mail-Wechsel automatisch nachgezogen. Den Status änderst du über die Aktions-Buttons in der Liste. Jede Änderung wird im Audit-Log und im ChangeLog des Teilnehmers mit Datum und deinem Namen protokolliert.'
                : 'You can fix first name, last name and email address here (e.g. after a typo during manual creation) and adjust event-specific fields. When changing the email, the app verifies that the address belongs to the Deloitte tenant and that the person exists there — external addresses are not allowed. Phone, Department, Location and Job Title come from the M365 profile and are read-only — they are refreshed automatically when the email changes. The status is changed via the action buttons in the list. Every change is logged in the audit log and in the attendee\'s ChangeLog with date and your name.'}
            </p>

            {(() => {
              // Vorname / Nachname / E-Mail sind seit v9.7 editierbar (mit
              // Deloitte-Domain- und Tenant-Existenz-Check beim Speichern).
              // Die uebrigen Profil-Felder bleiben read-only — sie kommen
              // aus dem M365-Profil und werden bei einer Mail-Aenderung
              // mit den Profil-Daten der neuen Person ueberschrieben.
              const editableStammFields: Array<{ key: string; label: string; type?: string }> = [
                { key: 'Vorname', label: isDe ? 'Vorname' : 'First name' },
                { key: 'Nachname', label: isDe ? 'Nachname' : 'Last name' },
                { key: 'ParticipantEmail', label: 'E-Mail', type: 'email' },
              ];
              const readOnlyFields: Array<{ key: string; label: string }> = [
                { key: 'Anrede', label: isDe ? 'Anrede' : 'Salutation' },
                { key: 'Phone', label: isDe ? 'Telefon' : 'Phone' },
                { key: 'Department', label: 'Department' },
                { key: 'Location', label: isDe ? 'Standort' : 'Location' },
                { key: 'JobTitle', label: 'Job Title' },
                { key: 'Status', label: 'Status' },
              ];
              const renderReadOnly = (label: string, value: string): React.ReactNode => (
                <div>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-500)', marginBottom: 4 }}>
                    {label}
                  </label>
                  <div style={{
                    width: '100%', padding: '8px 12px',
                    background: 'var(--dex-gray-50, #fafafa)',
                    border: '1px solid var(--dex-gray-200, #e5e7eb)',
                    borderRadius: 6, fontSize: '0.88rem',
                    color: value ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
                    minHeight: 38, lineHeight: 1.5,
                  }}>
                    {value || (isDe ? '— nicht gesetzt —' : '— not set —')}
                  </div>
                </div>
              );
              const renderEditable = (key: string, label: string, type?: string): React.ReactNode => (
                <div>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                    {label}
                  </label>
                  <input
                    className="form-input"
                    type={type || 'text'}
                    value={editForm[key] || ''}
                    onChange={e => setEditForm(prev => ({ ...prev, [key]: e.target.value }))}
                    style={{ fontSize: '0.88rem' }}
                  />
                </div>
              );
              return (
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                  {/* Anrede zuerst (read-only) */}
                  <div>{renderReadOnly(isDe ? 'Anrede' : 'Salutation', editForm.Anrede || '')}</div>
                  {/* Vorname + Nachname editierbar */}
                  {editableStammFields.filter(f => f.key !== 'ParticipantEmail').map(f => (
                    <div key={f.key}>
                      {renderEditable(f.key, f.label, f.type)}
                    </div>
                  ))}
                  {/* E-Mail editierbar (volle Breite) */}
                  <div style={{ gridColumn: '1 / -1' }}>
                    {renderEditable('ParticipantEmail', 'E-Mail', 'email')}
                    <p style={{ margin: '4px 0 0', fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                      {isDe
                        ? 'Nur Deloitte-Adressen (@deloitte.de / @deloitte.com). Beim Speichern wird die Person im Tenant verifiziert.'
                        : 'Only Deloitte addresses (@deloitte.de / @deloitte.com). The person is verified in the tenant on save.'}
                    </p>
                  </div>
                  {/* Restliche Profil-Felder read-only */}
                  {readOnlyFields.filter(f => f.key !== 'Anrede').map(f => (
                    <div key={f.key}>
                      {renderReadOnly(f.label, editForm[f.key] || '')}
                    </div>
                  ))}

                  {/* B2Run-Starter-Typ (Funstarter / Durchstarter). Hardcoded
                      SP-Spalte auf der Teilnehmerliste (kein regulärer
                      Custom-Field-Eintrag), daher explizit hier gerendert.
                      Updates BEIDE intern getrackten Felder zugleich
                      (StarterType + PreferredStarterType) — die getrennte
                      Speicherung von „aktuell vs. Wunsch" ist Implementierungs-
                      Detail für die Warteliste-Nachrück-Logik und braucht im
                      Edit-Modal keine UI-Komplexität. v10.15+ */}
                  {selectedEvent.durchstarterCapacity !== undefined
                    && selectedEvent.funstarterCapacity !== undefined
                    && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0) && (
                    <div style={{ gridColumn: '1 / -1', marginTop: 12, paddingTop: 16, borderTop: '1px solid var(--dex-gray-200)' }}>
                      <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                        {isDe ? 'B2Run-Starter-Typ' : 'B2Run starter type'}
                      </label>
                      <select
                        value={editForm.StarterType || ''}
                        onChange={e => {
                          const v = e.target.value;
                          // Beide Felder synchron halten — der Aktuelle wechselt
                          // mit, der Wunsch ebenfalls (User-Erwartung: „ich ändere
                          // den Starter-Typ" = beides ändert sich).
                          setEditForm(prev => ({ ...prev, StarterType: v, PreferredStarterType: v }));
                        }}
                        className="form-input"
                        style={{ maxWidth: 320 }}
                      >
                        <option value="">{isDe ? '— bitte wählen —' : '— please select —'}</option>
                        <option value="Durchstarter">Durchstarter</option>
                        <option value="Funstarter">Funstarter</option>
                      </select>
                    </div>
                  )}

                  {/* Custom Fields des Events — DAS ist der editierbare Teil.
                      Renderer abhängig vom Field-Type (text/number/select/
                      checkbox). Multi-Select speichert Werte als " | "-
                      getrennten String, identisch zum Registrierungs-Pfad. */}
                  {selectedEvent.eventSpecificFields && selectedEvent.eventSpecificFields.length > 0 && (
                    <div style={{ gridColumn: '1 / -1', marginTop: 12, paddingTop: 16, borderTop: '1px solid var(--dex-gray-200)' }}>
                      <h4 style={{ margin: '0 0 4px', fontSize: '0.92rem', color: 'var(--dex-gray-800)' }}>
                        {isDe ? 'Event-spezifische Felder (editierbar)' : 'Event-specific fields (editable)'}
                      </h4>
                      <p style={{ margin: '0 0 12px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                        {isDe
                          ? 'Nur diese Felder werden gespeichert.'
                          : 'Only these fields will be saved.'}
                      </p>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                        {selectedEvent.eventSpecificFields.map(cf => {
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const sp = (cf as any).spInternalName || '';
                          if (!sp) return null;
                          const value = editForm[sp] || '';
                          const setVal = (v: string): void => setEditForm(prev => ({ ...prev, [sp]: v }));
                          const labelEl = (
                            <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                              {cf.label}{cf.required && <span style={{ color: 'var(--dex-red, #c00)' }}> *</span>}
                            </label>
                          );

                          // Single-Select-Dropdown
                          if (cf.type === 'select' && !cf.multi && cf.options && cf.options.length > 0) {
                            return (
                              <div key={cf.id}>
                                {labelEl}
                                <select
                                  className="form-select"
                                  value={value}
                                  onChange={e => setVal(e.target.value)}
                                  style={{ width: '100%' }}
                                >
                                  <option value="">{isDe ? '— bitte wählen —' : '— please choose —'}</option>
                                  {cf.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                                </select>
                              </div>
                            );
                          }

                          // v11.89: Multi-Select-Dropdown (vorher Checkbox-Liste).
                          // Werte werden weiterhin ' | '-getrennt gespeichert.
                          if (cf.type === 'select' && cf.multi && cf.options && cf.options.length > 0) {
                            const selected = value.split(' | ').map(s => s.trim()).filter(Boolean);
                            return (
                              <div key={cf.id} style={{ gridColumn: '1 / -1' }}>
                                {labelEl}
                                <MultiSelectDropdown
                                  options={cf.options}
                                  value={selected}
                                  onChange={next => setVal(next.join(' | '))}
                                  placeholder={isDe ? '— bitte wählen —' : '— please choose —'}
                                />
                              </div>
                            );
                          }

                          // Checkbox (true/false)
                          if (cf.type === 'checkbox') {
                            const isChecked = value === 'true' || value === '1';
                            return (
                              <div key={cf.id}>
                                {labelEl}
                                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, padding: '6px 10px', background: isChecked ? 'rgba(134,188,37,0.12)' : 'var(--dex-gray-50)', borderRadius: 6, cursor: 'pointer', fontSize: '0.85rem' }}>
                                  <input
                                    type="checkbox"
                                    checked={isChecked}
                                    onChange={e => setVal(e.target.checked ? 'true' : 'false')}
                                    style={{ accentColor: 'var(--dex-green)' }}
                                  />
                                  {isChecked ? (isDe ? 'Ja' : 'Yes') : (isDe ? 'Nein' : 'No')}
                                </label>
                              </div>
                            );
                          }

                          // Number
                          if (cf.type === 'number') {
                            return (
                              <div key={cf.id}>
                                {labelEl}
                                <input
                                  className="form-input"
                                  type="number"
                                  value={value}
                                  onChange={e => setVal(e.target.value)}
                                  style={{ width: '100%' }}
                                />
                              </div>
                            );
                          }

                          // Default: text-Input (auch für 'text', 'user', 'roommate')
                          return (
                            <div key={cf.id}>
                              {labelEl}
                              <input
                                className="form-input"
                                value={value}
                                onChange={e => setVal(e.target.value)}
                                style={{ width: '100%' }}
                              />
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                </div>
              );
            })()}

            {editError && (
              <div style={{ marginTop: 12, padding: '8px 12px', background: 'rgba(218,41,28,0.08)', border: '1px solid var(--dex-red, #c00)', borderRadius: 6, fontSize: '0.85rem', color: 'var(--dex-red, #c00)' }}>
                {editError}
              </div>
            )}

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 20 }}>
              <button
                type="button"
                className="btn btn-outline"
                onClick={closeEditModal}
                disabled={isSavingEdit}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={saveEdit}
                disabled={isSavingEdit}
                style={{ opacity: isSavingEdit ? 0.6 : 1 }}
              >
                {isSavingEdit ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Speichern' : 'Save')}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* v19.30 (Feature A): Bearbeiten der Hauptevent-Custom-Felder einer
          konsolidierten Zeile. Schreibt in die Registrierung der Person auf
          der Hauptevent-Subsite — gleiche Persistenz wie das Teilnehmer-Edit. */}
      {mainFieldsEditReg && selectedEvent && (
        <Modal
          open={!!mainFieldsEditReg}
          onClose={() => { if (!mainFieldsEditSaving) closeMainFieldsEdit(); }}
          maxWidth={760}
          dismissable={!mainFieldsEditSaving}
          ariaLabel={isDe ? 'Hauptevent-Felder bearbeiten' : 'Edit main-event fields'}
        >
          <div className="flex-between">
            <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Pencil size={18} />{' '}
              {isDe ? 'Felder des Hauptevents bearbeiten' : 'Edit main-event fields'}
              {' — '}
              <span style={{ color: 'var(--dex-green-dark)' }}>{mainFieldsEditName}</span>
            </h3>
            <button
              onClick={closeMainFieldsEdit}
              style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}
              aria-label={isDe ? 'Schließen' : 'Close'}
              disabled={mainFieldsEditSaving}
            ><X size={20} /></button>
          </div>
          <p style={{ margin: 0, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
            {isDe
              ? 'Du bearbeitest hier die Antworten auf die Felder des Hauptevents (die hellblauen Spalten „Felder des Hauptevents"). Sie werden in der Registrierung der Person auf dem Hauptevent gespeichert — nicht pro Sub-Event. Jede Änderung wird im Änderungsprotokoll mit deinem Namen und Datum festgehalten.'
              : 'You are editing the answers to the main-event fields (the light-blue „Main-event fields" columns). They are stored in the person’s registration on the main event — not per sub-event. Every change is recorded in the audit log with your name and date.'}
          </p>
          {(() => {
            const parentFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'document' && f.label && f.label.trim());
            if (parentFields.length === 0) {
              return (
                <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
                  {isDe ? 'Dieses Hauptevent hat keine bearbeitbaren Felder.' : 'This main event has no editable fields.'}
                </p>
              );
            }
            return (
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                {parentFields.map(cf => {
                  const value = mainFieldsEditForm[cf.id] || '';
                  const setVal = (v: string): void => setMainFieldsEditForm(prev => ({ ...prev, [cf.id]: v }));
                  const labelEl = (
                    <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                      {cf.label}{cf.required && <span style={{ color: 'var(--dex-red, #c00)' }}> *</span>}
                    </label>
                  );
                  if (cf.type === 'select' && !cf.multi && cf.options && cf.options.length > 0) {
                    return (
                      <div key={cf.id}>
                        {labelEl}
                        <select className="form-select" value={value} onChange={e => setVal(e.target.value)} style={{ width: '100%' }}>
                          <option value="">{isDe ? '— bitte wählen —' : '— please choose —'}</option>
                          {cf.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                        </select>
                      </div>
                    );
                  }
                  if (cf.type === 'select' && cf.multi && cf.options && cf.options.length > 0) {
                    const selected = value.split(' | ').map(s => s.trim()).filter(Boolean);
                    return (
                      <div key={cf.id} style={{ gridColumn: '1 / -1' }}>
                        {labelEl}
                        <MultiSelectDropdown
                          options={cf.options}
                          value={selected}
                          onChange={next => setVal(next.join(' | '))}
                          placeholder={isDe ? '— bitte wählen —' : '— please choose —'}
                        />
                      </div>
                    );
                  }
                  if (cf.type === 'checkbox') {
                    const isChecked = value === 'true' || value === '1';
                    return (
                      <div key={cf.id}>
                        {labelEl}
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, padding: '6px 10px', background: isChecked ? 'rgba(134,188,37,0.12)' : 'var(--dex-gray-50)', borderRadius: 6, cursor: 'pointer', fontSize: '0.85rem' }}>
                          <input
                            type="checkbox"
                            checked={isChecked}
                            onChange={e => setVal(e.target.checked ? 'true' : 'false')}
                            style={{ accentColor: 'var(--dex-green)' }}
                          />
                          {isChecked ? (isDe ? 'Ja' : 'Yes') : (isDe ? 'Nein' : 'No')}
                        </label>
                      </div>
                    );
                  }
                  if (cf.type === 'number') {
                    return (
                      <div key={cf.id}>
                        {labelEl}
                        <input className="form-input" type="number" value={value} onChange={e => setVal(e.target.value)} style={{ width: '100%' }} />
                      </div>
                    );
                  }
                  return (
                    <div key={cf.id}>
                      {labelEl}
                      <input className="form-input" value={value} onChange={e => setVal(e.target.value)} style={{ width: '100%' }} />
                    </div>
                  );
                })}
              </div>
            );
          })()}
          {mainFieldsEditError && (
            <div style={{ padding: '8px 12px', background: 'rgba(218,41,28,0.08)', border: '1px solid var(--dex-red, #c00)', borderRadius: 6, fontSize: '0.85rem', color: 'var(--dex-red, #c00)' }}>
              {mainFieldsEditError}
            </div>
          )}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
            <button type="button" className="btn btn-outline" onClick={closeMainFieldsEdit} disabled={mainFieldsEditSaving}>
              {isDe ? 'Abbrechen' : 'Cancel'}
            </button>
            <button
              type="button"
              className="btn btn-primary"
              onClick={saveMainFieldsEdit}
              disabled={mainFieldsEditSaving}
              style={{ opacity: mainFieldsEditSaving ? 0.6 : 1 }}
            >
              {mainFieldsEditSaving ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Speichern' : 'Save')}
            </button>
          </div>
        </Modal>
      )}

      {/* v19.30 (Feature B): Abmelde-Modal mit Sub-Event-Auswahl. Listet alle
          Sub-Events, für die die Person aktiv angemeldet ist, je mit Checkbox
          plus „Alle"-Schalter. Beim Bestätigen werden die gewählten Sub-Events
          abgemeldet (inkl. Mail/Outlook/Nachrücken/ID-Reorder). */}
      {deregModal && selectedEvent && (() => {
        const allChecked = deregModal.items.length > 0 && deregModal.items.every(i => deregSelected.has(i.child.id));
        const someChecked = deregModal.items.some(i => deregSelected.has(i.child.id));
        const selectedCount = deregModal.items.filter(i => deregSelected.has(i.child.id)).length;
        const toggleAll = (): void => {
          if (allChecked) setDeregSelected(new Set());
          else setDeregSelected(new Set(deregModal.items.map(i => i.child.id)));
        };
        const toggleOne = (cid: string): void => {
          setDeregSelected(prev => {
            const next = new Set(prev);
            if (next.has(cid)) next.delete(cid); else next.add(cid);
            return next;
          });
        };
        return (
          <Modal
            open={!!deregModal}
            onClose={() => { if (!deregBusy) closeDeregModal(); }}
            maxWidth={640}
            dismissable={!deregBusy}
            ariaLabel={isDe ? 'Teilnehmer abmelden' : 'Deregister attendee'}
          >
            <div className="flex-between">
              <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 8 }}>
                <Trash2 size={18} />{' '}
                {isDe ? 'Abmelden' : 'Deregister'}
                {' — '}
                <span style={{ color: 'var(--dex-green-dark)' }}>{deregModal.name}</span>
              </h3>
              <button
                onClick={closeDeregModal}
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}
                aria-label={isDe ? 'Schließen' : 'Close'}
                disabled={deregBusy}
              ><X size={20} /></button>
            </div>
            {/* Orange Sicherheits-Hinweis */}
            <div style={{ display: 'flex', gap: 10, padding: '10px 12px', background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 8, fontSize: '0.82rem', color: 'var(--dex-gray-800)' }}>
              <span style={{ color: 'var(--dex-orange, #ed8b00)', flexShrink: 0, marginTop: 1 }}><AlertCircle size={18} /></span>
              <span>
                {isDe
                  ? <>Die ausgewählten Anmeldungen werden <strong>verbindlich abgemeldet</strong>. Pro Sub-Event bekommt die Person (sofern nicht deaktiviert) eine Abmelde-Bestätigung, der Outlook-Termin wird zurückgezogen, frei werdende Plätze rücken nach und die Teilnehmer-IDs werden neu vergeben. Dieser Schritt lässt sich nicht automatisch rückgängig machen.</>
                  : <>The selected registrations will be <strong>cancelled for good</strong>. For each sub-event the person receives (unless disabled) a cancellation confirmation, the Outlook invite is withdrawn, freed seats are filled from the waitlist and participant IDs are reassigned. This step cannot be undone automatically.</>}
              </span>
            </div>
            {deregModal.items.length === 0 ? (
              <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
                {isDe ? 'Diese Person ist in keinem Sub-Event aktiv angemeldet.' : 'This person is not actively registered in any sub-event.'}
              </p>
            ) : (
              <>
                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-800)', cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={allChecked}
                    ref={el => { if (el) el.indeterminate = !allChecked && someChecked; }}
                    onChange={toggleAll}
                    style={{ accentColor: 'var(--dex-green)' }}
                    disabled={deregBusy}
                  />
                  {isDe ? 'Alle Sub-Events auswählen' : 'Select all sub-events'}
                </label>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 320, overflowY: 'auto' }}>
                  {deregModal.items.map(({ child, reg }) => {
                    const checked = deregSelected.has(child.id);
                    return (
                      <label
                        key={child.id}
                        style={{
                          display: 'flex', alignItems: 'center', gap: 10, padding: '8px 10px',
                          background: checked ? 'rgba(218,41,28,0.05)' : 'var(--dex-gray-50)',
                          border: `1px solid ${checked ? 'rgba(218,41,28,0.35)' : 'var(--dex-gray-200)'}`,
                          borderRadius: 8, cursor: deregBusy ? 'default' : 'pointer', fontSize: '0.85rem',
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleOne(child.id)}
                          style={{ accentColor: 'var(--dex-red, #c00)' }}
                          disabled={deregBusy}
                        />
                        <span style={{ flex: 1, fontWeight: 500 }}>{shortSubEventTitle(child.title, selectedEvent.title) || child.title}</span>
                        <span className={`badge ${reg.Status === 'Eingecheckt' ? 'badge-green' : 'badge-gray'}`}>{translateStatus(reg.Status, isDe)}</span>
                      </label>
                    );
                  })}
                </div>
              </>
            )}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
              <button type="button" className="btn btn-outline" onClick={closeDeregModal} disabled={deregBusy}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={runDeregModal}
                disabled={deregBusy || selectedCount === 0}
                style={{ opacity: (deregBusy || selectedCount === 0) ? 0.6 : 1, background: 'var(--dex-red, #c00)', borderColor: 'var(--dex-red, #c00)' }}
              >
                {deregBusy
                  ? (isDe ? 'Melde ab…' : 'Cancelling…')
                  : (isDe ? `Abmelden (${selectedCount})` : `Deregister (${selectedCount})`)}
              </button>
            </div>
          </Modal>
        );
      })()}

      {/* v9.37: Vorschau-Modal für die QR-Code-Mail. Rendert das wirklich
          versendete Mail-HTML in einem sandboxed iframe — analog zur Live-
          Preview im Event-Wizard unter Kommunikation. Editieren ist hier
          NICHT vorgesehen, der Body wird zentral aus der QR-Code-Vorlage
          gebaut. */}
      {qrPreviewOpen && (
        <Modal
          open={qrPreviewOpen}
          onClose={() => setQrPreviewOpen(false)}
          maxWidth={720}
          padding={0}
          ariaLabel="Vorschau: QR-Code-Mail"
        >
            <div style={{
              padding: '14px 18px', borderBottom: '1px solid var(--dex-gray-200)',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12,
            }}>
              <div style={{ minWidth: 0, flex: 1 }}>
                <h3 style={{ margin: 0, fontSize: '1.05rem' }}>Vorschau: QR-Code-Mail</h3>
                <p style={{ margin: '2px 0 0', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                  So sieht die Mail aus, die jeder angemeldete Teilnehmer bekommt — der QR-Code in der Vorschau ist auf dich ausgestellt.
                </p>
                <p style={{ margin: '6px 0 0', fontSize: '0.8rem', color: 'var(--dex-gray-700)' }}>
                  <strong>Betreff:</strong> <span style={{ color: 'var(--dex-gray-600)' }}>{qrPreviewSubject}</span>
                </p>
              </div>
              <button
                type="button"
                onClick={() => setQrPreviewOpen(false)}
                style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-500)', padding: 4 }}
                aria-label="Schließen"
              >
                <X size={22} />
              </button>
            </div>
            <div style={{ flex: 1, overflow: 'hidden', background: '#f5f5f5', padding: 12 }}>
              <iframe
                title={isDe ? 'QR-Code-Mail-Vorschau' : 'QR code email preview'}
                srcDoc={qrPreviewHtml}
                sandbox=""
                style={{ width: '100%', height: '100%', minHeight: 480, border: 'none', borderRadius: 6, background: '#fff' }}
              />
            </div>
            <div style={{ padding: '12px 18px', borderTop: '1px solid var(--dex-gray-200)', display: 'flex', justifyContent: 'flex-end' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setQrPreviewOpen(false)}
                style={{ fontSize: '0.85rem' }}
              >
                Schließen
              </button>
            </div>
        </Modal>
      )}

      {/* v17.10: Step 1 — Zielgruppen-Picker für Massenmail. Erscheint vor
          dem RichText-Editor. */}
      {massmailMode === 'pick' && selectedEvent && (() => {
        const closeAll = (): void => { setMassmailMode('closed'); setMassmailPasteRaw(''); };
        const proceed = (): void => {
          if (massmailAudience === 'custom' && massmailStatuses.size === 0) return;
          if (massmailAudience === 'nachruecker') setMassmailMode('paste');
          else { setShowEmailModal(true); setMassmailMode('editor'); }
        };
        const STATUS_OPTIONS = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
        const toggleStatus = (st: string): void => {
          setMassmailStatuses(prev => {
            const next = new Set(prev);
            if (next.has(st)) next.delete(st); else next.add(st);
            return next;
          });
        };
        const Row = (props: { value: MassmailAudience; label: string; desc: string }): React.ReactElement => (
          <label style={{
            display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
            borderRadius: 8, border: `1px solid ${massmailAudience === props.value ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
            background: massmailAudience === props.value ? 'rgba(134,188,37,0.08)' : '#fff',
            cursor: 'pointer', marginBottom: 8,
          }}>
            <input
              type="radio"
              name="massmail-target"
              checked={massmailAudience === props.value}
              onChange={() => setMassmailAudience(props.value)}
              style={{ marginTop: 3 }}
            />
            <div>
              <div style={{ fontWeight: 600, fontSize: '0.92rem' }}>{props.label}</div>
              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>{props.desc}</div>
            </div>
          </label>
        );
        return (
          <Modal open={true} onClose={closeAll} maxWidth={560} padding={24} ariaLabel="Empfänger wählen">
            <h3 style={{ margin: '0 0 14px', fontSize: '1.1rem' }}>An wen soll die Mail gehen?</h3>
            <Row value="active" label="Teilnehmer (alle aktiven)" desc="Status: Angemeldet, QR versendet, Eingecheckt — Default für die ueblichen Info-Mails." />
            <Row value="activePlusWait" label="Teilnehmer + Warteliste" desc="Alle aktiven UND Wartelistler — z.B. wenn sich noch Plätze frei machen und du auch die Warteliste vorwarnen willst." />
            <Row value="waitOnly" label="Nur Warteliste" desc={'Nur Wartelistler — z.B. Info „Es wird wahrscheinlich keinen Platz mehr geben".'} />
            <Row value="nachruecker" label="Nachrücker (Manueller Abgleich)" desc="Du fügst im nächsten Schritt eine Liste von E-Mails ein (Verteiler, Vorname Nachname Email, beliebig formatiert) — die App extrahiert die Adressen und schickt die Mail an alle aktiven Teilnehmer, die NICHT in deiner Liste stehen." />
            {/* v22.9: Eigene Status-Auswahl — einzelne Status getrennt anhaken. */}
            <div style={{
              borderRadius: 8, border: `1px solid ${massmailAudience === 'custom' ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
              background: massmailAudience === 'custom' ? 'rgba(134,188,37,0.08)' : '#fff',
              marginBottom: 8, padding: 10,
            }}>
              <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                <input type="radio" name="massmail-target" checked={massmailAudience === 'custom'} onChange={() => setMassmailAudience('custom')} style={{ marginTop: 3 }} />
                <div>
                  <div style={{ fontWeight: 600, fontSize: '0.92rem' }}>Eigene Auswahl (nach Status)</div>
                  <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>Häkchen setzen, welche Status die Mail bekommen sollen — z.B. nur „QR versendet“.</div>
                </div>
              </label>
              {massmailAudience === 'custom' && (
                <div style={{ marginTop: 10, paddingLeft: 28, display: 'flex', flexDirection: 'column', gap: 7 }}>
                  {STATUS_OPTIONS.map(st => {
                    const count = registrations.filter(r => r.Status === st).length;
                    return (
                      <label key={st} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                        <input type="checkbox" checked={massmailStatuses.has(st)} onChange={() => toggleStatus(st)} />
                        <span style={{ fontWeight: 500 }}>{st}</span>
                        <span style={{ color: 'var(--dex-gray-500)' }}>({count})</span>
                      </label>
                    );
                  })}
                  {massmailStatuses.size === 0 && (
                    <span style={{ fontSize: '0.78rem', color: 'var(--dex-orange-dark, #b35a00)' }}>Bitte mindestens einen Status anhaken.</span>
                  )}
                </div>
              )}
            </div>
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 14 }}>
              <button type="button" className="btn btn-secondary" onClick={closeAll} style={{ fontSize: '0.85rem' }}>Abbrechen</button>
              <button type="button" className="btn btn-primary" onClick={proceed} disabled={massmailAudience === 'custom' && massmailStatuses.size === 0} style={{ fontSize: '0.85rem' }}>Weiter</button>
            </div>
          </Modal>
        );
      })()}

      {/* v17.10: Step 2 (nur für 'nachruecker') — Paste-Eingabe + Extraktion */}
      {massmailMode === 'paste' && selectedEvent && (() => {
        const closeAll = (): void => { setMassmailMode('closed'); setMassmailPasteRaw(''); };
        const back = (): void => { setMassmailMode('pick'); };
        // E-Mail-Adressen aus dem Rohtext extrahieren — robust gegen Vorname
        // Nachname <mail@…> / mail@…; sep / Outlook-Verteiler-Dumps.
        const extractEmails = (raw: string): string[] => {
          if (!raw) return [];
          const matches = raw.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g) || [];
          const seen = new Set<string>();
          const out: string[] = [];
          for (const m of matches) {
            const e = m.toLowerCase();
            if (!seen.has(e)) { seen.add(e); out.push(e); }
          }
          return out;
        };
        const pasted = extractEmails(massmailPasteRaw);
        const active = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
        const pastedSet = new Set(pasted);
        const missing = active.filter(r => !pastedSet.has((r.ParticipantEmail || '').toLowerCase()));
        const continueAction = (): void => {
          if (missing.length === 0) { showAlert('Alle aktiven Teilnehmer stehen bereits in deiner Liste — niemand zum Anschreiben uebrig.'); return; }
          setShowEmailModal(true);
          setMassmailMode('editor');
        };
        return (
          <Modal open={true} onClose={closeAll} maxWidth={680} padding={24} ariaLabel="Nachrücker — Liste einfügen">
            <h3 style={{ margin: '0 0 8px', fontSize: '1.1rem' }}>Nachrücker — bestehende Empfänger-Liste einfügen</h3>
            <p style={{ margin: '0 0 14px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              Hau alles rein, was du hast — Verteiler-Export, Outlook-To-Liste, Vorname Nachname &lt;mail@deloitte.de&gt;-Format, kommagetrennt, semikolongetrennt, Zeilenumbruch — die App pickt die E-Mail-Adressen automatisch raus.
            </p>
            <textarea
              value={massmailPasteRaw}
              onChange={e => setMassmailPasteRaw(e.target.value)}
              placeholder={'Max Mustermann <mmustermann@deloitte.de>; anna.schmidt@deloitte.de; ...'}
              style={{ width: '100%', minHeight: 160, fontFamily: 'monospace', fontSize: '0.82rem', padding: 8, border: '1px solid var(--dex-gray-300)', borderRadius: 6, resize: 'vertical' }}
            />
            <div style={{ marginTop: 10, padding: 10, borderRadius: 6, background: 'var(--dex-gray-50, #fafafa)', fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
              <strong>{pasted.length}</strong> Adressen aus dem Text extrahiert.<br />
              <strong>{active.length}</strong> aktive Teilnehmer im Event.<br />
              <strong style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>{missing.length}</strong> Teilnehmer NICHT in deiner Liste — die werden angeschrieben.
            </div>
            {missing.length > 0 && (
              <details style={{ marginTop: 8, padding: 0, borderRadius: 6, background: 'rgba(237,139,0,0.06)', border: '1px solid var(--dex-orange, #ed8b00)', fontSize: '0.82rem' }}>
                <summary style={{ padding: '8px 12px', cursor: 'pointer', fontWeight: 600, color: 'var(--dex-orange-dark, #b35a00)' }}>
                  Empfänger anzeigen ({missing.length})
                </summary>
                <div style={{ maxHeight: 320, overflowY: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.8rem' }}>
                    <thead>
                      <tr style={{ borderBottom: '1px solid var(--dex-orange, #ed8b00)', background: 'rgba(255,255,255,0.6)' }}>
                        <th style={{ textAlign: 'left', padding: 6 }}>Vorname</th>
                        <th style={{ textAlign: 'left', padding: 6 }}>Nachname</th>
                        <th style={{ textAlign: 'left', padding: 6 }}>Position</th>
                        <th style={{ textAlign: 'left', padding: 6 }}>Email</th>
                      </tr>
                    </thead>
                    <tbody>
                      {missing.map(r => {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const anyR = r as any;
                        return (
                          <tr key={r.Id} style={{ borderBottom: '1px solid rgba(237,139,0,0.15)' }}>
                            <td style={{ padding: 6 }}>{r.Vorname || '-'}</td>
                            <td style={{ padding: 6 }}>{r.Nachname || '-'}</td>
                            <td style={{ padding: 6, color: 'var(--dex-gray-600)' }}>{anyR.JobTitle || '-'}</td>
                            <td style={{ padding: 6, color: 'var(--dex-gray-600)' }}>{r.ParticipantEmail}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </details>
            )}
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 14 }}>
              <button type="button" className="btn btn-secondary" onClick={back} style={{ fontSize: '0.85rem' }}>Zurück</button>
              <button type="button" className="btn btn-secondary" onClick={closeAll} style={{ fontSize: '0.85rem' }}>Abbrechen</button>
              <button type="button" className="btn btn-primary" disabled={missing.length === 0} onClick={continueAction} style={{ fontSize: '0.85rem' }}>Weiter zum Mail-Editor</button>
            </div>
          </Modal>
        );
      })()}

      {/* v17.12: Excel-Export-Zielgruppen-Picker. */}
      {excelTargetModal && selectedEvent && (() => {
        const closeAll = (): void => setExcelTargetModal(null);
        // v20.4: Im Klammer-Modus entscheidet das Modal, WAS exportiert wird —
        // konsolidierte Matrix und/oder einzelne Sub-Event-Blätter.
        const consolidatedExportPossible = isConsolidatedMode && excelTargetModal.mode === 'deloitte' && consolidatedChildren.length > 0;
        const proceed = (): void => {
          const mode = excelTargetModal.mode;
          setExcelTargetModal(null);
          if (consolidatedExportPossible && (excelIncludeMatrix || excelSubIds.size > 0)) {
            exportConsolidatedExcel(excelAudience, excelIncludeMatrix, Array.from(excelSubIds));
          } else {
            exportCsv(mode, excelAudience);
          }
        };
        const toggleSubId = (id: string): void => {
          setExcelSubIds(prev => {
            const next = new Set(prev);
            if (next.has(id)) next.delete(id); else next.add(id);
            return next;
          });
        };
        const Row = (props: { value: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled'; label: string; desc: string }): React.ReactElement => (
          <label style={{
            display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
            borderRadius: 8, border: `1px solid ${excelAudience === props.value ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
            background: excelAudience === props.value ? 'rgba(134,188,37,0.08)' : '#fff',
            cursor: 'pointer', marginBottom: 8,
          }}>
            <input type="radio" name="excel-target" checked={excelAudience === props.value} onChange={() => setExcelAudience(props.value)} style={{ marginTop: 3 }} />
            <div>
              <div style={{ fontWeight: 600, fontSize: '0.92rem' }}>{props.label}</div>
              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>{props.desc}</div>
            </div>
          </label>
        );
        return (
          <Modal open={true} onClose={closeAll} maxWidth={520} padding={24} ariaLabel="Excel-Export Zielgruppe">
            <h3 style={{ margin: '0 0 14px', fontSize: '1.1rem' }}>Excel-Export — wen sollen wir exportieren?</h3>
            <Row value="active" label="Teilnehmer (alle aktiven)" desc="Status: Angemeldet, QR versendet, Eingecheckt — Default für den Check-In / die Vor-Ort-Liste." />
            <Row value="activePlusWait" label="Teilnehmer + Warteliste" desc="Alle aktiven + Wartelistler in einem Sheet, sortiert nach TeilnehmerID." />
            <Row value="waitOnly" label="Nur Warteliste" desc="Nur die Wartelistler — z.B. für Briefing." />
            <Row value="withCancelled" label="Alles inkl. Abmeldungen" desc="Alle Einträge inklusive abgemeldeter Personen — der Status steht pro Zeile in der Status-Spalte." />
            {/* v20.4: Klammer-Modus — wählen, was in die Datei kommt. */}
            {consolidatedExportPossible && (
              <div style={{ marginTop: 14, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
                <div style={{ fontWeight: 700, fontSize: '0.92rem', marginBottom: 8 }}>
                  {isDe ? 'Was soll in die Datei?' : 'What goes into the file?'}
                </div>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, padding: '8px 10px', borderRadius: 8, border: `1px solid ${excelIncludeMatrix ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`, background: excelIncludeMatrix ? 'rgba(134,188,37,0.08)' : '#fff', cursor: 'pointer', marginBottom: 8 }}>
                  <input type="checkbox" checked={excelIncludeMatrix} onChange={e => setExcelIncludeMatrix(e.target.checked)} style={{ marginTop: 3 }} />
                  <span>
                    <span style={{ display: 'block', fontWeight: 600, fontSize: '0.9rem' }}>{isDe ? 'Konsolidierte Matrix' : 'Consolidated matrix'}</span>
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                      {isDe ? 'Eine Zeile pro Person — mit den übergreifenden Feldern und pro Sub-Event dem Status + den Sub-Event-Antworten (wie die Tabelle in der Klammer-Ansicht).' : 'One row per person — with the cross-cutting fields and per sub-event the status + the sub-event answers (like the table in the consolidated view).'}
                    </span>
                  </span>
                </label>
                <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)', margin: '6px 0 6px' }}>
                  {isDe ? 'Zusätzlich einzelne Sub-Event-Blätter:' : 'Additionally individual sub-event sheets:'}
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 180, overflowY: 'auto' }}>
                  {consolidatedChildren.map(child => {
                    const checked = excelSubIds.has(child.id);
                    const short = shortSubEventTitle(child.title, selectedEvent.title) || child.title || '?';
                    return (
                      <label key={child.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '6px 10px', borderRadius: 8, border: `1px solid ${checked ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`, background: checked ? 'rgba(134,188,37,0.08)' : '#fff', cursor: 'pointer' }}>
                        <input type="checkbox" checked={checked} onChange={() => toggleSubId(child.id)} />
                        <span style={{ fontSize: '0.88rem' }}>{short}</span>
                        <span style={{ marginLeft: 'auto', fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                          {(subEventRegsByEventId[child.id] || []).length} {isDe ? 'Einträge' : 'entries'}
                        </span>
                      </label>
                    );
                  })}
                </div>
                <div style={{ display: 'flex', gap: 12, marginTop: 6 }}>
                  <button type="button" onClick={() => setExcelSubIds(new Set(consolidatedChildren.map(c => c.id)))} style={{ background: 'none', border: 'none', color: 'var(--dex-green-dark)', fontSize: '0.78rem', fontWeight: 600, cursor: 'pointer', padding: 0 }}>
                    {isDe ? 'Alle auswählen' : 'Select all'}
                  </button>
                  <button type="button" onClick={() => setExcelSubIds(new Set())} style={{ background: 'none', border: 'none', color: 'var(--dex-gray-500)', fontSize: '0.78rem', fontWeight: 600, cursor: 'pointer', padding: 0 }}>
                    {isDe ? 'Keine' : 'None'}
                  </button>
                </div>
              </div>
            )}
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 14 }}>
              <button type="button" className="btn btn-secondary" onClick={closeAll} style={{ fontSize: '0.85rem' }}>Abbrechen</button>
              <button type="button" className="btn btn-primary" onClick={proceed} style={{ fontSize: '0.85rem' }}>Excel herunterladen</button>
            </div>
          </Modal>
        );
      })()}

      {/* ===== MASSENMAIL MODAL (HtmlEditorModal mit Toolbar, Variablen, Live-Preview) ===== */}
      {showEmailModal && selectedEvent && (() => {
        // v17.10: Empfänger-Filter abhängig vom gewählten massmailAudience.
        const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
        const recipients = (() => {
          if (massmailAudience === 'custom') {
            return registrations.filter(r => massmailStatuses.has(r.Status));
          }
          if (massmailAudience === 'waitOnly') {
            return registrations.filter(r => r.Status === 'Warteliste');
          }
          if (massmailAudience === 'activePlusWait') {
            return registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0 || r.Status === 'Warteliste');
          }
          if (massmailAudience === 'nachruecker') {
            // Aktive minus die in der eingefügten Liste enthaltenen E-Mails.
            const matches = (massmailPasteRaw || '').match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g) || [];
            const pastedSet = new Set(matches.map(m => m.toLowerCase()));
            return registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0 && !pastedSet.has((r.ParticipantEmail || '').toLowerCase()));
          }
          return registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0);
        })();
        const orgNames = (selectedEvent.organizers || []).join(', ');
        const previewVars: Record<string, string> = {
          EventTitle: selectedEvent.title,
          Organizer: orgNames,
        };
        const customLogo = (() => {
          try {
            const o = JSON.parse(selectedEvent.emailTemplateOverrides || '{}');
            return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
          } catch { return ''; }
        })();
        const sendAction = async (): Promise<void> => {
          if (!eventServiceRef || !selectedEvent) return;
          if (recipients.length === 0) { showAlert('Keine Empfänger in der gewählten Auswahl.'); return; }
          if (!(await confirmDialog(`E-Mail an ${recipients.length} Teilnehmer senden?`, { confirmLabel: isDe ? 'Senden' : 'Send' }))) return;
          setEmailSending(true);
          // Variablen einmalig auflösen (Massenmail geht an alle zusammen)
          const resolvedSubject = replacePlaceholders(emailSubject, previewVars);
          const resolvedHeading = replacePlaceholders(emailHeading, previewVars);
          const resolvedBody = replacePlaceholders(emailBody, previewVars);
          // v22.11: editierbare Unter-Überschrift (leer = "Event <Titel>").
          const resolvedSubheading = massmailSubheading.trim()
            ? replacePlaceholders(massmailSubheading, previewVars)
            : `Event ${selectedEvent.title}`;
          const fullBody = wrapTemplate('#86bc25', resolvedHeading, resolvedSubheading, resolvedBody);
          const allEmails = recipients.map(r => r.ParticipantEmail).join(';');
          // v17.10: Organizer immer auf CC (falls nicht ohnehin schon
          // unter den Empfängern). Dedup per lowercase, semicolon-join.
          const orgEmails = (selectedEvent.organizerEmails || []).filter(Boolean);
          const recipientSet = new Set(recipients.map(r => (r.ParticipantEmail || '').toLowerCase()));
          const ccList = orgEmails.filter(e => e && !recipientSet.has(e.toLowerCase()));
          const ccString = ccList.length > 0 ? ccList.join(';') : undefined;
          try {
            await eventServiceRef.queueEmail(
              resolvedSubject, allEmails, 'Alle Teilnehmer', fullBody,
              'Massenmail', selectedEvent.title, selectedEvent.id,
              ccString,
            );
            setEmailSending(false);
            const ccInfo = ccString ? ` (Organizer auf CC: ${ccList.length})` : ' (Organizer schon in Empfängerliste)';
            showAlert(`E-Mail an ${recipients.length} Empfänger in die Warteschlange eingetragen.${ccInfo}`);
            setShowEmailModal(false);
            setMassmailMode('closed');
            setMassmailPasteRaw('');
          } catch {
            setEmailSending(false);
            showAlert('Fehler beim Eintragen der E-Mail.');
          }
        };
        // v22.11: „Briefumschlag"-Kopf über der Vorschau — wie bei der
        // Einladungsmail (An: Empfängergruppe, Betreff: aufgelöster Subject).
        const audienceLabel = massmailAudience === 'custom'
          ? Array.from(massmailStatuses).join(', ')
          : massmailAudience === 'waitOnly' ? 'Nur Warteliste'
          : massmailAudience === 'activePlusWait' ? 'Teilnehmer + Warteliste'
          : massmailAudience === 'nachruecker' ? 'Nachrücker (manueller Abgleich)'
          : 'Alle aktiven Teilnehmer';
        const previewToLine = `${recipients.length} Empfänger — ${audienceLabel} · Organizer in CC`;
        const previewSubjectLine = replacePlaceholders(emailSubject, previewVars);
        return (
          <HtmlEditorModal
            open={showEmailModal}
            onClose={() => !emailSending && setShowEmailModal(false)}
            title={`Massenmail an ${recipients.length} Teilnehmer: ${selectedEvent.title}`}
            value={emailBody}
            onChange={setEmailBody}
            previewMode="email"
            emailSubject={emailSubject}
            onEmailSubjectChange={setEmailSubject}
            emailHeading={emailHeading}
            onEmailHeadingChange={setEmailHeading}
            emailSubheading={massmailSubheading}
            onEmailSubheadingChange={setMassmailSubheading}
            previewToLine={previewToLine}
            previewSubjectLine={previewSubjectLine}
            emailHeadingColor="#86bc25"
            previewVars={previewVars}
            insertableVars={[
              { key: '{{EventTitle}}', label: 'Event' },
              { key: '{{Organizer}}', label: 'Organizer' },
            ]}
            imageBase64={customLogo}
            headerExtra={(
              <div style={{ padding: 12, background: 'var(--dex-gray-50, #fafafa)', border: '1px solid var(--dex-gray-200)', borderRadius: 'var(--dex-radius)', marginBottom: 4 }}>
                <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginBottom: 8 }}>
                  {isDe
                    ? <>Geht an <strong>{recipients.length}</strong> Empfänger (die oben gewählte Gruppe). Organizer kommen automatisch auf CC.</>
                    : <>Goes to <strong>{recipients.length}</strong> recipients (the group selected above). Organizers are automatically on CC.</>}
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                  <button type="button" className="btn btn-secondary" onClick={saveMassmailDraft} disabled={emailSending} style={{ fontSize: '0.78rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                    <Check size={14} /> {isDe ? 'Entwurf speichern' : 'Save draft'}
                  </button>
                  {massmailDraftSaved && (
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, fontSize: '0.78rem' }}>
                      <Check size={14} /> {isDe ? 'Gespeichert' : 'Saved'}
                    </span>
                  )}
                  <button type="button" className="btn btn-outline" onClick={() => { sendMassmailTestToOrganizers().catch(() => { /* */ }); }} disabled={emailSending || massmailTesting} style={{ fontSize: '0.78rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                    <Send size={14} /> {massmailTesting ? (isDe ? 'Sendet…' : 'Sending…') : (isDe ? 'Testmail an Organizer' : 'Test email to organizers')}
                  </button>
                  <span style={{ flex: 1 }} />
                  <button type="button" onClick={resetMassmailDraft} disabled={emailSending} style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600, fontSize: '0.74rem', textDecoration: 'underline' }}>
                    {isDe ? 'Zurücksetzen' : 'Reset'}
                  </button>
                </div>
                {massmailTestMsg && (
                  <div style={{ marginTop: 8, fontSize: '0.78rem', color: (massmailTestMsg.indexOf('verschickt') >= 0 || massmailTestMsg.indexOf('sent') >= 0) ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-orange-dark, #b35a00)' }}>
                    {massmailTestMsg}
                  </div>
                )}
                <div style={{ marginTop: 8, fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                  {isDe
                    ? 'Dein Text wird zusätzlich automatisch gespeichert und beim nächsten Öffnen wiederhergestellt.'
                    : 'Your text is also saved automatically and restored next time you open it.'}
                </div>
              </div>
            )}
            extraAction={{
              label: emailSending ? 'Wird eingetragen…' : `An ${recipients.length} Teilnehmer senden`,
              onClick: sendAction,
              disabled: emailSending || !emailSubject.trim() || !emailBody.trim() || recipients.length === 0,
              icon: <Send size={16} />,
            }}
          />
        );
      })()}

      {/* ===== EINLADUNGSMAIL MODAL (v11.40) ===== */}
      {showInviteModal && selectedEvent && (() => {
        const audienceEmails = (selectedEvent.audienceFilter || [])
          .map(s => (s || '').trim())
          .filter(Boolean);
        const myEmail = currentUser.email || '';
        const myDisplayName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || myEmail;
        const targetEmails = inviteTarget === 'organizer' ? [myEmail].filter(Boolean) : audienceEmails;
        // v11.43: Organizer-Mails als CC mitschicken — damit alle Organizer
        // sehen, dass die Einladung raus ist und ggf. auf Rückfragen
        // antworten können. Duplikate gegenüber TO werden rausgefiltert
        // (z.B. wenn der Sender selbst Organizer ist und 'An mich' wählt).
        const toLcSet = new Set(targetEmails.map(e => (e || '').toLowerCase()));
        const ccEmails = (selectedEvent.organizerEmails || [])
          .map(s => (s || '').trim())
          .filter(Boolean)
          .filter(e => !toLcSet.has(e.toLowerCase()));
        // v11.41: Blocked-Check für den aktuell gewählten Empfänger-Modus.
        // 'organizer'-Modus blockt eigentlich nie — die eigene Mail ist immer
        // eine Person, kein Verteiler — aber wir laufen das defensiv mit.
        const blockedInTargets = getBlockedInviteRecipients(targetEmails);
        const blockedInAudience = getBlockedInviteRecipients(audienceEmails);
        const recipientLabel = inviteTarget === 'organizer'
          ? (isDe ? `An mich (${myEmail})` : `To me (${myEmail})`)
          : (isDe
            ? `An Mailverteiler (${audienceEmails.length === 0 ? 'leer' : audienceEmails.length + ' Empfänger'})`
            : `To mail distribution (${audienceEmails.length === 0 ? 'empty' : audienceEmails.length + ' recipients'})`);
        const orgNames = (selectedEvent.organizers || []).join(', ');
        const appUrl = `${siteUrl}/SitePages/DEX.aspx?env=WebView`;
        const previewVars: Record<string, string> = {
          EventTitle: selectedEvent.title,
          Organizer: orgNames,
          Link: appUrl,
        };
        const customLogo = (() => {
          try {
            const o = JSON.parse(selectedEvent.emailTemplateOverrides || '{}');
            return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
          } catch { return ''; }
        })();
        const sendAction = async (): Promise<void> => {
          if (!eventServiceRef || !selectedEvent) return;
          if (targetEmails.length === 0) {
            showAlert(isDe
              ? (inviteTarget === 'audience'
                ? 'Es ist kein Mailverteiler auf dem Event hinterlegt. Bitte zuerst in Schritt 3 (Sichtbarkeit) Empfänger ergänzen.'
                : 'Keine eigene E-Mail-Adresse verfügbar.')
              : (inviteTarget === 'audience'
                ? 'No mail distribution list configured on the event. Please add recipients in step 3 (Visibility) first.'
                : 'No own email address available.'));
            return;
          }
          // v11.41: Hart blocken — Einladungsmail darf nie an pauschale
          // Standort-/All-Verteiler ('deall', 'all', 'de.<stadt>') gehen.
          if (blockedInTargets.length > 0) {
            const lines = blockedInTargets.map(b => `• ${b.email}  (${b.reason})`).join('\n');
            showAlert(isDe
              ? `Die Einladungs-Mail darf NICHT an pauschale Standort- oder All-Verteiler verschickt werden.\n\nFolgende Empfänger sind blockiert:\n\n${lines}\n\nBitte entferne diese Adressen aus dem Mailverteiler in Schritt 3 des Event-Edits oder nutze die Option „An mich (zum Weiterleiten)".`
              : `The invitation email must NOT be sent to entire location or all-distribution lists.\n\nThe following recipients are blocked:\n\n${lines}\n\nPlease remove these addresses from the mail distribution in step 3 of event edit, or use the option "To me (for forwarding)".`);
            return;
          }
          const confirmMsg = isDe
            ? (inviteTarget === 'organizer'
              ? `Einladungs-Mail an dich selbst (${myEmail}) senden? Du kannst sie anschließend aus Outlook an deinen Verteiler weiterleiten.`
              : `Einladungs-Mail an ${audienceEmails.length} Empfänger des Mailverteilers senden?\n\n${audienceEmails.join(', ')}`)
            : (inviteTarget === 'organizer'
              ? `Send invitation email to yourself (${myEmail})? You can then forward it from Outlook to your distribution list.`
              : `Send invitation email to ${audienceEmails.length} recipients of the mail distribution?\n\n${audienceEmails.join(', ')}`);
          if (!(await confirmDialog(confirmMsg, { confirmLabel: isDe ? 'Senden' : 'Send' }))) return;
          setInviteSending(true);
          const resolvedSubject = replacePlaceholders(inviteSubject, previewVars);
          const resolvedHeading = replacePlaceholders(inviteHeading, previewVars);
          const resolvedBody = replacePlaceholders(inviteBody, previewVars);
          // v22.5: editierbare Unter-Überschrift verwenden (leer = „Event <Titel>").
          const resolvedSubheading = inviteSubheading && inviteSubheading.trim()
            ? replacePlaceholders(inviteSubheading, previewVars)
            : `Event ${selectedEvent.title}`;
          const fullBody = wrapTemplate('#86bc25', resolvedHeading, resolvedSubheading, resolvedBody);
          const allEmails = targetEmails.join(';');
          const ccString = ccEmails.join(';');
          const recipientName = inviteTarget === 'organizer' ? myDisplayName : (isDe ? 'Mailverteiler' : 'Mail distribution');
          try {
            await eventServiceRef.queueEmail(
              resolvedSubject, allEmails, recipientName, fullBody,
              'Einladung', selectedEvent.title, selectedEvent.id,
              ccString || undefined,
            );
            setInviteSending(false);
            showAlert(isDe
              ? `Einladungs-Mail an ${targetEmails.length} Empfänger in die Warteschlange eingetragen.`
              : `Invitation email queued for ${targetEmails.length} recipient(s).`);
            setShowInviteModal(false);
          } catch {
            setInviteSending(false);
            showAlert(isDe ? 'Fehler beim Eintragen der E-Mail.' : 'Error queueing the email.');
          }
        };
        const headerExtra = (
          <div style={{
            padding: 12,
            background: 'var(--dex-gray-50, #fafafa)',
            border: '1px solid var(--dex-gray-200)',
            borderRadius: 'var(--dex-radius)',
            marginBottom: 4,
          }}>
            <div style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 8 }}>
              {isDe ? 'Empfänger' : 'Recipient'}
            </div>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, cursor: 'pointer', fontSize: '0.82rem' }}>
              <input
                type="radio"
                name="inviteTarget"
                checked={inviteTarget === 'organizer'}
                onChange={() => setInviteTarget('organizer')}
                style={{ marginTop: 3 }}
              />
              <span>
                <strong>{isDe ? 'An mich — zum Weiterleiten' : 'To me — for forwarding'}</strong>
                <br />
                <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem' }}>
                  {myEmail}
                </span>
              </span>
            </label>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: audienceEmails.length === 0 ? 'not-allowed' : 'pointer', fontSize: '0.82rem', opacity: audienceEmails.length === 0 ? 0.55 : 1 }}>
              <input
                type="radio"
                name="inviteTarget"
                checked={inviteTarget === 'audience'}
                onChange={() => setInviteTarget('audience')}
                disabled={audienceEmails.length === 0}
                style={{ marginTop: 3 }}
              />
              <span style={{ flex: 1 }}>
                <strong>
                  {isDe
                    ? `An Mailverteiler des Events (${audienceEmails.length})`
                    : `To event mail distribution (${audienceEmails.length})`}
                </strong>
                <br />
                <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem', wordBreak: 'break-word' }}>
                  {audienceEmails.length === 0
                    ? (isDe
                      ? 'Kein Mailverteiler auf dem Event hinterlegt — in Schritt 3 (Sichtbarkeit) im Event-Edit ergänzen.'
                      : 'No mail distribution configured — add recipients in step 3 (Visibility) of event edit.')
                    : audienceEmails.join(', ')}
                </span>
                {blockedInAudience.length > 0 && (
                  <div style={{
                    marginTop: 6, padding: '6px 8px',
                    background: '#fef3f2', border: '1px solid #c9302c',
                    borderRadius: 6, color: '#7a1f1c',
                    fontSize: '0.75rem', lineHeight: 1.4,
                  }}>
                    <strong>
                      {isDe ? '⚠ Blockierte Empfänger im Mailverteiler:' : '⚠ Blocked recipients in the distribution list:'}
                    </strong>
                    <ul style={{ margin: '4px 0 0 16px', padding: 0 }}>
                      {blockedInAudience.map(b => (
                        <li key={b.email}><code>{b.email}</code> — {b.reason}</li>
                      ))}
                    </ul>
                    <div style={{ marginTop: 4 }}>
                      {isDe
                        ? 'Pauschale Standort- oder All-Verteiler sind für Einladungs-Mails nicht zulässig. Bitte aus dem Mailverteiler entfernen (Event-Edit, Schritt 3) — sonst wird das Senden blockiert.'
                        : 'Entire location or all-distribution lists are not allowed for invitation emails. Please remove from the distribution list (event edit, step 3) — otherwise sending is blocked.'}
                    </div>
                  </div>
                )}
              </span>
            </label>
            {ccEmails.length > 0 && (
              <div style={{
                marginTop: 10, paddingTop: 8,
                borderTop: '1px dashed var(--dex-gray-200)',
                fontSize: '0.78rem', color: 'var(--dex-gray-600)',
              }}>
                <strong style={{ color: 'var(--dex-gray-700)' }}>{isDe ? 'CC' : 'CC'}: </strong>
                <span style={{ wordBreak: 'break-word' }}>{ccEmails.join(', ')}</span>
                <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 3 }}>
                  {isDe
                    ? 'Alle Organizer dieses Events werden automatisch in CC gesetzt.'
                    : 'All organizers of this event are automatically added in CC.'}
                </div>
              </div>
            )}
            {/* v22.5/v22.6: Entwurf speichern (Button) + Auto-Speichern-Hinweis
                + Zurücksetzen. */}
            <div style={{ marginTop: 10, paddingTop: 8, borderTop: '1px dashed var(--dex-gray-200)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={saveInviteDraft}
                  style={{ fontSize: '0.78rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                >
                  <Check size={14} />
                  {isDe ? 'Entwurf speichern' : 'Save draft'}
                </button>
                {inviteDraftSaved && (
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, fontSize: '0.78rem' }}>
                    <Check size={14} /> {isDe ? 'Gespeichert' : 'Saved'}
                  </span>
                )}
                <span style={{ flex: 1 }} />
                <button
                  type="button"
                  onClick={resetInviteDraft}
                  style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600, fontSize: '0.74rem', textDecoration: 'underline' }}
                >
                  {isDe ? 'Auf Standardtext zurücksetzen' : 'Reset to default text'}
                </button>
              </div>
              <div style={{ marginTop: 6, fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                {isDe
                  ? 'Dein Text wird zusätzlich automatisch gespeichert und beim nächsten Öffnen wiederhergestellt.'
                  : 'Your text is also saved automatically and restored next time you open it.'}
              </div>
            </div>
          </div>
        );
        const previewToLine = inviteTarget === 'organizer'
          ? myEmail
          : (isDe ? `${audienceEmails.length} Empfänger des Mailverteilers` : `${audienceEmails.length} mail distribution recipients`);
        const previewSubjectLine = replacePlaceholders(inviteSubject, previewVars);
        return (
          <HtmlEditorModal
            open={showInviteModal}
            onClose={() => !inviteSending && setShowInviteModal(false)}
            title={isDe
              ? `Einladungsmail: ${selectedEvent.title}`
              : `Invitation email: ${selectedEvent.title}`}
            value={inviteBody}
            onChange={setInviteBody}
            previewMode="email"
            emailSubject={inviteSubject}
            onEmailSubjectChange={setInviteSubject}
            emailHeading={inviteHeading}
            onEmailHeadingChange={setInviteHeading}
            emailSubheading={inviteSubheading}
            onEmailSubheadingChange={setInviteSubheading}
            emailHeadingColor="#86bc25"
            previewToLine={previewToLine}
            previewSubjectLine={previewSubjectLine}
            previewVars={previewVars}
            insertableVars={[
              { key: '{{EventTitle}}', label: isDe ? 'Event-Titel' : 'Event title' },
              { key: '{{Link}}', label: isDe ? 'Anmelde-Link' : 'Registration link' },
              { key: '{{Organizer}}', label: 'Organizer' },
            ]}
            imageBase64={customLogo}
            headerExtra={headerExtra}
            extraAction={{
              label: inviteSending
                ? (isDe ? 'Wird eingetragen…' : 'Queueing…')
                : (isDe ? `Senden — ${recipientLabel}` : `Send — ${recipientLabel}`),
              onClick: sendAction,
              disabled: inviteSending
                || !inviteSubject.trim()
                || !inviteBody.trim()
                || targetEmails.length === 0,
              icon: <Send size={16} />,
            }}
          />
        );
      })()}

      {/* ===== OUTLOOK-DECLINE-CHECK MODAL (Admin only) ===== */}
      {showDeclineModal && declineResult && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 2000,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
          }}
          onClick={() => setShowDeclineModal(false)}
        >
          <div
            className="card"
            style={{ background: '#fff', maxWidth: 720, width: '100%', maxHeight: '80vh', overflow: 'auto', padding: 24 }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0 }}>Outlook-Absagen vs. Anmeldungen</h3>
              <button className="btn btn-secondary" style={{ padding: '4px 10px' }} onClick={() => setShowDeclineModal(false)}>
                Schließen
              </button>
            </div>
            {declineResult.error ? (
              <p style={{ color: 'var(--dex-red)', whiteSpace: 'pre-line' }}>{declineResult.error}</p>
            ) : declineResult.declinedAndRegistered.length === 0 ? (
              <p style={{ color: 'var(--dex-gray-600)' }}>
                Keine Diskrepanzen gefunden. {declineResult.declinedTotal > 0
                  ? `Es gibt ${declineResult.declinedTotal} Outlook-Absage(n), aber keiner davon ist in der Teilnehmerliste noch aktiv.`
                  : 'Niemand hat den Outlook-Termin abgelehnt.'}
              </p>
            ) : (
              <>
                <p style={{ color: 'var(--dex-gray-700)' }}>
                  <strong>{declineResult.declinedAndRegistered.length}</strong> Teilnehmer haben den Outlook-Termin abgelehnt,
                  stehen aber in der Teilnehmerliste noch als aktiv:
                </p>
                <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
                  <button
                    className="btn btn-secondary"
                    style={{ fontSize: '0.8rem' }}
                    onClick={() => {
                      const emails = declineResult.declinedAndRegistered.map(d => d.email).join('; ');
                      navigator.clipboard.writeText(emails).then(() => {
                        setDeclineCopied(true);
                        setTimeout(() => setDeclineCopied(false), 2000);
                      }).catch(() => showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{emails}</span>, { title: isDe ? 'E-Mail-Adressen manuell kopieren' : 'Copy email addresses manually' }));
                    }}
                  >
                    <Copy size={14} /> {declineCopied ? 'Kopiert!' : 'E-Mails kopieren'}
                  </button>
                </div>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                  <thead>
                    <tr style={{ background: 'var(--dex-gray-50)' }}>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>ID</th>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>Name</th>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>E-Mail</th>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {declineResult.declinedAndRegistered.map(d => {
                      const displayName = (d.reg.Vorname && d.reg.Nachname)
                        ? `${d.reg.Vorname} ${d.reg.Nachname}`
                        : (d.reg.ParticipantName || d.name);
                      return (
                        <tr key={d.email}>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{d.reg.TeilnehmerID ?? '-'}</td>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{displayName}</td>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{d.email}</td>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{d.reg.Status}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 12 }}>
                  Insgesamt {declineResult.declinedTotal} Outlook-Absage(n) erfasst.
                </p>
              </>
            )}
          </div>
        </div>
      )}
      </>)}

      {/* v11.0: Modal für Teilnehmer-Attachments. Liste der hochgeladenen
          Dateien mit Download-Link + Lösch-Button (Admin/Organizer kann
          fremde Uploads löschen). Plus optional eigener Upload-Button für
          den Admin (z.B. Bestätigungsbescheinigung im Namen des
          Teilnehmers anhängen). */}
      {attachmentsModalReg && (() => {
        const reg = attachmentsModalReg;
        const list = attachmentsByReg[reg.Id] || [];
        const close = (): void => setAttachmentsModalReg(null);
        // v19.0: Dokument-Feld-Attachments tragen einen `dxf-<fieldId>--`-Präfix.
        // Für die Anzeige den Präfix + Timestamp strippen und das Feld-Label
        // ermitteln, damit der Organizer sieht, zu welchem Dokument-Feld die
        // Datei gehört.
        const docFields = (selectedEvent?.eventSpecificFields || []).filter(f => f.type === 'document');
        const fieldLabelForFile = (fileName: string): string => {
          const m = fileName.match(/^dxf-([a-zA-Z0-9]+)--/);
          if (!m) return '';
          const df = docFields.find(f => (f.id || '').replace(/[^a-zA-Z0-9]/g, '') === m[1]);
          return df ? df.label : '';
        };
        const prettyFileName = (fileName: string): string =>
          fileName
            .replace(/^dxf-[a-zA-Z0-9]+--\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}_/, '')
            .replace(/^dxf-[a-zA-Z0-9]+--/, '')
            .replace(/^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}_/, '');
        const refreshOne = async (regId: number): Promise<void> => {
          if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
          try {
            const fresh = await eventServiceRef.listRegistrationAttachments(selectedEvent.subsiteUrl, regId);
            setAttachmentsByReg(prev => ({ ...prev, [regId]: fresh }));
          } catch { /* */ }
        };
        const onDelete = async (fileName: string): Promise<void> => {
          if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
          if (!(await confirmDialog(isDe ? `Datei „${fileName}" wirklich löschen?` : `Really delete file „${fileName}"?`, { danger: true, confirmLabel: isDe ? 'Löschen' : 'Delete' }))) return;
          setAttachmentsBusy(true);
          try {
            await eventServiceRef.deleteRegistrationAttachment(selectedEvent.subsiteUrl, reg.Id, fileName);
            await refreshOne(reg.Id);
          } finally { setAttachmentsBusy(false); }
        };
        const onAdd = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
          const f = e.target.files && e.target.files[0];
          e.target.value = '';
          if (!f || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
          if (f.size > 10 * 1024 * 1024) {
            showAlert(isDe ? 'Datei ist größer als 10 MB.' : 'File is larger than 10 MB.');
            return;
          }
          setAttachmentsBusy(true);
          try {
            await eventServiceRef.addRegistrationAttachment(selectedEvent.subsiteUrl, reg.Id, f);
            await refreshOne(reg.Id);
          } finally { setAttachmentsBusy(false); }
        };
        const fullName = `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantEmail || '–';
        return (
          <Modal
            open={true}
            onClose={close}
            dismissable={!attachmentsBusy}
            maxWidth={560}
            padding={24}
            ariaLabel={isDe ? 'Hochgeladene Dateien' : 'Uploaded files'}
          >
              <h3 style={{ margin: '0 0 4px', fontSize: '1.1rem' }}>
                {isDe ? 'Hochgeladene Dateien' : 'Uploaded files'}
              </h3>
              <p style={{ margin: '0 0 14px', fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                {fullName}{reg.ParticipantEmail ? ` · ${reg.ParticipantEmail}` : ''}
              </p>
              {list.length === 0 ? (
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', fontStyle: 'italic', margin: '12px 0' }}>
                  {isDe ? 'Noch keine Dateien hochgeladen.' : 'No files uploaded yet.'}
                </p>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 16 }}>
                  {list.map(f => (
                    <div key={f.fileName} style={{
                      display: 'flex', alignItems: 'center', gap: 10,
                      padding: '8px 12px', borderRadius: 6,
                      background: 'rgba(134,188,37,0.08)',
                      border: '1px solid rgba(134,188,37,0.30)',
                      fontSize: '0.85rem',
                    }}>
                      <FileText size={16} />
                      <a
                        href={f.serverRelativeUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{ flex: 1, color: 'var(--dex-gray-800)', textDecoration: 'none', wordBreak: 'break-all' }}
                      >
                        {fieldLabelForFile(f.fileName) && (
                          <span style={{ display: 'inline-block', fontSize: '0.68rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', background: 'rgba(134,188,37,0.15)', borderRadius: 4, padding: '1px 6px', marginRight: 6 }}>
                            {fieldLabelForFile(f.fileName)}
                          </span>
                        )}
                        {prettyFileName(f.fileName)}
                      </a>
                      <a
                        href={f.serverRelativeUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.72rem', padding: '2px 10px', textDecoration: 'none' }}
                      >
                        <Download size={12} /> {isDe ? 'Download' : 'Download'}
                      </a>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.72rem', padding: '2px 10px', color: 'var(--dex-red, #c00)' }}
                        disabled={attachmentsBusy}
                        onClick={() => onDelete(f.fileName)}
                        title={isDe ? 'Löschen' : 'Delete'}
                      >
                        <X size={12} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
              <div style={{ display: 'flex', gap: 8, justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap' }}>
                <label className="btn btn-outline" style={{ fontSize: '0.82rem', padding: '6px 14px', cursor: attachmentsBusy ? 'wait' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  <Plus size={14} /> {attachmentsBusy ? (isDe ? 'Wird übertragen…' : 'Uploading…') : (isDe ? 'Datei hinzufügen' : 'Add file')}
                  <input
                    type="file"
                    accept="application/pdf,image/*,.doc,.docx"
                    style={{ display: 'none' }}
                    onChange={onAdd}
                    disabled={attachmentsBusy}
                  />
                </label>
                <button className="btn btn-primary" onClick={close} disabled={attachmentsBusy}>
                  {isDe ? 'Schließen' : 'Close'}
                </button>
              </div>
          </Modal>
        );
      })()}

      {/* v11.36: Fortschritts-Overlay für die ID-Neuvergabe (mit %). */}
      {reorderProgress !== null && (
        <Modal
          open={reorderProgress !== null}
          onClose={() => { /* progress overlay — nicht schließbar */ }}
          dismissable={false}
          maxWidth={420}
          padding={28}
          ariaLabel="ID-Neuvergabe"
        >
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontWeight: 600, marginBottom: 14 }}>{reorderProgressLabel || 'IDs werden neu vergeben…'}</div>
            <div style={{ height: 12, borderRadius: 6, background: 'var(--dex-gray-200, #e5e5e5)', overflow: 'hidden' }}>
              <div
                style={{
                  height: '100%', width: `${reorderProgress}%`,
                  background: 'var(--dex-green, #86bc25)', borderRadius: 6,
                  transition: 'width 0.25s ease',
                }}
              />
            </div>
            <div style={{ marginTop: 10, fontSize: '1.1rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {reorderProgress}%
            </div>
            <div style={{ marginTop: 6, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
              Bitte warten — das Fenster nicht schließen.
            </div>
          </div>
        </Modal>
      )}

      {/* v11.70: kein Modal mehr — der Hinweis wird inline ueber der
          Teilnehmerliste angezeigt (siehe Render-Block oberhalb der
          Teilnehmer-Tabelle). */}

      {/* v23.2: Duplikat-Abmelde-Modal — beim Abmelden einer doppelt
          angemeldeten Person fragt die App, ob die Zeile STILL entfernt werden
          soll (Duplikat löschen, ohne Mail/Outlook/Nachrücken — die Person
          bleibt über ihre zweite Zeile angemeldet) oder normal abgemeldet. */}
      {dupCancelReg && selectedEvent && (() => {
        const reg = dupCancelReg;
        const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
        const teamLabel = reg.TeamName ? `„${reg.TeamName}"` : (reg.TeamId ? (isDe ? 'Team ohne Namen' : 'unnamed team') : (isDe ? 'Einzel-Anmeldung' : 'individual registration'));
        return (
          <Modal
            open={true}
            onClose={() => { if (!dupCancelBusy) setDupCancelReg(null); }}
            dismissable={!dupCancelBusy}
            maxWidth={540}
            padding={24}
            ariaLabel={isDe ? 'Doppel-Anmeldung entfernen' : 'Remove duplicate registration'}
          >
            <h3 style={{ marginTop: 0, marginBottom: 8, color: 'var(--dex-red, #c00)' }}>
              {isDe ? 'Doppel-Anmeldung entfernen' : 'Remove duplicate registration'}
            </h3>
            <p style={{ marginTop: 0, fontSize: '0.88rem', lineHeight: 1.5 }}>
              {isDe
                ? <><strong>{name}</strong> ({reg.ParticipantEmail}) ist mehrfach für dieses Event angemeldet. Du entfernst gerade die Zeile <strong>{teamLabel}</strong>.</>
                : <><strong>{name}</strong> ({reg.ParticipantEmail}) is registered more than once for this event. You are removing the row <strong>{teamLabel}</strong>.</>}
            </p>
            <p style={{ fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              {isDe
                ? 'Bei einer Dublette ist die stille Entfernung richtig: Die Zeile wird gelöscht, ohne dass eine Abmelde-Mail oder eine Outlook-Absage rausgeht und ohne dass jemand von der Warteliste nachrückt — die Person bleibt über ihre andere Anmeldung regulär dabei. Nur falls es KEINE Dublette ist, sondern eine echte Abmeldung, wähle „Normal abmelden".'
                : 'For a duplicate, silent removal is the right choice: the row is deleted without a cancellation email or Outlook removal and without promoting anyone from the waitlist — the person stays registered via their other entry. Only if this is NOT a duplicate but a real cancellation, choose „Cancel normally".'}
            </p>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 18 }}>
              <button
                type="button"
                className="btn btn-primary"
                style={{ background: 'var(--dex-red, #c00)', borderColor: 'var(--dex-red, #c00)' }}
                disabled={dupCancelBusy}
                onClick={async () => {
                  setDupCancelBusy(true);
                  await performSilentDuplicateDelete(reg);
                  setDupCancelBusy(false);
                  setDupCancelReg(null);
                  showAlert(isDe ? 'Doppelte Anmeldung still entfernt.' : 'Duplicate registration silently removed.', { variant: 'success' });
                }}
              >
                {dupCancelBusy ? (isDe ? 'Wird entfernt…' : 'Removing…') : (isDe ? 'Duplikat still entfernen (empfohlen)' : 'Silently remove duplicate (recommended)')}
              </button>
              <button
                type="button"
                className="btn btn-secondary"
                disabled={dupCancelBusy}
                onClick={async () => {
                  setDupCancelBusy(true);
                  await performStandardCancel(reg);
                  setDupCancelBusy(false);
                  setDupCancelReg(null);
                }}
              >
                {isDe ? 'Normal abmelden (mit Mail & Nachrücken)' : 'Cancel normally (with email & promotion)'}
              </button>
              <button
                type="button"
                className="btn btn-secondary"
                style={{ border: 'none' }}
                disabled={dupCancelBusy}
                onClick={() => setDupCancelReg(null)}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
            </div>
          </Modal>
        );
      })()}

      {/* v11.36: Überbuchungs-Entscheidungs-Modal (Bestätigen / Platz behalten) */}
      {overbookModal && selectedEvent && (
        <Modal
          open={true}
          onClose={() => setOverbookModal(null)}
          dismissable={!obBusy}
          maxWidth={560}
          padding={24}
          ariaLabel="Überbuchung"
        >
            {overbookModal.mode === 'confirm' ? (
              <>
                <h3 style={{ marginTop: 0 }}>
                  Auf Warteliste bestätigen ({overbookModal.targets.length})
                </h3>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                  {overbookModal.targets.length === 1
                    ? `${(overbookModal.targets[0].Vorname && overbookModal.targets[0].Nachname) ? `${overbookModal.targets[0].Vorname} ${overbookModal.targets[0].Nachname}` : overbookModal.targets[0].ParticipantName} wird auf die Warteliste der Gruppe zurückgesetzt. Im Audit-Log wird vermerkt, dass die Person fälschlich angemeldet war.`
                    : `${overbookModal.targets.length} Personen werden auf die Warteliste zurückgesetzt. Im Audit-Log jeder Person wird der Vorgang vermerkt.`}
                </p>
                <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.88rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="checkbox" checked={obWithMail} onChange={e => setObWithMail(e.target.checked)} disabled={obBusy} />
                  Mit Entschuldigungs-Mail (Deloitte-Layout, in die Mail-Queue)
                </label>
                {obWithMail && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, margin: '0 0 10px', fontSize: '0.82rem' }}>
                    <span style={{ color: 'var(--dex-gray-600)' }}>Sprache:</span>
                    {(['DE', 'EN'] as const).map(lng => (
                      <button
                        key={lng}
                        type="button"
                        className="btn btn-secondary"
                        disabled={obBusy}
                        onClick={() => setObMailLang(lng)}
                        style={{
                          fontSize: '0.75rem', padding: '3px 12px',
                          ...(obMailLang === lng ? { background: 'var(--dex-green, #86bc25)', color: '#fff', fontWeight: 600 } : {}),
                        }}
                      >
                        {lng === 'DE' ? 'Deutsch' : 'English'}
                      </button>
                    ))}
                  </div>
                )}
                {obWithMail && overbookModal.targets.length === 1 && (
                  <div style={{ marginBottom: 10 }}>
                    <input
                      className="form-input"
                      value={obMailSubject}
                      onChange={e => setObMailSubject(e.target.value)}
                      disabled={obBusy}
                      style={{ width: '100%', marginBottom: 6, padding: '6px 10px', fontSize: '0.82rem' }}
                    />
                    <textarea
                      value={obMailBody}
                      onChange={e => setObMailBody(e.target.value)}
                      disabled={obBusy}
                      rows={5}
                      style={{ width: '100%', fontFamily: 'monospace', fontSize: '0.72rem', padding: 8 }}
                    />
                    <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', margin: '4px 0 8px' }}>
                      Vorschlagstext — editierbar. Wird in die Mail-Queue gelegt, nicht direkt versendet.
                    </p>
                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)', marginBottom: 4 }}>Vorschau (echte Deloitte-Mail):</div>
                    <div
                      style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 6, maxHeight: 280, overflow: 'auto', background: '#fff' }}
                      dangerouslySetInnerHTML={{
                        __html: obMailBody
                          .replace(/\{\{LOGO_URL\}\}/g, getCachedLogoBase64() || '')
                          .replace(/\{\{ORB_URL\}\}/g, getCachedOrbBase64() || ''),
                      }}
                    />
                  </div>
                )}
                {obWithMail && overbookModal.targets.length > 1 && (
                  // v13.0: Preview teilt sich den obMailBody-State mit der
                  // Modal-Open-useEffect — beide rendern den Body der
                  // ersten Person. Vorher wurde buildOverbookApologyEmail
                  // synchron im Render aufgerufen; seit der Template-DB-
                  // Lookup async ist, geht das nicht mehr direkt im JSX.
                  <div style={{ marginBottom: 10 }}>
                    <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', margin: '0 0 6px' }}>
                      Bei &bdquo;Alle&ldquo; wird der Standardtext je Person personalisiert versendet (eigene Wartelisten-Position). Vorschau am Beispiel der ersten Person:
                    </p>
                    <div
                      style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 6, maxHeight: 260, overflow: 'auto', background: '#fff' }}
                      dangerouslySetInnerHTML={{
                        __html: obMailBody
                          .replace(/\{\{LOGO_URL\}\}/g, getCachedLogoBase64() || '')
                          .replace(/\{\{ORB_URL\}\}/g, getCachedOrbBase64() || ''),
                      }}
                    />
                  </div>
                )}
                <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.88rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="checkbox" checked={obRemoveCalendar} onChange={e => setObRemoveCalendar(e.target.checked)} disabled={obBusy} />
                  Vom Kalendereintrag abmelden (falls vorhanden)
                </label>
              </>
            ) : (
              <>
                <h3 style={{ marginTop: 0 }}>Platz behalten ({overbookModal.targets.length})</h3>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                  Wie soll die Person ihren Platz behalten?
                </p>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: '0.85rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="radio" name="obkeep" checked={obKeepVariant === 'firstWaitlist'} onChange={() => setObKeepVariant('firstWaitlist')} disabled={obBusy} style={{ marginTop: 3 }} />
                  <span><strong>Erste(r) auf der Warteliste</strong> — rückt beim nächsten frei werdenden Platz der Gruppe garantiert als Erste(r) nach (risikoarm).</span>
                </label>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: '0.85rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="radio" name="obkeep" checked={obKeepVariant === 'active'} onChange={() => setObKeepVariant('active')} disabled={obBusy} style={{ marginTop: 3 }} />
                  <span><strong>Bleibt angemeldet</strong> (als Letzte(r)) — Gruppe bleibt +1, der nächste frei werdende Platz wird einmal nicht nachgerückt, bis die Überzahl absorbiert ist.</span>
                </label>
              </>
            )}
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 18 }}>
              <button className="btn btn-secondary" onClick={() => setOverbookModal(null)} disabled={obBusy}>
                Abbrechen
              </button>
              <button className="btn btn-primary" onClick={() => { runOverbookResolution().catch(() => { /* */ }); }} disabled={obBusy}>
                {obBusy ? 'Wird ausgeführt…' : (overbookModal.mode === 'confirm' ? 'Bestätigen & IDs neu vergeben' : 'Übernehmen & IDs neu vergeben')}
              </button>
            </div>
        </Modal>
      )}

      {adminAddMemberDialog && selectedEvent && (() => {
        // v17.2: Quick-Pick aus bereits registrierten Personen ohne Team —
        // damit der Organizer nicht via Graph-Suche jeden neu picken muss,
        // wenn die Person ohnehin schon angemeldet ist.
        const teamlessActiveLocal = registrations.filter(r =>
          (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt')
          && !r.TeamId);
        const closeDlg = (): void => {
          setAdminAddMemberDialog(null);
          setAdminAddMemberPick(null);
          setAdminAddMemberQuery('');
          setAdminAddMemberResults([]);
          setAdminAddMemberConsent(false);
          setAdminAddMemberError('');
          setAdminAddMemberBusy(false);
          setAdminAddTeamlessPicks(new Set());
          setAdminAddLeadRegId(null);
          setAdminAddSendMail(false);
          setAdminAddCcOrganizer(false);
          setAdminAddNotifyOthers(false);
          setAdminAddNotifyScope('all');
          setAdminAddNewPersonMail(true);
        };
        // v17.4: Logik zur Auswertung der Multi-Pick + (optionalem) Graph-Pick.
        const hasMultiPicks = adminAddTeamlessPicks.size > 0;
        const hasGraphPick = !!adminAddMemberPick;
        // Wenn ausschliesslich teamlose Picks: keine Consent-Box (Person hat
        // bei der eigenen Anmeldung bereits zugestimmt). Sobald aber eine
        // NEUE Person via Graph-Suche dabei ist, bleibt die Consent-Pflicht.
        const onlyTeamlessPicks = hasMultiPicks && !hasGraphPick;
        const consentRequired = !onlyTeamlessPicks && hasGraphPick;
        // v22.40: Auswahl-/Kapazitäts-Zählung für Belegung-Anzeige,
        // Über-Kapazitäts-Sperre und Button-Aktivierung.
        const totalPicks = adminAddTeamlessPicks.size + (hasGraphPick ? 1 : 0);
        const freeSlots = adminAddMemberDialog.freeSlots;
        const atCap = freeSlots > 0 && totalPicks >= freeSlots;
        const overCap = freeSlots > 0 && totalPicks > freeSlots;
        const submit = async (): Promise<void> => {
          if (!adminAddMemberDialog || adminAddMemberBusy) return;
          if (!hasMultiPicks && !hasGraphPick) return;
          if (consentRequired && !adminAddMemberConsent) return;
          setAdminAddMemberBusy(true);
          setAdminAddMemberError('');
          try {
            const tid = adminAddMemberDialog.teamId;
            const tName = adminAddMemberDialog.teamName || undefined;
            let assignedCount = 0;
            // 1) Teamlose Picks zuordnen (PATCH only). v22.40: Wenn die
            // „Info-Mail"-Checkbox an ist, bekommt jede zugeordnete Person
            // die Mail direkt — die Empfänger-Daten kommen aus der bereits
            // geladenen Registrierungs-Zeile (kein erneutes Eingeben nötig).
            for (const regId of Array.from(adminAddTeamlessPicks)) {
              const isLead = adminAddLeadRegId === regId;
              const reg = teamlessActiveLocal.find(p => p.Id === regId);
              try {
                const ok = await assignTeamlessToTeam(selectedEvent.id, tid, tName, regId, isLead, {
                  sendMail: adminAddSendMail,
                  recipientEmail: reg?.ParticipantEmail,
                  recipientFirstName: reg?.Vorname,
                  recipientLastName: reg?.Nachname,
                  ccEmail: (adminAddSendMail && adminAddCcOrganizer) ? currentUser.email : undefined,
                });
                if (ok) assignedCount++;
              } catch (err) { console.warn('[DEX] assignTeamlessToTeam failed for', regId, err); }
            }
            // 2) Falls noch ein Graph-Pick dabei: addTeamMember (neuer Insert).
            // v22.49: Kommunikation an die neue Person optional
            // (adminAddNewPersonMail); die „übrige Mitglieder"-Info wird hier
            // unterdrückt und unten zentral (mit Reichweite alle/Lead) gesteuert.
            if (hasGraphPick && adminAddMemberPick) {
              const res = await addTeamMember(selectedEvent.id, tid, tName, adminAddMemberPick, undefined, {
                suppressMemberMail: !adminAddNewPersonMail,
                suppressOthersMail: true,
                ccEmail: (adminAddNewPersonMail && adminAddCcOrganizer) ? currentUser.email : undefined,
              });
              if (!res.ok) {
                if (res.reason && res.reason.startsWith('already-registered')) {
                  setAdminAddMemberError('Person bereits beim Event angemeldet — Picker aus „Bereits angemeldet"-Liste benutzen.');
                } else if (res.reason === 'team-full') {
                  setAdminAddMemberError('Das Team ist bereits voll.');
                } else {
                  setAdminAddMemberError('Hinzufügen fehlgeschlagen.');
                }
                setAdminAddMemberBusy(false);
                return;
              }
              assignedCount++;
            }
            // v22.49: „Neues Mitglied"-Info an die übrigen Team-Mitglieder
            // (Reichweite alle / nur Lead), sofern gewählt. excludeEmails =
            // die gerade neu hinzugefügten Personen (nicht sich selbst melden).
            if (adminAddNotifyOthers) {
              const assignedRegs = Array.from(adminAddTeamlessPicks)
                .map(id => teamlessActiveLocal.find(p => p.Id === id))
                .filter(Boolean) as SPRegistration[];
              const newNames = assignedRegs.map(r => `${r.Vorname || ''} ${r.Nachname || ''}`.trim() || r.ParticipantName || r.ParticipantEmail);
              const excludeEmails = assignedRegs.map(r => r.ParticipantEmail || '').filter(Boolean);
              if (hasGraphPick && adminAddMemberPick) {
                newNames.push(adminAddMemberPick.displayName || adminAddMemberPick.email);
                excludeEmails.push(adminAddMemberPick.email);
              }
              if (newNames.length > 0) {
                try { await notifyExistingTeamMembers(selectedEvent.id, tid, tName, newNames, excludeEmails, adminAddNotifyScope); }
                catch (err) { console.warn('[DEX] notifyExistingTeamMembers failed:', err); }
              }
            }
            const teamLabel = tName ? `„${tName}"` : 'das Team';
            const toastMsg = adminAddSendMail
              ? `${assignedCount} ${assignedCount === 1 ? 'Person' : 'Personen'} ${teamLabel} zugeordnet — Info-Mail wird versendet.`
              : `${assignedCount} ${assignedCount === 1 ? 'Person' : 'Personen'} ${teamLabel} zugeordnet (ohne Mail-Versand).`;
            setTeamsToast(toastMsg);
            // TODO v17.5: Wenn adminAddSendMail=true UND assignTeamlessToTeam-
            // Pfad genutzt wurde, hier explizit eine „Du bist jetzt im Team
            // <Name>"-Mail queuen. Aktuell läuft die Mail nur ueber den
            // addTeamMember-Pfad (Graph-Pick) automatisch.
            window.setTimeout(() => setTeamsToast(''), 4500);
            const regs = await getAllRegistrations(selectedEvent.id);
            setRegistrations(regs);
            closeDlg();
          } catch {
            setAdminAddMemberError('Hinzufügen fehlgeschlagen.');
            setAdminAddMemberBusy(false);
          }
        };
        return (
          <Modal
            open={true}
            onClose={closeDlg}
            dismissable={!adminAddMemberBusy}
            maxWidth={540}
            ariaLabel="Person zum Team hinzufügen"
          >
              <h3 style={{ margin: 0, fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
                {adminAddMemberDialog.isNewTeam
                  ? 'Neues Team anlegen — Mitglieder zuordnen'
                  : adminAddMemberDialog.teamName
                    ? `Mitglieder zum Team „${adminAddMemberDialog.teamName}" hinzufügen`
                    : 'Mitglieder zum Team hinzufügen'}
              </h3>
              <div style={{ fontSize: '0.85rem', color: overCap ? 'var(--dex-red, #c00)' : 'var(--dex-gray-600)' }}>
                {/* v22.40: Belegung berücksichtigt die aktuelle Auswahl live
                    (bisherige Belegung + ausgewählte Personen). */}
                {(selectedEvent.teamSize || 0) > 0
                  ? `Team-Belegung: ${((selectedEvent.teamSize || 0) - freeSlots) + totalPicks}/${selectedEvent.teamSize}${totalPicks > 0 ? ' (inkl. Auswahl)' : ''}${overCap ? ' — zu viele ausgewählt!' : ''}`
                  : 'Belegung wird nach dem Hinzufügen aktualisiert.'}
              </div>
              {/* v17.1: Team-Name-Eingabe nur im „Neues Team anlegen"-Flow.
                  Optional — wenn leer, bekommt das Team beim Insert keinen
                  Namen, der Lead kann ihn aber später nicht mehr setzen,
                  daher direkt hier abfragen. */}
              {adminAddMemberDialog.isNewTeam && (
                <div style={{ marginTop: 4 }}>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                    Team-Name {selectedEvent.askTeamName ? <span style={{ color: 'var(--dex-red, #c00)' }}>*</span> : <span style={{ color: 'var(--dex-gray-400)', fontWeight: 400 }}>(optional)</span>}
                  </label>
                  <input
                    type="text"
                    className="form-input"
                    placeholder="z.B. „Borntowin"
                    value={adminAddMemberDialog.teamName}
                    onChange={e => setAdminAddMemberDialog(d => d ? { ...d, teamName: e.target.value } : d)}
                    style={{ width: '100%' }}
                  />
                </div>
              )}
              {/* v17.4: Consent-Box nur, wenn eine wirklich NEUE Person
                  via Graph hinzugefügt wird. Bei reiner Team-Zuordnung
                  schon-angemeldeter Personen brauchen wir keine zusätzliche
                  Zustimmung — die haben sie bei der eigenen Anmeldung
                  bereits gegeben. */}
              {consentRequired ? (
                <div style={{
                  padding: '14px 16px',
                  background: 'rgba(237,139,0,0.10)',
                  border: '2px solid var(--dex-orange, #ed8b00)',
                  borderRadius: 8,
                  color: '#7a4a00',
                  fontSize: '0.88rem',
                  lineHeight: 1.5,
                }}>
                  <div style={{ fontWeight: 700, marginBottom: 6 }}>
                    Vorab die Zustimmung des Mitglieds einholen
                  </div>
                  <div>
                    {'Mit dem Hinzufügen meldest du diese Person an. Sie erhält automatisch '}
                    {'eine Anmeldebestätigung per Mail, einen Outlook-Termin und sieht das '}
                    {'Event in „Meine Events". Bitte stelle sicher, dass die Person ihrer '}
                    {'Anmeldung '}<strong>vorher zugestimmt</strong>{' hat.'}
                  </div>
                </div>
              ) : (onlyTeamlessPicks && (
                <div style={{
                  padding: '10px 14px',
                  background: 'rgba(33,150,243,0.06)',
                  border: '1px solid var(--dex-info, #2196f3)',
                  borderRadius: 8,
                  color: 'var(--dex-gray-700)',
                  fontSize: '0.82rem',
                  lineHeight: 1.5,
                }}>
                  Du ordnest bereits-angemeldete Teilnehmer einem Team zu — keine neue Anmeldung, keine Bestätigungsmail an die Personen (es sei denn du hakst &bdquo;Info-Mail an die zugeordneten&hellip;&ldquo; unten an).
                </div>
              ))}
              <div>
                {/* v22.45: Drei klare Abschnitte — 1. Bestehende Teilnehmer
                    (bereits angemeldet, nur zuordnen), 2. Neue Teilnehmer (per
                    Suche stellvertretend anmelden), 3. Kommunikation ans Team. */}
                {teamlessActiveLocal.length > 0 && (
                  <div style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-green-dark, #4a7c1f)', marginBottom: 4 }}>
                    1 · Bestehende Teilnehmer
                  </div>
                )}
                {/* v17.4: Multi-Select aus bereits registrierten Personen
                    ohne Team. Checkbox-Liste; bei Mehrfach-Auswahl
                    erscheint zusätzlich die Lead-Radio-Auswahl. */}
                {teamlessActiveLocal.length > 0 && (
                  <div style={{ marginBottom: 12, padding: 10, border: '1px dashed var(--dex-orange, #ed8b00)', borderRadius: 6, background: 'rgba(237,139,0,0.04)' }}>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 600, marginBottom: 6 }}>
                      Bereits angemeldet ohne Team ({teamlessActiveLocal.length}) — mehrere auswählbar:
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4, maxHeight: 280, overflowY: 'auto' }}>
                      {teamlessActiveLocal.map(p => {
                        const nm = `${p.Vorname || ''} ${p.Nachname || ''}`.trim() || p.ParticipantName || p.ParticipantEmail;
                        const isPicked = adminAddTeamlessPicks.has(p.Id);
                        const isLead = adminAddLeadRegId === p.Id;
                        return (
                          <div
                            key={p.Id}
                            style={{
                              display: 'flex', alignItems: 'center', gap: 8,
                              padding: '6px 10px',
                              border: `1px solid ${isPicked ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-200)'}`,
                              borderRadius: 6,
                              background: isPicked ? 'rgba(237,139,0,0.08)' : '#fff',
                            }}
                          >
                            <input
                              type="checkbox"
                              checked={isPicked}
                              // v22.40: Über-Kapazitäts-Sperre — nicht mehr als
                              // freie Plätze auswählbar; bereits Gewählte bleiben
                              // abwählbar.
                              disabled={!isPicked && atCap}
                              onChange={e => {
                                setAdminAddTeamlessPicks(prev => {
                                  const next = new Set(prev);
                                  if (e.target.checked) next.add(p.Id);
                                  else next.delete(p.Id);
                                  return next;
                                });
                                // Wenn Lead deselektiert wurde: Lead zurücksetzen.
                                if (!e.target.checked && adminAddLeadRegId === p.Id) setAdminAddLeadRegId(null);
                              }}
                              style={{ flexShrink: 0, cursor: (!isPicked && atCap) ? 'not-allowed' : 'pointer' }}
                            />
                            <img
                              src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(p.ParticipantEmail)}&size=S`}
                              alt={nm}
                              onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                              style={{ width: 24, height: 24, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                            />
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontSize: '0.82rem', fontWeight: 500 }}>{nm}</div>
                              <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{p.ParticipantEmail}</div>
                            </div>
                            {isPicked && (
                              <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: '0.72rem', color: 'var(--dex-gray-700)', cursor: 'pointer' }}>
                                <input
                                  type="radio"
                                  name="lead-pick"
                                  checked={isLead}
                                  onChange={() => setAdminAddLeadRegId(p.Id)}
                                  style={{ margin: 0 }}
                                />
                                Lead
                              </label>
                            )}
                          </div>
                        );
                      })}
                    </div>
                    {adminAddTeamlessPicks.size > 0 && (
                      <div style={{ marginTop: 8, display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.78rem', color: 'var(--dex-gray-700)', flexWrap: 'wrap' }}>
                        <strong>{adminAddTeamlessPicks.size}</strong> ausgewählt
                        {!adminAddLeadRegId && adminAddTeamlessPicks.size > 0 && (
                          <span style={{ color: 'var(--dex-gray-500)' }}>
                            — bitte einen Lead markieren (oder leer = kein Lead).
                          </span>
                        )}
                        {atCap && (
                          <span style={{ color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 600 }}>
                            — Team voll ({freeSlots} {freeSlots === 1 ? 'Platz' : 'Plätze'}).
                          </span>
                        )}
                      </div>
                    )}
                    <div style={{ marginTop: 8, fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                      Oder weiter unten via Suche eine zusätzliche neue Person hinzufügen.
                    </div>
                  </div>
                )}
                {/* v22.45: 2 · Neue Teilnehmer — Person, die noch NICHT beim
                    Event angemeldet ist, per Suche stellvertretend hinzufügen.
                    (Kommunikation/Info-Mail folgt als Abschnitt 3 weiter unten.) */}
                <div style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-green-dark, #4a7c1f)', marginTop: teamlessActiveLocal.length > 0 ? 14 : 0, marginBottom: 2 }}>
                  {teamlessActiveLocal.length > 0 ? '2 · Neue Teilnehmer' : 'Neue Teilnehmer'}
                </div>
                <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginBottom: 6 }}>
                  Jemand, der noch nicht beim Event angemeldet ist — per Suche hinzufügen (wird stellvertretend angemeldet).
                </div>
                {adminAddMemberPick ? (
                  <div style={{
                    display: 'inline-flex', alignItems: 'center', gap: 10,
                    padding: '6px 10px 6px 6px',
                    border: '1px solid var(--dex-gray-200)',
                    borderRadius: 'var(--dex-radius)',
                    background: 'var(--dex-gray-50, #f7f7f7)',
                    maxWidth: '100%',
                  }}>
                    <img
                      src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(adminAddMemberPick.email)}&size=S`}
                      alt={adminAddMemberPick.displayName}
                      onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                      style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                    />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: '0.9rem' }}>{adminAddMemberPick.displayName}</div>
                      <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem' }}>{adminAddMemberPick.email}</div>
                    </div>
                    <button
                      type="button"
                      onClick={() => { setAdminAddMemberPick(null); setAdminAddMemberQuery(''); setAdminAddMemberResults([]); }}
                      title={isDe ? 'Auswahl entfernen' : 'Remove selection'}
                      style={{
                        background: 'var(--dex-gray-200)', border: 'none', color: 'var(--dex-gray-700)',
                        width: 22, height: 22, borderRadius: '50%', cursor: 'pointer',
                        fontSize: '0.9rem', lineHeight: 1,
                      }}
                    >×</button>
                  </div>
                ) : (
                  <div style={{ position: 'relative' }}>
                    <input
                      className="form-input"
                      value={adminAddMemberQuery}
                      // v22.40: Suche sperren, wenn das Team durch die Auswahl
                      // bereits voll ist (keine zusätzliche neue Person mehr).
                      disabled={atCap}
                      placeholder={atCap ? 'Team voll — keine weitere Person' : 'Name oder E-Mail eingeben…'}
                      onChange={e => {
                        const val = e.target.value;
                        setAdminAddMemberQuery(val);
                        if (adminAddMemberQueryTimer.current) clearTimeout(adminAddMemberQueryTimer.current);
                        if (val.length >= 2) {
                          adminAddMemberQueryTimer.current = setTimeout(async () => {
                            setAdminAddMemberSearching(true);
                            try {
                              const res = await searchUsers(val, adminAddMemberIncludeIntl);
                              setAdminAddMemberResults(res.map(r => ({ email: r.email, displayName: r.displayName })));
                            } catch { setAdminAddMemberResults([]); }
                            setAdminAddMemberSearching(false);
                          }, 300);
                        } else {
                          setAdminAddMemberResults([]);
                        }
                      }}
                    />
                    <InternationalSearchToggle
                      checked={adminAddMemberIncludeIntl}
                      onChange={setAdminAddMemberIncludeIntl}
                      isDe={isDe}
                    />
                    {(adminAddMemberResults.length > 0 || adminAddMemberSearching) && (
                      <div style={{
                        position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 10,
                        background: '#fff', border: '1px solid var(--dex-gray-200)',
                        borderRadius: 6, marginTop: 4, maxHeight: 220, overflowY: 'auto',
                        boxShadow: '0 4px 12px rgba(0,0,0,0.12)',
                      }}>
                        {adminAddMemberSearching && (
                          <div style={{ padding: 10, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
                            Suche…
                          </div>
                        )}
                        {adminAddMemberResults.map(r => (
                          <button
                            key={r.email}
                            type="button"
                            onClick={() => { setAdminAddMemberPick(r); setAdminAddMemberResults([]); setAdminAddMemberQuery(''); }}
                            style={{
                              display: 'flex', alignItems: 'center', gap: 8,
                              width: '100%', padding: '6px 10px', border: 'none',
                              background: '#fff', cursor: 'pointer', textAlign: 'left',
                            }}
                          >
                            <img
                              src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(r.email)}&size=S`}
                              alt={r.displayName}
                              onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                              style={{ width: 28, height: 28, borderRadius: '50%' }}
                            />
                            <div>
                              <div style={{ fontSize: '0.85rem', fontWeight: 600 }}>{r.displayName}</div>
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{r.email}</div>
                            </div>
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                )}
                {/* v22.45/v22.49: 3 · Kommunikation an das Team. */}
                {(adminAddTeamlessPicks.size > 0 || hasGraphPick) && (
                  <>
                    <div style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-green-dark, #4a7c1f)', marginTop: 14, marginBottom: 4 }}>
                      3 · Kommunikation an das Team
                    </div>
                    {/* a) Neue Person (Graph-Pick): Anmeldebestätigung + Outlook
                        optional (Default an — echte Neu-Anmeldung). */}
                    {hasGraphPick && (
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, fontSize: '0.82rem', cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={adminAddNewPersonMail}
                          onChange={e => setAdminAddNewPersonMail(e.target.checked)}
                          style={{ marginTop: 2 }}
                        />
                        <span>
                          Anmeldebestätigung &amp; Kalendereinladung an die neue Person senden
                          <span style={{ display: 'block', fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                            Default an — die Person wird ja neu angemeldet. Abwählen = still hinzufügen.
                          </span>
                        </span>
                      </label>
                    )}
                    {/* b) Info-Mail an die zugeordneten (bereits angemeldeten) Personen. */}
                    {adminAddTeamlessPicks.size > 0 && (
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, fontSize: '0.82rem', cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={adminAddSendMail}
                          onChange={e => setAdminAddSendMail(e.target.checked)}
                          style={{ marginTop: 2 }}
                        />
                        <span>
                          Info-Mail an die zugeordneten Team-Mitglieder versenden
                          <span style={{ display: 'block', fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                            Default aus — die Person ist ja bereits beim Event angemeldet.
                          </span>
                        </span>
                      </label>
                    )}
                    {(adminAddSendMail || (hasGraphPick && adminAddNewPersonMail)) && (
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, marginLeft: 24, fontSize: '0.82rem', cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={adminAddCcOrganizer}
                          onChange={e => setAdminAddCcOrganizer(e.target.checked)}
                          style={{ marginTop: 2 }}
                        />
                        <span>
                          Bestätigungsmail als Kopie (CC) an mich
                          <span style={{ display: 'block', fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                            {currentUser.email} bekommt diese Mail(s) in Kopie.
                          </span>
                        </span>
                      </label>
                    )}
                    {/* c) Übrige Team-Mitglieder informieren — Reichweite alle / nur Lead. */}
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 4, fontSize: '0.82rem', cursor: 'pointer' }}>
                      <input
                        type="checkbox"
                        checked={adminAddNotifyOthers}
                        onChange={e => setAdminAddNotifyOthers(e.target.checked)}
                        style={{ marginTop: 2 }}
                      />
                      <span>
                        Auch die übrigen Team-Mitglieder informieren
                        <span style={{ display: 'block', fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                          Schickt den bisherigen Mitgliedern eine „neues Mitglied“-Info.
                        </span>
                      </span>
                    </label>
                    {adminAddNotifyOthers && (
                      <div style={{ display: 'flex', gap: 16, marginLeft: 24, marginBottom: 8, fontSize: '0.82rem' }}>
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: 'pointer' }}>
                          <input type="radio" name="notifyScope" checked={adminAddNotifyScope === 'all'} onChange={() => setAdminAddNotifyScope('all')} style={{ margin: 0 }} />
                          Alle Mitglieder
                        </label>
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: 'pointer' }}>
                          <input type="radio" name="notifyScope" checked={adminAddNotifyScope === 'lead'} onChange={() => setAdminAddNotifyScope('lead')} style={{ margin: 0 }} />
                          Nur den Team-Lead
                        </label>
                      </div>
                    )}
                  </>
                )}
              </div>
              {/* v17.4: Consent-Checkbox nur bei wirklich neuer Person via Graph. */}
              {consentRequired && (
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
                  <input
                    type="checkbox"
                    checked={adminAddMemberConsent}
                    onChange={e => setAdminAddMemberConsent(e.target.checked)}
                    style={{ marginTop: 3 }}
                  />
                  <span>
                    <span style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                    Ich bestätige, dass die Person ihrer Anmeldung zugestimmt hat.
                  </span>
                </label>
              )}
              {adminAddMemberError && (
                <div style={{ padding: 10, borderRadius: 6, background: 'rgba(220,38,38,0.10)', color: '#b91c1c', fontSize: '0.85rem' }}>
                  {adminAddMemberError}
                </div>
              )}
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={closeDlg}
                  disabled={adminAddMemberBusy}
                >
                  Abbrechen
                </button>
                {(() => {
                  // v17.1: Bei „Neues Team anlegen" + askTeamName=true ist
                  // der Team-Name Pflicht (analog Self-Registration-Flow).
                  const needName = !!adminAddMemberDialog.isNewTeam && !!selectedEvent.askTeamName;
                  const nameOk = !needName || (adminAddMemberDialog.teamName.trim().length > 0);
                  // v22.40-Bugfix: Vorher verlangte `disabled` zwingend einen
                  // Graph-Pick + Consent — dadurch war der Button bei reiner
                  // Zuordnung bereits-angemeldeter Personen NIE klickbar. Jetzt:
                  // mindestens eine Auswahl (teamlos ODER Graph), Consent nur bei
                  // echtem Graph-Neu-Pick, nicht über Kapazität, Name ok.
                  const consentOk = !consentRequired || adminAddMemberConsent;
                  const disabled = totalPicks === 0 || !consentOk || adminAddMemberBusy || !nameOk || overCap;
                  const title = !nameOk ? 'Bitte einen Team-Namen eingeben.'
                    : overCap ? `Zu viele ausgewählt — nur noch ${freeSlots} Platz/Plätze frei.`
                    : totalPicks === 0 ? 'Bitte mindestens eine Person auswählen.'
                    : (!consentOk ? 'Bitte die Zustimmung bestätigen.' : '');
                  return (
                    <button
                      type="button"
                      className="btn btn-primary"
                      onClick={() => { submit().catch(() => { /* */ }); }}
                      disabled={disabled}
                      title={title}
                    >
                      {adminAddMemberBusy
                        ? 'Wird gespeichert…'
                        : (adminAddMemberDialog.isNewTeam
                          ? `Team anlegen${totalPicks > 0 ? ` (${totalPicks})` : ''}`
                          : `Hinzufügen${totalPicks > 0 ? ` (${totalPicks})` : ''}`)}
                    </button>
                  );
                })()}
              </div>
          </Modal>
        );
      })()}
      {/* v17.8: Floating Jump-Buttons. Erscheinen sobald der User durch die
          Teilnehmer-Tabelle scrollt — sparen Zeit bei langen Listen. */}
      {/* v17.11.1: JumpButtons immer rendern wenn ein Event selektiert ist —
          das interne Show-Gating (scrollY > 300) reicht. Früher Threshold
          activeRegs>10 war zu hoch, für Test-Events mit wenig TN
          erschienen die Buttons nie. */}
      {selectedEvent && (
        <JumpButtons hasWaitlist={waitlistRegs.length > 0} />
      )}
    </div>
  );
}

/**
 * v17.8: Floating Jump-Buttons rechts unten. Erscheinen sobald der User
 * den Viewport >300 px nach unten gescrollt hat. Bietet:
 *  - Nach oben springen (window.scrollTo {top:0})
 *  - Zur Warteliste springen (scrollIntoView auf #admin-waitlist-anchor)
 *
 * Nur sichtbar wenn die Teilnehmer-Tabelle >10 Einträge hat (kurze Listen
 * brauchen keine Sprung-Hilfe).
 */
/**
 * v17.13: Floating Jump-Buttons. Im SPFx-Webpart-Kontext scrollt nicht
 * window, sondern ein SP-interner Container — deshalb sind die Buttons
 * jetzt IMMER sichtbar (kein scrollY-Gating), und der Click sucht den
 * tatsächlich scrollenden Vorfahren des Targets statt window.scrollTo.
 */
function JumpButtons(props: { hasWaitlist: boolean }): React.ReactElement {
  const { hasWaitlist } = props;
  /** Sucht den ersten scroll-baren Vorfahren — typischerweise der
   *  SP-Page-Body. Fallback auf document.scrollingElement / window. */
  const findScrollParent = (el: HTMLElement | null): HTMLElement | Window => {
    let cur: HTMLElement | null = el ? el.parentElement : null;
    while (cur && cur !== document.body && cur !== document.documentElement) {
      const cs = window.getComputedStyle(cur);
      const overflowY = cs.overflowY;
      const isScrollable = (overflowY === 'auto' || overflowY === 'scroll' || overflowY === 'overlay');
      if (isScrollable && cur.scrollHeight > cur.clientHeight) return cur;
      cur = cur.parentElement;
    }
    return (document.scrollingElement as HTMLElement) || document.documentElement || window;
  };
  const scrollToTop = (): void => {
    // Versuche window, documentElement, body und alle scrollbaren Vorfahren.
    try { window.scrollTo({ top: 0, behavior: 'smooth' }); } catch { /* */ }
    try { if (document.scrollingElement) (document.scrollingElement as HTMLElement).scrollTo({ top: 0, behavior: 'smooth' }); } catch { /* */ }
    try { if (document.documentElement) document.documentElement.scrollTop = 0; } catch { /* */ }
    try { if (document.body) document.body.scrollTop = 0; } catch { /* */ }
    // SP-Page-Container — typischer Klassenname in modernen SP-Pages.
    const candidates = ['.SPPageChromeAppDiv', '#spPageCanvasContent', '[data-automation-id="contentScrollRegion"]', '.spAppAriaRegion', 'main'];
    for (const sel of candidates) {
      const el = document.querySelector(sel) as HTMLElement | null;
      if (el && el.scrollHeight > el.clientHeight) {
        try { el.scrollTo({ top: 0, behavior: 'smooth' }); } catch { el.scrollTop = 0; }
      }
    }
  };
  const scrollToWaitlist = (): void => {
    const el = document.getElementById('admin-waitlist-anchor');
    if (!el) return;
    try { el.scrollIntoView({ behavior: 'smooth', block: 'start' }); } catch { el.scrollIntoView(); }
    // Zusätzlich: scroll-Parent suchen und manuell scrollen (Fallback
    // wenn scrollIntoView durch den SP-Container gefangen wird).
    const parent = findScrollParent(el);
    if (parent !== window && parent instanceof HTMLElement) {
      const rect = el.getBoundingClientRect();
      const parentRect = parent.getBoundingClientRect();
      try { parent.scrollTo({ top: parent.scrollTop + (rect.top - parentRect.top) - 20, behavior: 'smooth' }); } catch { /* */ }
    }
  };
  return (
    <div style={{
      position: 'fixed',
      // v18: Buttons horizontal ZENTRIERT ueber dem Content (vorher links am
      // Rand „im Raum hängend"). Als Zeile nebeneinander, mittig unten —
      // liegt damit in der horizontalen Mitte der Teilnehmer-Tabelle /
      // Spaltenüberschriften statt am Seitenrand.
      left: '50%', transform: 'translateX(-50%)',
      bottom: 20, zIndex: 900,
      display: 'flex', flexDirection: 'row', gap: 8,
    }}>
      {hasWaitlist && (
        <button
          type="button"
          onClick={scrollToWaitlist}
          style={{
            background: 'var(--dex-orange, #ed8b00)', color: '#fff',
            border: 'none', padding: '10px 16px', borderRadius: 999,
            cursor: 'pointer', fontSize: '0.85rem', fontWeight: 600,
            boxShadow: '0 4px 12px rgba(0,0,0,0.18)',
            display: 'inline-flex', alignItems: 'center', gap: 6,
          }}
        >
          ↓ Zur Warteliste
        </button>
      )}
      <button
        type="button"
        onClick={scrollToTop}
        style={{
          background: 'var(--dex-green, #86bc25)', color: '#fff',
          border: 'none', padding: '10px 16px', borderRadius: 999,
          cursor: 'pointer', fontSize: '0.85rem', fontWeight: 600,
          boxShadow: '0 4px 12px rgba(0,0,0,0.18)',
          display: 'inline-flex', alignItems: 'center', gap: 6,
        }}
      >
        ↑ Nach oben
      </button>
    </div>
  );
}
