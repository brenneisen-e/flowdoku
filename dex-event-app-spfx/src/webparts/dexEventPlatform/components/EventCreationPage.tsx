/**
 * Event-Erstellung (nur fuer Organizer / SuperAdmin)
 *
 * Erstellt ein Event in der DEX_Events-Liste und eine
 * separate Teilnehmerliste mit Item-Level Permissions.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { EventService, CustomField } from '../services/EventService';
import { eventCreatedEmail, buildOutlookBody, stripOutlookWrapper, parseOutlookHeadings, replacePlaceholders, getCachedOrbBase64 } from '../services/EmailTemplates';
import { exportSummaryAsPdf, exportSummaryAsDoc, SummaryData } from '../services/EventSummaryExport';
import { EventType, AgendaItem } from '../types';
import { Trash2, Send, Plus, X, Users, Check } from './Icons';
import { RichText } from '@pnp/spfx-controls-react/lib/controls/richText';
import { HtmlEditorModal } from './HtmlEditorModal';
import { RegisterPreviewModal } from './RegisterPreviewModal';
import { InfoTooltip } from './InfoTooltip';
import BulkUserImportModal from './BulkUserImportModal';
import Modal from './Modal';
import InternationalSearchToggle from './InternationalSearchToggle';
import { generateSelfCheckInToken } from '../utils/selfCheckIn';
import { downloadSelfCheckInPdf } from '../utils/selfCheckInPdf';
import { buildOutlookLocation } from '../utils/eventFormat';
import { Icon } from '@fluentui/react/lib/Icon';
import DatePicker, { registerLocale } from 'react-datepicker';
import { de } from 'date-fns/locale';
import 'react-datepicker/dist/react-datepicker.css';

// Deutsche Locale registrieren
registerLocale('de', de);

// Curated Fluent UI Icons fuer Agenda-Programmpunkte (nur bestaetigt vorhandene MDL2 Icons)
const AGENDA_ICONS: Array<{ name: string; label: string; category: string }> = [
  // Vorträge & Meetings
  { name: 'Microphone', label: 'Vortrag', category: 'meeting' },
  { name: 'People', label: 'Meeting', category: 'meeting' },
  { name: 'Group', label: 'Team', category: 'meeting' },
  { name: 'Presentation', label: 'Präsentation', category: 'meeting' },
  { name: 'Chat', label: 'Diskussion', category: 'meeting' },
  // Pausen & Essen
  { name: 'Cafe', label: 'Kaffeepause', category: 'break' },
  { name: 'EatDrink', label: 'Essen', category: 'break' },
  { name: 'Brunch', label: 'Brunch', category: 'break' },
  // Aktivitäten
  { name: 'Running', label: 'Sport/Lauf', category: 'activity' },
  { name: 'Trophy', label: 'Award', category: 'activity' },
  { name: 'Balloons', label: 'Feier', category: 'activity' },
  { name: 'MusicInCollection', label: 'Musik', category: 'activity' },
  { name: 'PartyLeader', label: 'Networking', category: 'activity' },
  // Organisation
  { name: 'Calendar', label: 'Termin', category: 'org' },
  { name: 'Clock', label: 'Uhrzeit', category: 'org' },
  { name: 'CheckMark', label: 'Check-in', category: 'org' },
  { name: 'MapPin', label: 'Ort', category: 'org' },
  { name: 'Car', label: 'Anfahrt', category: 'org' },
  { name: 'Bus', label: 'Bus/Transfer', category: 'org' },
  { name: 'Airplane', label: 'Flug', category: 'org' },
  { name: 'Hotel', label: 'Hotel', category: 'org' },
  // Workshops & Arbeit
  { name: 'Edit', label: 'Workshop', category: 'work' },
  { name: 'Lightbulb', label: 'Ideen', category: 'work' },
  { name: 'Code', label: 'Tech', category: 'work' },
  { name: 'ReadingMode', label: 'Schulung', category: 'work' },
  { name: 'Page', label: 'Dokument', category: 'work' },
  // Allgemein
  { name: 'Info', label: 'Info', category: 'general' },
  { name: 'FavoriteStar', label: 'Highlight', category: 'general' },
  { name: 'Heart', label: 'Social', category: 'general' },
  { name: 'Camera', label: 'Foto', category: 'general' },
  { name: 'Flag', label: 'Flagge', category: 'general' },
];

// Erweiterte Icon-Liste fuer "Show All"
const EXTENDED_ICONS: Array<{ name: string; label: string; category: string }> = [
  { name: 'Home', label: 'Home', category: 'general' },
  { name: 'Mail', label: 'Mail', category: 'general' },
  { name: 'Phone', label: 'Telefon', category: 'general' },
  { name: 'Send', label: 'Senden', category: 'general' },
  { name: 'Attach', label: 'Anhang', category: 'general' },
  { name: 'Link', label: 'Link', category: 'general' },
  { name: 'Globe', label: 'Web', category: 'general' },
  { name: 'Lock', label: 'Sicherheit', category: 'org' },
  { name: 'Sunny', label: 'Wetter', category: 'activity' },
  { name: 'Ringer', label: 'Glocke', category: 'general' },
  { name: 'Contact', label: 'Kontakt', category: 'meeting' },
  { name: 'AddFriend', label: 'Person hinzufügen', category: 'meeting' },
  { name: 'TeamFavorite', label: 'Team-Favorit', category: 'meeting' },
  { name: 'Handshake', label: 'Handshake', category: 'meeting' },
  { name: 'Medical', label: 'Medizin', category: 'org' },
  { name: 'Shield', label: 'Schutz', category: 'org' },
  { name: 'Settings', label: 'Einstellungen', category: 'org' },
  { name: 'Toolbox', label: 'Werkzeug', category: 'work' },
  { name: 'Chart', label: 'Diagramm', category: 'work' },
  { name: 'BarChart4', label: 'Statistik', category: 'work' },
  { name: 'TaskList', label: 'Aufgaben', category: 'work' },
  { name: 'ClipboardList', label: 'Checkliste', category: 'work' },
  { name: 'Video', label: 'Video', category: 'activity' },
  { name: 'Photo2', label: 'Foto', category: 'activity' },
  { name: 'Game', label: 'Spiel', category: 'activity' },
  { name: 'Rocket', label: 'Rakete', category: 'activity' },
  { name: 'Emoji2', label: 'Spaß', category: 'activity' },
  { name: 'Gift', label: 'Geschenk', category: 'activity' },
  { name: 'Ferry', label: 'Schiff', category: 'org' },
  { name: 'Train', label: 'Zug', category: 'org' },
  { name: 'Walk', label: 'Fußweg', category: 'org' },
  { name: 'Weights', label: 'Fitness', category: 'activity' },
];

/**
 * Komprimiert ein Bild clientseitig via Canvas.
 * Max 1200px Breite, JPEG 80% Qualität.
 */
async function compressImage(file: File, maxWidth: number = 1200, quality: number = 0.8): Promise<File> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      let { width, height } = img;
      if (width > maxWidth) {
        height = Math.round((height * maxWidth) / width);
        width = maxWidth;
      }
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      if (!ctx) { resolve(file); return; }
      ctx.drawImage(img, 0, 0, width, height);
      canvas.toBlob(
        (blob) => {
          if (!blob || blob.size >= file.size) {
            // Komprimierung bringt nichts oder ist grösser → Original verwenden
            resolve(file);
            return;
          }
          const compressed = new File([blob], file.name.replace(/\.[^.]+$/, '.jpg'), { type: 'image/jpeg' });
          resolve(compressed);
        },
        'image/jpeg',
        quality
      );
    };
    img.onerror = () => reject(new Error('Bild konnte nicht geladen werden'));
    img.src = URL.createObjectURL(file);
  });
}

interface CustomFieldInput {
  id: string;
  label: string;
  type: 'text' | 'select' | 'number' | 'checkbox' | 'user' | 'roommate' | 'document'; // v19.0: document = Datei-Upload
  required: boolean;
  // Optionen als Array (incl. leerer Slots fuer "frisch hinzugefuegte" Eintraege)
  options: string[];
  visible: boolean;
  externalLinks?: Array<{ label: string; url: string }>;
  /** v18.41: Nur People-Picker (user/roommate): ausgewählte Person bei
   *  An-/Abmelde-Mail auf CC setzen (nicht im Outlook-Termin). */
  ccOnEmails?: boolean;
  /** v7.11: Bei type=select erlaubt true Mehrfachauswahl (Checkbox-Liste statt
   *  Single-Dropdown). Wert wird " | "-getrennt gespeichert. */
  multi?: boolean;
  /** v7.20: Optionale Beschreibung — landet als "i"-Tooltip neben dem
   *  Feld-Label im Registrierungsformular. */
  helpText?: string;
  /** v18.18: 'tooltip' (Default) = "i"-Hover-Box neben dem Label;
   *  'inline' = nicht-fetter Erklär-Text direkt unter dem Label. */
  helpTextStyle?: 'tooltip' | 'inline';
  /** v7.21: Sichtbarkeitsbedingung — Feld nur anzeigen wenn das Quell-Feld
   *  einen der `values` als Antwort hat. */
  showIf?: { fieldId: string; values: string[] };
  /** v10.24: Bei aktiver Split-Capacity Feld nur fuer eine der zwei
   *  Gruppen sichtbar machen ('A' = Durchstarter / Gruppe A, 'B' =
   *  Funstarter / Gruppe B). 'all' / undefined = beide Gruppen. */
  onlyForGroup?: 'all' | 'A' | 'B';
  /** v11.94: Nur fuer type='checkbox' — Text neben der Checkbox im
   *  Registrierungsformular (Default „Ja, bestätigen" / „Yes, confirm"). */
  confirmLabel?: string;
  /** v17.20: Englische Varianten — nur relevant wenn der Organizer im
   *  selben Schritt 5 den Toggle „Deutsch und Englisch ermöglichen"
   *  gesetzt hat. */
  labelEn?: string;
  helpTextEn?: string;
  confirmLabelEn?: string;
  optionsEn?: string[];
}

// v18.22: Pro-Event-Override eines Mail-Templates (Subject/Headings/Body +
// Formatierung). Vorher als gleiche Inline-Form an ~7 Stellen wiederholt —
// jetzt ein zentraler Alias, damit neue Felder nur hier ergänzt werden.
type EmailOverrideEntry = {
  subject: string;
  heading: string;
  subheading?: string;
  bodyHtml: string;
  headingColor?: string;
  headingFontSize?: string;
  /** v18.22: Fett/Kursiv für die Überschrift (h1). */
  headingBold?: boolean;
  headingItalic?: boolean;
  /** v18.22: Unter-Überschrift (h2) frei formatierbar. */
  subheadingColor?: string;
  subheadingFontSize?: string;
  subheadingBold?: boolean;
  subheadingItalic?: boolean;
};

// v17.22: Einziger Serializer fuer Custom-Fields → CustomFields-JSON.
// Vorher dreimal copy-paste (Create-Save, Edit-Save, Sub-Event-Save), was
// dazu fuehrte, dass der Sub-Event-Pfad die v17.20-EN-Varianten nicht
// mitnahm. Zentral hier, damit alle drei Pfade identisch persistieren.
//
// Wichtig (v17.22-Fix): DE-Optionen UND EN-Optionen werden POSITIONAL
// gepaart gefiltert — vorher wurde `options` per `.filter(Boolean)` von
// Leereintraegen befreit, `optionsEn` aber nicht, wodurch das Index-Mapping
// zwischen DE und EN bei leeren Slots verrutschte (leere/falsche EN-Labels
// auf der Anmeldeseite).
function serializeCustomFields(
  fields: CustomFieldInput[],
  bilingual: boolean
): CustomField[] {
  return fields
    .filter(f => f.label && f.label.trim().length > 0)
    .map(f => {
      let optionsOut: string[] | undefined;
      let optionsEnOut: string[] | undefined;
      if (f.type === 'select') {
        const pairs = (f.options || [])
          .map((o, i) => ({ de: (o || '').trim(), en: ((f.optionsEn || [])[i] || '').trim() }))
          .filter(p => p.de.length > 0);
        optionsOut = pairs.map(p => p.de);
        if (bilingual && pairs.some(p => p.en.length > 0)) {
          optionsEnOut = pairs.map(p => p.en);
        }
      }
      return {
        id: f.id,
        label: f.label.trim(),
        type: f.type,
        required: !!f.required,
        visible: f.visible !== false,
        ...(f.helpText && f.helpText.trim() ? { helpText: f.helpText.trim() } : {}),
        // v18.18: nur persistieren wenn 'inline' (Default 'tooltip' = weglassen).
        ...(f.helpTextStyle === 'inline' ? { helpTextStyle: 'inline' as const } : {}),
        ...(f.showIf && f.showIf.fieldId && f.showIf.values && f.showIf.values.length > 0
          ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } }
          : {}),
        ...(optionsOut ? { options: optionsOut, ...(f.multi ? { multi: true } : {}) } : {}),
        ...(f.onlyForGroup && f.onlyForGroup !== 'all' ? { onlyForGroup: f.onlyForGroup } : {}),
        ...(f.type === 'checkbox' && f.confirmLabel && f.confirmLabel.trim()
          ? { confirmLabel: f.confirmLabel.trim() }
          : {}),
        // v17.20: Englische Varianten — nur wenn der Bilingual-Toggle an ist
        // UND der Organizer Text eingegeben hat.
        ...(bilingual && f.labelEn && f.labelEn.trim() ? { labelEn: f.labelEn.trim() } : {}),
        ...(bilingual && f.helpTextEn && f.helpTextEn.trim() ? { helpTextEn: f.helpTextEn.trim() } : {}),
        ...(bilingual && f.type === 'checkbox' && f.confirmLabelEn && f.confirmLabelEn.trim()
          ? { confirmLabelEn: f.confirmLabelEn.trim() }
          : {}),
        ...(optionsEnOut ? { optionsEn: optionsEnOut } : {}),
        ...(f.externalLinks && f.externalLinks.length > 0
          ? { externalLinks: f.externalLinks.map(x => ({ label: x.label, url: x.url })) }
          : {}),
        // v18.41: CC-bei-Mail nur für People-Picker-Felder persistieren.
        ...((f.type === 'user' || f.type === 'roommate') && f.ccOnEmails ? { ccOnEmails: true } : {}),
      } as CustomField;
    });
}

// v19.22: Sub-Event-Titel auf den reinen Sub-Namen kuerzen (Parent-Praefix
// entfernen) — gleiche Logik wie im Admin Center. Sub-Event-Titel werden oft
// als „<Hauptevent> | <Sub-Name>" gespeichert; in den Tabs/Listen reicht der
// Sub-Name (z.B. „HER SPACE").
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

function StepBadge({ n }: { n: number }): React.ReactElement {
  return (
    <span style={{
      display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
      width: 22, height: 22, borderRadius: '50%',
      background: 'var(--dex-green)', color: '#fff',
      fontSize: '0.72rem', fontWeight: 700, flexShrink: 0,
    }}>{n}</span>
  );
}

// v8.0: Multi-Select-Dropdown fuer den Standortfilter (loest die alten
// Pillen-Buttons ab — kompakter und mit Suche bei vielen Optionen).
function LocationMultiSelect({
  options, selected, onChange, isDe,
}: {
  options: string[];
  selected: string[];
  onChange: (next: string[]) => void;
  isDe: boolean;
}): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  const [query, setQuery] = React.useState('');
  const ref = React.useRef<HTMLDivElement>(null);

  // Click-Outside zum Schliessen
  React.useEffect(() => {
    if (!open) return undefined;
    const handler = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false);
        setQuery('');
      }
    };
    window.addEventListener('mousedown', handler);
    return () => window.removeEventListener('mousedown', handler);
  }, [open]);

  const toggle = (loc: string): void => {
    if (selected.indexOf(loc) >= 0) onChange(selected.filter(l => l !== loc));
    else onChange([...selected, loc]);
  };

  const filtered = query.trim()
    ? options.filter(o => o.toLowerCase().indexOf(query.trim().toLowerCase()) >= 0)
    : options;

  return (
    <div ref={ref} style={{ position: 'relative', maxWidth: 520 }}>
      <button
        type="button"
        onClick={() => setOpen(v => !v)}
        style={{
          width: '100%', minHeight: 42, padding: '6px 12px',
          background: '#fff',
          border: `1.5px solid ${open ? 'var(--dex-green)' : 'var(--dex-gray-300)'}`,
          borderRadius: 8, cursor: 'pointer', textAlign: 'left',
          display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 6,
          fontSize: '0.88rem', color: 'var(--dex-gray-800)',
          transition: 'border-color 0.15s ease',
        }}
      >
        {selected.length === 0 ? (
          <span style={{ color: 'var(--dex-gray-400)' }}>
            {isDe ? 'Standorte auswählen…' : 'Select locations…'}
          </span>
        ) : (
          selected.map(loc => (
            <span
              key={loc}
              style={{
                display: 'inline-flex', alignItems: 'center', gap: 4,
                padding: '3px 4px 3px 10px',
                background: 'var(--dex-green)', color: '#fff',
                borderRadius: 999, fontSize: '0.78rem',
              }}
            >
              {loc}
              <span
                role="button"
                aria-label={`${loc} entfernen`}
                onClick={e => { e.stopPropagation(); toggle(loc); }}
                style={{
                  width: 18, height: 18, borderRadius: '50%',
                  background: 'rgba(255,255,255,0.25)',
                  display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                  fontSize: '0.85rem', lineHeight: 1, cursor: 'pointer',
                }}
              >×</span>
            </span>
          ))
        )}
        <span style={{ marginLeft: 'auto', color: 'var(--dex-gray-500)', fontSize: '0.7rem' }}>
          {open ? '▲' : '▼'}
        </span>
      </button>
      {open && (
        <div style={{
          position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0,
          background: '#fff', border: '1px solid var(--dex-gray-200)',
          borderRadius: 8, boxShadow: '0 8px 24px rgba(0,0,0,0.12)',
          maxHeight: 320, overflowY: 'auto', zIndex: 50,
        }}>
          {options.length > 6 && (
            <div style={{ padding: 8, borderBottom: '1px solid var(--dex-gray-100)', position: 'sticky', top: 0, background: '#fff' }}>
              <input
                autoFocus
                value={query}
                onChange={e => setQuery(e.target.value)}
                placeholder={isDe ? 'Suchen…' : 'Search…'}
                style={{
                  width: '100%', padding: '6px 10px',
                  border: '1px solid var(--dex-gray-200)',
                  borderRadius: 6, fontSize: '0.85rem',
                }}
              />
            </div>
          )}
          {filtered.length === 0 ? (
            <div style={{ padding: 12, fontSize: '0.82rem', color: 'var(--dex-gray-400)' }}>
              {isDe ? 'Keine Treffer.' : 'No matches.'}
            </div>
          ) : (
            filtered.map(loc => {
              const isChecked = selected.indexOf(loc) >= 0;
              return (
                <label
                  key={loc}
                  style={{
                    display: 'flex', alignItems: 'center', gap: 10,
                    padding: '8px 12px', cursor: 'pointer', fontSize: '0.88rem',
                    background: isChecked ? 'rgba(134,188,37,0.08)' : 'transparent',
                  }}
                  onMouseEnter={e => { if (!isChecked) (e.currentTarget as HTMLLabelElement).style.background = 'var(--dex-gray-50)'; }}
                  onMouseLeave={e => { if (!isChecked) (e.currentTarget as HTMLLabelElement).style.background = 'transparent'; }}
                >
                  <input
                    type="checkbox"
                    checked={isChecked}
                    onChange={() => toggle(loc)}
                    style={{ width: 16, height: 16, accentColor: 'var(--dex-green)', cursor: 'pointer' }}
                  />
                  <span style={{ color: 'var(--dex-gray-800)' }}>{loc}</span>
                </label>
              );
            })
          )}
          {selected.length > 0 && (
            <div style={{
              padding: 8, borderTop: '1px solid var(--dex-gray-100)',
              position: 'sticky', bottom: 0, background: '#fff',
              display: 'flex', justifyContent: 'flex-end',
            }}>
              <button
                type="button"
                onClick={() => onChange([])}
                style={{
                  background: 'none', border: 'none', cursor: 'pointer',
                  fontSize: '0.78rem', color: 'var(--dex-gray-500)',
                  textDecoration: 'underline', padding: '4px 8px',
                }}
              >
                {isDe ? 'Auswahl leeren' : 'Clear selection'}
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/**
 * v16.4: Audience-Liste (kommasepariert) in eine flache, ';'-separierte
 * Liste von Member-E-Mails aufloesen. Jede '@'-Eintrag wird via
 * getGroupMembers (Graph) probiert — wenn die Aufloesung eine
 * Mitglieder-Liste liefert, werden alle deren E-Mails uebernommen, sonst
 * wird der Eintrag als direkte User-E-Mail behandelt. Lowercase + dedupliziert.
 *
 * Wird beim Event-Save aufgerufen und in die SP-Spalte
 * `AudienceResolvedEmails` geschrieben. matchesAudience im
 * EventListPage prueft Sichtbarkeit zur Laufzeit gegen diese Liste.
 */
async function resolveAudienceMembersToCsv(
  audienceCsv: string,
  getGroupMembers: (groupEmail: string) => Promise<{ groupName: string; members: Array<{ email: string }> } | null>,
): Promise<string> {
  const items = (audienceCsv || '').split(',').map(s => s.trim()).filter(Boolean);
  if (items.length === 0) return '';
  const out = new Set<string>();
  for (const item of items) {
    if (item.indexOf('@') < 0) continue; // Gruppen-Patterns (DEKOELN etc.) bleiben Runtime-Match
    try {
      const grp = await getGroupMembers(item);
      if (grp && grp.members && grp.members.length > 0) {
        for (const m of grp.members) {
          const e = (m.email || '').toLowerCase().trim();
          if (e) out.add(e);
        }
      } else {
        // Keine Member zurueckgeliefert → behandle als direkte User-Adresse.
        out.add(item.toLowerCase());
      }
    } catch {
      out.add(item.toLowerCase());
    }
  }
  return Array.from(out).join(';');
}

export default function EventCreationPage(): React.ReactElement {
  const { goBack, selectedEventId, currentPage, setNavigationGuard } = useNavigation();
  const { events, childEventsOf, createEvent, updateEvent, deleteEvent, deleteEventItemOnly, refreshEvents } = useEvents();
  const { currentUser } = useCurrentUser();
  const { searchUsers, searchGroups, getGroupMembers, searchUsersByLocation, canCreateEvents, roles, siteUrl } = useRoles();
  // v18.4: Power User = Experten-Organizer; auf Wizard-Seite 1 als
  // Hilfe-Ansprechpartner im einklappbaren Layover unten rechts angezeigt.
  const powerUsers = React.useMemo(
    () => (roles || []).filter(r => r.isPowerUser),
    [roles]
  );
  // v18.6: Power-User-Hilfe als zugeklappter „?"-Ball unten rechts —
  // Klick klappt das Panel auf, X klappt es wieder zum Ball zu.
  const [powerUserHelpOpen, setPowerUserHelpOpen] = React.useState(false);
  // v13.0: Frühe Permission-Prüfung — vorher konnte ein Demo-User die
  // Seite öffnen und das Save würde erst beim SP-Write scheitern. Mit
  // Guard zurück zur Start-Seite, falls keine Organizer-Rechte.
  React.useEffect(() => {
    if (!canCreateEvents) goBack();
  }, [canCreateEvents, goBack]);
  // Audience-Suche (Personen + Verteiler/Security-Groups)
  const [audienceSearch, setAudienceSearch] = React.useState('');
  const [audienceResults, setAudienceResults] = React.useState<Array<{ kind: 'user' | 'group'; email: string; displayName: string }>>([]);
  const [isSearchingAudience, setIsSearchingAudience] = React.useState(false);
  const [audienceIncludeIntl, setAudienceIncludeIntl] = React.useState(false);
  const audienceTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  // Modal: Members einer Gruppe anzeigen
  const [memberModalOpen, setMemberModalOpen] = React.useState(false);
  const [memberModalGroupEmail, setMemberModalGroupEmail] = React.useState('');
  const [memberModalGroupName, setMemberModalGroupName] = React.useState('');
  const [memberModalLoading, setMemberModalLoading] = React.useState(false);
  const [memberModalMembers, setMemberModalMembers] = React.useState<Array<{ email: string; displayName: string }>>([]);
  const [memberModalError, setMemberModalError] = React.useState('');
  // Modal: Massenimport für Sichtbarkeits-Filter (Audience). Logik + UI sind in
  // BulkUserImportModal gekapselt — gleiche Komponente wie für die drei Team-Felder.
  const [bulkAudienceOpen, setBulkAudienceOpen] = React.useState(false);
  // Audience-Chip-Pagination + Inline-Suche
  const [audienceShowAll, setAudienceShowAll] = React.useState(false);
  const [audienceChipSearch, setAudienceChipSearch] = React.useState('');

  // Nutzungsbedingungen: Beim Erstellen eines neuen Events muss der Organizer
  // zuerst eine Bestaetigungs-Maske mit den Nutzungs- und Datenschutz-
  // bedingungen akzeptieren. Nicht relevant beim Bearbeiten bestehender Events.
  const [tcAccepted, setTcAccepted] = React.useState(false);
  const [tcCheckbox, setTcCheckbox] = React.useState(false);
  const [tcExpanded, setTcExpanded] = React.useState(false);

  const addAudienceItem = (value: string): void => {
    const list = audience.split(',').map(s => s.trim()).filter(Boolean);
    if (list.indexOf(value) >= 0) return;
    list.push(value);
    setAudience(list.join(', '));
  };
  const removeAudienceItem = (value: string): void => {
    const list = audience.split(',').map(s => s.trim()).filter(Boolean).filter(x => x !== value);
    setAudience(list.join(', '));
  };


  const openMembersModal = async (groupEmail: string): Promise<void> => {
    setMemberModalOpen(true);
    setMemberModalGroupEmail(groupEmail);
    setMemberModalGroupName(groupEmail);
    setMemberModalMembers([]);
    setMemberModalError('');
    setMemberModalLoading(true);
    const res = await getGroupMembers(groupEmail);
    setMemberModalLoading(false);
    if (!res) {
      setMemberModalError('Mitglieder konnten nicht geladen werden. Vermutlich fehlt die Berechtigung "Group.Read.All" im SharePoint App Catalog (Admin Consent erforderlich).');
      return;
    }
    setMemberModalGroupName(res.groupName || groupEmail);
    setMemberModalMembers(res.members);
  };
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';

  // Edit-Modus: wenn wir auf 'edit-event' sind und eine selectedEventId haben
  const isEditMode = currentPage === 'edit-event' && !!selectedEventId;
  const editEvent = isEditMode ? events.find(e => e.id === selectedEventId) : null;

  // ========== Zeitzonen-Handling (Europe/Berlin, browser-TZ-unabhaengig) ==========
  //
  // Hintergrund: Der datetime-local-Input liefert einen naiven String ohne TZ-Suffix
  // (z.B. "2026-04-23T19:00"). Wenn wir diesen mit `new Date(str).toISOString()`
  // konvertieren, interpretiert JavaScript den String in der BROWSER-Zeitzone. Bei
  // einem Browser auf UTC oder in einer VM/Citrix mit falscher TZ fuehrt das zu einem
  // 2h-Shift: 19:00 wird als UTC interpretiert statt als MESZ, SP speichert 19:00Z
  // statt 17:00Z, und Outlook zeigt dann 21:00 MESZ.
  //
  // Fix: Die App interpretiert ALLE Event-Zeiten explizit als Europe/Berlin, egal
  // welche Zeitzone der Browser hat. Wir nutzen Intl.DateTimeFormat um den Offset
  // fuer einen konkreten Zeitpunkt zu bestimmen (DST-aware).

  /** Gibt den Offset von Europe/Berlin zu UTC an dem gegebenen Zeitpunkt in ms zurueck.
   *  Im Winter: +3600000 (+1h). Im Sommer: +7200000 (+2h). */
  const berlinOffsetMs = (dateUtc: Date): number => {
    const dtf = new Intl.DateTimeFormat('en-US', {
      timeZone: 'Europe/Berlin',
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
    });
    const parts = dtf.formatToParts(dateUtc);
    const get = (type: string): number => parseInt(parts.find(p => p.type === type)?.value || '0', 10);
    let h = get('hour');
    if (h === 24) h = 0; // en-US hour12:false liefert manchmal 24 statt 0
    const asIfUtc = Date.UTC(get('year'), get('month') - 1, get('day'), h, get('minute'), get('second'));
    return asIfUtc - dateUtc.getTime();
  };

  /** datetime-local-String ("2026-04-23T19:00") als Europe/Berlin interpretieren
   *  und nach UTC-ISO konvertieren ("2026-04-23T17:00:00.000Z"). */
  const berlinLocalToUtcIso = (localStr: string): string => {
    if (!localStr) return '';
    // Parse den String erstmal als ob er UTC waere -> das sind UTC-Zahlen die den Berlin-Werten entsprechen
    const asUtc = new Date(localStr.length === 16 ? localStr + ':00Z' : localStr + 'Z');
    if (isNaN(asUtc.getTime())) return '';
    // Der echte UTC-Zeitpunkt ist asUtc minus Berlin-Offset an diesem Zeitpunkt
    const offset = berlinOffsetMs(asUtc);
    return new Date(asUtc.getTime() - offset).toISOString();
  };

  /** UTC-ISO ("2026-04-23T17:00:00.000Z") nach datetime-local in Europe/Berlin
   *  ("2026-04-23T19:00") konvertieren — fuer das Input-Feld. */
  const isoToLocal = (iso: string): string => {
    if (!iso) return '';
    const d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    const dtf = new Intl.DateTimeFormat('en-US', {
      timeZone: 'Europe/Berlin',
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', hour12: false,
    });
    const parts = dtf.formatToParts(d);
    const get = (type: string): string => parts.find(p => p.type === type)?.value || '00';
    let hour = get('hour');
    if (hour === '24') hour = '00';
    return `${get('year')}-${get('month')}-${get('day')}T${hour}:${get('minute')}`;
  };

  // Deadline-Datum als Ende-des-Tages (23:59 Europe/Berlin) speichern, damit:
  //  a) Die Uhrzeit-Anzeige in der EventCard nicht mehr "02:00" zeigt
  //  b) Die Deadline-Pruefung "new Date(deadline) < new Date()" wirklich den
  //     gesamten ausgewaehlten Tag als gueltig behandelt (statt nur bis UTC-Mitternacht).
  const deadlineToEndOfDayIso = (dateStr: string): string | null => {
    if (!dateStr) return null;
    // dateStr im Format "YYYY-MM-DD" (date-Input) - wir behandeln als 23:59 Europe/Berlin
    const localStr = dateStr.length === 10 ? `${dateStr}T23:59` : dateStr;
    const utcIso = berlinLocalToUtcIso(localStr);
    return utcIso || null;
  };

  const [title, setTitle] = React.useState(editEvent ? editEvent.title : '');
  // Mehrere Organizer werden mit '; ' getrennt gespeichert (innerhalb eines Namens
  // kann ',' vorkommen, z.B. 'Maerzluft, Petra').
  //
  // Auto-Heal bei Längen-Mismatch (Legacy-Korruption aus v10.0–v10.2-Closure-Bug):
  // pad auf max(names.length, emails.length) statt truncate auf min. Dadurch
  // verliert der User keine Organizer-Einträge beim Edit-Load — fehlende Emails
  // bleiben als leere Strings erhalten und können vom User einzeln per Picker
  // nachgepflegt werden. Die Warning-Box (siehe weiter unten in der UI) macht das
  // Mismatch sichtbar.
  const [organizer, setOrganizer] = React.useState(() => {
    if (!editEvent) return `${currentUser.firstName} ${currentUser.surname}`;
    const names = editEvent.organizers || [];
    const emails = editEvent.organizerEmails || [];
    const max = Math.max(names.length, emails.length);
    const padded: string[] = [];
    for (let i = 0; i < max; i++) padded.push(names[i] || '');
    return padded.join('; ');
  });
  const [organizerResults, setOrganizerResults] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const [organizerSearch, setOrganizerSearch] = React.useState('');
  // Beim Edit: organizerEmails aus dem gespeicherten Event uebernehmen, nicht auf currentUser
  // zuruecksetzen. Sonst ueberschreibt ein Edit+Save die gesamte Organizer-Email-Liste mit
  // nur der Mail des aktuellen Editors — alle anderen Organizer wuerden stumm aus der
  // Late-Cancel- / Organizer-Mail-Verteilung rausfallen.
  //
  // Auto-Heal: wenn organizers (Names) und organizerEmails unterschiedliche Längen haben
  // (Symptom des Closure-Bugs aus v10.0–v10.2: in dem Fenster wurden Emails per prev =>
  // korrekt akkumuliert, Names aber nur einmal geschrieben → mehr Emails als Namen
  // gespeichert), wird das längere Array auf die Länge des kürzeren abgeschnitten. Sonst
  // produziert die Index-basierte Render-Logik Phantom-Chips ohne Namen oder zeigt Fotos
  // zur falschen Email — und der Picker grayt Personen aus, die sichtbar gar nicht in
  // der Liste sind, weil ihre Email noch im Array liegt.
  const [organizerEmails, setOrganizerEmails] = React.useState<string[]>(() => {
    if (!editEvent || !editEvent.organizerEmails || editEvent.organizerEmails.length === 0) {
      return editEvent && editEvent.organizers && editEvent.organizers.length > 0
        ? editEvent.organizers.map(() => '')
        : [currentUser.email];
    }
    const names = editEvent.organizers || [];
    const emails = editEvent.organizerEmails;
    if (names.length === emails.length) return emails.slice();
    if (names.length !== emails.length) {
      console.warn(
        `[DEX] EventCreationPage: organizers/organizerEmails Längen-Mismatch (${names.length} vs ${emails.length}) — `
        + `padding auf max=${Math.max(names.length, emails.length)} mit leeren Slots. `
        + `Ursache: Legacy-Daten aus v10.0–v10.2-Closure-Bug. User muss fehlende Emails per Picker nachfüllen, dann Save heilt.`
      );
    }
    const max = Math.max(names.length, emails.length);
    const padded: string[] = [];
    for (let i = 0; i < max; i++) padded.push(emails[i] || '');
    return padded;
  });
  // isSearchingOrganizer entfaellt seit v4.8.0 — Filter laeuft sync gegen den
  // bereits geladenen DEX_Roles-State, kein Async-Spinner mehr noetig.
  const isSearchingOrganizer = false;
  const organizerTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  // v10.16: Optionaler Ansprechpartner pro Event. Reines Anzeige-Feld
  // (kein App-Login, keine Permissions) — z.B. die Person vor Ort die
  // Teilnehmer bei Fragen anrufen sollen. Erscheint auf Registration- +
  // MyEvents-Seite zusätzlich zu den Organizern.
  const [contactName, setContactName] = React.useState<string>(editEvent ? (editEvent.contactName || '') : '');
  const [contactEmail, setContactEmail] = React.useState<string>(editEvent ? (editEvent.contactEmail || '') : '');
  const [contactInfo, setContactInfo] = React.useState<string>(editEvent ? (editEvent.contactInfo || '') : '');

  // v6.19: QR-Code-Scanner pro Event (beliebiger Deloitte-User, kein Admin/Organizer nötig).
  // Getrennte State-Arrays für Namen + Emails (index-synchron). Sucht via Graph-API.
  const [qrScannerNames, setQrScannerNames] = React.useState<string[]>(
    editEvent && editEvent.qrScannerNames ? editEvent.qrScannerNames.slice() : []
  );
  const [qrScannerEmails, setQrScannerEmails] = React.useState<string[]>(
    editEvent && editEvent.qrScannerEmails ? editEvent.qrScannerEmails.slice() : []
  );
  const [qrScannerSearch, setQrScannerSearch] = React.useState('');
  const [qrScannerResults, setQrScannerResults] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  // v9.18: Debounce-Timer fuer Graph-Search (statt nur Role-Filter)
  const qrScannerTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  // v9.18/v9.20: Co-Organizer-State obsolet — Organizer-Picker selbst nimmt
  // jetzt alle Deloitte-User per Graph-Search. Felder bleiben fuer
  // Backward-Compat: Events vor v9.20 koennen noch _coOrganizers haben,
  // Access-Checks lesen sie weiterhin.
  const coOrganizerNames: string[] = (editEvent && editEvent.coOrganizerNames) ? editEvent.coOrganizerNames.slice() : [];
  const coOrganizerEmails: string[] = (editEvent && editEvent.coOrganizerEmails) ? editEvent.coOrganizerEmails.slice() : [];

  // v9.21: Test-Team pro Event — Personen die das Event im Entwurfsmodus
  // sehen + sich anmelden duerfen.
  const [testTeamNames, setTestTeamNames] = React.useState<string[]>(
    editEvent && editEvent.testTeamNames ? editEvent.testTeamNames.slice() : []
  );
  const [testTeamEmails, setTestTeamEmails] = React.useState<string[]>(
    editEvent && editEvent.testTeamEmails ? editEvent.testTeamEmails.slice() : []
  );
  const [testTeamSearch, setTestTeamSearch] = React.useState('');
  const [testTeamResults, setTestTeamResults] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const testTeamTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  const [organizerIncludeIntl, setOrganizerIncludeIntl] = React.useState(false);
  const [testTeamIncludeIntl, setTestTeamIncludeIntl] = React.useState(false);
  const [qrScannerIncludeIntl, setQrScannerIncludeIntl] = React.useState(false);

  // Massenimport-Modale für die drei Team-Felder (Co-Organizer, Test-Team,
  // Check-In Team). Pattern analog zum Audience-Massenimport (Sichtbarkeits-
  // Reiter), aber pro Team-Liste eigenes Modal mit eigenem Import-Target.
  const [bulkOrganizerOpen, setBulkOrganizerOpen] = React.useState(false);
  const [bulkTestTeamOpen, setBulkTestTeamOpen] = React.useState(false);
  const [bulkQrScannerOpen, setBulkQrScannerOpen] = React.useState(false);

  // v9.21: Active-From-Datum (optional) — Event auto-aktiv ab diesem Zeitpunkt.
  const [activeFrom, setActiveFrom] = React.useState(editEvent ? (editEvent.activeFrom || '') : '');
  const [location, setLocation] = React.useState(editEvent ? editEvent.location : '');
  // Strukturierte Adresse (Straße, Hausnummer, PLZ, Ort) - separat zum freien Location-Feld
  const [addrStreet, setAddrStreet] = React.useState(editEvent?.locationAddress?.street || '');
  const [addrHouseNo, setAddrHouseNo] = React.useState(editEvent?.locationAddress?.houseNo || '');
  const [addrZip, setAddrZip] = React.useState(editEvent?.locationAddress?.zip || '');
  const [addrCity, setAddrCity] = React.useState(editEvent?.locationAddress?.city || '');
  // v18.40: Überschreibbarer Outlook-Termin-Ort. Leer = automatischer Standard
  // (Veranstaltungsort + Adresse). Gefüllt = manueller Wert. Beim Edit nur als
  // Override vorbelegen, wenn der gespeicherte Wert vom Auto-Standard abweicht —
  // sonst bleibt das Feld leer und der Ort zieht weiter automatisch nach.
  const [outlookLocationOverride, setOutlookLocationOverride] = React.useState<string>(() => {
    if (!editEvent) return '';
    const auto = buildOutlookLocation(editEvent.location, editEvent.locationAddress);
    const stored = editEvent.outlookLocation || '';
    return (stored && stored !== auto) ? stored : '';
  });
  const [locationFilter, setLocationFilter] = React.useState(
    editEvent ? editEvent.locationAudience.join(', ') : ''
  );
  const [audience, setAudience] = React.useState(
    editEvent && editEvent.audienceFilter ? editEvent.audienceFilter.join(', ') : ''
  );
  // Default fuer neue Events: 'OR' — konsistent mit EventContext-Read-Fallback
  // und konservativer (UND-Verknuepfung kann Mitarbeiter unbeabsichtigt
  // ausschliessen). Bestehende Events behalten ihren gespeicherten Wert.
  const [filterMode, setFilterMode] = React.useState<'AND' | 'OR'>(
    editEvent ? editEvent.filterMode : 'OR'
  );
  const [description, setDescription] = React.useState(editEvent ? editEvent.description : '');
  // EventType wird nicht mehr als UI-Feld abgefragt (v5.2) — neue Events:
  // aus Template abgeleitet (b2run → 'B2Run', sonst → 'Other'). Bei Edit:
  // den gespeicherten Wert beibehalten. Die Variable wird weiterhin fuer
  // Card-Gradient + B2Run-spezifische Admin-Funktionen gebraucht.
  const [storedEventType] = React.useState<EventType>(editEvent ? editEvent.type : 'Other');
  const [startDate, setStartDate] = React.useState(editEvent ? isoToLocal(editEvent.startDate) : '');
  const [endDate, setEndDate] = React.useState(editEvent ? isoToLocal(editEvent.endDate) : '');
  const [registrationDeadline, setRegistrationDeadline] = React.useState(
    editEvent ? isoToLocal(editEvent.registrationDeadline) : ''
  );
  const [lastDeregisterDate, setLastDeregisterDate] = React.useState(editEvent ? isoToLocal(editEvent.lastDeregisterDate) : '');
  // v9.22: Auto-Fill der Deadlines wenn Start-Datum gesetzt wird und die
  // Deadlines noch leer sind. Default-Logik:
  //   - RegistrationDeadline: 7 Tage vor Event-Start
  //   - LastDeregisterDate: 3 Tage vor Event-Start
  // Der Organizer kann beides ueberschreiben — wir aktualisieren NICHT,
  // wenn der User schon einen Wert gesetzt hat.
  const autoFillRanRef = React.useRef(false);
  React.useEffect(() => {
    if (autoFillRanRef.current) return;
    if (!startDate) return;
    if (registrationDeadline || lastDeregisterDate) return;
    try {
      const start = new Date(startDate);
      if (isNaN(start.getTime())) return;
      const fmt = (d: Date): string => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}T${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
      const reg = new Date(start.getTime() - 7 * 24 * 60 * 60 * 1000);
      const lastCancel = new Date(start.getTime() - 3 * 24 * 60 * 60 * 1000);
      setRegistrationDeadline(fmt(reg));
      setLastDeregisterDate(fmt(lastCancel));
      autoFillRanRef.current = true;
    } catch { /* */ }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [startDate]);
  const [maxParticipants, setMaxParticipants] = React.useState(
    editEvent && editEvent.maxParticipants ? editEvent.maxParticipants.toString() : ''
  );
  const [unlimitedParticipants, setUnlimitedParticipants] = React.useState(
    !editEvent || !editEvent.maxParticipants || editEvent.maxParticipants === 0
  );
  const [waitlistEnabled, setWaitlistEnabled] = React.useState(
    editEvent && typeof editEvent.waitlistEnabled !== 'undefined' ? editEvent.waitlistEnabled : true
  );
  const [eventImageUrl, setEventImageUrl] = React.useState(editEvent ? (editEvent.imageUrl || '') : '');
  const [imageFile, setImageFile] = React.useState<File | null>(null);
  const [imagePreview, setImagePreview] = React.useState(editEvent ? (editEvent.imageUrl || '') : '');
  // v11.20: Re-sync useEffect aus v11.19 wieder rausgenommen — der hat
  // den Wizard-State mit stale-editEvent-Daten ueberschrieben (re-sync 2
  // mit helpText="" wurde im Maintainer-DevTools beobachtet, obwohl SP
  // nachweislich helpText="Test123" hatte). Das Aufrufen von
  // setCustomFields aus dem Effect heraus war zu fragil. Stattdessen
  // verlassen wir uns wieder auf den useState-Initializer + zusaetzlich
  // ein detaillierteres Save-Log um zu sehen was *wirklich* an SP geht.
  const [customFields, setCustomFields] = React.useState<CustomFieldInput[]>(
    editEvent ? editEvent.eventSpecificFields.map(f => ({
      id: f.id, label: f.label, type: f.type, required: f.required,
      options: f.options ? [...f.options] : [], visible: true,
      ...(f.multi ? { multi: true } : {}),
      ...(f.helpText ? { helpText: f.helpText } : {}),
      ...(f.helpTextStyle === 'inline' ? { helpTextStyle: 'inline' as const } : {}),
      ...(f.showIf ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } } : {}),
      ...(f.onlyForGroup ? { onlyForGroup: f.onlyForGroup } : {}),
      // v11.94: confirmLabel beim Edit-Mount mit-übernehmen.
      ...(f.confirmLabel ? { confirmLabel: f.confirmLabel } : {}),
      // v17.20: EN-Varianten beim Edit-Mount mit-übernehmen.
      ...(f.labelEn ? { labelEn: f.labelEn } : {}),
      ...(f.helpTextEn ? { helpTextEn: f.helpTextEn } : {}),
      ...(f.confirmLabelEn ? { confirmLabelEn: f.confirmLabelEn } : {}),
      ...(f.optionsEn && f.optionsEn.length > 0 ? { optionsEn: [...f.optionsEn] } : {}),
      ...(f.externalLinks && f.externalLinks.length > 0 ? { externalLinks: f.externalLinks.map(x => ({ ...x })) } : {}),
      // v18.41: CC-bei-Mail-Flag beim Edit-Mount mit-übernehmen.
      ...(f.ccOnEmails ? { ccOnEmails: true } : {}),
    })) : []
  );
  const [outlookBody, setOutlookBody] = React.useState(editEvent ? stripOutlookWrapper(editEvent.outlookBody || '') : '');
  // Outlook-Termin-Header: beide Ueberschriften sind pro Event editierbar.
  // Default: eventTitle + formatiertes Startdatum. Parsed aus bestehendem
  // OutlookBody, falls der User sie schon angepasst hat.
  const [outlookHeading, setOutlookHeading] = React.useState(() => {
    if (editEvent) {
      const p = parseOutlookHeadings(editEvent.outlookBody || '');
      if (p.heading) return p.heading;
    }
    return editEvent ? (editEvent.title || '') : '';
  });
  const [outlookSubheading, setOutlookSubheading] = React.useState(() => {
    if (editEvent) {
      const p = parseOutlookHeadings(editEvent.outlookBody || '');
      if (p.subheading && p.subheading !== 'Event Details') return p.subheading;
    }
    return '';
  });
  // v18.42: Betreff des Outlook-Termins (leer = Event-Titel). Per-Tab gespiegelt
  // wie outlookHeading; persistiert in der DEX_Events-Spalte OutlookSubject.
  const [outlookSubject, setOutlookSubject] = React.useState<string>(editEvent?.outlookSubject || '');
  // v18.44: abweichendes Outlook-Datum (Top-Level). Leer = Event-Start/-Ende.
  // Als ISO gespeichert (wie Sub-Event-Datum); DatePicker konvertiert via isoToLocal.
  const [outlookStartOverride, setOutlookStartOverride] = React.useState<string>(editEvent?.outlookStart || '');
  const [outlookEndOverride, setOutlookEndOverride] = React.useState<string>(editEvent?.outlookEnd || '');
  // Modal-State fuer den HTML-Editor (Outlook-Body + E-Mail-Templates)
  const [htmlEditorOpen, setHtmlEditorOpen] = React.useState(false);
  const [htmlEditorMode, setHtmlEditorMode] = React.useState<'outlook' | 'email' | 'description'>('outlook');
  const [htmlEditorTemplateType, setHtmlEditorTemplateType] = React.useState<string>('');
  const [emailLanguage, setEmailLanguage] = React.useState(
    editEvent
      ? (editEvent.emailLanguage || (locale === 'de' ? 'DE' : 'EN'))
      : (locale === 'de' ? 'DE' : 'EN')
  );
  const [disableEmails, setDisableEmails] = React.useState(editEvent ? !!editEvent.disableEmails : false);
  // v19.21: granulare Sub-Schalter unter dem Master „Bestätigungs-E-Mails":
  // einzeln die Anmelde- bzw. Abmelde-Bestätigung abschaltbar (Top-Level-Event).
  const [disableRegistrationEmail, setDisableRegistrationEmail] = React.useState(editEvent ? !!editEvent.disableRegistrationEmail : false);
  const [disableCancellationEmail, setDisableCancellationEmail] = React.useState(editEvent ? !!editEvent.disableCancellationEmail : false);
  const [disableOutlook, setDisableOutlook] = React.useState(editEvent ? !!editEvent.disableOutlook : false);
  // v14.4: Acknowledgement, dass bei Top-Level-Kommunikation = AUS die
  // Teilnehmer sich für mindestens ein Sub-Event anmelden müssen. Vorausgewählt
  // für Events, die schon mit deaktivierter Kommunikation gespeichert sind
  // (alter Lauf ist bereits durch die Gate-Logik durchgekommen). Bei neuen
  // Events / frisch umgeschaltetem Toggle bleibt der Haken aus, der Save
  // ist dann blockiert bis bestätigt.
  const [mainCommDisabledAck, setMainCommDisabledAck] = React.useState<boolean>(
    !!editEvent && (!!editEvent.disableEmails || !!editEvent.disableOutlook),
  );
  // v14.5: Toggle „Anmeldung für mindestens ein Sub-Event verpflichtend".
  // Wird im RegistrationForm erzwungen — der Submit-Button blockiert, bis
  // der Teilnehmer ein Sub-Event angehakt hat. Sinnvoll wenn die Haupt-
  // Event-Kommunikation aus ist und alle Mails/Outlook-Termine nur über
  // die Sub-Events laufen.
  // v15.3: setRequireSubEventSelection wird nicht mehr direkt von der UI
  // aufgerufen — der Flag wird beim Save aus dem subEventsOnlyMode-Toggle
  // in Schritt 2 abgeleitet. State bleibt als Read-only für die Save-
  // Logik (siehe handleSubmit).
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [requireSubEventSelection, setRequireSubEventSelection] = React.useState<boolean>(
    !!editEvent && !!editEvent.requireSubEventSelection,
  );
  // v14.8: „Nur Sub-Events"-Modus. Wenn true, ist die Hauptevent-Anmeldung im
  // Teilnehmerformular ausgeblendet. Impliziert requireSubEventSelection=true.
  const [subEventsOnlyMode, setSubEventsOnlyMode] = React.useState<boolean>(
    !!editEvent && !!editEvent.subEventsOnlyMode,
  );
  // v14.8: Organizer-konfigurierbarer Begriff für die untergeordneten Events
  // (Standard „Sub-Event" / „Sub-Events", alternativ Workshop / Session etc.).
  // v15.9: separater `customTermMode`-Flag, damit „Eigene Bezeichnung…"
  // im Dropdown auch dann angeklebt bleibt, wenn beide Inputs noch leer
  // sind (sonst kippt die Heuristik unten auf 'subevent' zurück und die
  // Custom-Inputs verschwinden bevor der User tippen kann).
  const [customTermMode, setCustomTermMode] = React.useState<boolean>(false);
  const [childTermSingular, setChildTermSingular] = React.useState<string>(
    (editEvent && editEvent.childEventTermSingular) || '',
  );
  const [childTermPlural, setChildTermPlural] = React.useState<string>(
    (editEvent && editEvent.childEventTermPlural) || '',
  );
  // v8.5: Organizer-BCC-Modi (Anmeldung + Abmeldung).
  const [notifyOrgRegisterMode, setNotifyOrgRegisterMode] = React.useState<'never' | 'always' | 'fromDate'>(
    editEvent ? (editEvent.notifyOrgRegisterMode || 'never') : 'never'
  );
  const [notifyOrgRegisterFromDate, setNotifyOrgRegisterFromDate] = React.useState<string>(
    editEvent ? (editEvent.notifyOrgRegisterFromDate || '') : ''
  );
  const [notifyOrgCancelMode, setNotifyOrgCancelMode] = React.useState<'never' | 'always' | 'afterDeadline'>(
    // v10.17+: Default für neue Events ist 'afterDeadline' (Erst nach der
    // letzten Abmeldemöglichkeit) — sonst flutet jede Stornierung den
    // Organizer-Posteingang. User-Wunsch.
    editEvent ? (editEvent.notifyOrgCancelMode || 'never') : 'afterDeadline'
  );
  // v8.6: Exclude-Liste — explizit ausgeschlossene User (ueberschreiben den
  // Sichtbarkeits-Filter). UI: Modal "Sichtbare Personen anzeigen".
  const [excludedUsers, setExcludedUsers] = React.useState<string[]>(
    editEvent ? (editEvent.excludedUsers || []) : []
  );
  const [excludeModalOpen, setExcludeModalOpen] = React.useState(false);
  // v11.88: Demo-Auswahl-Modal — der „Demo"-Button oeffnet einen Dialog
  // mit vier Vorlagen-Karten (Standard, Mit Gruppen, Mit Sub-Event,
  // Mit Sub-Event + Team). Klick auf eine Karte fuellt das Formular
  // mit der jeweiligen Variante und schliesst das Modal.
  const [showDemoVariantModal, setShowDemoVariantModal] = React.useState<boolean>(false);
  // v17.21: Modal nach erfolgreichem Speichern — fragt den Organizer, ob er
  // eine A4-Zusammenfassung des Events herunterladen moechte. Pending-Payload
  // haelt die Info fuer den `dex-event-submit-success`-Dispatch, der erst
  // gefeuert wird, wenn der User im Modal eine Auswahl getroffen hat.
  const [showSummaryModal, setShowSummaryModal] = React.useState<boolean>(false);
  const [pendingSuccessDispatch, setPendingSuccessDispatch] = React.useState<{
    title: string; eventId: string; type: 'create' | 'update';
  } | null>(null);
  // v17.22: Unmount-Safety. Der Success-Dispatch (dex-event-submit-success,
  // treibt Erfolgs-Banner + Auto-Navigation in DexEventPlatform) laeuft erst,
  // wenn der User im Summary-Modal eine Auswahl trifft. Verlaesst er den
  // Wizard vorher (Header-Navigation, Browser-Back, Tab-Eviction), wuerde der
  // Dispatch sonst verloren gehen — Folge: kein Banner, kein Redirect, User
  // denkt der Save sei fehlgeschlagen. Dieser Ref + Cleanup-Effect feuert den
  // Dispatch beim Unmount nach, falls er noch aussteht.
  const pendingSuccessDispatchRef = React.useRef<{ title: string; eventId: string; type: 'create' | 'update' } | null>(null);
  React.useEffect(() => {
    return () => {
      const pending = pendingSuccessDispatchRef.current;
      if (pending) {
        pendingSuccessDispatchRef.current = null;
        try {
          window.dispatchEvent(new CustomEvent('dex-event-submit-success', { detail: pending }));
        } catch { /* */ }
      }
    };
  }, []);
  const [excludeResolvedUsers, setExcludeResolvedUsers] = React.useState<Array<{ email: string; displayName: string; firstName: string; lastName: string; jobTitle: string; location: string; source: string }>>([]);
  const [excludeResolving, setExcludeResolving] = React.useState(false);
  const [excludeSearch, setExcludeSearch] = React.useState('');
  const [excludeIncludeIntl, setExcludeIncludeIntl] = React.useState(false);
  // v8.8: Tabellen-Filter (pro Spalte) und Sortierung
  const [excludeFilters, setExcludeFilters] = React.useState<{ email: string; lastName: string; firstName: string; jobTitle: string; location: string }>({
    email: '', lastName: '', firstName: '', jobTitle: '', location: '',
  });
  const [excludeSortBy, setExcludeSortBy] = React.useState<'email' | 'lastName' | 'firstName' | 'jobTitle' | 'location'>('lastName');
  const [excludeSortDir, setExcludeSortDir] = React.useState<'asc' | 'desc'>('asc');
  const [excludePage, setExcludePage] = React.useState(0); // v8.9: 0-indexed Seite (200 pro Seite)
  const EXCLUDE_PAGE_SIZE = 200;
  // v9.16: neue Events starten standardmaessig als Test-Event — der Organizer
  // kann sich erst alles in Ruhe anschauen, das Test-Team probiert die
  // Anmeldung durch, und erst wenn alles passt wird der Schalter rausgenommen.
  const [isFictive, setIsFictive] = React.useState(editEvent ? !!editEvent.isFictive : true);
  // v18.9: Organizer-Anzeige (Chips mit Name + Foto) auf Anmelde-Seite +
  // „Meine Events" ausblenden. Rein visuell — Rechte/Mails unberührt.
  const [hideOrganizer, setHideOrganizer] = React.useState(editEvent ? !!editEvent.hideOrganizer : false);
  // Nur im Edit-Modus: standardmaessig wird der Outlook-Termin NICHT angefasst,
  // damit bei kleinen Aenderungen (z.B. Description) nicht unnoetig eine
  // "Updated meeting"-Benachrichtigung an alle Teilnehmer geht. Der Organizer
  // muss die Checkbox aktiv setzen wenn er moechte dass Titel/Start/Ende im
  // Outlook-Termin aktualisiert werden.
  const [triggerOutlookUpdate, setTriggerOutlookUpdate] = React.useState(false);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [emailTemplates, setEmailTemplates] = React.useState<Array<{ id: number; templateType: string; language: string; subject: string; heading: string; headingColor: string; bodyHtml: string }>>([]);
  const [emailTemplateOverrides, setEmailTemplateOverrides] = React.useState<Record<string, EmailOverrideEntry>>(
    editEvent?.emailTemplateOverrides ? (() => {
      try {
        const parsed = JSON.parse(editEvent.emailTemplateOverrides);
        // v11.39: Alle Piggyback-Keys rausstrippen — sie werden in separaten
        // States gehalten (emailLogoPreview, outlookLogoPreview, testTeamEmails
        // etc.) und beim Speichern frisch dazugemerged. Wenn sie hier
        // mitgeschleppt werden, überschreibt der spread `...emailTemplateOverrides`
        // am Ende von handleSubmit die frisch berechneten Werte und das
        // Entfernen z.B. eines Test-Team-Mitglieds bleibt ohne Wirkung.
        const {
          _eventLogo, _outlookLogo, _b2run,
          _qrScanners, _coOrganizers, _testTeam,
          _splitDisplayOrderReversed,
          _requireSubEventSelection,
          _subEventsOnlyMode, _childEventTerm,
          _inheritFlags, _hideOrganizer, _headerImageLayout,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          ...rest
        } = parsed as Record<string, unknown>;
        // Variablen nur destrukturiert, um sie aus `rest` zu entfernen.
        void _eventLogo; void _outlookLogo; void _b2run;
        void _qrScanners; void _coOrganizers; void _testTeam;
        void _splitDisplayOrderReversed; void _requireSubEventSelection;
        void _subEventsOnlyMode; void _childEventTerm;
        void _inheritFlags; void _hideOrganizer; void _headerImageLayout;
        return rest as Record<string, EmailOverrideEntry>;
      } catch { return {}; }
    })() : {}
  );
  // editingTemplate state entfaellt seit Modal-Migration v4.7.0
  // Custom Event-Logo fuer E-Mails (ersetzt das DEX-Orb in E-Mails).
  const [emailLogoPreview, setEmailLogoPreview] = React.useState(() => {
    if (!editEvent?.emailTemplateOverrides) return '';
    try { const o = JSON.parse(editEvent.emailTemplateOverrides); return o._eventLogo || ''; } catch { return ''; }
  });
  // Custom Event-Logo fuer Outlook-Termin (ersetzt das DEX-Orb im Termin-Body).
  // Separat vom Mail-Logo, damit man z.B. in Mails das neutrale DEX-Logo lassen
  // und im Outlook-Termin ein event-spezifisches Bild anzeigen kann.
  const [outlookLogoPreview, setOutlookLogoPreview] = React.useState(() => {
    if (!editEvent?.emailTemplateOverrides) return '';
    try { const o = JSON.parse(editEvent.emailTemplateOverrides); return o._outlookLogo || ''; } catch { return ''; }
  });
  // v18.73: Header-Bild (Event-Bild = {{ORB_URL}}) Größe + Innenabstand pro
  // Event. Gilt für Mail- UND Outlook-Termin-Kopf. Persistiert als Piggyback
  // `_headerImageLayout` in EmailTemplateOverrides. Default = bisheriges
  // Layout (Breite 180, Innenabstand 30/30).
  const [headerImageLayout, setHeaderImageLayout] = React.useState<{ width: number; paddingV: number; paddingH: number }>(() => {
    const def = { width: 180, paddingV: 30, paddingH: 30 };
    if (!editEvent?.emailTemplateOverrides) return def;
    try {
      const o = JSON.parse(editEvent.emailTemplateOverrides);
      const il = o._headerImageLayout || {};
      return {
        width: typeof il.width === 'number' && il.width > 0 ? il.width : 180,
        paddingV: typeof il.paddingV === 'number' && il.paddingV >= 0 ? il.paddingV : 30,
        paddingH: typeof il.paddingH === 'number' && il.paddingH >= 0 ? il.paddingH : 30,
      };
    } catch { return def; }
  });
  // v19.20: Snapshot des initialen Header-Bild-Layouts (Breite/Innenabstand)
  // beim Edit-Mount. Eine reine Layout-Änderung verändert NICHT den rohen
  // Outlook-Body-Text (das Layout wird erst beim Wrappen via buildOutlookBody
  // angewendet) — der Outlook-Änderungs-Detektor hätte sie deshalb übersehen.
  // Wir vergleichen das aktuelle Layout gegen diesen Snapshot, damit eine
  // Größen-/Abstands-Änderung das „Outlook-Termin aktualisieren?"-Modal genauso
  // öffnet wie eine Textänderung. useRef fixiert den Wert beim ersten Render.
  const initialHeaderImageLayoutRef = React.useRef<{ width: number; paddingV: number; paddingH: number }>(headerImageLayout);
  // v18.73: Piggyback-Konfig für den Save (leer wenn alles auf Default steht —
  // dann wird der Key gar nicht geschrieben). Wird in Create- UND Edit-Pfad
  // sowie in die Sub-Event-Overrides gemerged.
  const headerImageLayoutConfig = (headerImageLayout.width !== 180 || headerImageLayout.paddingV !== 30 || headerImageLayout.paddingH !== 30)
    ? { _headerImageLayout: { width: headerImageLayout.width, paddingV: headerImageLayout.paddingV, paddingH: headerImageLayout.paddingH } }
    : {};
  const [dragFieldId, setDragFieldId] = React.useState<string | null>(null);
  // v18.55: Pro-Feld Ein-/Ausklapp-Status für Schritt 5 (Felder). Default =
  // eingeklappt (kompakte Karte: nur Nummer + Label + Typ + Pflicht + Aktionen);
  // Detail-Einstellungen (Hilfetext, Optionen, Bedingung, CC, EN-Variante …)
  // erst beim Aufklappen. Neu hinzugefügte Felder starten aufgeklappt.
  const [fieldExpandOverride, setFieldExpandOverride] = React.useState<Record<string, boolean>>({});
  const toggleFieldExpand = (id: string, current: boolean): void =>
    setFieldExpandOverride(prev => ({ ...prev, [id]: !current }));
  const [dragOverFieldId, setDragOverFieldId] = React.useState<string | null>(null);
  // v9.28: Reorder-Mode toggelt die Hoch/Runter-Pfeile pro Custom-Field.
  // Standardmäßig aus — sonst sieht das Feld-Listing zu unruhig aus.
  const [reorderMode, setReorderMode] = React.useState(false);
  // v9.28: Modal für neuen Quiz-Bereich (statt window.prompt)
  const [newSectionModalOpen, setNewSectionModalOpen] = React.useState(false);
  const [newSectionName, setNewSectionName] = React.useState('');
  const [newSectionError, setNewSectionError] = React.useState('');
  const [agenda, setAgenda] = React.useState<AgendaItem[]>(
    editEvent && editEvent.agenda ? [...editEvent.agenda] : []
  );
  const [transferTimes, setTransferTimes] = React.useState<Array<{id: string; location: string; meetingPoint: string; address: string; date: string; departureTime: string; arrivalTime: string; description: string}>>(
    editEvent?.transferTimes?.map(t => ({...t, meetingPoint: t.meetingPoint || '', address: t.address || '', arrivalTime: t.arrivalTime || '', description: t.description || ''})) || []
  );
  const [documents, setDocuments] = React.useState<Array<{name: string; file?: File; url: string; size: number}>>(
    editEvent?.documents?.map(d => ({...d, size: d.size || 0})) || []
  );
  // Snapshot der beim Edit-Start vorhandenen Dokument-Namen, um beim Speichern
  // entfernte Attachments aus SharePoint loeschen zu koennen.
  const [initialDocumentNames] = React.useState<string[]>(
    editEvent?.documents?.map(d => d.name) || []
  );
  const [quiz, setQuiz] = React.useState<Array<{id: string; question: string; options: string[]; correctIndices: number[]; imageBase64?: string; section?: string}>>(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    editEvent?.quiz?.map(q => ({...q, correctIndices: q.correctIndices || [(q as any).correctIndex || 0], imageBase64: (q as any).imageBase64, section: (q as any).section})) || []
  );
  const quizClusterSize = editEvent?.quizClusterSize || 1;
  // Sub-Events-Drafts im UI. Seit v6.4: Sub-Events sind eigene DEX_Events-Items.
  // Beim Edit laden wir die bestehenden Child-Events und mappen sie auf Drafts.
  // Beim Save werden Drafts mit `dbId` als updateEvent, ohne als createEvent geschrieben;
  // in der DB verbliebene Child-Events, die nicht mehr im Draft sind, werden gelöscht.
  interface SubEventDraft {
    id: string;                     // Synthetische Client-ID (für React-Keys); nicht = DB-Id
    dbId?: string;                  // DEX_Events-Id wenn das Sub-Event bereits persistiert wurde
    title: string;
    description?: string;
    location?: string;
    startDate: string;
    endDate: string;
    maxParticipants?: number;
    registrationDeadline?: string;
    disableEmails?: boolean;
    // v19.22: granulare An-/Abmelde-Mail-Schalter jetzt auch pro Sub-Event.
    disableRegistrationEmail?: boolean;
    disableCancellationEmail?: boolean;
    disableOutlook?: boolean;
    /** Per-Sub-Event Custom-Fields (v10.11+). Ersetzt die hardcoded Funstarter/
     *  Durchstarter-Frage bei B2Run — wer eine zusätzliche Auswahl-Frage pro
     *  Sub-Event will, definiert sie hier individuell. Default: leeres Array
     *  (= Sub-Event ohne zusätzliche Frage, Teilnehmer wählen nur den Termin).
     *  Wird unabhängig vom Hauptevent-customFields gespeichert; mit dem Button
     *  „Vom Hauptevent kopieren" kann der Organizer die Felder duplizieren. */
    customFields?: CustomFieldInput[];
    /** v11.57: Pro-Sub-Event Kommunikations-Einstellungen (Step 5 im Wizard).
     *  Jeder Sub-Event kann eigene Mail-Sprache, Versand-Schalter, Logo-Bilder
     *  und Outlook-Termin-Texte haben. Wird beim Speichern in das jeweilige
     *  DEX_Events-Item geschrieben (siehe persistSubEventsForParent). */
    emailLanguage?: string;
    emailLogoBase64?: string;
    outlookLogoBase64?: string;
    outlookBody?: string;
    outlookHeading?: string;
    outlookSubheading?: string;
    /** v18.42: Betreff des Sub-Event-Outlook-Termins (leer = Sub-Event-Titel). */
    outlookSubject?: string;
    /** v18.44: abweichendes Outlook-Datum/-Ort des Sub-Events (ISO/Text). Leer = übernommen. */
    outlookStart?: string;
    outlookEnd?: string;
    outlookLocation?: string;
    /** v14.4: Pro-Sub-Event Mail-Text-Overrides (Anmeldung / Warteliste /
     *  Abmeldung / Nachruecken). Erlaubt es, jedem Sub-Event eigene Subjects,
     *  Headings und Bodies zu geben — Frage von 2026-05 (3 Sub-Events sollen
     *  jeweils eigene An-/Abmelde-Mails versenden können). Vorher landeten
     *  Änderungen auf einem Sub-Tab fälschlicherweise im Top-Level-Override
     *  → die Sub-Events feuerten die Haupt-Event-Texte ab. */
    emailTemplateOverrides?: Record<string, EmailOverrideEntry>;
    /** v11.57: Snapshot der initialen Outlook-relevanten Felder, um beim Save
     *  zu erkennen, ob die Teilnehmer einen Update-Termin bekommen sollen. */
    initialOutlookEventId?: string;
    initialCalendarLink?: string;
    initialTitle?: string;
    initialStartDate?: string;
    initialEndDate?: string;
    initialOutlookBody?: string;
    /** v15.0 (legacy, ungenutzt ab v15.3): Inheritance-Flags fuer
     *  pro-Sub-Event-Tabs. Mit v15.3 sind Sub-Events vollwertige Events
     *  mit eigener Konfiguration — Inherit-Flags wurden ersatzlos
     *  gestrichen. Felder bleiben optional im Interface, damit
     *  bestehende Piggyback-JSONs ohne Crash gelesen werden koennen. */
    inheritLocationFromParent?: boolean;
    inheritCapacityFromParent?: boolean;
    inheritCustomFieldsFromParent?: boolean;
    /** v15.3: pro Sub-Event eigene strukturierte Adresse (analog
     *  Hauptevent). Persistiert als JSON in `LocationAddress`-Spalte. */
    locationAddress?: { street: string; houseNo: string; zip: string; city: string };
    /** v15.3: pro Sub-Event eigene Agenda. Persistiert als JSON in
     *  `Agenda`-Spalte. */
    agenda?: AgendaItem[];
    /** v15.3: pro Sub-Event eigene Transferzeiten. Persistiert als JSON
     *  in `Transfers`-Spalte. */
    transferTimes?: Array<{ id: string; location: string; meetingPoint: string; address: string; date: string; departureTime: string; arrivalTime: string; description: string }>;
    /** v15.3: pro Sub-Event eigene Abmeldefrist. */
    lastDeregisterDate?: string;
    /** v15.3: pro Sub-Event eigener Standortfilter (comma-separated). */
    locationFilter?: string;
    /** v15.3: pro Sub-Event eigene Zielgruppe (audience-String). */
    audience?: string;
    /** v15.3: pro Sub-Event eigene Filterverknuepfung. */
    filterMode?: 'AND' | 'OR';
    /** v15.3: pro Sub-Event eigene Warteliste an/aus. */
    waitlistEnabled?: boolean;
    /** v15.3: pro Sub-Event eigene Anrede-Abfrage. */
    askSalutation?: boolean;
  }
  const [subEvents, setSubEvents] = React.useState<SubEventDraft[]>(() => {
    if (!editEvent) return [];
    const kids = childEventsOf(editEvent.id);
    return kids.map(k => {
      // v11.57: Pro-Sub-Event Logo-Bilder aus EmailTemplateOverrides
      // (Piggyback-Pattern, gleich wie Top-Level-Event).
      // v14.4: zusätzlich die Mail-Text-Overrides (Anmeldung/Warteliste/
      // Abmeldung/Nachruecken) — vorher landeten Edits auf Sub-Event-Tabs
      // versehentlich beim Haupt-Event.
      let emailLogo = '';
      let outlookLogo = '';
      let subOverrides: Record<string, EmailOverrideEntry> = {};
      // v15.0: Inheritance-Flags aus dem Piggyback-JSON lesen. Wenn der
      // Flag nicht persistiert wurde (alte Events) faellt die App auf
      // datenbasierte Heuristik zurueck (siehe weiter unten).
      let inheritFlagsRaw: { capacity?: boolean; fields?: boolean; location?: boolean } | undefined;
      try {
        const ov = JSON.parse(k.emailTemplateOverrides || '{}') as Record<string, unknown>;
        emailLogo = (ov?._eventLogo as string) || '';
        outlookLogo = (ov?._outlookLogo as string) || '';
        inheritFlagsRaw = (ov?._inheritFlags as { capacity?: boolean; fields?: boolean; location?: boolean } | undefined);
        // Piggyback-Keys (mit Unterstrich-Prefix) rausstrippen, der Rest sind
        // die echten Mail-Template-Overrides pro TemplateType.
        const filtered: Record<string, EmailOverrideEntry> = {};
        for (const key of Object.keys(ov)) {
          if (key.startsWith('_')) continue;
          const val = ov[key] as Partial<EmailOverrideEntry> | undefined;
          if (val && (val.subject || val.heading || val.bodyHtml || val.headingColor || val.headingFontSize || val.subheading !== undefined || val.headingBold !== undefined || val.headingItalic !== undefined || val.subheadingColor || val.subheadingFontSize || val.subheadingBold !== undefined || val.subheadingItalic !== undefined)) {
            filtered[key] = {
              subject: val.subject || '',
              heading: val.heading || '',
              bodyHtml: val.bodyHtml || '',
              // v18.19/v18.22: Überschrift-Farbe/-Größe/-Stil + Subheading-
              // Formatierung mit-übernehmen.
              ...(val.subheading !== undefined ? { subheading: val.subheading } : {}),
              ...(val.headingColor ? { headingColor: val.headingColor } : {}),
              ...(val.headingFontSize ? { headingFontSize: val.headingFontSize } : {}),
              ...(val.headingBold !== undefined ? { headingBold: val.headingBold } : {}),
              ...(val.headingItalic !== undefined ? { headingItalic: val.headingItalic } : {}),
              ...(val.subheadingColor ? { subheadingColor: val.subheadingColor } : {}),
              ...(val.subheadingFontSize ? { subheadingFontSize: val.subheadingFontSize } : {}),
              ...(val.subheadingBold !== undefined ? { subheadingBold: val.subheadingBold } : {}),
              ...(val.subheadingItalic !== undefined ? { subheadingItalic: val.subheadingItalic } : {}),
            };
          }
        }
        subOverrides = filtered;
      } catch { /* */ }
      const parsedHeads = parseOutlookHeadings(k.outlookBody || '');
      // v15.0: Inheritance-Heuristik fuer Bestands-Events: wenn das
      // Piggyback-Flag fehlt UND das jeweilige Datenfeld nicht-leer ist,
      // gilt es als „eigener Wert" (nicht vom Hauptevent geerbt). Wenn
      // das Feld leer ist, default = uebernehmen.
      const inheritCap = inheritFlagsRaw && typeof inheritFlagsRaw.capacity === 'boolean'
        ? inheritFlagsRaw.capacity
        : !(k.maxParticipants && k.maxParticipants > 0);
      const inheritFields = inheritFlagsRaw && typeof inheritFlagsRaw.fields === 'boolean'
        ? inheritFlagsRaw.fields
        : !((k.eventSpecificFields || []).length > 0);
      const inheritLoc = inheritFlagsRaw && typeof inheritFlagsRaw.location === 'boolean'
        ? inheritFlagsRaw.location
        : !(k.location && k.location.trim().length > 0);
      return {
      id: k.id,
      dbId: k.id,
      title: k.title,
      description: k.description,
      location: k.location,
      startDate: k.startDate,
      endDate: k.endDate,
      maxParticipants: k.maxParticipants || 0,
      registrationDeadline: k.registrationDeadline,
      disableEmails: k.disableEmails,
      disableRegistrationEmail: k.disableRegistrationEmail,
      disableCancellationEmail: k.disableCancellationEmail,
      disableOutlook: k.disableOutlook,
      // v11.57: pro-Sub-Event Kommunikations-Felder laden
      emailLanguage: k.emailLanguage || (locale === 'de' ? 'DE' : 'EN'),
      emailLogoBase64: emailLogo,
      outlookLogoBase64: outlookLogo,
      emailTemplateOverrides: subOverrides,
      outlookBody: stripOutlookWrapper(k.outlookBody || ''),
      outlookHeading: parsedHeads.heading || k.title || '',
      outlookSubheading: parsedHeads.subheading && parsedHeads.subheading !== 'Event Details' ? parsedHeads.subheading : '',
      outlookSubject: k.outlookSubject || '',
      outlookStart: k.outlookStart || '',
      outlookEnd: k.outlookEnd || '',
      outlookLocation: k.outlookLocation || '',
      // v11.57: Snapshot der initialen Outlook-relevanten Felder
      initialOutlookEventId: k.outlookEventId || '',
      // v11.61: CalendarLink (iCalUId) als Outlook-Existenz-Indikator. Der
      // Flow schreibt OutlookEventId nicht — auf erfolgreichen Sub-Events
      // ist nur CalendarLink gefuellt.
      initialCalendarLink: k.calendarLink || '',
      initialTitle: k.title || '',
      initialStartDate: k.startDate || '',
      initialEndDate: k.endDate || '',
      initialOutlookBody: k.outlookBody || '',
      customFields: (k.eventSpecificFields || []).map(f => ({
        id: f.id,
        label: f.label,
        type: f.type as CustomFieldInput['type'],
        required: !!f.required,
        options: f.options || [],
        // EventSpecificField hat kein 'visible'-Feld — default auf true.
        // Sichtbarkeit ist im Storage immer „shown" (default), nur im Wizard
        // kann der User Felder ausblenden.
        visible: true,
        externalLinks: f.externalLinks,
        multi: f.multi,
        helpText: f.helpText,
        helpTextStyle: f.helpTextStyle,
        showIf: f.showIf,
      })),
      // v15.3: pro-Sub-Event Felder aus dem Event-Datenmodell laden. Alle
      // Sub-Events haben jetzt eigene Adresse, Agenda, Transferzeiten,
      // Deadline, Standortfilter, Audience, Filter-Modus, Warteliste und
      // Anrede-Toggle — wie der Hauptevent.
      locationAddress: k.locationAddress ? {
        street: k.locationAddress.street || '',
        houseNo: k.locationAddress.houseNo || '',
        zip: k.locationAddress.zip || '',
        city: k.locationAddress.city || '',
      } : { street: '', houseNo: '', zip: '', city: '' },
      agenda: (k.agenda || []) as AgendaItem[],
      transferTimes: (k.transferTimes || []).map(tt => ({
        id: tt.id,
        location: tt.location || '',
        meetingPoint: tt.meetingPoint || '',
        address: tt.address || '',
        date: tt.date || '',
        departureTime: tt.departureTime || '',
        arrivalTime: tt.arrivalTime || '',
        description: tt.description || '',
      })),
      lastDeregisterDate: k.lastDeregisterDate || '',
      // Form-Felder fuer Standortfilter / Mailverteiler sind comma-separated
      // Strings, persistiert im Event aber als Arrays — siehe Top-Level-Mapping.
      locationFilter: (k.locationAudience || []).join(', '),
      audience: (k.audienceFilter || []).join(', '),
      filterMode: (k.filterMode === 'AND' ? 'AND' : 'OR') as 'AND' | 'OR',
      waitlistEnabled: typeof k.waitlistEnabled === 'boolean' ? k.waitlistEnabled : true,
      askSalutation: !!k.askSalutation,
      // v15.0 (legacy): Inheritance-Flags werden seit v15.3 nicht mehr
      // ausgewertet. Bleiben in den geparsten Drafts, weil das Schema
      // sie noch erlaubt — Wirkung gleich Null.
      inheritLocationFromParent: inheritLoc,
      inheritCapacityFromParent: inheritCap,
      inheritCustomFieldsFromParent: inheritFields,
    };
    });
  });
  // v11.57: aktiv ausgewaehlter Tab in Step 6 (Kommunikation, v11.80 Renumbering). 0 = Haupt-Event,
  // N>0 = subEvents[N-1]. Beim Tab-Wechsel werden die Step-5-Felder zwischen
  // dem Top-Level-State und der jeweiligen Sub-Event-Slice gespiegelt — siehe
  // switchCommTab-Helper weiter unten.
  const [activeCommTabIdx, setActiveCommTabIdx] = React.useState<number>(0);
  // v15.0: pro-Event-Tabs in Schritt 3 (Ort), Schritt 4 (Kapazitaet) und
  // Schritt 6 (Felder). 0 = Haupt-Event, N>0 = subEvents[N-1]. Im
  // Gegensatz zu Step 6 (Kommunikation) gibt es hier KEIN Mirror-Pattern,
  // weil die per-Tab-Werte direkt im subEvents[]-State (location,
  // maxParticipants, customFields) bzw. in den Top-Level-States stehen —
  // jedes Eingabefeld liest/schreibt direkt aus seinem Zielort.
  const [activeLocationTabIdx, setActiveLocationTabIdx] = React.useState<number>(0);
  const [activeCapacityTabIdx, setActiveCapacityTabIdx] = React.useState<number>(0);
  const [activeFieldsTabIdx, setActiveFieldsTabIdx] = React.useState<number>(0);
  // v11.60: synchroner Spiegel von subEvents fuer Save/Detect-Pfade. React-
  // State-Updates sind async — wenn flushActiveCommTabToState() per
  // setSubEvents(prev=>...) den aktiven Tab in die jeweilige Slice
  // zurueckschreibt, sieht das direkt danach laufende
  // detectOutlookRelevantChanges() (und auch persistSubEventsForParent
  // unter handleSubmit) noch die alte Array aus dem Closure. Ergebnis:
  // Modal kommt nicht, und die Sub-Event-Aenderung wird beim Schreiben
  // wieder mit dem Original ueberschrieben. Der Ref haelt die jeweils
  // aktuellste Array synchron — alle Code-Pfade, die nach einem Flush
  // lesen, gehen ueber `subEventsRef.current`.
  const subEventsRef = React.useRef<typeof subEvents>(subEvents);
  React.useEffect(() => { subEventsRef.current = subEvents; }, [subEvents]);
  // Snapshot der beim Edit-Start vorhandenen Sub-Event-DB-IDs, um beim Save
  // entfernte Sub-Events zu löschen.
  const [initialSubEventDbIds] = React.useState<string[]>(() => {
    if (!editEvent) return [];
    return childEventsOf(editEvent.id).map(k => k.id);
  });
  // Snapshot der initialen Outlook-Metadaten pro Sub-Event (DisableOutlook +
  // OutlookEventId + SubsiteUrl). Wird beim Save gebraucht, um zu erkennen, ob
  // ein Sub-Event nachträglich von „Outlook deaktiviert" auf „Outlook
  // aktiviert" gedreht wurde — der DEX_CreateOutlookEvent-Flow lauscht nur
  // auf NEUE DEX_Events-Items (GetOnNewItems-Trigger), deshalb muss das
  // betroffene Sub-Event in diesem Fall gelöscht und neu angelegt werden,
  // damit überhaupt ein Outlook-Termin entsteht.
  const [initialSubEventOutlookMeta] = React.useState<Record<string, { disableOutlook: boolean; outlookEventId: string; subsiteUrl: string; registrationListName: string }>>(() => {
    if (!editEvent) return {};
    const acc: Record<string, { disableOutlook: boolean; outlookEventId: string; subsiteUrl: string; registrationListName: string }> = {};
    for (const k of childEventsOf(editEvent.id)) {
      acc[k.id] = {
        disableOutlook: !!k.disableOutlook,
        outlookEventId: k.outlookEventId || '',
        subsiteUrl: k.subsiteUrl || '',
        // v11.69: Subsite-Events nutzen immer die Standard-Teilnehmerliste
        // "Teilnehmer" (siehe REG_LIST_NAME in EventService). Wird beim
        // Recreate-Pfad an `createEvent({ existingRegistrationListName })`
        // mitgegeben, damit der Reuse-Branch in createEvent greift.
        registrationListName: 'Teilnehmer',
      };
    }
    return acc;
  });
  // v11.57: Snapshot der initialen Outlook-relevanten Felder des Top-Level-
  // Events (Title, Start, End, OutlookBody). Wird beim Save mit den aktuellen
  // Werten verglichen — Aenderung loest das Update-Confirm-Modal aus.
  // Im Ref, weil wir das einmal beim Mount fixieren und nicht bei Re-Renders
  // neu setzen wollen.
  const initialOutlookSnapshot = React.useRef<{ title: string; startDate: string; endDate: string; outlookBody: string; outlookLocation: string; outlookSubject: string; outlookStart: string; outlookEnd: string }>({
    title: editEvent?.title || '',
    startDate: editEvent?.startDate || '',
    endDate: editEvent?.endDate || '',
    outlookBody: editEvent?.outlookBody || '',
    // v18.44: abweichendes Outlook-Datum in den Snapshot — eine Override-Änderung
    // soll das Update-Modal öffnen.
    outlookStart: editEvent?.outlookStart || '',
    outlookEnd: editEvent?.outlookEnd || '',
    // v18.34/v18.40: effektiver Ort in den Snapshot (gespeicherte Override ODER
    // Auto). Eine reine Ort-Aenderung soll das Outlook-Update-Modal oeffnen.
    outlookLocation: editEvent?.outlookLocation || buildOutlookLocation(editEvent?.location, editEvent?.locationAddress),
    // v18.42: Betreff in den Snapshot — eine reine Betreff-Aenderung soll das
    // Outlook-Update-Modal ebenfalls oeffnen.
    outlookSubject: editEvent?.outlookSubject || '',
  });
  // v11.57: Update-Confirm-Modal-State. Beim Save mit Outlook-relevanten
  // Aenderungen oeffnen wir das Modal und warten auf die Entscheidung des
  // Organizers. v11.63: Statt einem globalen "Outlook-Update senden ja/nein"
  // listet das Modal jetzt jedes geaenderte Event einzeln (Hauptevent +
  // betroffene Sub-Events) und der Organizer setzt pro Event einen Haken.
  const [outlookConfirmOpen, setOutlookConfirmOpen] = React.useState(false);
  // v11.63: Snapshot der Detect-Items zum Modal-Open-Zeitpunkt. Jeder Eintrag
  // beschreibt ein Event (Hauptevent oder Sub-Event) mit Outlook-relevanten
  // Aenderungen — Title, Start, End oder OutlookBody — und listet, welche
  // Felder sich geaendert haben (fuer die Anzeige als Sub-Text pro Item).
  type OutlookConfirmItem = {
    kind: 'top' | 'sub';
    eventId: string;
    title: string;
    changedFields: Array<'title' | 'startDate' | 'endDate' | 'outlookBody' | 'location' | 'subject' | 'layout'>;
    /** v11.68: Sub-Event hat noch keinen Outlook-Termin (kein CalendarLink in
     *  DEX_Events). Body-/Titel-Change wird beim Save in DEX_Events
     *  persistiert, aber es kann KEIN UpdateEvent gequeuet werden — es gibt
     *  keinen Outlook-Termin, an den die Teilnehmer eine Notification kriegen
     *  koennten. Im Modal wird das Item statt mit Checkbox als
     *  Info-Eintrag mit Erklaerung gerendert. */
    noOutlookYet?: boolean;
  };
  const [outlookConfirmItems, setOutlookConfirmItems] = React.useState<OutlookConfirmItem[]>([]);
  // v11.63: Pro Event-ID, ob die Checkbox im Modal angehakt ist.
  // true = UpdateEvent in Queue + OutlookDirty=false setzen.
  // false (oder nicht im Map) = kein UpdateEvent, OutlookDirty=true setzen.
  const [outlookConfirmChecks, setOutlookConfirmChecks] = React.useState<Record<string, boolean>>({});
  // v11.63: Top-Level-Outlook-Update-Entscheidung. true = nach erfolgreichem
  // updateEvent ein DEX_Outlook 'UpdateEvent' in die Queue schreiben.
  const pendingOutlookUpdateForTopRef = React.useRef<boolean>(false);
  // Sub-Event-IDs, fuer die ein DEX_Outlook 'UpdateEvent' angefordert wurde.
  const pendingOutlookUpdateForSubEventsRef = React.useRef<string[]>([]);
  // v11.69: Sub-Event-IDs, fuer die ein *Recreate* des DEX_Events-Items
  // angefordert wurde (Outlook-Termin nachtraeglich anlegen ohne Teilnehmer-
  // Verlust). Werden in `persistSubEventsForParent` aufgegriffen: das alte
  // DEX_Events-Item wird per `deleteEventItemOnly` entfernt (Subsite +
  // Teilnehmerliste bleiben unangetastet), dann wird per
  // `createEvent({ existingSubsiteUrl, existingRegistrationListName })` ein
  // neues DEX_Events-Item angelegt, das die alte Subsite wiederverwendet.
  // Der `DEX_CreateOutlookEvent`-Flow triggert auf das neue Item und legt den
  // Outlook-Termin an.
  const pendingOutlookRecreateForSubEventsRef = React.useRef<string[]>([]);
  // v11.63: Pro Event-ID der gewuenschte OutlookDirty-Wert.
  // Nur Eventd-IDs, die im Modal waren, werden hier gesetzt — alle anderen
  // bleiben unberuehrt (kein OutlookDirty-Patch).
  const pendingOutlookDirtyWriteRefs = React.useRef<Record<string, boolean>>({});
  // v11.57 (kompatibel): Schreibwert fuer OutlookDirty fuer das Top-Level-
  // Event im naechsten updateEvent-Call. null = nicht setzen, false/true =
  // setzen. Wird aus pendingOutlookDirtyWriteRefs[topId] abgeleitet.
  const pendingOutlookDirtyWriteRef = React.useRef<boolean | null>(null);
  // Bereiche, die per "+ Bereich"-Button angelegt aber noch nicht mit einer
  // Frage belegt wurden. Sobald eine Frage per Drag&Drop reinkommt, ergibt sich
  // der Section-Name aus dem question.section-Feld selbst — pendingSections
  // hält nur die noch leeren Zwischen-Buckets.
  const [pendingSections, setPendingSections] = React.useState<string[]>([]);
  const [draggedQuestionId, setDraggedQuestionId] = React.useState<string | null>(null);
  // v6.35: Handbuch-Previews können einen bestimmten Wizard-Schritt gezielt
  // zeigen, indem sie vor dem Mount `window.__dexPreviewInitialStep = <n>`
  // setzen (0..6). Nur für Read-only-Previews; in der echten App ist das
  // Flag nie gesetzt, dann bleibt der Default 0 (Step 1 "Grundlagen").
  const [currentStep, setCurrentStep] = React.useState<number>(() => {
    if (typeof window !== 'undefined') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const init = (window as any).__dexPreviewInitialStep;
      if (typeof init === 'number' && init >= 0 && init <= 7) return init;
    }
    return 0;
  });
  const [selectedTemplate, setSelectedTemplate] = React.useState<'blank' | 'b2run'>('blank');
  // EventType wird bei neuen Events aus dem Template abgeleitet; bei Edit
  // bleibt der gespeicherte Wert erhalten.
  const eventType: EventType = editEvent ? storedEventType : (selectedTemplate === 'b2run' ? 'B2Run' : 'Other');
  const [b2runStartblocks, setB2runStartblocks] = React.useState<string[]>([]);
  const [newStartblock, setNewStartblock] = React.useState<string>('');
  const [durchstarterCapacity, setDurchstarterCapacity] = React.useState<string>(
    editEvent && typeof editEvent.durchstarterCapacity === 'number' ? String(editEvent.durchstarterCapacity) : ''
  );
  const [funstarterCapacity, setFunstarterCapacity] = React.useState<string>(
    editEvent && typeof editEvent.funstarterCapacity === 'number' ? String(editEvent.funstarterCapacity) : ''
  );
  // v10.20: frei waehlbare Bezeichnungen fuer die zwei Kapazitaets-Gruppen.
  // Default leer; wenn der User die Split-Capacity einschaltet ohne Label
  // zu setzen, fallen Wizard und RegistrationPage auf 'Durchstarter' /
  // 'Funstarter' zurueck (Backward-Compat fuer B2Run-Events vor v10.20).
  const [splitLabelA, setSplitLabelA] = React.useState<string>(
    (editEvent && editEvent.splitLabelA) || ''
  );
  const [splitLabelB, setSplitLabelB] = React.useState<string>(
    (editEvent && editEvent.splitLabelB) || ''
  );
  // v10.20: Warteliste-Modus bei aktiver Split-Capacity. Default false =
  // getrennte Wartelisten pro Gruppe (alter B2Run-Stil). true = eine
  // gemeinsame Warteliste, FIFO ueber beide Gruppen hinweg.
  const [splitSharedWaitlist, setSplitSharedWaitlist] = React.useState<boolean>(
    !!editEvent?.splitSharedWaitlist
  );
  // v11.25: Anzeige-Reihenfolge der zwei Gruppen-Karten in der Registrierung
  // umkehren. Pure UI-Toggle — interne Daten (splitLabelA/B, Kapazitäten,
  // StarterType auf Anmeldungen) bleiben unangetastet.
  const [splitDisplayOrderReversed, setSplitDisplayOrderReversed] = React.useState<boolean>(
    !!editEvent?.splitDisplayOrderReversed
  );
  // v11.0: Teilnehmer-Upload aktivieren + optionaler Hinweistext
  const [allowAttendeeUpload, setAllowAttendeeUpload] = React.useState<boolean>(
    !!editEvent?.allowAttendeeUpload
  );
  const [attendeeUploadHint, setAttendeeUploadHint] = React.useState<string>(
    editEvent?.attendeeUploadHint || ''
  );
  const [attendeeUploadLabel, setAttendeeUploadLabel] = React.useState<string>(
    editEvent?.attendeeUploadLabel || ''
  );
  // v11.80: Anrede im Registrierungsformular abfragen (Default false). Wenn
  // false, wird das Anrede-Dropdown ausgeblendet und ein leerer String als
  // Anrede gespeichert. Wird im neuen Schritt 5 (Felder) konfiguriert.
  const [askSalutation, setAskSalutation] = React.useState<boolean>(
    !!editEvent?.askSalutation
  );
  // v18.75: Sicherheitshinweis vor dem Absenden der Anmeldung (Schritt 5, ganz
  // unten). Default aus. Modus 'summary' = Auswahl-Übersicht (Haupt-/Sub-Events
  // mit De-/Selektieren), 'freetext' = eigener Hinweis-Text.
  const [confirmDialogEnabled, setConfirmDialogEnabled] = React.useState<boolean>(!!editEvent?.confirmDialogEnabled);
  const [confirmDialogMode, setConfirmDialogMode] = React.useState<string>(editEvent?.confirmDialogMode || 'summary');
  const [confirmDialogText, setConfirmDialogText] = React.useState<string>(editEvent?.confirmDialogText || '');
  // v18.35: Anmeldesprache vorgeben. '' = App-Sprache (Default), 'de' / 'en' =
  // Anmeldeseite (inkl. Disclaimer) immer in dieser Sprache anzeigen.
  const [registrationLanguage, setRegistrationLanguage] = React.useState<'' | 'de' | 'en'>(
    editEvent?.registrationLanguage === 'de' || editEvent?.registrationLanguage === 'en' ? editEvent.registrationLanguage : ''
  );
  // v18.33: Self-Check-in — Teilnehmer checken sich selbst per QR-Code ein.
  // Konfiguration in Schritt 3 (Kapazität & Sichtbarkeit). Beim Aktivieren
  // wird einmalig ein geheimer Token generiert (Schlüssel für statischen Link
  // + rotierenden HMAC-QR) und ein Erklär-Modal geöffnet.
  const [selfCheckInEnabled, setSelfCheckInEnabled] = React.useState<boolean>(
    !!editEvent?.selfCheckInEnabled
  );
  const [selfCheckInToken, setSelfCheckInToken] = React.useState<string>(
    editEvent?.selfCheckInToken || ''
  );
  const [selfCheckInFrom, setSelfCheckInFrom] = React.useState<string>(
    editEvent?.selfCheckInFrom || ''
  );
  const [selfCheckInTo, setSelfCheckInTo] = React.useState<string>(
    editEvent?.selfCheckInTo || ''
  );
  // Erklär-Modal beim Aktivieren des Self-Check-ins.
  const [showSelfCheckInModal, setShowSelfCheckInModal] = React.useState<boolean>(false);
  // v11.80: Team-Anmeldung — eine Person meldet ein ganzes Team an.
  // Konfiguration im neuen Schritt 4 (Team-Anmeldung). Die tatsächliche
  // Multi-Person-Anmelde-Logik folgt mit v11.81+; aktuell wird nur die
  // Konfiguration persistiert.
  const [teamRegistrationEnabled, setTeamRegistrationEnabled] = React.useState<boolean>(
    !!editEvent?.teamRegistrationEnabled
  );
  const [teamSize, setTeamSize] = React.useState<number>(
    typeof editEvent?.teamSize === 'number' && editEvent.teamSize > 0 ? editEvent.teamSize : 4
  );
  const [askTeamName, setAskTeamName] = React.useState<boolean>(
    !!editEvent?.askTeamName
  );
  // v11.81: Erweiterte Team-Konfiguration — Beitritts-Modus, Sichtbarkeit
  // offener Slots, Lead-Approval. Die tatsächliche Team-Anmelde-Logik
  // (Multi-Person-Form, Mails, Outlook) folgt mit v11.82+.
  const [teamPartialAllowed, setTeamPartialAllowed] = React.useState<boolean>(
    !!editEvent?.teamPartialAllowed
  );
  const [teamOpenSlotsVisible, setTeamOpenSlotsVisible] = React.useState<boolean>(
    !!editEvent?.teamOpenSlotsVisible
  );
  const [teamJoinRequiresApproval, setTeamJoinRequiresApproval] = React.useState<boolean>(
    !!editEvent?.teamJoinRequiresApproval
  );
  // v17.20: Bilingual-Toggle — wenn an, kann der Organizer pro Custom-Field
  // (Label, Help-Text, Checkbox-Confirm-Text, Dropdown-Optionen) eine
  // englische Variante hinterlegen. Wird im Wizard-Schritt 5 ganz oben als
  // separater Toggle eingestellt; die EN-Inputs blenden pro Card auf, wenn
  // der Toggle aktiv ist.
  const [bilingualFields, setBilingualFields] = React.useState<boolean>(
    !!editEvent?.bilingualFields
  );
  // v6.15: Starter-Typ → Startblock-Zuordnung + Leistungsnachweis-Pflicht
  const [durchstarterStartblock, setDurchstarterStartblock] = React.useState<string>(
    editEvent?.durchstarterStartblock || ''
  );
  const [funstarterStartblock, setFunstarterStartblock] = React.useState<string>(
    editEvent?.funstarterStartblock || ''
  );
  // v10.24: setDurchstarterRequiresProof wird nicht mehr aufgerufen — der UI-
  // Toggle in Schritt 3 ist entfallen, das Feature wird durch Pro-Gruppe-
  // Custom-Fields in Schritt 4 ersetzt. State bleibt erhalten, damit
  // bestehende Events mit gesetztem Wert nicht beim Save den Wert verlieren
  // (durchstarterRequiresProof wird beim Persist mitgeschrieben falls
  // editEvent das Flag schon hatte).
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [durchstarterRequiresProof, setDurchstarterRequiresProof] = React.useState<boolean>(
    !!editEvent?.durchstarterRequiresProof
  );
  const [showPreview, setShowPreview] = React.useState(false);
  const [showRegisterPreview, setShowRegisterPreview] = React.useState(false);
  const [triedNext, setTriedNext] = React.useState(false);
  const [previewSections, setPreviewSections] = React.useState<Array<{ id: string; label: string }>>([
    { id: 'event', label: 'Event-Karte' },
    { id: 'personal', label: 'Personal Information' },
    { id: 'specific', label: 'Event specific Information' },
    { id: 'actions', label: 'Buttons' },
  ]);
  const [dragSectionId, setDragSectionId] = React.useState<string | null>(null);
  const [dragOverSectionId, setDragOverSectionId] = React.useState<string | null>(null);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [progress, setProgress] = React.useState(0);
  const [progressLabel, setProgressLabel] = React.useState('');
  const [showEmailModal, setShowEmailModal] = React.useState(false);
  // v8.9: Cache der aufgeloesten Verteiler-Members fuer den Sichtbarkeits-
  // Check. Wird beim Oeffnen des Modals einmal befuellt, damit jeder
  // Such-Treffer in O(1) gegen die Mailgruppen-Mitgliedschaft geprueft
  // werden kann (statt pro Treffer alle Verteiler erneut auszuloesen).
  const [visibilityAudienceCache, setVisibilityAudienceCache] = React.useState<Set<string>>(new Set());
  const [visibilityCacheLoading, setVisibilityCacheLoading] = React.useState(false);
  const [emailSearch, setEmailSearch] = React.useState('');
  const [emailSearchResults, setEmailSearchResults] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const [emailSearchIncludeIntl, setEmailSearchIncludeIntl] = React.useState(false);
  const [isSearchingEmails, setIsSearchingEmails] = React.useState(false);
  const [submitted, setSubmitted] = React.useState(false);
  const [error, setError] = React.useState('');
  const [imageUploadError, setImageUploadError] = React.useState('');
  const [iconPickerOpen, setIconPickerOpen] = React.useState<string | null>(null);
  const [iconSearch, setIconSearch] = React.useState('');
  const [showAllIcons, setShowAllIcons] = React.useState(false);

  const locationOptions = ['Berlin', 'Dresden', 'Düsseldorf', 'Frankfurt', 'Görlitz', 'Halle', 'Hamburg', 'Hannover', 'Köln', 'Leipzig', 'Magdeburg', 'Mannheim', 'München', 'Nürnberg', 'Stuttgart', 'Walldorf', 'All'];

  const addCustomField = (): void => {
    const newId = `cf-${Date.now()}`;
    setCustomFields([...customFields, {
      id: newId, label: '', type: 'text',
      required: false, options: [], visible: true,
    }]);
    // v18.55: neues Feld direkt aufgeklappt, damit man es sofort ausfüllen kann.
    setFieldExpandOverride(prev => ({ ...prev, [newId]: true }));
  };

  /**
   * Deloitte-Standard-Vorschlaege als Katalog. Der Organizer waehlt ueber ein
   * Modal mit Checkboxen aus, welche dieser Felder hinzugefuegt werden sollen.
   * Ausgewaehlte Felder werden ans Ende der aktuellen customFields angehaengt.
   */
  // Bilingual: Labels + Optionen der Felder werden in der Event-Sprache (DE/EN)
  // angelegt, passend zum Locale beim Klick auf 'Vorgeschlagene Felder'.
  // v10.21: Catalog mit Kategorien — 'general' (default ausgeklappt) und
  // 'b2run' (default eingeklappt). Damit ersetzt das Suggested-Modal den
  // alten Template-Dropdown: User wählt fokussiert die Felder, die er
  // wirklich braucht, statt einen B2Run-Block auf einmal aufzuziehen.
  type SuggestedCategory = 'general' | 'b2run';
  // v10.23: jeder Suggested-Field-Eintrag hat ein Fluent-UI-Icon (visuelles
  // Erkennungsmerkmal in der Auswahl-Liste) und einen ausfuehrlicheren
  // Tooltip-Text — der erklaert dem Organizer, was das Feld in der App
  // bewirkt, ohne dass er es erst hinzufuegen muss.
  type SuggestedEntry = { key: string; label: string; description: string; category: SuggestedCategory; icon: string; tooltip?: string; build: (_now: number) => CustomFieldInput };
  const SUGGESTED_FIELDS_CATALOG: SuggestedEntry[] = isDe ? [
    {
      key: 'tshirt', category: 'general', icon: 'Tag',
      label: 'T-Shirt Größe',
      description: 'Dropdown mit Kein T-Shirt / XS–XXL',
      build: (n) => ({ id: `cf-${n}`, label: 'T-Shirt Größe', type: 'select', required: false, options: ['Habe bereits ein T-Shirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'allergies', category: 'general', icon: 'Warning',
      label: 'Allergien',
      description: 'Freitextfeld für Allergien/Unverträglichkeiten',
      build: (n) => ({ id: `cf-${n}`, label: 'Allergien', type: 'text', required: false, options: [], visible: true }),
    },
    {
      key: 'diet', category: 'general', icon: 'EatDrink',
      label: 'Essenspräferenzen',
      description: 'Dropdown: Keine Präferenzen / Vegetarisch / Vegan / Pescetarisch',
      build: (n) => ({ id: `cf-${n}`, label: 'Essenspräferenzen', type: 'select', required: false, options: ['Keine Präferenzen', 'Vegetarisch', 'Vegan', 'Pescetarisch'], visible: true }),
    },
    {
      key: 'hotel', category: 'general', icon: 'Hotel',
      label: 'Hotel benötigt',
      description: 'Checkbox: Teilnehmer benötigt ein Hotel',
      build: (n) => ({ id: `cf-${n}`, label: 'Hotel benötigt', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'roomtype', category: 'general', icon: 'Room',
      label: 'Zimmerart',
      description: 'Dropdown: Keine Präferenz / Einzelzimmer / Doppelzimmer',
      build: (n) => ({ id: `cf-${n}`, label: 'Zimmerart (falls Hotel benötigt)', type: 'select', required: false, options: ['Keine Präferenz', 'Einzelzimmer', 'Doppelzimmer'], visible: true }),
    },
    {
      key: 'roommate', category: 'general', icon: 'People',
      label: 'Bevorzugter Zimmerpartner',
      description: 'Personen-Suche; Match-Erkennung im Admin Center',
      build: (n) => ({ id: `cf-${n}`, label: 'Bevorzugter Zimmerpartner (bei Doppelzimmer)', type: 'roommate', required: false, options: [], visible: true }),
    },
    // B2Run-Pakete — nur fuer Lauf-Events relevant. Sektion ist im Modal
    // standardmaessig eingeklappt, damit der Standard-Organizer sie nicht
    // versehentlich aktiviert.
    {
      key: 'b2run_startblock', category: 'b2run', icon: 'Running',
      label: 'Startblock',
      description: 'Dropdown der Startbloecke. Optionen werden nachtraeglich im Wizard gepflegt.',
      build: (_n) => ({ id: `b2run_startblock`, label: 'Startblock', type: 'select', required: true, options: [], visible: true }),
    },
    {
      key: 'b2run_gruppe', category: 'b2run', icon: 'BulletedList',
      label: 'Gruppe',
      description: 'Dropdown: offene Klasse / Nordic Walker / Damen / Herren',
      build: (_n) => ({ id: `b2run_gruppe`, label: 'Gruppe', type: 'select', required: true, options: ['offene Klasse', 'Nordic Walker', 'Damen', 'Herren'], visible: true }),
    },
    {
      key: 'b2run_altersklasse', category: 'b2run', icon: 'Calendar',
      label: 'Altersklasse',
      description: 'Dropdown: unter 18 / 18-29 / 30-39 / 40-49 / 50-59 / 60+',
      build: (_n) => ({ id: `b2run_altersklasse`, label: 'Altersklasse', type: 'select', required: true, options: ['unter 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true }),
    },
    {
      key: 'b2run_infoservice', category: 'b2run', icon: 'CellPhone',
      label: 'Infoservice (SMS)',
      description: 'Checkbox: aktiviert die Mobilnummer-Pflicht fuer den B2Run-SMS-Service',
      build: (_n) => ({ id: `b2run_infoservice`, label: 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_mobilnummer', category: 'b2run', icon: 'Phone',
      label: 'Mobilnummer',
      description: 'Freitext, dynamisch Pflicht wenn Infoservice aktiv',
      build: (_n) => ({ id: `b2run_mobilnummer`, label: 'Mobilnummer (nur bei aktiviertem Infoservice)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } }),
    },
    {
      key: 'b2run_anonym', category: 'b2run', icon: 'Hide3',
      label: 'Anonym teilnehmen',
      description: 'Checkbox: Teilnehmer in Ergebnislisten anonymisieren',
      build: (_n) => ({ id: `b2run_anonym`, label: 'Anonym teilnehmen', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_laufshirt', category: 'b2run', icon: 'Sport',
      label: 'Deloitte-Laufshirt',
      description: 'Dropdown: vorhandenes Shirt / XS-XXL',
      build: (_n) => ({ id: `b2run_laufshirt`, label: 'Deloitte-Laufshirt', type: 'select', required: true, options: ['Habe bereits ein Laufshirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'b2run_datenschutz', category: 'b2run', icon: 'LockShield',
      label: 'AGB / Datenschutz',
      description: 'Pflicht-Checkbox mit Links zu B2Run-AGB und Datenschutzerklaerung',
      build: (_n) => ({
        id: `b2run_datenschutz`,
        label: 'Zustimmung AGB, Datenschutz & Bildaufnahmen',
        type: 'checkbox', required: true, options: [], visible: true,
        externalLinks: [
          { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
          { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
        ],
      }),
    },
  ] : [
    {
      key: 'tshirt', category: 'general', icon: 'Tag',
      label: 'T-Shirt size',
      description: 'Dropdown: No t-shirt needed / XS–XXL',
      build: (n) => ({ id: `cf-${n}`, label: 'T-Shirt size', type: 'select', required: false, options: ['I already have one', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'allergies', category: 'general', icon: 'Warning',
      label: 'Allergies',
      description: 'Free-text field for allergies / intolerances',
      build: (n) => ({ id: `cf-${n}`, label: 'Allergies', type: 'text', required: false, options: [], visible: true }),
    },
    {
      key: 'diet', category: 'general', icon: 'EatDrink',
      label: 'Dietary preferences',
      description: 'Dropdown: No preference / Vegetarian / Vegan / Pescetarian',
      build: (n) => ({ id: `cf-${n}`, label: 'Dietary preferences', type: 'select', required: false, options: ['No preference', 'Vegetarian', 'Vegan', 'Pescetarian'], visible: true }),
    },
    {
      key: 'hotel', category: 'general', icon: 'Hotel',
      label: 'Hotel required',
      description: 'Checkbox: participant needs a hotel room',
      build: (n) => ({ id: `cf-${n}`, label: 'Hotel required', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'roomtype', category: 'general', icon: 'Room',
      label: 'Room type',
      description: 'Dropdown: No preference / Single room / Double room',
      build: (n) => ({ id: `cf-${n}`, label: 'Room type (if hotel needed)', type: 'select', required: false, options: ['No preference', 'Single room', 'Double room'], visible: true }),
    },
    {
      key: 'roommate', category: 'general', icon: 'People',
      label: 'Preferred roommate',
      description: 'People search; match detection in the admin center',
      build: (n) => ({ id: `cf-${n}`, label: 'Preferred roommate (for double room)', type: 'roommate', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_startblock', category: 'b2run', icon: 'Running',
      label: 'Start block',
      description: 'Dropdown of start blocks. Options are added later in the wizard.',
      build: (_n) => ({ id: `b2run_startblock`, label: 'Start block', type: 'select', required: true, options: [], visible: true }),
    },
    {
      key: 'b2run_gruppe', category: 'b2run', icon: 'BulletedList',
      label: 'Category',
      description: 'Dropdown: Open class / Nordic Walker / Women / Men',
      build: (_n) => ({ id: `b2run_gruppe`, label: 'Category', type: 'select', required: true, options: ['Open class', 'Nordic Walker', 'Women', 'Men'], visible: true }),
    },
    {
      key: 'b2run_altersklasse', category: 'b2run', icon: 'Calendar',
      label: 'Age group',
      description: 'Dropdown: under 18 / 18-29 / 30-39 / 40-49 / 50-59 / 60+',
      build: (_n) => ({ id: `b2run_altersklasse`, label: 'Age group', type: 'select', required: true, options: ['under 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true }),
    },
    {
      key: 'b2run_infoservice', category: 'b2run', icon: 'CellPhone',
      label: 'Info service (SMS)',
      description: 'Checkbox: enables the mandatory mobile-number for the B2Run SMS service',
      build: (_n) => ({ id: `b2run_infoservice`, label: 'Use B2Run info service (SMS — mobile number required)', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_mobilnummer', category: 'b2run', icon: 'Phone',
      label: 'Mobile number',
      description: 'Free text, dynamically required when info service is active',
      build: (_n) => ({ id: `b2run_mobilnummer`, label: 'Mobile number (only if info service is enabled)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } }),
    },
    {
      key: 'b2run_anonym', category: 'b2run', icon: 'Hide3',
      label: 'Anonymous participation',
      description: 'Checkbox: anonymise attendee in result lists',
      build: (_n) => ({ id: `b2run_anonym`, label: 'Participate anonymously', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_laufshirt', category: 'b2run', icon: 'Sport',
      label: 'Deloitte running shirt',
      description: 'Dropdown: existing shirt / XS-XXL',
      build: (_n) => ({ id: `b2run_laufshirt`, label: 'Deloitte running shirt', type: 'select', required: true, options: ['I already have one', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'b2run_datenschutz', category: 'b2run', icon: 'LockShield',
      label: 'Terms / privacy',
      description: 'Required checkbox with links to B2Run terms and privacy policy',
      build: (_n) => ({
        id: `b2run_datenschutz`,
        label: 'I agree to the terms, privacy policy and photo/video recordings',
        type: 'checkbox', required: true, options: [], visible: true,
        externalLinks: [
          { label: 'Terms (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
          { label: 'Privacy (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
        ],
      }),
    },
  ];

  const [showSuggestedModal, setShowSuggestedModal] = React.useState(false);
  const [suggestedSelection, setSuggestedSelection] = React.useState<Record<string, boolean>>({});
  // v10.21: B2Run-Sektion im Suggested-Modal default eingeklappt — die meisten
  // Organizer brauchen sie nicht; soll nicht visuell uebernehmen.
  const [showB2runSuggested, setShowB2runSuggested] = React.useState(false);

  const openSuggestedModal = (): void => {
    // v9.17: Standard ist KEINS ausgewaehlt — User waehlt aktiv aus, was er
    // wirklich braucht. Vorher waren alle vorgewaehlt, was zu unbeabsichtigt
    // viele uebernommenen Feldern fuehrte.
    setSuggestedSelection({});
    setShowSuggestedModal(true);
  };

  const addSelectedSuggestedFields = (): void => {
    const selected = SUGGESTED_FIELDS_CATALOG.filter(s => suggestedSelection[s.key]);
    if (selected.length === 0) { setShowSuggestedModal(false); return; }
    const now = Date.now();
    const newFields: CustomFieldInput[] = selected.map((s, i) => s.build(now + i));
    // v10.21: B2Run-Felder haben deterministische IDs (b2run_startblock etc.).
    // Wenn ein Feld mit gleicher ID schon im customFields-Array steht, skippen
    // wir es — sonst entstehen Duplikate, wenn der User das Modal mehrfach
    // oeffnet. Allgemeine Felder (cf-<timestamp>) bekommen eindeutige IDs und
    // werden immer angehaengt.
    const existingIds = new Set(customFields.map(f => f.id));
    const dedupedNewFields = newFields.filter(f => !existingIds.has(f.id));
    setCustomFields([...customFields, ...dedupedNewFields]);
    setShowSuggestedModal(false);
  };

  const removeCustomField = (id: string): void => {
    setCustomFields(customFields.filter(f => f.id !== id));
  };

  const updateCustomField = (id: string, updates: Partial<CustomFieldInput>): void => {
    setCustomFields(customFields.map(f => f.id === id ? { ...f, ...updates } : f));
  };

  // === Sub-Event Custom-Field Helpers (v10.11+) =============================
  // Per-Sub-Event Custom-Fields ersetzen die hardcoded Funstarter/Durchstarter-
  // Frage. Pattern parallel zu den Hauptevent-Helpers — operieren aber auf dem
  // `customFields[]` eines spezifischen SubEventDraft (nach Client-`id`
  // identifiziert). Funktional minimaler als die Hauptevent-Variante (kein
  // Suggested-Modal, kein showIf für v1), reicht aber für „Auswahlfrage pro
  // Sub-Event mit individuellem Label + Optionen".
  const addSubEventCustomField = (subEventId: string): void => {
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: [...(se.customFields || []), {
        id: `cf_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
        label: '',
        type: 'select',
        required: false,
        options: ['', ''],
        visible: true,
      }],
    })));
  };
  const removeSubEventCustomField = (subEventId: string, fieldId: string): void => {
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: (se.customFields || []).filter(f => f.id !== fieldId),
    })));
  };
  const updateSubEventCustomField = (subEventId: string, fieldId: string, updates: Partial<CustomFieldInput>): void => {
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: (se.customFields || []).map(f => f.id === fieldId ? { ...f, ...updates } : f),
    })));
  };
  const copyParentFieldsToSubEvent = (subEventId: string): void => {
    // Dupliziert die Hauptevent-Felder ins Sub-Event mit frischen IDs (sonst
    // kollidieren Field-IDs zwischen Parent und Children, was bei Validierungs-
    // Logik und showIf-Refs zu Konflikten führen würde).
    const cloned: CustomFieldInput[] = customFields.map(f => ({
      ...f,
      id: `cf_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      options: f.options.slice(),
      externalLinks: f.externalLinks ? f.externalLinks.map(x => ({ ...x })) : undefined,
      showIf: undefined,  // showIf-Refs würden auf Parent-Field-IDs zeigen, droppen
    }));
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: cloned,
    })));
  };

  /**
   * Template-Auswahl: setzt EventType und Custom Fields automatisch.
   * B2Run: legt alle Pflichtfelder fuer die Anmeldung bei b2run.com an
   * (laut Excel "Deloitte_Teilnehmer_innen_B2Run_Koeln_2025_v4.xlsx").
   *
   * v10.21: Template-Dropdown im Wizard entfaellt; B2Run-Felder werden ueber
   * das Suggested-Felder-Modal einzeln gewaehlt. Diese Funktion bleibt fuer
   * eventuelle programmatische Aufrufer (Edit-Modus, Migrations-Skripte)
   * erhalten — sie wird im aktuellen UI nicht mehr aufgerufen.
   */
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const applyTemplate = (template: 'blank' | 'b2run'): void => {
    setSelectedTemplate(template);
    if (template === 'blank') {
      // v7.20-Fix: NICHT alle Fields loeschen — nur die B2Run-spezifischen
      // (Praefix "b2run_"). So gehen Custom-Felder, die der Organizer manuell
      // angelegt hat, beim Deselect des B2Run-Templates nicht verloren.
      setCustomFields(prev => prev.filter(f => !f.id.startsWith('b2run_')));
      setB2runStartblocks([]);
      return;
    }
    if (template === 'b2run') {
      // Custom Fields in der Reihenfolge der B2Run-Excel-Spalten
      // Hinweis: Strasse/PLZ/Stadt werden NICHT abgefragt (werden leer in der Excel stehen)
      // Locale-abhaengige Labels/Optionen. IDs bleiben konstant, damit die
      // B2Run-Logik (Infoservice -> Mobilnummer, CSV-Export etc.) unabhaengig
      // von der Sprache funktioniert.
      const fields: CustomFieldInput[] = isDe ? [
        { id: 'b2run_startblock', label: 'Startblock', type: 'select', required: true, options: [...b2runStartblocks], visible: true },
        { id: 'b2run_gruppe', label: 'Gruppe', type: 'select', required: true, options: ['offene Klasse', 'Nordic Walker', 'Damen', 'Herren'], visible: true },
        { id: 'b2run_altersklasse', label: 'Altersklasse', type: 'select', required: true, options: ['unter 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true },
        { id: 'b2run_infoservice', label: 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_mobilnummer', label: 'Mobilnummer (nur bei aktiviertem Infoservice)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } },
        { id: 'b2run_anonym', label: 'Anonym teilnehmen', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_laufshirt', label: 'Deloitte-Laufshirt', type: 'select', required: true, options: ['Habe bereits ein Laufshirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true },
        {
          id: 'b2run_datenschutz',
          label: 'Zustimmung AGB, Datenschutz & Bildaufnahmen',
          type: 'checkbox',
          required: true,
          options: [],
          visible: true,
          externalLinks: [
            { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
            { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
          ],
        },
      ] : [
        { id: 'b2run_startblock', label: 'Start block', type: 'select', required: true, options: [...b2runStartblocks], visible: true },
        { id: 'b2run_gruppe', label: 'Category', type: 'select', required: true, options: ['Open class', 'Nordic Walker', 'Women', 'Men'], visible: true },
        { id: 'b2run_altersklasse', label: 'Age group', type: 'select', required: true, options: ['under 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true },
        { id: 'b2run_infoservice', label: 'Use B2Run info service (SMS — mobile number required)', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_mobilnummer', label: 'Mobile number (only if info service is enabled)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } },
        { id: 'b2run_anonym', label: 'Participate anonymously', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_laufshirt', label: 'Deloitte running shirt', type: 'select', required: true, options: ['I already have one', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true },
        {
          id: 'b2run_datenschutz',
          label: 'I agree to the terms, privacy policy and photo/video recordings',
          type: 'checkbox',
          required: true,
          options: [],
          visible: true,
          externalLinks: [
            { label: 'Terms (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
            { label: 'Privacy (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
          ],
        },
      ];
      // v7.20-Fix: bestehende NON-b2run-Felder erhalten und die B2Run-Felder
      // anhaengen (vorher: setCustomFields(fields) hat alles ueberschrieben).
      setCustomFields(prev => {
        const nonB2run = prev.filter(f => !f.id.startsWith('b2run_'));
        return [...nonB2run, ...fields];
      });
    }
  };

  // Startbloecke-Aenderung direkt in das Custom Field uebernehmen
  React.useEffect(() => {
    if (selectedTemplate !== 'b2run' && !(isEditMode && customFields.some(f => f.id === 'b2run_startblock'))) return;
    setCustomFields(prev => prev.map(f =>
      f.id === 'b2run_startblock' ? { ...f, options: [...b2runStartblocks] } : f
    ));
  }, [b2runStartblocks]);

  // Edit-Mode: Wenn das Event B2Run-Custom-Fields hat, Startbloecke aus dem Field laden
  React.useEffect(() => {
    if (!isEditMode) return;
    const sb = customFields.find(f => f.id === 'b2run_startblock');
    if (sb && b2runStartblocks.length === 0 && sb.options && sb.options.length > 0) {
      const parts = sb.options.map(s => s.trim()).filter(Boolean);
      setB2runStartblocks(parts);
      setSelectedTemplate('b2run');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEditMode]);

  const addStartblock = (): void => {
    const trimmed = newStartblock.trim();
    if (!trimmed) return;
    if (b2runStartblocks.indexOf(trimmed) >= 0) { setNewStartblock(''); return; }
    setB2runStartblocks([...b2runStartblocks, trimmed]);
    setNewStartblock('');
  };

  const removeStartblock = (block: string): void => {
    setB2runStartblocks(b2runStartblocks.filter(b => b !== block));
  };

  // Bei B2Run: maxParticipants automatisch aus Summe von Durchstarter + Funstarter berechnen
  const isB2runTemplate = selectedTemplate === 'b2run' || (isEditMode && customFields.some(f => f.id === 'b2run_startblock'));
  // Seit v6.5: explizite Checkbox in Schritt 3 ("Lauf-Event mit getrennten
  // Starter-Kapazitäten") statt versteckt über das Template gesteuert.
  // Initial-Wert: beim Edit aus vorhandenen Kapazitäten abgeleitet, bei neuem
  // Event true wenn B2Run-Template gewählt wurde.
  const [useSplitCapacities, setUseSplitCapacities] = React.useState<boolean>(() => {
    if (editEvent) {
      return typeof editEvent.durchstarterCapacity === 'number'
        && typeof editEvent.funstarterCapacity === 'number'
        && (editEvent.durchstarterCapacity > 0 || editEvent.funstarterCapacity > 0);
    }
    return selectedTemplate === 'b2run';
  });
  // Automatisch aktivieren wenn B2Run-Template nachträglich gewählt wird.
  React.useEffect(() => {
    if (!editEvent && selectedTemplate === 'b2run') setUseSplitCapacities(true);
  }, [selectedTemplate, editEvent]);

  React.useEffect(() => {
    if (!useSplitCapacities) return;
    const d = parseInt(durchstarterCapacity, 10) || 0;
    const f = parseInt(funstarterCapacity, 10) || 0;
    const sum = d + f;
    if (sum > 0) setMaxParticipants(String(sum));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [durchstarterCapacity, funstarterCapacity, useSplitCapacities]);

  // v11.88: Helpers fuer Datums-Formatierung — werden von allen Demo-
  // Varianten + dem alten fillDemo geteilt.
  const fmtDatetime = (d: Date): string => {
    const pad = (n: number): string => (n < 10 ? '0' : '') + n;
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  };
  const fmtDate = (d: Date): string => {
    const pad = (n: number): string => (n < 10 ? '0' : '') + n;
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  };

  // v11.88: Reset-Helfer — setzt alle Team-, Split- und sonstigen
  // Variant-spezifischen Felder auf neutralen Default zurueck, damit die
  // Demo-Varianten nicht versehentlich Zustand der vorigen Variante erben.
  const resetDemoVariantBaseState = (): void => {
    setUseSplitCapacities(false);
    setSplitLabelA('Teilnehmergruppe 1');
    setSplitLabelB('Teilnehmergruppe 2');
    setDurchstarterCapacity('0');
    setFunstarterCapacity('0');
    setSplitSharedWaitlist(false);
    setTeamRegistrationEnabled(false);
    setTeamSize(4);
    setAskTeamName(false);
    setTeamPartialAllowed(false);
    setTeamOpenSlotsVisible(false);
    setTeamJoinRequiresApproval(false);
    setAskSalutation(false);
    setConfirmDialogEnabled(false); // v18.75: Sicherheitshinweis-Default
    setConfirmDialogMode('summary');
    setConfirmDialogText('');
    setSubEvents([]);
    setCustomFields([]);
    setAgenda([]);
    setTransferTimes([]);
    setLocationFilter('');
    setAudience('');
    setEventImageUrl('');
    setContactName('');
    setContactEmail('');
    setContactInfo('');
    setWaitlistEnabled(true);
    setEmailLanguage('DE');
  };

  // v11.88: Berechnet einen Demo-Termin relativ zu „heute".
  // daysAhead = Anzahl Tage in der Zukunft, hour/minute = Start.
  const demoDate = (daysAhead: number, hour: number, minute: number): Date => {
    const d = new Date();
    d.setDate(d.getDate() + daysAhead);
    d.setHours(hour, minute, 0, 0);
    return d;
  };

  // v11.88: Vier Demo-Vorlagen — vom „Demo"-Button-Modal aufgerufen.
  // Jede Variante fuellt das Formular vollstaendig (inkl. Reset der
  // Felder, die diese Variante NICHT setzt).
  const loadDemoStandard = (): void => {
    resetDemoVariantBaseState();
    const start = demoDate(14, 10, 0);
    const end = demoDate(14, 12, 0);
    const deadline = demoDate(13, 23, 59);
    setTitle('Demo-Meeting Standard');
    setDescription('Beispielhaftes einfaches Meeting ohne Gruppen und ohne Sub-Events.');
    setLocation('Heinrich Campus Düsseldorf, 6. Etage');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    setRegistrationDeadline(fmtDate(deadline));
    setLastDeregisterDate(fmtDate(deadline));
    setMaxParticipants('50');
    setUseSplitCapacities(false);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    const tDemo = Date.now();
    setCustomFields([
      { id: `cf-${tDemo}`, label: 'Essenspräferenz', type: 'select', required: true,
        options: ['Vegetarisch', 'Vegan', 'Keine Einschränkungen'], visible: true },
    ]);
    setCurrentStep(0);
  };

  const loadDemoGroups = (): void => {
    resetDemoVariantBaseState();
    const start = demoDate(14, 9, 0);
    const end = demoDate(14, 17, 0);
    const deadline = demoDate(12, 23, 59);
    setTitle('Demo-Workshop mit Gruppen');
    setDescription('Workshop mit zwei Teilnehmer-Gruppen (Vormittag/Nachmittag) und gemeinsamer Warteliste.');
    setLocation('Deloitte Office Köln');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    setRegistrationDeadline(fmtDate(deadline));
    setLastDeregisterDate(fmtDate(deadline));
    setMaxParticipants('50');
    setUseSplitCapacities(true);
    setSplitLabelA('Vormittag');
    setSplitLabelB('Nachmittag');
    setDurchstarterCapacity('25');
    setFunstarterCapacity('25');
    setSplitSharedWaitlist(true);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    setCurrentStep(0);
  };

  const loadDemoSubEvent = (): void => {
    resetDemoVariantBaseState();
    const start = demoDate(21, 9, 0);
    const end = demoDate(21, 17, 0);
    const deadline = demoDate(18, 23, 59);
    setTitle('Demo-Conference mit Dinner');
    setDescription('Hauptkonferenz + abendliches Dinner als getrenntes Sub-Event mit eigener Anmeldung.');
    setLocation('Deloitte Office Hamburg');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    setRegistrationDeadline(fmtDate(deadline));
    setLastDeregisterDate(fmtDate(deadline));
    setMaxParticipants('100');
    setUseSplitCapacities(false);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    const tDemo = Date.now();
    setCustomFields([
      { id: `cf-${tDemo}`, label: 'Hotel-Buchung', type: 'select', required: false,
        options: ['Ja, ich brauche ein Hotel', 'Nein, ich reise abends ab'], visible: true },
    ]);
    const dinnerStart = demoDate(21, 18, 0);
    const dinnerEnd = demoDate(21, 22, 0);
    setSubEvents([
      {
        id: `se-${tDemo}`,
        title: 'Networking-Dinner',
        startDate: fmtDatetime(dinnerStart),
        endDate: fmtDatetime(dinnerEnd),
        registrationDeadline: '',
        location: 'Restaurant Fischmarkt',
        description: 'Optionales Networking-Dinner im Anschluss an die Konferenz.',
        maxParticipants: 60,
        disableEmails: false,
        disableOutlook: false,
        customFields: [],
      },
    ]);
    setCurrentStep(0);
  };

  const loadDemoSubEventTeam = (): void => {
    resetDemoVariantBaseState();
    const start = demoDate(14, 18, 0);
    const end = demoDate(14, 22, 0);
    const deadline = demoDate(9, 23, 59);
    setTitle('Demo-Kneipenquiz mit Team-Anmeldung');
    setDescription('Quizabend, bei dem ganze Teams über das Anmeldeformular angemeldet werden.');
    setLocation('Heinrich Campus Düsseldorf, 6. Etage, Dachterrasse');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    setRegistrationDeadline(fmtDate(deadline));
    setLastDeregisterDate(fmtDate(deadline));
    setMaxParticipants('80');
    setUseSplitCapacities(false);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    setTeamRegistrationEnabled(true);
    setTeamSize(4);
    setAskTeamName(true);
    setTeamPartialAllowed(true);
    setTeamOpenSlotsVisible(true);
    setTeamJoinRequiresApproval(false);
    const tDemo = Date.now();
    setCustomFields([
      { id: `cf-${tDemo}`, label: 'Essenspräferenz', type: 'select', required: true,
        options: ['Vegetarisch', 'Vegan', 'Keine Einschränkungen'], visible: true },
    ]);
    const briefStart = demoDate(14, 17, 0);
    const briefEnd = demoDate(14, 17, 30);
    setSubEvents([
      {
        id: `se-${tDemo}`,
        title: 'Vorbereitungs-Briefing (Quizmaster)',
        startDate: fmtDatetime(briefStart),
        endDate: fmtDatetime(briefEnd),
        registrationDeadline: '',
        location: 'Heinrich Campus Düsseldorf, 6. Etage, Dachterrasse',
        description: 'Kurzes Briefing für die Quizmaster-Helfer vor dem Event.',
        maxParticipants: 10,
        disableEmails: false,
        disableOutlook: false,
        customFields: [],
      },
    ]);
    setCurrentStep(0);
  };

  // v11.88: Variant-Map fuer den Demo-Button. Key entspricht der Karten-
  // Auswahl im Modal, Value ist die Loader-Funktion oben.
  const DEMO_VARIANTS: Record<'standard' | 'groups' | 'subevent' | 'subeventTeam', () => void> = {
    standard: loadDemoStandard,
    groups: loadDemoGroups,
    subevent: loadDemoSubEvent,
    subeventTeam: loadDemoSubEventTeam,
  };


  const moveCustomField = (id: string, direction: 'up' | 'down'): void => {
    const idx = customFields.findIndex(f => f.id === id);
    if (idx < 0) return;
    const newIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (newIdx < 0 || newIdx >= customFields.length) return;
    const updated = [...customFields];
    [updated[idx], updated[newIdx]] = [updated[newIdx], updated[idx]];
    setCustomFields(updated);
  };

  // ===== Quiz helpers =====
  const addQuizQuestion = (): void => {
    setQuiz([...quiz, { id: `q-${Date.now()}`, question: '', options: ['', ''], correctIndices: [0] }]);
  };
  const removeQuizQuestion = (id: string): void => {
    setQuiz(quiz.filter(q => q.id !== id));
  };
  const updateQuizQuestion = (id: string, updates: Partial<{question: string; options: string[]; correctIndices: number[]; imageBase64: string | undefined; section: string | undefined}>): void => {
    setQuiz(quiz.map(q => q.id === id ? { ...q, ...updates } : q));
  };

  // ===== Agenda helpers =====
  const addAgendaItem = (): void => {
    setAgenda([...agenda, {
      id: `ag-${Date.now()}`,
      date: startDate ? startDate.slice(0, 10) : '',
      time: '',
      endTime: '',
      icon: 'Calendar',
      title: '',
      description: '',
    }]);
  };

  const removeAgendaItem = (id: string): void => {
    setAgenda(agenda.filter(a => a.id !== id));
  };

  const updateAgendaItem = (id: string, updates: Partial<AgendaItem>): void => {
    setAgenda(agenda.map(a => a.id === id ? { ...a, ...updates } : a));
  };

  /**
   * Persistiert die Sub-Event-Drafts nach dem Parent-Save. Seit v6.4 sind Sub-Events
   * eigene DEX_Events-Items mit gesetztem parentEventId.
   *
   * - Drafts **ohne dbId** → `createEvent({ ..., parentEventId })`
   * - Drafts **mit dbId** und Werte-Diff → `updateEvent(dbId, patch)`
   * - Initial vorhandene Sub-Event-DB-IDs, die **nicht** mehr als Draft existieren → `deleteEvent(id)` (kaskadierend inkl. Subsite + Kalendertermin)
   *
   * Alle Sub-Events erben Metadaten (Organizer, Audience, Email-Language,
   * Templates, Logos) vom Parent — eigenständig sind nur Titel, Daten, Ort und Kapazität.
   */
  const persistSubEventsForParent = async (parentEventId: string): Promise<void> => {
    const keptDbIds = new Set<string>();
    // v11.87: Sub-Event-Progress-Callback aus dem aufrufenden handleSubmit
    // einspeisen. Der Caller setzt window.__dexSubEventProgress vor dem
    // Aufruf und entfernt es danach. Wenn nicht gesetzt: no-op.
    const subOnProgress = (stage: string): void => {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const cb = (window as any).__dexSubEventProgress;
        if (typeof cb === 'function') cb(stage);
      } catch { /* */ }
    };
    // Sub-Events erben Organizer + OrganizerEmail vom Parent. Einmal sanitisieren
    // statt pro Iteration, identisch für alle Children.
    const sanitizedOrgPair = sanitizeOrganizerPairs();
    // v11.60: aus dem Ref iterieren — der React-State ist beim Save evtl.
    // noch nicht propagiert, weil flushActiveCommTabToState() per setState
    // erst async wirkt. Der Ref haelt synchron die letzten Tab-Werte.
    for (const draft of subEventsRef.current) {
      if (!draft.title || !draft.title.trim()) continue; // leere Drafts ignorieren
      // v11.57: Pro-Sub-Event Kommunikations-Felder. Wenn der Organizer fuer
      // den Sub-Event eigene Werte in Step 5 gesetzt hat, verwenden wir die;
      // sonst fallback auf die Top-Level-Werte (Backward-Compat fuer
      // Sub-Events ohne eigene Communication-Einstellungen).
      const subEmailLang = draft.emailLanguage || emailLanguage;
      const subOutlookBodyRaw = (typeof draft.outlookBody === 'string' && draft.outlookBody !== '') ? draft.outlookBody : '';
      const subOutlookHeading = draft.outlookHeading || draft.title || '';
      const subOutlookSub = draft.outlookSubheading || '';
      const subOutlookSubject = (draft.outlookSubject || '').trim();
      const subEmailLogo = draft.emailLogoBase64 || '';
      const subOutlookLogo = draft.outlookLogoBase64 || '';
      // Outlook-Body wrappen, wenn vorhanden — sonst leer lassen, der Flow
      // bzw. Create-Pfad setzt einen Default-Body.
      let wrappedSubOutlookBody = '';
      if (subOutlookBodyRaw) {
        const vars = {
          EventTitle: draft.title.trim(),
          Organizer: organizer,
          Location: draft.location || '',
          Address: '',
          StartDate: draft.startDate ? new Date(draft.startDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
          EndDate: draft.endDate ? new Date(draft.endDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
        };
        const resolvedBody = replacePlaceholders(subOutlookBodyRaw, vars);
        const resolvedHead = subOutlookHeading ? replacePlaceholders(subOutlookHeading, vars) : draft.title.trim();
        const resolvedSub2 = subOutlookSub ? replacePlaceholders(subOutlookSub, vars) : undefined;
        // v18.73: Sub-Events erben das Header-Bild-Layout des Hauptevents.
        const wrapped = buildOutlookBody(resolvedHead, resolvedBody, resolvedSub2, { imageWidth: headerImageLayout.width, imagePaddingV: headerImageLayout.paddingV, imagePaddingH: headerImageLayout.paddingH });
        wrappedSubOutlookBody = wrapped.replace(/\{\{ORB_URL\}\}/g, subOutlookLogo || getCachedOrbBase64() || '');
      }
      // Sub-Event-EmailTemplateOverrides: Logo-Piggybacks (Top-Level-Pattern)
      // + ab v14.4 die echten Mail-Text-Overrides pro Sub-Event
      // (Anmeldung/Warteliste/Abmeldung/Nachruecken).
      const subDraftOverrides = draft.emailTemplateOverrides || {};
      // v15.3: Inheritance-Flags entfallen — Sub-Events sind seit v15.3
      // vollwertige Events mit eigener Konfiguration. Der Piggyback-Key
      // `_inheritFlags` wird nicht mehr geschrieben.
      const subOverridesMerged: Record<string, unknown> = {
        ...subDraftOverrides,
        ...(subEmailLogo ? { _eventLogo: subEmailLogo } : {}),
        ...(subOutlookLogo ? { _outlookLogo: subOutlookLogo } : {}),
        // v18.73: Sub-Event erbt das Header-Bild-Layout des Hauptevents, damit
        // auch die Sub-Event-Mails den gleichen Bild-Kopf nutzen.
        ...headerImageLayoutConfig,
      };
      const subEmailOverrides = Object.keys(subOverridesMerged).length > 0
        ? JSON.stringify(subOverridesMerged)
        : '';
      // v15.3: Sub-Event-eigene strukturierte Adresse serialisieren (analog
      // zum Hauptevent-Top-Level-Pattern). Wenn alle vier Komponenten leer
      // sind, wird ein leerer String gespeichert.
      const draftAddr = draft.locationAddress || { street: '', houseNo: '', zip: '', city: '' };
      const draftHasAddress = !!(draftAddr.street || draftAddr.houseNo || draftAddr.zip || draftAddr.city);
      const draftLocationAddress = draftHasAddress ? JSON.stringify(draftAddr) : '';
      const draftAgendaJson = JSON.stringify(draft.agenda || []);
      const draftTransfersJson = JSON.stringify(draft.transferTimes || []);
      const childPayload = {
        title: draft.title.trim(),
        type: 'Other',
        status: 'Active',
        description: draft.description || '',
        location: draft.location || '',
        locationAddress: draftLocationAddress,
        locationFilter: draft.locationFilter || '',
        audience: draft.audience || '',
        filterMode: (draft.filterMode === 'AND' ? 'AND' : 'OR'),
        startDate: draft.startDate || '',
        endDate: draft.endDate || '',
        registrationDeadline: draft.registrationDeadline || '',
        lastDeregisterDate: draft.lastDeregisterDate || '',
        maxParticipants: draft.maxParticipants || 0,
        waitlistEnabled: typeof draft.waitlistEnabled === 'boolean' ? draft.waitlistEnabled : true,
        eventImageUrl: '',
        organizer: sanitizedOrgPair.orgString,
        organizerEmail: sanitizedOrgPair.orgEmailString,
        outlookEventId: '',
        outlookBody: wrappedSubOutlookBody,
        outlookSubject: subOutlookSubject || undefined,
        outlookStart: (draft.outlookStart || '') || undefined,
        outlookEnd: (draft.outlookEnd || '') || undefined,
        outlookLocation: (draft.outlookLocation || '') || undefined,
        agenda: draftAgendaJson,
        transfers: draftTransfersJson,
        documents: '[]',
        funZone: '[]',
        quizClusterSize: 1,
        emailLanguage: subEmailLang,
        emailTemplateOverrides: subEmailOverrides,
        disableEmails: !!draft.disableEmails,
        // v19.22: granulare An-/Abmelde-Mail-Schalter pro Sub-Event persistieren.
        disableRegistrationEmail: !!draft.disableRegistrationEmail,
        disableCancellationEmail: !!draft.disableCancellationEmail,
        disableOutlook: !!draft.disableOutlook,
        isFictive: isFictive,
        askSalutation: !!draft.askSalutation,
        customFields: draft.customFields || [],
        parentEventId: parentEventId,
      };
      if (draft.dbId) {
        // v11.69: Recreate-Pfad via Modal-Auswahl. Wenn der Organizer im
        // Outlook-Confirm-Modal ein Sub-Event mit `noOutlookYet=true`
        // angehakt hat (es existiert noch kein Outlook-Termin), wird das
        // bestehende DEX_Events-Item per `deleteEventItemOnly` entfernt und
        // eine NEUE DEX_Events-Zeile angelegt — wobei die bestehende
        // Subsite + Teilnehmerliste an die neue Zeile gekoppelt werden.
        // Damit triggert der `DEX_CreateOutlookEvent`-Flow (GetOnNewItems)
        // auf dem neuen Item und legt den Outlook-Termin an. Die alte
        // Subsite mit ALLEN Anmeldungen bleibt unangetastet.
        const initialMeta = initialSubEventOutlookMeta[draft.dbId];
        if (pendingOutlookRecreateForSubEventsRef.current.includes(draft.dbId)) {
          const subsiteUrlForReuse = initialMeta?.subsiteUrl || '';
          const regListNameForReuse = initialMeta?.registrationListName || 'Teilnehmer';
          if (subsiteUrlForReuse && regListNameForReuse) {
            try {
              // DEX_Events-Item entfernen — Subsite + Teilnehmerliste
              // bleiben unangetastet (deleteEventItemOnly ist explizit
              // non-cascading).
              await deleteEventItemOnly(draft.dbId);
            } catch { /* Delete-Fehler: trotzdem versuchen, neu anzulegen */ }
            // Reuse-Payload: bestehende Subsite + Teilnehmerliste an die
            // neue DEX_Events-Zeile koppeln. disableOutlook explizit false,
            // outlookEventId leer, damit der Flow sauber neu schreibt.
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const reusePayload: any = {
              ...childPayload,
              disableOutlook: false,
              outlookEventId: '',
              existingSubsiteUrl: subsiteUrlForReuse,
              existingRegistrationListName: regListNameForReuse,
              onProgress: subOnProgress,
            };
            try {
              await createEvent(reusePayload);
            } catch (err) {
              console.warn('[DEX][v11.69] Sub-Event-Recreate mit Subsite-Reuse fehlgeschlagen:', draft.dbId, err);
            }
            // NICHT zu keptDbIds hinzufuegen — das alte Item wurde geloescht
            // und die neue Zeile hat eine andere ID.
            continue;
          } else {
            console.warn('[DEX][v11.69] Recreate angefordert aber keine subsiteUrl/registrationListName vorhanden — Sub-Event:', draft.dbId, 'meta:', initialMeta);
            // Fall durch zum normalen Update-Pfad — wenigstens die Felder
            // werden persistiert; Outlook-Termin entsteht aber nicht.
          }
        }
        // Spezialfall: Outlook nachträglich aktivieren via DisableOutlook-
        // Toggle (alter Pfad vor v11.69). Wenn der User auf einem
        // **bestehenden** Sub-Event die "Outlook erstellen"-Checkbox
        // einschaltet (DisableOutlook: true → false) und bisher kein
        // Outlook-Termin angelegt wurde (OutlookEventId leer), muss das
        // Sub-Event neu angelegt werden — der Power-Automate-Flow
        // `DEX_CreateOutlookEvent` triggert ausschließlich auf NEUE
        // DEX_Events-Items (GetOnNewItems). Ein reines MERGE-Update würde
        // den Flow nie anstoßen → kein Outlook-Termin.
        const wasOutlookDisabled = !!initialMeta?.disableOutlook;
        const nowOutlookEnabled = !draft.disableOutlook;
        const hadOutlookEventId = !!(initialMeta?.outlookEventId);
        const needsOutlookRecreate = wasOutlookDisabled && nowOutlookEnabled && !hadOutlookEventId;
        if (needsOutlookRecreate) {
          // v11.69: Seit dem Subsite-Reuse-Pfad muss hier KEINE destruktive
          // Loesch-Aktion mehr passieren. Wir entfernen nur die DEX_Events-
          // Zeile und legen sie mit `existingSubsiteUrl` neu an — alle
          // Anmeldungen, TeilnehmerIDs und die Subsite bleiben unangetastet.
          // Daher auch kein window.confirm mehr noetig.
          const subsiteUrlForReuse = initialMeta?.subsiteUrl || '';
          const regListNameForReuse = initialMeta?.registrationListName || 'Teilnehmer';
          if (subsiteUrlForReuse && regListNameForReuse) {
            try {
              await deleteEventItemOnly(draft.dbId);
            } catch { /* Delete-Fehler: trotzdem versuchen, neu anzulegen */ }
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const reusePayload: any = {
              ...childPayload,
              disableOutlook: false,
              outlookEventId: '',
              existingSubsiteUrl: subsiteUrlForReuse,
              existingRegistrationListName: regListNameForReuse,
              onProgress: subOnProgress,
            };
            try {
              await createEvent(reusePayload);
            } catch (err) {
              console.warn('[DEX][v11.69] Legacy-Toggle-Recreate fehlgeschlagen:', draft.dbId, err);
            }
            continue;
          } else {
            // Edge-Case: kein subsiteUrl auf dem alten Item bekannt (sehr
            // alte Events). In dem Fall fallen wir auf den destruktiven
            // Legacy-Pfad zurueck — mit Confirm.
            const msg = isDe
              ? `Beim Sub-Event „${draft.title}" wurde „Outlook-Termin erstellen" nachträglich aktiviert, aber es konnte keine bestehende Teilnehmer-Subsite ermittelt werden.\n\n`
                + `Wenn du jetzt fortfährst, wird das Sub-Event komplett neu aufgesetzt — vorhandene Anmeldungen, Teilnehmer-IDs und die Teilnehmer-Subsite gehen verloren (landen 93 Tage im Papierkorb).\n\n`
                + `Trotzdem fortfahren?`
              : `On sub-event "${draft.title}" you turned on "Create Outlook event" after the fact, but no existing participant subsite could be determined.\n\n`
                + `If you continue, the sub-event will be re-created from scratch — existing registrations, participant IDs and the participant subsite will be lost (recycled for 93 days).\n\n`
                + `Continue anyway?`;
            const confirmed = window.confirm(msg);
            if (confirmed) {
              try {
                await deleteEvent(draft.dbId);
              } catch { /* delete-Fehler darf Re-Create nicht blockieren */ }
              await createEvent({ ...childPayload, onProgress: subOnProgress });
              continue;
            } else {
              draft.disableOutlook = true;
              childPayload.disableOutlook = true;
            }
          }
        }
        keptDbIds.add(draft.dbId);
        // Update bestehender Sub-Event: nur geänderte Felder patchen. CustomFields
        // werden als JSON-String serialisiert — v17.22: zentraler
        // serializeCustomFields-Helper, damit der Sub-Event-Pfad dieselben
        // EN-Varianten + Options-Pairing erhaelt wie Top-Level (vorher droppte
        // dieser Pfad labelEn/helpTextEn/confirmLabelEn/optionsEn still).
        const cfJson = JSON.stringify(serializeCustomFields(draft.customFields || [], bilingualFields));
        // v11.57: Sub-Event-Kommunikations-Felder mit-persistieren — bisher
        // wurden Mail-Sprache, Outlook-Body, Logos pro Sub-Event nur am
        // Top-Level gespeichert. Mit den Tabs in Step 5 kann jeder Sub-Event
        // jetzt seine eigene Konfiguration haben.
        // v11.57: Outlook-Update-Flag pro Sub-Event setzen, wenn dieses
        // Sub-Event als „outlookDirty" markiert wurde (vom Confirm-Modal
        // entschieden).
        const subUpdates: Record<string, unknown> = {
          'Title': childPayload.title,
          'Description': childPayload.description,
          'Location': childPayload.location,
          'StartDate': childPayload.startDate || null,
          'EndDate': childPayload.endDate || null,
          'RegistrationDeadline': childPayload.registrationDeadline || null,
          'MaxParticipants': childPayload.maxParticipants,
          'DisableEmails': childPayload.disableEmails,
          'DisableOutlook': childPayload.disableOutlook,
          'OutlookSubject': subOutlookSubject,
          'OutlookStart': (draft.outlookStart || '') || null,
          'OutlookEnd': (draft.outlookEnd || '') || null,
          'OutlookLocation': (draft.outlookLocation || '') || '',
          'EmailLanguage': childPayload.emailLanguage,
          'OutlookBody': childPayload.outlookBody,
          'EmailTemplateOverrides': childPayload.emailTemplateOverrides,
          'EmailImageBase64': subEmailLogo || '',
          'CustomFields': cfJson,
        };
        // OutlookDirty + Update wird vom Aufrufer (handleSubmit) anhand des
        // jeweiligen Sub-Event-Snapshots gesteuert — siehe pendingSubUpdates.
        await updateEvent(draft.dbId, subUpdates);
      } else {
        await createEvent({ ...childPayload, onProgress: subOnProgress });
      }
    }
    // Entfernte Sub-Events aufräumen: deleteEvent löscht kaskadierend auch
    // die Subsite (Teilnehmerliste) und queued einen Outlook-DeleteEvent.
    for (const oldId of initialSubEventDbIds) {
      if (!keptDbIds.has(oldId)) {
        try { await deleteEvent(oldId); } catch { /* Delete-Fehler darf Save nicht blockieren */ }
      }
    }
  };

  /**
   * v11.57: Tab-Wechsel im Schritt 7 (Kommunikation). Der aktuelle Step-6-
   * UI-State (emailLanguage, outlookBody, disableEmails, disableOutlook,
   * Logo-Previews, Outlook-Heading) wird in das ausgehende Tab-Slot
   * geschrieben, danach werden die Felder aus dem neuen Tab-Slot geladen.
   *  - Slot 0 = Top-Level-Event-State (die normalen `emailLanguage`,
   *    `outlookBody` etc. — also der gleiche Speicherort wie heute).
   *  - Slot N>0 = subEvents[N-1] (die Felder aus SubEventDraft).
   * Die Step-5-UI bleibt unveraendert an die Top-Level-States gebunden — wir
   * spiegeln nur beim Tab-Wechsel hin und zurueck.
   */
  const switchCommTab = (nextIdx: number): void => {
    if (nextIdx === activeCommTabIdx) return;
    // 1) Aktuellen UI-State in das ausgehende Slot schreiben.
    if (activeCommTabIdx > 0) {
      const fromIdx = activeCommTabIdx - 1;
      // v11.60: synchron in den Ref schreiben (siehe flushActiveCommTabToState).
      const flushed = subEventsRef.current.map((s, i) => i === fromIdx ? {
        ...s,
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        disableOutlook,
        // v14.4: Mail-Text-Overrides pro Sub-Event mitspiegeln.
        emailTemplateOverrides: { ...emailTemplateOverrides },
      } : s);
      subEventsRef.current = flushed;
      setSubEvents(flushed);
    } else {
      // Slot 0 = Top-Level. Der UI-State wird hier direkt vom Top-Level-State
      // gehalten — kein Snapshot noetig, weil setEmailLanguage etc. den Wert
      // schon dort haelt. Beim Zurueck-Wechsel auf Tab 0 setzen wir die
      // Top-Level-States aus dem `topLevelCommSnapshot`-Ref (siehe unten).
      topLevelCommSnapshot.current = {
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        disableOutlook,
        emailTemplateOverrides: { ...emailTemplateOverrides },
      };
    }
    // 2) Werte aus dem Ziel-Slot in die Step-5-UI laden.
    if (nextIdx === 0) {
      const snap = topLevelCommSnapshot.current;
      if (snap) {
        setEmailLanguage(snap.emailLanguage);
        setEmailLogoPreview(snap.emailLogoBase64 || '');
        setOutlookLogoPreview(snap.outlookLogoBase64 || '');
        setOutlookBody(snap.outlookBody || '');
        setOutlookHeading(snap.outlookHeading || '');
        setOutlookSubheading(snap.outlookSubheading || '');
        setOutlookSubject(snap.outlookSubject || '');
        setDisableEmails(!!snap.disableEmails);
        setDisableRegistrationEmail(!!snap.disableRegistrationEmail);
        setDisableCancellationEmail(!!snap.disableCancellationEmail);
        setDisableOutlook(!!snap.disableOutlook);
        setEmailTemplateOverrides(snap.emailTemplateOverrides || {});
      }
    } else {
      const sub = subEvents[nextIdx - 1];
      if (sub) {
        setEmailLanguage(sub.emailLanguage || (locale === 'de' ? 'DE' : 'EN'));
        setEmailLogoPreview(sub.emailLogoBase64 || '');
        setOutlookLogoPreview(sub.outlookLogoBase64 || '');
        setOutlookBody(sub.outlookBody || '');
        setOutlookHeading(sub.outlookHeading || sub.title || '');
        setOutlookSubheading(sub.outlookSubheading || '');
        setOutlookSubject(sub.outlookSubject || '');
        setDisableEmails(!!sub.disableEmails);
        setDisableRegistrationEmail(!!sub.disableRegistrationEmail);
        setDisableCancellationEmail(!!sub.disableCancellationEmail);
        setDisableOutlook(!!sub.disableOutlook);
        setEmailTemplateOverrides(sub.emailTemplateOverrides || {});
      }
    }
    setActiveCommTabIdx(nextIdx);
  };
  // v11.57: Snapshot des Top-Level-Step-5-States. Wird beim Wechsel auf einen
  // Sub-Event-Tab gesetzt und beim Zurueckspringen wieder eingespielt.
  const topLevelCommSnapshot = React.useRef<{
    emailLanguage: string;
    emailLogoBase64: string;
    outlookLogoBase64: string;
    outlookBody: string;
    outlookHeading: string;
    outlookSubheading: string;
    outlookSubject: string;
    disableEmails: boolean;
    disableRegistrationEmail: boolean;
    disableCancellationEmail: boolean;
    disableOutlook: boolean;
    emailTemplateOverrides: Record<string, EmailOverrideEntry>;
  } | null>(null);
  // v11.57: Bevor wir submitten, muessen die Werte des aktuell sichtbaren
  // Tabs ins zugehoerige Slot zurueckgeschrieben werden — sonst gehen die
  // letzten Aenderungen verloren.
  const flushActiveCommTabToState = (): void => {
    if (activeCommTabIdx > 0) {
      const fromIdx = activeCommTabIdx - 1;
      // v11.60: synchron in den Ref schreiben — sonst sieht die direkt
      // anschliessende Detect-/Persist-Logik noch die alte Array.
      const flushed = subEventsRef.current.map((s, i) => i === fromIdx ? {
        ...s,
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        disableOutlook,
        emailTemplateOverrides: { ...emailTemplateOverrides },
      } : s);
      subEventsRef.current = flushed;
      setSubEvents(flushed);
    }
    // Slot 0 (Top-Level) wird ohnehin direkt von den State-Variablen gespeist
    // — kein Snapshot-Flush noetig (resolveTopLevelCommState liest auf Tab 0
    // direkt aus dem State, der Snapshot wird nur für Sub-Tab-Pfade benutzt).
  };

  /**
   * v11.93: Liefert die Top-Level-Kommunikations-Werte zuverlässig — egal auf
   * welchem Tab der User gerade in Schritt 6 steht. Bug-Hintergrund: Die
   * Logo-/Body-/Heading-States werden zwischen Top-Level und Sub-Event-Tabs
   * hin- und hergespiegelt; wenn der User auf einem Sub-Tab Speichern klickt,
   * stehen in den State-Variablen die Sub-Event-Werte — der Top-Level-Save
   * würde diese fälschlich aufs Haupt-Event schreiben (Logo, Outlook-Body
   * etc.). Diese Helper-Funktion entscheidet auf Basis von activeCommTabIdx,
   * ob die aktuellen State-Variablen schon Top-Level sind (Tab 0) oder ob aus
   * dem topLevelCommSnapshot resolved werden muss.
   */
  const resolveTopLevelCommState = (): {
    emailLanguage: string;
    emailLogoBase64: string;
    outlookLogoBase64: string;
    outlookBody: string;
    outlookHeading: string;
    outlookSubheading: string;
    outlookSubject: string;
    disableEmails: boolean;
    disableRegistrationEmail: boolean;
    disableCancellationEmail: boolean;
    disableOutlook: boolean;
    emailTemplateOverrides: Record<string, EmailOverrideEntry>;
  } => {
    if (activeCommTabIdx === 0) {
      return {
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        disableOutlook,
        emailTemplateOverrides,
      };
    }
    const snap = topLevelCommSnapshot.current;
    if (snap) return snap;
    // Fallback (sollte praktisch nicht eintreten): wir sind auf einem Sub-Tab,
    // hatten aber noch keinen Snapshot — verwenden die aktuellen State-Werte,
    // damit zumindest kein Crash entsteht.
    return {
      emailLanguage,
      emailLogoBase64: emailLogoPreview,
      outlookLogoBase64: outlookLogoPreview,
      outlookBody,
      outlookHeading,
      outlookSubheading,
      outlookSubject,
      disableEmails,
      disableRegistrationEmail,
      disableCancellationEmail,
      disableOutlook,
      emailTemplateOverrides,
    };
  };

  /**
   * Save-Side-Sanity: Organizer-Names und -Emails 1:1 paaren bevor sie nach SP
   * geschrieben werden. Pairs ohne BEIDE (Name + Email) fallen raus — verhindert
   * dass eine Mismatch-State (z.B. „Spiegel, Mirjam" gepaart mit
   * „egenctuerk@deloitte.de") in DEX_Events landet. Bisher wurden organizer und
   * organizerEmails unabhängig serialisiert, dadurch konnten Drift-States aus
   * Closure-Bugs / Edit-Pfaden / Move-Bugs in die Persistenz durchschlagen.
   *
   * Returnt sauber serialisierte Strings (`Organizer` mit '; '-Trenner,
   * `OrganizerEmail` mit ';'-Trenner) — exakt das Format das DEX_Events erwartet
   * und der OutlookEventCreate-Flow + DEX_SEND_MAIL als Recipient-Liste lesen.
   */
  const sanitizeOrganizerPairs = React.useCallback((): { orgString: string; orgEmailString: string; droppedCount: number } => {
    const names = (organizer || '').split(';').map(s => s.trim());
    const emails = (organizerEmails || []).map(e => (e || '').trim());
    const max = Math.max(names.length, emails.length);
    const pairs: Array<{ n: string; e: string }> = [];
    let dropped = 0;
    for (let i = 0; i < max; i++) {
      const n = (names[i] || '').trim();
      const e = (emails[i] || '').trim();
      if (n && e) pairs.push({ n, e });
      else if (n || e) dropped++;
    }
    return {
      orgString: pairs.map(p => p.n).join('; '),
      orgEmailString: pairs.map(p => p.e).join(';'),
      droppedCount: dropped,
    };
  }, [organizer, organizerEmails]);

  const handleSubmit = async (): Promise<void> => {
    // v9.14: Beschreibung ist jetzt optional. Nur Title bleibt Pflicht.
    if (!title) return;

    // v18.36: Harte Datums-Validierung als letzter Riegel — das Enddatum darf
    // NIE vor (oder gleich) dem Startdatum liegen. Outlook lehnt solche Termine
    // ab und der DEX_CreateOutlookEvent-Flow failt dann mit HTTP 400
    // („At least one property failed validation"). Gilt fuer das Hauptevent UND
    // jedes Sub-Event — Sub-Events liefen bisher ohne Datums-Pruefung durch.
    const dateProblems: string[] = [];
    if (startDate && endDate && new Date(endDate) <= new Date(startDate)) {
      dateProblems.push(isDe ? 'Hauptevent' : 'Main event');
    }
    subEventsRef.current.forEach(s => {
      if (s.title && s.title.trim() && s.startDate && s.endDate && new Date(s.endDate) <= new Date(s.startDate)) {
        dateProblems.push(isDe ? `Sub-Event „${s.title}"` : `Sub-event „${s.title}"`);
      }
    });
    if (dateProblems.length > 0) {
      // eslint-disable-next-line no-alert
      alert(isDe
        ? `Das Enddatum darf nicht vor dem Startdatum liegen. Bitte korrigiere das Datum bei: ${dateProblems.join(', ')}.`
        : `The end date must not be before the start date. Please fix the date for: ${dateProblems.join(', ')}.`);
      return;
    }

    // v11.93: Top-Level-Kommunikations-Werte sauber resolven (s. Helper-
    // Doku oben). Sonst würden, falls beim Speichern ein Sub-Event-Tab
    // aktiv ist, die Sub-Event-States (Logo, Outlook-Body, Headings,
    // etc.) fälschlich auf das Haupt-Event geschrieben.
    const topComm = resolveTopLevelCommState();
    const effEmailLanguage = topComm.emailLanguage;
    const effEmailLogo = topComm.emailLogoBase64;
    const effOutlookLogo = topComm.outlookLogoBase64;
    const effOutlookBody = topComm.outlookBody;
    const effOutlookHeading = topComm.outlookHeading;
    const effOutlookSubheading = topComm.outlookSubheading;
    const effOutlookSubject = topComm.outlookSubject;
    const effDisableEmails = topComm.disableEmails;
    // v19.22: granulare An-/Abmelde-Mail-Schalter des Hauptevents top-level
    // auflösen (auf Sub-Tabs hält der State den Sub-Wert → resolveTopLevelCommState).
    const effDisableRegistrationEmail = topComm.disableRegistrationEmail;
    const effDisableCancellationEmail = topComm.disableCancellationEmail;
    const effDisableOutlook = topComm.disableOutlook;

    // v14.4 / v14.5: Wenn das Hauptevent Sub-Events hat UND die
    // Kommunikation auf Top-Level abgestellt ist, muss entweder der
    // „Sub-Event verpflichtend"-Toggle aktiv sein (erzwingt es im
    // Anmeldeformular) ODER der Organizer den Ack-Haken gesetzt haben.
    // Sonst landen Teilnehmer ohne Bestätigungs-Mail und ohne Kalender-
    // Termin in der Liste.
    const hasSubs = subEventsRef.current.some(s => s.title && s.title.trim());
    if (hasSubs && (effDisableEmails || effDisableOutlook) && !requireSubEventSelection && !mainCommDisabledAck) {
      // eslint-disable-next-line no-alert
      alert(isDe
        ? 'Du hast die Kommunikation für das Hauptevent deaktiviert. Bitte aktiviere in Schritt 6 (Kommunikation, Tab „Haupt-Event") entweder den Toggle „Anmeldung für mindestens ein Sub-Event verpflichtend" ODER bestätige den Ack-Haken — sonst landen Teilnehmer stumm in der Liste.'
        : 'You disabled communication for the main event. Please either enable the toggle „Require selecting at least one sub-event" in step 6 OR tick the acknowledgement — otherwise attendees land silently in the list.');
      return;
    }

    setIsSubmitting(true);
    setError('');
    setProgress(0);

    // Schritt 1: Bild wird spaeter (nach Event-Erstellung) als Item-Attachment hochgeladen.
    // Bestehende URL beibehalten (z.B. bei Edit ohne neues Bild).
    setProgress(5);
    setProgressLabel('Event wird vorbereitet...');
    const imageUrl = eventImageUrl;
    setProgress(15);

    if (isEditMode && selectedEventId) {
      setProgressLabel('Event wird aktualisiert...');
      // v11.18: Debug-Trace JSON-stringified, sodass der Output ohne
      // Array-Aufklappen direkt als Text in der Console steht.
      // eslint-disable-next-line no-console
      console.log('[DEX][edit-save] customFields state at save:\n' + JSON.stringify(
        customFields.map(f => ({
          id: f.id,
          label: f.label,
          type: f.type,
          required: f.required,
          helpText: f.helpText,
          helpTextStyle: f.helpTextStyle,
          onlyForGroup: f.onlyForGroup,
          showIf: f.showIf,
          externalLinks: f.externalLinks,
        })),
        null,
        2
      ));
      // Sanitize: paart Organizer-Names + -Emails 1:1, droppt unvollständige
      // Pairs — verhindert Mismatch-State in DEX_Events.
      const sanitizedOrgPairEdit = sanitizeOrganizerPairs();
      // Event aktualisieren - nur bekannte Felder senden
      const updates: Record<string, unknown> = {
        'Title': title,
        'Description': description,
        'Location': location,
        'LocationAddress': (addrStreet || addrHouseNo || addrZip || addrCity)
          ? JSON.stringify({ street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity })
          : '',
        // v18.34/v18.40: Outlook-Ort = manuelle Überschreibung, sonst Auto aus
        // Veranstaltungsort + Adresse. Flow mappt OutlookLocation 1:1.
        'OutlookLocation': outlookLocationOverride.trim() || buildOutlookLocation(location, { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity }),
        'LocationFilter': locationFilter,
        'Audience': audience,
        // v16.4: Audience-DLs vor-aufgeloest mitschreiben.
        'AudienceResolvedEmails': await resolveAudienceMembersToCsv(audience, getGroupMembers),
        'FilterMode': filterMode,
        'StartDate': startDate ? berlinLocalToUtcIso(startDate) : null,
        'EndDate': endDate ? berlinLocalToUtcIso(endDate) : null,
        'RegistrationDeadline': deadlineToEndOfDayIso(registrationDeadline),
        'MaxParticipants': unlimitedParticipants ? 0 : (Number(maxParticipants) || 0),
        'EventImageUrl': imageUrl,
        'Organizer': sanitizedOrgPairEdit.orgString,
        'OrganizerEmail': sanitizedOrgPairEdit.orgEmailString,
        // v10.16: optionaler Ansprechpartner (Anzeige-Feld, kein Login)
        'ContactName': contactName.trim(),
        'ContactEmail': contactEmail.trim(),
        'ContactInfo': contactInfo.trim(),
        // v17.22: zentraler serializeCustomFields-Helper (Options-Pairing +
        // EN-Varianten konsistent zu allen Pfaden).
        'CustomFields': JSON.stringify(serializeCustomFields(customFields, bilingualFields)),
      };

      // Optionale Felder - immer senden damit Loeschungen wirken
      updates['LastDeregisterDate'] = deadlineToEndOfDayIso(lastDeregisterDate);
      // Outlook-Body: Variablen werden bereits hier aufgeloest (gleicher Body fuer alle Teilnehmer).
      const outlookVars: Record<string, string> = {
        EventTitle: title,
        Organizer: organizer,
        Location: location,
        Address: [addrStreet, addrHouseNo].filter(Boolean).join(' ') + ((addrZip || addrCity) ? ', ' + [addrZip, addrCity].filter(Boolean).join(' ') : ''),
        StartDate: startDate ? new Date(startDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
        EndDate: endDate ? new Date(endDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
      };
      // v7.4: Auch wenn der User keinen Outlook-Body eingegeben hat, IMMER
      // das Outlook-Mail-Layout (buildOutlookBody) verwenden + einen
      // Default-Body einsetzen, der den Empfänger an die Organizer
      // verweist. Sonst kommt der Termin ganz ohne Body — wirkt
      // unprofessionell und der Teilnehmer hat keinen Ansprechpartner
      // bei organisatorischen Rückfragen.
      const orgNames = organizer.split(';').map(s => s.trim()).filter(Boolean).join(', ');
      const escHtml = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      // v9.8: Default-Body enthaelt jetzt auch den Abmelde-Hinweis analog zur
      // Anmeldebestaetigungs-Mail. Sonst weiss der Empfaenger nicht, wie er
      // sich abmelden kann — die Outlook-Decline-Funktion triggert zwar einen
      // Reminder-Flow, aber der eigentliche App-Abmelde-Pfad ist sauberer.
      const APP_URL_OL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView';
      const defaultOutlookBody = effEmailLanguage === 'EN'
        ? `<p>You are registered for the event <strong>${escHtml(title)}</strong>.</p>`
          + `<p>If you are unable to attend, please cancel your registration in time via the <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;My Events&ldquo;).</p>`
          + `<p>For organizational questions please contact <strong>${escHtml(orgNames || 'the organizer')}</strong>.</p>`
        : `<p>Du bist für das Event <strong>${escHtml(title)}</strong> angemeldet.</p>`
          + `<p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;Meine Events&ldquo;) ab.</p>`
          + `<p>Bei organisatorischen Fragen wende dich bitte an <strong>${escHtml(orgNames || 'den Organizer')}</strong>.</p>`;
      const resolvedBody = effOutlookBody
        ? replacePlaceholders(effOutlookBody, outlookVars)
        : defaultOutlookBody;
      const resolvedOlHeading = effOutlookHeading ? replacePlaceholders(effOutlookHeading, outlookVars) : title;
      const resolvedOlSub = effOutlookSubheading ? replacePlaceholders(effOutlookSubheading, outlookVars) : undefined;
      // v18.73: Header-Bild Größe + Innenabstand (event-weit) in den Outlook-Body.
      const wrappedOutlook = buildOutlookBody(resolvedOlHeading, resolvedBody, resolvedOlSub, { imageWidth: headerImageLayout.width, imagePaddingV: headerImageLayout.paddingV, imagePaddingH: headerImageLayout.paddingH });
      // v11.93: Top-Level-Logo aus dem Resolver — sonst würde beim Speichern
      // aus einem Sub-Tab das falsche Logo aufs Haupt-Event geschrieben.
      updates['OutlookBody'] = wrappedOutlook.replace(/\{\{ORB_URL\}\}/g, effOutlookLogo || getCachedOrbBase64() || '');
      // v18.42: Betreff des Outlook-Termins mit-persistieren (leer = Titel via Flow-Fallback).
      updates['OutlookSubject'] = effOutlookSubject.trim();
      // v18.44: abweichendes Outlook-Datum mit-persistieren (leer = Event-Datum via Flow-Fallback).
      updates['OutlookStart'] = outlookStartOverride || null;
      updates['OutlookEnd'] = outlookEndOverride || null;
      updates['Agenda'] = JSON.stringify(agenda);
      updates['Transfers'] = JSON.stringify(transferTimes);
      updates['FunZone'] = JSON.stringify(quiz);
      updates['QuizClusterSize'] = Math.min(Math.max(1, quizClusterSize || 1), 4);
      updates['EmailLanguage'] = effEmailLanguage;
      // v18.35: erzwungene Anmeldeseiten-Sprache mit-persistieren ('' = App-Sprache).
      updates['RegistrationLanguage'] = registrationLanguage || '';
      // v6.15: B2Run-Config (Starter-Typ → Startblock, Leistungsnachweis-Pflicht)
      // wird in EmailTemplateOverrides._b2run piggyback gespeichert, damit keine
      // neue SP-Spalte nötig ist.
      const b2runExtraConfig = (durchstarterStartblock || funstarterStartblock || durchstarterRequiresProof)
        ? { _b2run: {
            ...(durchstarterStartblock ? { durchstarterStartblock } : {}),
            ...(funstarterStartblock ? { funstarterStartblock } : {}),
            ...(durchstarterRequiresProof ? { durchstarterRequiresProof: true } : {}),
          } }
        : {};
      // v6.19: QR-Code-Scanner piggyback in EmailTemplateOverrides._qrScanners
      const qrScannerConfig = qrScannerEmails.length > 0
        ? { _qrScanners: qrScannerNames.map((n, i) => ({ name: n, email: qrScannerEmails[i] || '' })).filter(x => x.email) }
        : {};
      // v9.18: Co-Organizer-Liste piggyback in EmailTemplateOverrides._coOrganizers
      const coOrganizerConfig = coOrganizerEmails.length > 0
        ? { _coOrganizers: coOrganizerNames.map((n, i) => ({ name: n, email: coOrganizerEmails[i] || '' })).filter(x => x.email) }
        : {};
      // v9.21: Test-Team-Liste piggyback in EmailTemplateOverrides._testTeam
      const testTeamConfig = testTeamEmails.length > 0
        ? { _testTeam: testTeamNames.map((n, i) => ({ name: n, email: testTeamEmails[i] || '' })).filter(x => x.email) }
        : {};
      // v11.25: Display-Reihenfolge-Toggle als piggyback in
      // EmailTemplateOverrides._splitDisplayOrderReversed.
      const splitDispRevConfig = splitDisplayOrderReversed && useSplitCapacities
        ? { _splitDisplayOrderReversed: true }
        : {};
      // v11.93: Top-Level-Logos aus dem Resolver lesen, NICHT direkt aus
      // den State-Variablen — sonst wird beim Speichern aus einem Sub-Tab
      // das Sub-Logo aufs Haupt-Event geschrieben.
      // v14.4: Mail-Text-Overrides ebenfalls aus dem Resolver — vorher wurden
      // beim Speichern auf einem Sub-Tab die Sub-Overrides fälschlich aufs
      // Hauptevent gemerged.
      const topOverrides = topComm.emailTemplateOverrides || {};
      // v14.8: subEventsOnlyMode impliziert requireSubEventSelection — wenn die
      // Hauptevent-Anmeldung gar nicht mehr angeboten wird, MUSS jeder Teilnehmer
      // mindestens einen Sub-Event auswählen, sonst kommt keine Anmeldung zustande.
      const effRequireSubEventSelection = requireSubEventSelection || subEventsOnlyMode;
      const requireSubEventConfig = effRequireSubEventSelection
        ? { _requireSubEventSelection: true }
        : {};
      // v14.8: Sub-Events-Only-Modus + Custom-Bezeichnung als Piggyback.
      const subEventsOnlyConfig = subEventsOnlyMode
        ? { _subEventsOnlyMode: true }
        : {};
      const childTermConfig = (childTermSingular.trim() || childTermPlural.trim())
        ? { _childEventTerm: { singular: childTermSingular.trim(), plural: childTermPlural.trim() } }
        : {};
      // v18.9: Organizer-Anzeige ausblenden (Piggyback).
      const hideOrganizerConfig = hideOrganizer ? { _hideOrganizer: true } : {};
      updates['EmailTemplateOverrides'] = (Object.keys(topOverrides).length > 0 || effEmailLogo || effOutlookLogo || Object.keys(b2runExtraConfig).length > 0 || Object.keys(qrScannerConfig).length > 0 || Object.keys(coOrganizerConfig).length > 0 || Object.keys(testTeamConfig).length > 0 || Object.keys(splitDispRevConfig).length > 0 || Object.keys(requireSubEventConfig).length > 0 || Object.keys(subEventsOnlyConfig).length > 0 || Object.keys(childTermConfig).length > 0 || Object.keys(hideOrganizerConfig).length > 0 || Object.keys(headerImageLayoutConfig).length > 0)
        ? JSON.stringify({
            ...(effEmailLogo ? { _eventLogo: effEmailLogo } : {}),
            ...(effOutlookLogo ? { _outlookLogo: effOutlookLogo } : {}),
            ...b2runExtraConfig,
            ...qrScannerConfig,
            ...coOrganizerConfig,
            ...testTeamConfig,
            ...splitDispRevConfig,
            ...requireSubEventConfig,
            ...subEventsOnlyConfig,
            ...childTermConfig,
            ...hideOrganizerConfig,
            // v18.73: Header-Bild-Layout (Breite + Innenabstand) — event-weit.
            ...headerImageLayoutConfig,
            ...topOverrides,
          })
        : '';
      // v9.21: ActiveFrom als SP-DateTime
      updates['ActiveFrom'] = activeFrom ? new Date(activeFrom).toISOString() : null;
      // Custom-Mail-Logo in EmailImageBase64 (SP-Spalte) — der Flow ersetzt
      // {{ORB_URL}} in Mails damit. Wenn leer: Flow faellt auf _Config
      // DefaultImageBase64 (DEX-Orb) zurueck.
      updates['EmailImageBase64'] = effEmailLogo || '';
      updates['DisableEmails'] = effDisableEmails;
      // v19.22: granulare An-/Abmelde-Mail-Schalter des Hauptevents (top-level
      // aufgelöst, damit ein Save von einem Sub-Tab nicht den Sub-Wert aufs
      // Hauptevent schreibt). Pro Sub-Event werden sie in persistSubEventsForParent
      // geschrieben.
      updates['DisableRegistrationEmail'] = effDisableRegistrationEmail;
      updates['DisableCancellationEmail'] = effDisableCancellationEmail;
      updates['DisableOutlook'] = effDisableOutlook;
      // v11.57: OutlookDirty schreiben. Wenn Outlook-relevante Aenderungen
      // anstehen und der Organizer im Update-Confirm-Modal die Checkbox
      // *nicht* gesetzt hat, bleibt der Flag dirty=true; bei Checkbox=true
      // (= UpdateEvent wird gequeued) wird dirty wieder auf false gesetzt.
      // Wenn keine Outlook-relevante Aenderung vorlag (z.B. nur Beschreibung
      // angepasst), wird dirty NICHT angefasst — der Wert bleibt wie er war.
      // Den eigentlichen Wert setzen wir aus dem Modal-State unten (siehe
      // pendingOutlookDirtyWrite).
      if (pendingOutlookDirtyWriteRef.current !== null) {
        updates['OutlookDirty'] = pendingOutlookDirtyWriteRef.current;
      }
      updates['NotifyOrgRegisterMode'] = notifyOrgRegisterMode === 'always' ? 'Always' : notifyOrgRegisterMode === 'fromDate' ? 'FromDate' : 'Never';
      updates['NotifyOrgRegisterFromDate'] = notifyOrgRegisterMode === 'fromDate' && notifyOrgRegisterFromDate ? berlinLocalToUtcIso(notifyOrgRegisterFromDate) : null;
      updates['NotifyOrgCancelMode'] = notifyOrgCancelMode === 'always' ? 'Always' : notifyOrgCancelMode === 'afterDeadline' ? 'AfterDeadline' : 'Never';
      updates['ExcludedUsers'] = excludedUsers.filter(Boolean).join(';');
      updates['IsFictive'] = isFictive;
      if (useSplitCapacities) {
        updates['DurchstarterCapacity'] = parseInt(durchstarterCapacity, 10) || 0;
        updates['FunstarterCapacity'] = parseInt(funstarterCapacity, 10) || 0;
        // v10.20: frei waehlbare Bezeichnungen mitschreiben — leer = Default-
        // Fallback in der Registration-UI ('Durchstarter' / 'Funstarter').
        updates['SplitLabelA'] = (splitLabelA || '').trim();
        updates['SplitLabelB'] = (splitLabelB || '').trim();
        updates['SplitSharedWaitlist'] = !!splitSharedWaitlist;
      } else {
        // Split deaktiviert: Kapazitäten nullen + Labels leer setzen, damit
        // die Registration-Logik nicht irrtümlich den Split-Pfad nimmt.
        updates['DurchstarterCapacity'] = null;
        updates['FunstarterCapacity'] = null;
        updates['SplitLabelA'] = '';
        updates['SplitLabelB'] = '';
        updates['SplitSharedWaitlist'] = false;
      }
      // v11.0: Teilnehmer-Upload-Setting
      updates['AllowAttendeeUpload'] = !!allowAttendeeUpload;
      updates['AttendeeUploadHint'] = (attendeeUploadHint || '').trim();
      updates['AttendeeUploadLabel'] = (attendeeUploadLabel || '').trim();
      // v11.80: Anrede-Toggle + Team-Anmeldung-Konfiguration mit-persistieren.
      updates['AskSalutation'] = !!askSalutation;
      // v18.75: Sicherheitshinweis vor dem Absenden mit-persistieren.
      updates['ConfirmDialogEnabled'] = !!confirmDialogEnabled;
      updates['ConfirmDialogMode'] = confirmDialogEnabled ? (confirmDialogMode || 'summary') : '';
      updates['ConfirmDialogText'] = confirmDialogEnabled && confirmDialogMode === 'freetext' ? confirmDialogText : '';
      // v18.33: Self-Check-in mit-persistieren. Token nur schreiben, wenn aktiv.
      updates['SelfCheckInEnabled'] = !!selfCheckInEnabled;
      updates['SelfCheckInToken'] = selfCheckInEnabled ? (selfCheckInToken || '') : '';
      updates['SelfCheckInFrom'] = selfCheckInEnabled && selfCheckInFrom ? selfCheckInFrom : null;
      updates['SelfCheckInTo'] = selfCheckInEnabled && selfCheckInTo ? selfCheckInTo : null;
      updates['TeamRegistrationEnabled'] = !!teamRegistrationEnabled;
      updates['TeamSize'] = teamRegistrationEnabled && teamSize > 0 ? teamSize : null;
      updates['AskTeamName'] = !!askTeamName;
      // v11.81: Erweiterte Team-Konfiguration mit-persistieren.
      updates['TeamPartialAllowed'] = !!(teamRegistrationEnabled && teamPartialAllowed);
      updates['TeamOpenSlotsVisible'] = !!(teamRegistrationEnabled && teamOpenSlotsVisible);
      updates['TeamJoinRequiresApproval'] = !!(teamRegistrationEnabled && teamOpenSlotsVisible && teamJoinRequiresApproval);
      // v17.20: Bilingual-Toggle persistieren.
      updates['BilingualFields'] = !!bilingualFields;

      // v11.22: feinere Progress-Stufen waehrend Edit-Save. Vorher
      // sprang es bei 50% sehr lange auf der Stelle, weil zwischen
      // setProgress(50) und setProgress(100) die Dokument-Sync,
      // updateEvent, Berechtigungs-Sync, Sub-Event-Persistierung,
      // Teilnehmer-Spalten-Sync, Bild-Upload und Outlook-Update
      // nacheinander liefen — alles ohne Zwischen-Tick.
      setProgress(40);
      setProgressLabel(isDe ? 'Dokumente werden synchronisiert...' : 'Syncing documents...');

      // Dokument-Sync: entfernte Attachments loeschen + neue hochladen.
      // Wichtig: erst loeschen, dann uploaden (SharePoint verbietet Duplikat-Namen).
      if (selectedEventId) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ctx = (window as any).__dexSpfxContext;
        if (ctx) {
          const svc = new EventService(ctx);
          // Namen die weiterhin vorhanden sind (als bestehende Attachments, ohne file)
          const keptOriginalNames = new Set(
            documents.filter(d => !d.file).map(d => d.name)
          );
          // Namen die im initialen Snapshot waren, jetzt aber nicht mehr
          const toDelete = initialDocumentNames.filter(name => !keptOriginalNames.has(name));
          for (const fileName of toDelete) {
            try {
              await svc.deleteEventDocument(Number(selectedEventId), fileName);
            } catch { /* einzelner Delete-Fehler darf Save nicht blockieren */ }
          }
          // Neue Dokumente hochladen
          for (const doc of documents) {
            if (doc.file) {
              try {
                await svc.uploadEventDocument(Number(selectedEventId), doc.file);
              } catch { /* einzelner Upload-Fehler darf Save nicht blockieren */ }
            }
          }
        }
      }

      setProgress(55);
      setProgressLabel(isDe ? 'Event-Daten werden gespeichert...' : 'Saving event data...');

      // v11.20: Direkt vor dem updateEvent-Call loggen was am SP-Server
      // landet. Damit sehen wir im Browser-DevTools:
      //   1. ob updates['CustomFields'] als JSON-String den helpText
      //      enthaelt (= Save sendet's korrekt → SP-Persist OK).
      //   2. oder ob updates['CustomFields'] ohne helpText/onlyForGroup
      //      ankommt (= State zum Save-Zeitpunkt war schon kaputt).
      // eslint-disable-next-line no-console
      console.log('[DEX][edit-save] updates.CustomFields about to POST:', updates['CustomFields']);
      const success = await updateEvent(selectedEventId, updates);
      if (success) {
        setProgress(65);
        setProgressLabel(isDe ? 'Berechtigungen werden gesetzt...' : 'Setting permissions...');
        // v9.35: Berechtigungs-Sync — beim Edit können neue Co-Organizer hinzugekommen
        // sein, die bisher nur in EmailTemplateOverrides._coOrganizers stehen, aber
        // noch keine SharePoint-Berechtigung auf Subsite + Teilnehmerliste haben.
        // ensureOrganizerPermissions ist idempotent: bestehende Rechte werden nicht
        // doppelt vergeben, neue kommen sauber dazu.
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const ctxPerm = (window as any).__dexSpfxContext;
          if (ctxPerm && editEvent?.subsiteUrl) {
            const svcPerm = new EventService(ctxPerm);
            const allOrgEmailsForPerm = [
              organizerEmails.join(';'),
              coOrganizerEmails.join(';'),
            ].filter(Boolean).join(';');
            if (allOrgEmailsForPerm) {
              await svcPerm.ensureOrganizerPermissions(editEvent.subsiteUrl, allOrgEmailsForPerm);
            }
          }
        } catch (err) { console.warn('[DEX] Permission-Sync für Organizer fehlgeschlagen:', err); }

        setProgress(75);
        setProgressLabel(isDe ? 'Sub-Events werden gespeichert...' : 'Saving sub-events...');
        // Sub-Events persistieren (create/update/delete pro Draft). Seit v6.4.
        try { await persistSubEventsForParent(selectedEventId); }
        catch (err) { console.warn('[DEX] Sub-Events persistieren fehlgeschlagen:', err); }

        setProgress(82);
        setProgressLabel(isDe ? 'Teilnehmerlisten-Spalten werden geprüft...' : 'Verifying participant list columns...');
        // Custom-Fields-Columns auf der Teilnehmerliste auto-sync: falls
        // neue Custom-Fields ohne spInternalName hinzugekommen sind oder
        // SP-Spalten fehlen, jetzt anlegen + spInternalName ins Event
        // zurueckschreiben.
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const ctx = (window as any).__dexSpfxContext;
          if (ctx && editEvent?.subsiteUrl) {
            const svc = new EventService(ctx);
            const cfForFix = customFields
              .filter(f => f.label && f.label.trim().length > 0)
              .map(f => ({
                id: f.id, label: f.label.trim(), type: f.type, required: f.required, visible: f.visible,
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                spInternalName: (f as any).spInternalName || '',
                ...(f.type === 'select' ? { options: f.options.map(o => o.trim()).filter(Boolean), ...(f.multi ? { multi: true } : {}) } : {}),
                ...(f.onlyForGroup && f.onlyForGroup !== 'all' ? { onlyForGroup: f.onlyForGroup } : {}),
                // v11.15: externalLinks (AGB/Datenschutz-URLs etc.) beim Save
                // mit-persistieren — vorher haben alle drei Persist-Pfade
                // (Edit-Save, Create-Save, Sub-Event-Save) sie gedroppt.
                ...(f.externalLinks && f.externalLinks.length > 0
                  ? { externalLinks: f.externalLinks.map(x => ({ label: x.label, url: x.url })) }
                  : {}),
                // v11.21 KRITISCHER BUG-FIX: helpText UND showIf wurden hier
                // beim cfForFix-Mapping nicht mit-uebernommen — der zweite
                // updateEvent-Call (siehe `merged`-JSON unten, ueberschreibt
                // CustomFields nochmal mit der spInternalName-Anreicherung)
                // hat die helpText- und showIf-Properties wieder vom SP-
                // Item entfernt. Folge: jede gespeicherte Beschreibung war
                // direkt nach dem Save wieder weg, obwohl der erste
                // updateEvent sie korrekt geschrieben hatte.
                ...(f.helpText && f.helpText.trim() ? { helpText: f.helpText.trim() } : {}),
                // v18.20 BUG-FIX (gleiches Muster wie v11.21): helpTextStyle
                // wurde hier nicht mit-uebernommen — der zweite updateEvent-Call
                // (merged-JSON unten) hat die Property wieder vom SP-Item
                // entfernt. Folge: „Text unter dem Feld-Titel" war direkt nach
                // dem Speichern wieder weg und das Feld zeigte die „i"-Box.
                ...(f.helpTextStyle === 'inline' ? { helpTextStyle: 'inline' as const } : {}),
                ...(f.showIf && f.showIf.fieldId && f.showIf.values && f.showIf.values.length > 0
                  ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } }
                  : {}),
                // v11.94: confirmLabel auch im cfForFix-Mapping mit-übernehmen,
                // sonst überschreibt der zweite updateEvent-Call die Property weg.
                ...(f.type === 'checkbox' && f.confirmLabel && f.confirmLabel.trim()
                  ? { confirmLabel: f.confirmLabel.trim() }
                  : {}),
                // v18.41 (gleiches Muster wie v11.21/v18.20): ccOnEmails muss
                // auch im zweiten Write mit, sonst droppt der spInternalName-
                // Merge die Property direkt nach dem Speichern wieder.
                ...((f.type === 'user' || f.type === 'roommate') && f.ccOnEmails ? { ccOnEmails: true } : {}),
              }));
            // v11.6 BUG-FIX: vorher wurde hier `isB2runTemplate` (= b2run_*-
            // Custom-Fields vorhanden) als Indikator genutzt. Das war falsch,
            // sobald die generische Split-Capacity ohne B2Run-Template
            // genutzt wird — dann hat das Event Split-Kapazitäten + StarterType-
            // Werte in der Teilnehmerliste, aber `isB2runTemplate=false`. Der
            // Fix-Lauf hat daraufhin StarterType + PreferredStarterType
            // gelöscht und die Teilnehmer-Daten weggeworfen. Korrekter Check:
            // entweder altes B2Run-Template ODER Split-Capacity aktiv.
            const splitActive = useSplitCapacities && ((parseInt(durchstarterCapacity, 10) || 0) > 0 || (parseInt(funstarterCapacity, 10) || 0) > 0);
            const fixResult = await svc.fixRegistrationListColumns(editEvent.subsiteUrl, {
              isB2Run: isB2runTemplate || splitActive,
              hasQuiz: quiz.length > 0,
              customFields: cfForFix,
            });
            if (fixResult.customFieldMap && Object.keys(fixResult.customFieldMap).length > 0) {
              // v19.20: ROBUSTER FIX für den wiederkehrenden „zweiter
              // CustomFields-Write droppt Properties"-Bug (siehe CLAUDE.md).
              // Früher wurde dieser zweite Write aus dem manuell gepflegten
              // cfForFix-Mapping gebaut — jede dort vergessene Property wurde
              // damit direkt nach dem Speichern wieder vom SP-Item entfernt
              // (zuletzt die EN-Varianten labelEn/helpTextEn/confirmLabelEn/
              // optionsEn; historisch multi/ccOnEmails/helpText/showIf/…).
              // Jetzt nehmen wir den KANONISCHEN serializeCustomFields-Output
              // (der ALLE Properties korrekt persistiert) und ergänzen pro
              // Feld nur noch spInternalName. So kann die Property-Liste nie
              // wieder veralten — cfForFix dient ab jetzt ausschließlich dem
              // Spalten-Fix-Aufruf oben, nicht mehr der Persistenz.
              const spById: Record<string, string> = {};
              for (const f of customFields) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                spById[f.id] = (f as any).spInternalName || '';
              }
              const merged = serializeCustomFields(customFields, bilingualFields).map(f => ({
                ...f,
                spInternalName: fixResult.customFieldMap![f.id] || spById[f.id] || '',
              }));
              await updateEvent(selectedEventId, { 'CustomFields': JSON.stringify(merged) });
            }
          }
        } catch (err) { console.warn('[DEX] Auto-fix Teilnehmer-Columns fehlgeschlagen:', err); }

        setProgress(90);
        // Bild als Attachment hochladen (falls neues Bild gewaehlt wurde)
        if (imageFile) {
          try {
            setProgressLabel(isDe ? 'Bild wird hochgeladen...' : 'Uploading image...');
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ctx = (window as any).__dexSpfxContext;
            if (ctx) {
              const svc = new EventService(ctx);
              const compressed = await compressImage(imageFile);
              const uploadedUrl = await svc.uploadEventImageAsAttachment(Number(selectedEventId), compressed);
              if (uploadedUrl) {
                await svc.updateEventImageUrl(Number(selectedEventId), uploadedUrl);
                // Events neu laden, damit die UI das frische Bild ohne Hard-Refresh anzeigt
                // (updateEvent oben hat schon einmal geladen, aber zu dem Zeitpunkt war
                // EventImageUrl noch der alte Wert)
                await refreshEvents();
              } else {
                setImageUploadError('Bild-Upload fehlgeschlagen.');
              }
            }
          } catch (err) {
            console.warn('[DEX] Bild-Upload fehlgeschlagen', err);
            setImageUploadError('Bild-Upload fehlgeschlagen.');
          }
        }
        setProgress(95);
        setProgressLabel(isDe ? 'Outlook wird aktualisiert...' : 'Updating Outlook...');
        // v11.63: Outlook-Updates pro Event entscheiden — der Organizer hat
        // im Confirm-Modal pro betroffenem Event (Top + Sub) einzeln ent-
        // schieden. Top-Event bekommt UpdateEvent nur, wenn explizit
        // angehakt (pendingOutlookUpdateForTopRef.current). OutlookDirty
        // wurde fuer das Top-Event schon im updateEvent-Call oben mit
        // pendingOutlookDirtyWriteRef geschrieben.
        if (!disableOutlook && pendingOutlookUpdateForTopRef.current) {
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ctx = (window as any).__dexSpfxContext;
            if (ctx && editEvent?.subsiteUrl) {
              const svc = new EventService(ctx);
              await svc.queueOutlookEvent('', selectedEventId, title, 'UpdateEvent');
            }
          } catch { /* Outlook-Update optional */ }
        }
        // v11.63: Sub-Event-Outlook-Updates pro angehaktem Sub-Event.
        if (pendingOutlookUpdateForSubEventsRef.current.length > 0) {
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ctx = (window as any).__dexSpfxContext;
            if (ctx) {
              const svc = new EventService(ctx);
              for (const subId of pendingOutlookUpdateForSubEventsRef.current) {
                const subDraft = subEventsRef.current.find(s => s.dbId === subId);
                const subTitle = subDraft?.title || '';
                try {
                  await svc.queueOutlookEvent('', subId, subTitle, 'UpdateEvent');
                  await updateEvent(subId, { 'OutlookDirty': false });
                } catch { /* einzelne Sub-Update-Fehler nicht eskalieren */ }
              }
            }
          } catch { /* Sub-Outlook-Updates optional */ }
        }
        // v11.63: Sub-Events, die im Modal waren aber NICHT angehakt wurden,
        // bekommen OutlookDirty=true, damit beim naechsten Wizard-Lauf der
        // Hinweis erscheint. Aus pendingOutlookDirtyWriteRefs lesen — Top-
        // Level haben wir oben bereits ueber pendingOutlookDirtyWriteRef
        // erledigt, hier nur Sub-Events.
        try {
          const dirtyMap = pendingOutlookDirtyWriteRefs.current || {};
          const checkedSubIds = new Set(pendingOutlookUpdateForSubEventsRef.current);
          for (const subId of Object.keys(dirtyMap)) {
            if (subId === selectedEventId) continue; // Top-Level schon erledigt
            if (checkedSubIds.has(subId)) continue;  // bereits auf false gesetzt
            if (dirtyMap[subId] === true) {
              try { await updateEvent(subId, { 'OutlookDirty': true }); }
              catch { /* */ }
            }
          }
        } catch { /* */ }
        setProgress(100);
        setProgressLabel('Änderungen gespeichert!');
        // v9.45: Soft-Refresh analog zum Create-Pfad. Wizard verlassen via
        // CustomEvent, DexEventPlatform navigiert + zeigt Banner. Refresh wird
        // beim Update sofort getriggert (loadEvents in updateEvent hat schon
        // gefeuert) — kein delayed refresh nötig wie beim Create.
        try {
          // v17.4: Nach erfolgreichem Save den Initial-Snapshot auf den
          // aktuellen Stand setzen, damit die Navigation-Guard (Unsaved-
          // Changes-Confirm) anschliessend nicht falsch ausloest. Sonst
          // sieht der User bei jedem Zurueck-Klick nach Save das Modal,
          // obwohl alles persistiert ist.
          initialFormSnapshotRef.current = computeFormSnapshot();
          setNavigationGuard(null);
          // v17.21: Statt sofort den Wizard zu verlassen, oeffnet sich erst
          // das Summary-Export-Modal. Der eigentliche Success-Dispatch laeuft
          // erst, wenn der User dort eine Auswahl getroffen hat (PDF / Word /
          // Nein, danke).
          pendingSuccessDispatchRef.current = { title, eventId: String(selectedEventId), type: 'update' };
          setPendingSuccessDispatch({ title, eventId: String(selectedEventId), type: 'update' });
          setShowSummaryModal(true);
        } catch { /* */ }
        setIsSubmitting(false);
      } else {
        setIsSubmitting(false);
        setProgress(0);
        setError('Event konnte nicht aktualisiert werden.');
      }
    } else {
      // Neues Event erstellen — v11.87: Progress wird per Callback-Stage vom
      // EventService getrieben, damit der Balken sich tatsaechlich an die
      // laufende SP-Operation koppelt und nicht stumm bei 92 % stehen bleibt.
      try {
      // Sub-Event-Anzahl bestimmt die Aufteilung des Bereichs 30 % - 90 %.
      // Bei N Sub-Events haben wir (1 Top + N Sub) Anlagen, die diesen
      // Bereich gleichmaessig fuellen. Ohne Sub-Events bleibt das Hauptevent
      // den ganzen Bereich fuer sich.
      const subEventDraftsCount = subEventsRef.current.filter(d => d.title && d.title.trim()).length;
      const totalAnlagen = 1 + subEventDraftsCount;
      const topStart = 30;
      const topEnd = topStart + (90 - topStart) / totalAnlagen;
      // Innerhalb des Top-Event-Slots (topStart..topEnd) werden die Stages
      // verteilt: subsite-creating, permissions, list-creating, list-done,
      // item-insert, done.
      const reportCreateStage = (stage: string): void => {
        const slot = topEnd - topStart;
        switch (stage) {
          case 'start':
            setProgress(Math.round(topStart));
            setProgressLabel('Event-Daten gespeichert — Teilnehmer-Subsite wird angelegt...');
            break;
          case 'subsite-creating':
            setProgress(Math.round(topStart + slot * 0.05));
            setProgressLabel('Teilnehmer-Subsite wird angelegt...');
            break;
          case 'subsite-done':
            setProgress(Math.round(topStart + slot * 0.35));
            setProgressLabel('Subsite angelegt — Berechtigungen werden gesetzt...');
            break;
          case 'permissions':
            setProgress(Math.round(topStart + slot * 0.45));
            setProgressLabel('Berechtigungen werden gesetzt...');
            break;
          case 'list-creating':
            setProgress(Math.round(topStart + slot * 0.55));
            setProgressLabel('Teilnehmerliste wird angelegt...');
            break;
          case 'list-done':
            setProgress(Math.round(topStart + slot * 0.80));
            setProgressLabel('Teilnehmerliste fertig — Views werden konfiguriert...');
            break;
          case 'item-insert':
            setProgress(Math.round(topStart + slot * 0.90));
            setProgressLabel('Event-Daten werden gespeichert...');
            break;
          case 'done':
            setProgress(Math.round(topEnd));
            setProgressLabel(subEventDraftsCount > 0
              ? `Haupt-Event angelegt — Sub-Event 1 wird angelegt...`
              : 'Haupt-Event angelegt — Berechtigungen und Aufräumarbeiten...');
            break;
          default:
            break;
        }
      };

      setProgress(10);
      setProgressLabel('Event-Daten werden vorbereitet...');

      // v16.4: Audience-DLs beim Save in Member-E-Mails aufloesen, damit der
      // Runtime-Sichtbarkeits-Check sie ohne weitere Graph-Calls treffen kann.
      const audienceResolved = await resolveAudienceMembersToCsv(audience, getGroupMembers);

      const sanitizedOrgPairCreate = sanitizeOrganizerPairs();
      const eventId = await createEvent({
        title,
        type: eventType,
        status: 'Active',
        description,
        location,
        locationAddress: (addrStreet || addrHouseNo || addrZip || addrCity)
          ? JSON.stringify({ street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity })
          : '',
        // v18.40: manueller Outlook-Ort (leer = Auto in createEvent).
        outlookLocation: outlookLocationOverride.trim() || undefined,
        outlookSubject: effOutlookSubject.trim() || undefined,
        outlookStart: outlookStartOverride || undefined,
        outlookEnd: outlookEndOverride || undefined,
        locationFilter,
        audience,
        audienceResolvedEmails: audienceResolved,
        filterMode,
        startDate: startDate ? berlinLocalToUtcIso(startDate) : '',
        endDate: endDate ? berlinLocalToUtcIso(endDate) : '',
        registrationDeadline: deadlineToEndOfDayIso(registrationDeadline) || '',
        lastDeregisterDate: deadlineToEndOfDayIso(lastDeregisterDate) || '',
        maxParticipants: unlimitedParticipants ? 0 : (Number(maxParticipants) || 0),
        waitlistEnabled,
        eventImageUrl: imageUrl,
        // Sanitize: paart Organizer-Names + -Emails 1:1, droppt unvollständige
        // Pairs (Name ohne Email oder umgekehrt) — verhindert Mismatch-State
        // in DEX_Events durch Drift während Edit/Closure-Bugs.
        organizer: sanitizedOrgPairCreate.orgString,
        organizerEmail: sanitizedOrgPairCreate.orgEmailString,
        contactName: contactName.trim(),
        contactEmail: contactEmail.trim(),
        contactInfo: contactInfo.trim(),
        outlookEventId: '',
        outlookBody: (() => {
          // v7.4: Auch wenn der User keinen Outlook-Body geschrieben hat,
          // immer das Outlook-Layout mit einem Default-Body erzeugen,
          // der auf den Organizer für organisatorische Fragen verweist
          // (analog zur Anmeldebestätigungs-Mail).
          const vars = {
            EventTitle: title,
            Organizer: organizer,
            Location: location,
            Address: [addrStreet, addrHouseNo].filter(Boolean).join(' ') + ((addrZip || addrCity) ? ', ' + [addrZip, addrCity].filter(Boolean).join(' ') : ''),
            StartDate: startDate ? new Date(startDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
            EndDate: endDate ? new Date(endDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
          };
          const orgNames = organizer.split(';').map(s => s.trim()).filter(Boolean).join(', ');
          const escHtml = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
          // v9.8: gleicher Default-Body wie im Update-Pfad — inkl. Abmelde-Hinweis
          // mit Link auf die App ("Meine Events"-Tab).
          const APP_URL_OL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView';
          const defaultBody = effEmailLanguage === 'EN'
            ? `<p>You are registered for the event <strong>${escHtml(title)}</strong>.</p>`
              + `<p>If you are unable to attend, please cancel your registration in time via the <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;My Events&ldquo;).</p>`
              + `<p>For organizational questions please contact <strong>${escHtml(orgNames || 'the organizer')}</strong>.</p>`
            : `<p>Du bist für das Event <strong>${escHtml(title)}</strong> angemeldet.</p>`
              + `<p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;Meine Events&ldquo;) ab.</p>`
              + `<p>Bei organisatorischen Fragen wende dich bitte an <strong>${escHtml(orgNames || 'den Organizer')}</strong>.</p>`;
          const resolvedBody = effOutlookBody ? replacePlaceholders(effOutlookBody, vars) : defaultBody;
          const resolvedHeading = effOutlookHeading ? replacePlaceholders(effOutlookHeading, vars) : title;
          const resolvedSub = effOutlookSubheading ? replacePlaceholders(effOutlookSubheading, vars) : undefined;
          // v18.73: Header-Bild Größe + Innenabstand (event-weit) in den Outlook-Body.
          const wrapped = buildOutlookBody(resolvedHeading, resolvedBody, resolvedSub, { imageWidth: headerImageLayout.width, imagePaddingV: headerImageLayout.paddingV, imagePaddingH: headerImageLayout.paddingH });
          // v11.93: Logo aus Top-Level-Resolver, sonst landet beim Speichern
          // aus einem Sub-Tab das Sub-Logo aufs Haupt-Event.
          return wrapped.replace(/\{\{ORB_URL\}\}/g, effOutlookLogo || getCachedOrbBase64() || '');
        })(),
        agenda: JSON.stringify(agenda),
        transfers: JSON.stringify(transferTimes),
        documents: '[]', // Dokumente werden nach erfolgreichem Upload gespeichert
        funZone: JSON.stringify(quiz),
        quizClusterSize: Math.min(Math.max(1, quizClusterSize || 1), 4),
        emailLanguage: effEmailLanguage,
        registrationLanguage: registrationLanguage || undefined,
        emailTemplateOverrides: (() => {
          const b2runExtra = (durchstarterStartblock || funstarterStartblock || durchstarterRequiresProof)
            ? { _b2run: {
                ...(durchstarterStartblock ? { durchstarterStartblock } : {}),
                ...(funstarterStartblock ? { funstarterStartblock } : {}),
                ...(durchstarterRequiresProof ? { durchstarterRequiresProof: true } : {}),
              } }
            : {};
          const qrExtra = qrScannerEmails.length > 0
            ? { _qrScanners: qrScannerNames.map((n, i) => ({ name: n, email: qrScannerEmails[i] || '' })).filter(x => x.email) }
            : {};
          // v9.18: Co-Organizer ebenso in EmailTemplateOverrides piggybacken.
          const coExtra = coOrganizerEmails.length > 0
            ? { _coOrganizers: coOrganizerNames.map((n, i) => ({ name: n, email: coOrganizerEmails[i] || '' })).filter(x => x.email) }
            : {};
          // v9.21: Test-Team ebenso piggybacken.
          const ttExtra = testTeamEmails.length > 0
            ? { _testTeam: testTeamNames.map((n, i) => ({ name: n, email: testTeamEmails[i] || '' })).filter(x => x.email) }
            : {};
          // v11.25: Display-Reihenfolge-Toggle (siehe Edit-Pfad oben).
          const splitDispRevExtra = splitDisplayOrderReversed && useSplitCapacities
            ? { _splitDisplayOrderReversed: true }
            : {};
          // v14.8: subEventsOnlyMode + custom childEventTerm piggybacken;
          // subEventsOnlyMode impliziert requireSubEventSelection.
          const effRequireSubEventSelection = requireSubEventSelection || subEventsOnlyMode;
          const reqSubEvtExtra = effRequireSubEventSelection
            ? { _requireSubEventSelection: true }
            : {};
          const subEvtsOnlyExtra = subEventsOnlyMode
            ? { _subEventsOnlyMode: true }
            : {};
          const childTermExtra = (childTermSingular.trim() || childTermPlural.trim())
            ? { _childEventTerm: { singular: childTermSingular.trim(), plural: childTermPlural.trim() } }
            : {};
          // v18.9: Organizer-Anzeige ausblenden (Piggyback).
          const hideOrganizerExtra = hideOrganizer ? { _hideOrganizer: true } : {};
          // v11.93: Top-Level-Logos aus dem Resolver lesen.
          const hasAny = Object.keys(emailTemplateOverrides).length > 0 || effEmailLogo || effOutlookLogo || Object.keys(b2runExtra).length > 0 || Object.keys(qrExtra).length > 0 || Object.keys(coExtra).length > 0 || Object.keys(ttExtra).length > 0 || Object.keys(splitDispRevExtra).length > 0 || Object.keys(reqSubEvtExtra).length > 0 || Object.keys(subEvtsOnlyExtra).length > 0 || Object.keys(childTermExtra).length > 0 || Object.keys(hideOrganizerExtra).length > 0 || Object.keys(headerImageLayoutConfig).length > 0;
          return hasAny
            ? JSON.stringify({
                ...(effEmailLogo ? { _eventLogo: effEmailLogo } : {}),
                ...(effOutlookLogo ? { _outlookLogo: effOutlookLogo } : {}),
                ...b2runExtra,
                ...splitDispRevExtra,
                ...qrExtra,
                ...coExtra,
                ...ttExtra,
                ...reqSubEvtExtra,
                ...subEvtsOnlyExtra,
                ...childTermExtra,
                ...hideOrganizerExtra,
                // v18.73: Header-Bild-Layout (Breite + Innenabstand) — event-weit.
                ...headerImageLayoutConfig,
                ...emailTemplateOverrides,
              })
            : '';
        })(),
        // v11.93: aus dem Top-Level-Resolver — Sub-Tab-Werte würden sonst
        // beim Save fälschlich aufs Haupt-Event übernommen.
        disableEmails: effDisableEmails,
        // v19.22: granulare An-/Abmelde-Mail-Schalter (Top-Level aufgelöst).
        disableRegistrationEmail: effDisableRegistrationEmail,
        disableCancellationEmail: effDisableCancellationEmail,
        disableOutlook: effDisableOutlook,
        notifyOrgRegisterMode,
        notifyOrgRegisterFromDate: notifyOrgRegisterMode === 'fromDate' && notifyOrgRegisterFromDate ? berlinLocalToUtcIso(notifyOrgRegisterFromDate) : '',
        notifyOrgCancelMode,
        excludedUsers,
        isFictive,
        durchstarterCapacity: useSplitCapacities ? (parseInt(durchstarterCapacity, 10) || 0) : undefined,
        funstarterCapacity: useSplitCapacities ? (parseInt(funstarterCapacity, 10) || 0) : undefined,
        splitLabelA: useSplitCapacities ? (splitLabelA || '').trim() : undefined,
        splitLabelB: useSplitCapacities ? (splitLabelB || '').trim() : undefined,
        splitSharedWaitlist: useSplitCapacities ? !!splitSharedWaitlist : undefined,
        allowAttendeeUpload: !!allowAttendeeUpload,
        attendeeUploadHint: (attendeeUploadHint || '').trim() || undefined,
        attendeeUploadLabel: (attendeeUploadLabel || '').trim() || undefined,
        // v18.33: Self-Check-in mit-durchreichen (Token + optionales Fenster).
        selfCheckInEnabled: !!selfCheckInEnabled,
        selfCheckInToken: selfCheckInEnabled ? (selfCheckInToken || undefined) : undefined,
        selfCheckInFrom: selfCheckInEnabled && selfCheckInFrom ? selfCheckInFrom : undefined,
        selfCheckInTo: selfCheckInEnabled && selfCheckInTo ? selfCheckInTo : undefined,
        // v11.80: Anrede-Toggle + Team-Anmelde-Konfiguration mit-durchreichen.
        askSalutation: !!askSalutation,
        // v18.75: Sicherheitshinweis vor dem Absenden.
        confirmDialogEnabled: !!confirmDialogEnabled,
        confirmDialogMode: confirmDialogEnabled ? (confirmDialogMode || 'summary') : '',
        confirmDialogText: confirmDialogEnabled && confirmDialogMode === 'freetext' ? confirmDialogText : '',
        teamRegistrationEnabled: !!teamRegistrationEnabled,
        teamSize: teamRegistrationEnabled && teamSize > 0 ? teamSize : undefined,
        askTeamName: !!askTeamName,
        // v11.81: Erweiterte Team-Konfiguration mit-durchreichen.
        teamPartialAllowed: !!(teamRegistrationEnabled && teamPartialAllowed),
        teamOpenSlotsVisible: !!(teamRegistrationEnabled && teamOpenSlotsVisible),
        teamJoinRequiresApproval: !!(teamRegistrationEnabled && teamOpenSlotsVisible && teamJoinRequiresApproval),
        // v17.20: Bilingual-Toggle durchreichen.
        bilingualFields: !!bilingualFields,
        // v17.22: zentraler serializeCustomFields-Helper.
        customFields: serializeCustomFields(customFields, bilingualFields),
        onProgress: reportCreateStage,
      });

      if (eventId) {
        // v11.87: Sub-Events bekommen den Bereich (topEnd..90) gleichmäßig
        // aufgeteilt — pro Sub-Event ein eigener Stage-Slot. persistSubEventsForParent
        // erhält einen Sub-Progress-Callback über ein Window-Event-Bus-ähnliches
        // Setup ist hier nicht nötig, weil wir die Schleife per index zählen.
        if (subEventDraftsCount > 0) {
          const subSlotSize = (90 - topEnd) / subEventDraftsCount;
          // Wir setzen pro Sub-Event-Start manuell den Progress und übergeben
          // optional einen onProgress-Callback an persistSubEventsForParent, um
          // den Sub-Site-Anlage-Fortschritt fein abzubilden. Da
          // persistSubEventsForParent intern über sequenzielle createEvent
          // läuft, koppeln wir den Sub-Progress an einen externen Counter.
          let processedSubIdx = 0;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__dexSubEventProgress = (stage: string): void => {
            const base = topEnd + subSlotSize * processedSubIdx;
            const slot = subSlotSize;
            switch (stage) {
              case 'start':
                setProgress(Math.round(base));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} wird vorbereitet...`);
                break;
              case 'subsite-creating':
                setProgress(Math.round(base + slot * 0.10));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} — Subsite wird angelegt...`);
                break;
              case 'subsite-done':
                setProgress(Math.round(base + slot * 0.45));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} — Berechtigungen werden gesetzt...`);
                break;
              case 'list-creating':
                setProgress(Math.round(base + slot * 0.60));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} — Teilnehmerliste wird angelegt...`);
                break;
              case 'list-done':
                setProgress(Math.round(base + slot * 0.80));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} — Views werden konfiguriert...`);
                break;
              case 'item-insert':
                setProgress(Math.round(base + slot * 0.90));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} — Event-Daten werden gespeichert...`);
                break;
              case 'done':
                processedSubIdx += 1;
                setProgress(Math.round(topEnd + subSlotSize * processedSubIdx));
                setProgressLabel(processedSubIdx >= subEventDraftsCount
                  ? 'Sub-Events fertig — Aufräumarbeiten...'
                  : `Sub-Event ${processedSubIdx} fertig — Sub-Event ${processedSubIdx + 1} wird angelegt...`);
                break;
              default:
                break;
            }
          };
          setProgress(Math.round(topEnd));
          setProgressLabel(`Sub-Event 1 von ${subEventDraftsCount} wird angelegt...`);
        } else {
          setProgress(Math.round(topEnd));
          setProgressLabel('Haupt-Event angelegt — Aufräumarbeiten...');
        }
        try { await persistSubEventsForParent(String(eventId)); }
        catch (err) { console.warn('[DEX] Sub-Events beim Create persistieren fehlgeschlagen:', err); }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        try { delete (window as any).__dexSubEventProgress; } catch { /* */ }
        setProgress(92);
        setProgressLabel('Dokumente und Bild werden hochgeladen...');
        // E-Mail an Organisator senden
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const ctx = (window as any).__dexSpfxContext;
          if (ctx) {
            const svc = new EventService(ctx);
            // Subsite-URL aus dem neu geladenen Event holen
            const allEvents = await svc.getEvents();
            const created = allEvents.find(e => String(e.Id) === String(eventId));
            const subsiteUrl = created?.SubsiteUrl || '';
            // Dokumente als Attachments an das Event-Item anfuegen (kein updateEvent noetig)
            if (eventId && documents.length > 0) {
              for (const doc of documents) {
                if (doc.file) {
                  await svc.uploadEventDocument(Number(eventId), doc.file);
                }
              }
            }
            // Event-Bild als Attachment hochladen + URL ins Item schreiben
            if (eventId && imageFile) {
              try {
                setProgressLabel('Bild wird hochgeladen...');
                const compressed = await compressImage(imageFile);
                const uploadedUrl = await svc.uploadEventImageAsAttachment(Number(eventId), compressed);
                if (uploadedUrl) {
                  await svc.updateEventImageUrl(Number(eventId), uploadedUrl);
                  // v9.41: KEIN refreshEvents direkt nach Create. Die Subsite ist
                  // gerade erst angelegt und die SP-API ist noch nicht konsistent —
                  // ein Refresh hier kann zu 400/404 auf den Subsite-Listen und im
                  // schlimmsten Fall zu einem React-Render-Crash führen. Refresh
                  // wird verschoben auf den Klick 'Events anzeigen' auf der
                  // Erfolgsseite, dann hat SP genug Zeit zum Propagieren.
                } else {
                  setImageUploadError('Bild-Upload fehlgeschlagen.');
                }
              } catch (err) {
                console.warn('[DEX] Bild-Upload fehlgeschlagen', err);
                setImageUploadError('Bild-Upload fehlgeschlagen.');
              }
            }
            // Event-Created Mail an alle Organizer senden.
            // {{Name}} in der Anrede = nur Vorname (nicht voller Name), darum
            // den Organizer-String anhand von ";" in Namen splitten und pro
            // Name das erste Token als Vorname nehmen. Paarweise zu den
            // organizerEmails - bei Laengen-Mismatch faellt wir auf den
            // ersten Namen zurueck.
            const allOrgEmails = organizerEmails.length > 0 ? organizerEmails : [currentUser.email];
            const orgNames = organizer.split(';').map(s => s.trim()).filter(Boolean);
            for (let i = 0; i < allOrgEmails.length; i++) {
              const orgEmail = allOrgEmails[i];
              const orgFullName = orgNames[i] || orgNames[0] || `${currentUser.firstName} ${currentUser.surname}`;
              const orgFirstName = orgFullName.split(/\s+/)[0] || orgFullName;
              const emailData = eventCreatedEmail(orgFirstName, title, subsiteUrl);
              svc.queueEmail(
                emailData.subject, orgEmail, orgFullName, emailData.body,
                'EventErstellt', title, String(eventId)
              ).catch(err => console.warn('[DEX]', err));
            }
          }
        } catch { /* E-Mail-Fehler ignorieren */ }
        // v9.45: Soft-Refresh statt Hard-Reload. Statt die Success-Page zu rendern
        // (wo zwischen Wizard und SuccessPage ein React #300 auftrat) ODER die
        // Page hart zu reloaden (was den User auf der Landing-Seite landete),
        // gehen wir den Mittelweg:
        //
        // 1. Wizard sofort verlassen via dispatchEvent('dex-event-submit-success',
        //    {title, eventId, type}) — DexEventPlatform hört darauf, navigiert zur
        //    Event-Liste und zeigt den grünen Erfolgs-Banner.
        // 2. setIsSubmitting(false) damit der Wizard unmounted (kein hängender
        //    Submit-State).
        // 3. Delayed refreshEvents (3 Sekunden später) lädt das frisch erstellte
        //    Event nach — SP hatte dann genug Zeit zum Propagieren und der Read
        //    auf die neue Subsite läuft sauber durch (gleicher Pfad wie der
        //    Aktualisieren-Button im Header — der hat nie Probleme).
        setProgress(100);
        setProgressLabel('Event erfolgreich erstellt!');
        try {
          // v17.4: gleicher Reset wie im Update-Pfad, damit der
          // Navigation-Guard nach erfolgreichem Create nicht stoert.
          initialFormSnapshotRef.current = computeFormSnapshot();
          setNavigationGuard(null);
          // v17.21: Summary-Export-Modal vor dem Submit-Success-Dispatch.
          pendingSuccessDispatchRef.current = { title, eventId: String(eventId), type: 'create' };
          setPendingSuccessDispatch({ title, eventId: String(eventId), type: 'create' });
          setShowSummaryModal(true);
        } catch { /* */ }
        setIsSubmitting(false);
        // Delayed Refresh — SP braucht typischerweise 2-5s bis frische Subsite-
        // Listen API-konsistent abrufbar sind. 3000ms ist ein guter Kompromiss.
        setTimeout(() => {
          refreshEvents().catch(err => console.warn('[DEX] post-create soft refresh fehlgeschlagen:', err));
        }, 3000);
      } else {
        setIsSubmitting(false);
        setProgress(0);
        setError('Event konnte nicht erstellt werden. Bitte versuche es erneut.');
      }
      } catch (err) {
        setIsSubmitting(false);
        setProgress(0);
        setError(err instanceof Error ? err.message : 'Event konnte nicht erstellt werden.');
      }
    }
  };

  // v11.57: detect ob beim aktuellen Edit-State Outlook-relevante Aenderungen
  // anstehen (Title, Start, End oder OutlookBody). Vergleich gegen den
  // Mount-Snapshot.
  // v11.63: liefert pro betroffenem Event (Hauptevent + Sub-Events) eine
  // Liste der konkret geaenderten Felder, damit der User im Modal pro
  // Eintrag sehen kann, was sich wirklich veraendert hat.
  const detectOutlookRelevantChanges = (): { items: OutlookConfirmItem[] } => {
    const items: OutlookConfirmItem[] = [];
    if (!editEvent) return { items };
    const snap = initialOutlookSnapshot.current;
    // v11.64: Datetime-Vergleich ueber Date.getTime(), nicht String. Sonst
    // kippt der Vergleich an Format-Unterschieden (snap kommt roh aus SP
    // mit „2026-09-24T16:00:00Z", currentStart geht durch
    // berlinLocalToUtcIso() und wird „2026-09-24T16:00:00.000Z" — gleicher
    // Zeitpunkt, anderer String). Das hat den Hauptevent in v11.63
    // faelschlich als „Startzeit, Endzeit geaendert" gemeldet.
    const sameInstant = (a: string, b: string): boolean => {
      if (a === b) return true;
      if (!a && !b) return true;
      if (!a || !b) return false;
      const da = new Date(a).getTime();
      const db = new Date(b).getTime();
      if (isNaN(da) || isNaN(db)) return a === b;
      return da === db;
    };
    const currentTitle = title || '';
    const currentStart = startDate ? berlinLocalToUtcIso(startDate) : '';
    const currentEnd = endDate ? berlinLocalToUtcIso(endDate) : '';
    // Outlook-Body vergleich anhand des „rohen" Body (ohne Wrapper), da der
    // Wrapper bei jedem Save neu gebaut wird und dadurch immer „aenderbar"
    // aussehen wuerde. Vergleich gegen den initial gestrippten Wert.
    const initialStripped = stripOutlookWrapper(snap.outlookBody || '');
    const currentStripped = activeCommTabIdx === 0 ? (outlookBody || '') : stripOutlookWrapper(snap.outlookBody || '');
    const currentTopLocation = outlookLocationOverride.trim() || buildOutlookLocation(location, { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity });
    const currentTopSubject = (resolveTopLevelCommState().outlookSubject || '').trim();
    // v19.20: globale Header-Bild-Layout-Änderung (Breite/Innenabstand) erkennen.
    // Das Layout steht NICHT im rohen Body (wird erst beim Wrappen angewendet),
    // betrifft aber den Hero-Bild-Kopf des Outlook-Termins — daher als eigenes
    // Änderungs-Feld „layout" werten, damit das Update-Modal aufgeht (und der
    // Grund klar als „Kopfbild" benannt wird, nicht irreführend als „Termin-Text").
    const initLayout = initialHeaderImageLayoutRef.current;
    const layoutChanged = headerImageLayout.width !== initLayout.width
      || headerImageLayout.paddingV !== initLayout.paddingV
      || headerImageLayout.paddingH !== initLayout.paddingH;
    const topChangedFields: Array<'title' | 'startDate' | 'endDate' | 'outlookBody' | 'location' | 'subject' | 'layout'> = [];
    if (currentTitle !== (snap.title || '')) topChangedFields.push('title');
    if (!sameInstant(currentStart, snap.startDate || '')) topChangedFields.push('startDate');
    if (!sameInstant(currentEnd, snap.endDate || '')) topChangedFields.push('endDate');
    if (currentStripped !== initialStripped) topChangedFields.push('outlookBody');
    if (layoutChanged) topChangedFields.push('layout');
    // v18.34: reine Ort-Aenderung gilt ebenfalls als Outlook-relevant.
    if (currentTopLocation !== (snap.outlookLocation || '')) topChangedFields.push('location');
    // v18.42: reine Betreff-Aenderung gilt ebenfalls als Outlook-relevant.
    if (currentTopSubject !== (snap.outlookSubject || '').trim()) topChangedFields.push('subject');
    // v18.44: abweichendes Outlook-Datum (Override) gilt als Termin-Aenderung.
    if ((outlookStartOverride || '') !== (snap.outlookStart || '') && topChangedFields.indexOf('startDate') < 0) topChangedFields.push('startDate');
    if ((outlookEndOverride || '') !== (snap.outlookEnd || '') && topChangedFields.indexOf('endDate') < 0) topChangedFields.push('endDate');
    // v11.61: Beide Pointer pruefen — DEX_CreateOutlookEvent setzt nur
    // CalendarLink auf Erfolg, OutlookEventId bleibt leer. Wer beides
    // leer hat, hatte nie einen Outlook-Termin.
    const topHasOutlook = !!editEvent.outlookEventId || !!editEvent.calendarLink;
    // v18.45 BUG-FIX: für das Hauptevent IMMER dessen Top-Level-DisableOutlook
    // prüfen — nicht das rohe `disableOutlook` (das hält beim Speichern auf einem
    // Sub-Event-Tab den Sub-Wert). Sonst wurde das Hauptevent fälschlich im
    // Update-Modal gelistet, obwohl dort Outlook deaktiviert ist (z.B. Event mit
    // Outlook nur auf Sub-Event-Ebene).
    const topDisableOutlook = resolveTopLevelCommState().disableOutlook;
    // v18.51: Im „Nur für Sub-Events"-Modus (subEventsOnlyMode) ist das
    // Hauptevent von der Teilnehmer-Anmeldung ausgenommen — niemand meldet sich
    // direkt fürs Hauptevent an. Ein Outlook-Update-Hinweis fürs Hauptevent ist
    // dann sinnlos und wird unterdrückt (Sub-Events bekommen weiter ihre Hinweise).
    if (topChangedFields.length > 0 && !topDisableOutlook && topHasOutlook && !subEventsOnlyMode) {
      items.push({
        kind: 'top',
        eventId: editEvent.id,
        title: currentTitle || editEvent.title || '',
        changedFields: topChangedFields,
      });
    }
    // Sub-Events: pro Sub-Event vergleichen.
    // v11.60: subEventsRef statt subEvents — der Flush hat die aktuellen
    // UI-Werte gerade synchron in den Ref geschrieben, der React-State
    // ist noch nicht propagiert.
    for (const s of subEventsRef.current) {
      if (!s.dbId) continue;
      const initTitle = s.initialTitle || '';
      const initStart = s.initialStartDate || '';
      const initEnd = s.initialEndDate || '';
      const initBodyStripped = stripOutlookWrapper(s.initialOutlookBody || '');
      const curBodyStripped = (s.outlookBody || '');
      const hasOutlookEvId = !!s.initialOutlookEventId || !!s.initialCalendarLink;
      const subChangedFields: Array<'title' | 'startDate' | 'endDate' | 'outlookBody' | 'layout'> = [];
      if ((s.title || '') !== initTitle) subChangedFields.push('title');
      // v11.64: auch hier semantischer Vergleich — gleiche Falle wie oben.
      if (!sameInstant(s.startDate || '', initStart)) subChangedFields.push('startDate');
      if (!sameInstant(s.endDate || '', initEnd)) subChangedFields.push('endDate');
      if (curBodyStripped !== initBodyStripped) subChangedFields.push('outlookBody');
      // v19.20: globale Header-Bild-Layout-Änderung betrifft auch die
      // Sub-Event-Outlook-Termine (gleicher Hero-Bild-Kopf) — als eigenes
      // „layout"-Feld werten, damit das Update-Modal sie mit auflistet.
      if (layoutChanged) subChangedFields.push('layout');
      // v11.66: Debug-Log fuer jeden Sub-Event, damit wir in der Browser-
      // Konsole nachvollziehen koennen, warum das Modal manchmal nicht
      // erscheint. v11.67: JSON.stringify damit der Browser die Werte
      // direkt anzeigt (statt nur „Object" mit Klick zum Aufklappen).
      // v11.79: nur noch sichtbar, wenn der Maintainer in der Console
      // `window.__dexDebug = true` setzt — sonst spammt das Log im
      // Normalbetrieb die DevTools voll.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      if (typeof window !== 'undefined' && (window as any).__dexDebug) {
        // eslint-disable-next-line no-console
        console.log('[DEX][outlook-detect][sub] ' + JSON.stringify({
          dbId: s.dbId,
          title: s.title,
          subChangedFields,
          disableOutlook: s.disableOutlook,
          hasOutlookEvId,
          initialOutlookEventId: s.initialOutlookEventId,
          initialCalendarLink: s.initialCalendarLink,
          bodyLenInitial: (s.initialOutlookBody || '').length,
          bodyLenCurrent: (s.outlookBody || '').length,
          bodyLenInitStripped: initBodyStripped.length,
          bodyMatch: curBodyStripped === initBodyStripped,
          titleMatch: (s.title || '') === initTitle,
        }));
      }
      if (subChangedFields.length > 0 && !s.disableOutlook) {
        items.push({
          kind: 'sub',
          eventId: s.dbId,
          title: s.title || '',
          changedFields: subChangedFields,
          // v11.68: ohne CalendarLink/OutlookEventId existiert kein Outlook-Termin
          // — Save persistiert den neuen Body in DEX_Events, aber wir koennen
          // kein UpdateEvent queuen. Modal rendert Info-Eintrag.
          noOutlookYet: !hasOutlookEvId,
        });
      }
    }
    // v14.8: Items aus persistiertem OutlookDirty-Flag nachziehen. Wenn ein
    // Sub-Event oder das Hauptevent in einer früheren Session als „Outlook-
    // Update ausstehend" markiert wurde (User hat den Haken damals nicht
    // gesetzt → OutlookDirty=true wurde in SP geschrieben), soll der nächste
    // Save trotzdem das Modal anbieten — auch ohne neue inhaltliche Änderung
    // in dieser Session. Sonst bleibt der Dirty-Flag ewig hängen und der
    // Yellow-Hint in Schritt 1 wird nie aufgelöst.
    const hasItemForEvent = (id: string): boolean => items.some(it => it.eventId === id);
    // Hauptevent
    // v18.50 BUG-FIX: auch im Dirty-Marker-Pfad das Top-Level-DisableOutlook
    // prüfen (nicht das rohe `disableOutlook`, das beim Speichern auf einem
    // Sub-Event-Tab den Sub-Wert hält) — sonst taucht das Hauptevent im
    // Update-Modal als „Frühere Änderung nicht synchronisiert" auf, obwohl
    // dort Outlook deaktiviert ist (Event mit Outlook nur auf Sub-Event-Ebene).
    // Gleiche Falle wie v18.45 im Changed-Fields-Pfad oben.
    if (editEvent.outlookDirty && !topDisableOutlook && !subEventsOnlyMode
        && (editEvent.outlookEventId || editEvent.calendarLink)
        && !hasItemForEvent(editEvent.id)) {
      items.push({
        kind: 'top',
        eventId: editEvent.id,
        title: title || editEvent.title || '',
        changedFields: [],
      });
    }
    // Sub-Events
    for (const s of subEventsRef.current) {
      if (!s.dbId) continue;
      if (s.disableOutlook) continue;
      const hasOutlookEvId = !!s.initialOutlookEventId || !!s.initialCalendarLink;
      if (!hasOutlookEvId) continue;
      const childEvt = childEventsOf(editEvent.id).find(c => c.id === s.dbId);
      if (childEvt && childEvt.outlookDirty && !hasItemForEvent(s.dbId)) {
        items.push({
          kind: 'sub',
          eventId: s.dbId,
          title: s.title || '',
          changedFields: [],
          noOutlookYet: false,
        });
      }
    }
    // v11.79: gated debug log — siehe oben.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if (typeof window !== 'undefined' && (window as any).__dexDebug) {
      // eslint-disable-next-line no-console
      console.log('[DEX][outlook-detect][result] ' + JSON.stringify({
        itemsCount: items.length,
        items,
        activeCommTabIdx,
        topOutlookBodyLen: (outlookBody || '').length,
        topInitialOutlookBodyLen: (snap.outlookBody || '').length,
        topBodyMatch: currentStripped === initialStripped,
      }));
    }
    return { items };
  };

  // v11.57: Wrapper-Funktion fuer den Save-Button. Im Edit-Modus mit
  // Outlook-relevanter Aenderung wird das Confirm-Modal gezeigt. Sonst
  // direkt handleSubmit.
  const attemptSubmit = (): void => {
    // Aktuelle Tab-Werte zurueck ins jeweilige Slot schreiben, damit
    // beim handleSubmit nichts verloren geht.
    flushActiveCommTabToState();
    if (!isEditMode || !editEvent) {
      pendingOutlookDirtyWriteRef.current = null;
      pendingOutlookDirtyWriteRefs.current = {};
      pendingOutlookUpdateForTopRef.current = false;
      pendingOutlookUpdateForSubEventsRef.current = [];
      pendingOutlookRecreateForSubEventsRef.current = [];
      handleSubmit().catch(() => { /* Errors werden in handleSubmit gesetzt */ });
      return;
    }
    const det = detectOutlookRelevantChanges();
    if (det.items.length > 0) {
      // v11.63: Snapshot der Items + leerer Check-Map (alle false). Pro
      // Event entscheidet der Organizer einzeln im Modal.
      setOutlookConfirmItems(det.items);
      setOutlookConfirmChecks({});
      setOutlookConfirmOpen(true);
      return;
    }
    // Keine Outlook-relevante Aenderung — dirty-Flag nicht anfassen.
    pendingOutlookDirtyWriteRef.current = null;
    pendingOutlookDirtyWriteRefs.current = {};
    // Wenn der User den expliziten Step-5-Schalter „Outlook-Termin
    // aktualisieren" angehakt hat, ueberschreibt das die Modal-Logik
    // und triggert ein manuelles UpdateEvent fuer das Top-Level — auch
    // wenn die Detect-Heuristik nichts Outlook-relevantes gefunden hat.
    pendingOutlookUpdateForTopRef.current = !!triggerOutlookUpdate;
    pendingOutlookUpdateForSubEventsRef.current = [];
    pendingOutlookRecreateForSubEventsRef.current = [];
    handleSubmit().catch(() => { /* */ });
  };

  // v11.57: Confirm-Modal-Handler.
  // v11.63: Liest aus outlookConfirmChecks ab, welche Events der Organizer
  // angehakt hat. Angehakte Events bekommen UpdateEvent + OutlookDirty=false,
  // nicht angehakte (aber im Detect-Items gelistete) bekommen
  // OutlookDirty=true. Events ausserhalb des Detect-Items bleiben unberuehrt.
  const confirmOutlookSave = (): void => {
    setOutlookConfirmOpen(false);
    const topId = editEvent ? editEvent.id : '';
    const topItem = outlookConfirmItems.find(it => it.kind === 'top');
    const subItems = outlookConfirmItems.filter(it => it.kind === 'sub');
    const topChecked = !!topItem && !!outlookConfirmChecks[topItem.eventId];
    // v11.69: Angehakte Sub-Events trennen in:
    //  - `normalUpdateSubIds`: Sub-Event hat bereits einen Outlook-Termin →
    //    DEX_Outlook 'UpdateEvent' in die Queue schreiben (bestehender Pfad).
    //  - `recreateSubIds`: Sub-Event hat noch keinen Outlook-Termin
    //    (`noOutlookYet`) → DEX_Events-Item per `deleteEventItemOnly` loeschen
    //    und mit `existingSubsiteUrl` neu anlegen, damit der
    //    DEX_CreateOutlookEvent-Flow triggert. Teilnehmer-Subsite + Liste
    //    bleiben unangetastet erhalten.
    const checkedSubItems = subItems.filter(it => !!outlookConfirmChecks[it.eventId]);
    const normalUpdateSubIds = checkedSubItems.filter(it => !it.noOutlookYet).map(it => it.eventId);
    const recreateSubIds = checkedSubItems.filter(it => !!it.noOutlookYet).map(it => it.eventId);
    pendingOutlookUpdateForTopRef.current = topChecked;
    pendingOutlookUpdateForSubEventsRef.current = normalUpdateSubIds;
    pendingOutlookRecreateForSubEventsRef.current = recreateSubIds;
    // Pro Event-ID den OutlookDirty-Schreibwert vormerken.
    // v11.69: noOutlookYet-Items werden — egal ob angehakt oder nicht — NICHT
    // dirty markiert. Bei angehakt erfolgt ein Recreate (neues Item hat von
    // Haus aus OutlookDirty=false), bei nicht angehakt existiert immer noch
    // kein Outlook-Termin der "aus-Sync" sein koennte → Marker waere falsch.
    const dirtyMap: Record<string, boolean> = {};
    for (const it of outlookConfirmItems) {
      if (it.noOutlookYet) continue;
      dirtyMap[it.eventId] = !outlookConfirmChecks[it.eventId];
    }
    pendingOutlookDirtyWriteRefs.current = dirtyMap;
    // Top-Level kompatibel halten: wenn das Top-Event im Modal war, wird
    // OutlookDirty entsprechend gesetzt; sonst null = nicht anfassen.
    if (topItem) {
      pendingOutlookDirtyWriteRef.current = !topChecked;
    } else {
      pendingOutlookDirtyWriteRef.current = null;
    }
    // setTriggerOutlookUpdate steuert in handleSubmit, ob der Top-Level-
    // Outlook-Branch ueberhaupt betreten wird. v11.63: nur true wenn das
    // Top-Event angehakt wurde ODER mindestens ein Sub-Event angehakt
    // wurde (damit der Sub-Event-Branch im handleSubmit getroffen wird).
    setTriggerOutlookUpdate(topChecked || normalUpdateSubIds.length > 0 || recreateSubIds.length > 0);
    // Verhindern dass topId als „angehakt" interpretiert wird ohne Modal.
    void topId;
    handleSubmit().catch(() => { /* */ });
  };
  const cancelOutlookSave = (): void => {
    setOutlookConfirmOpen(false);
    // Nichts speichern — User bleibt im Wizard.
  };

  // v11.57: bei Sub-Event-Anzahl-Aenderung Tab sicher in Range halten.
  React.useEffect(() => {
    if (activeCommTabIdx > subEvents.length) {
      setActiveCommTabIdx(0);
    }
  }, [subEvents.length, activeCommTabIdx]);
  // v15.0: gleiche Range-Garantie fuer die neuen Tab-Sets in den
  // Steps 3 (Ort), 4 (Kapazitaet) und 6 (Felder).
  React.useEffect(() => {
    if (activeLocationTabIdx > subEvents.length) setActiveLocationTabIdx(0);
    if (activeCapacityTabIdx > subEvents.length) setActiveCapacityTabIdx(0);
    if (activeFieldsTabIdx > subEvents.length) setActiveFieldsTabIdx(0);
  }, [subEvents.length, activeLocationTabIdx, activeCapacityTabIdx, activeFieldsTabIdx]);

  // v15: Templates laden wenn Step 6 (Kommunikation, currentStep === 5) erreicht
  // wird. Index hat sich verschoben, weil Team-Anmeldung jetzt NACH Kommunikation
  // kommt (siehe steps-Array).
  // WICHTIG: Dieser useEffect MUSS vor dem early return (if submitted) stehen,
  // da React die gleiche Anzahl Hooks bei jedem Render erwartet (Rules of Hooks).
  React.useEffect(() => {
    if (currentStep === 5 && emailTemplates.length === 0) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (ctx) {
        const svc = new EventService(ctx);
        svc.getAllEmailTemplates().then(setEmailTemplates).catch(() => { /* Templates nicht verfuegbar */ });
      }
    }
  }, [currentStep]);

  // v15.6: Hinweis-Banner für den Hauptevent-Tab in den Steps 3/4/5, wenn
  // subEventsOnlyMode aktiv ist. Der Hauptevent ist dann nicht buchbar — die
  // Einstellungen aus diesem Tab werden zur Laufzeit ignoriert. Der Tab bleibt
  // sichtbar (Konsistenz mit den restlichen Steps), wird aber ausgegraut, und
  // ein gelber Hinweis-Banner sagt explizit, dass der Organizer auf einen
  // [childTermPlural]-Tab wechseln soll, um die Konfiguration zu pflegen.
  //
  // WICHTIG: kein Hook — reines, render-loses Helper. Damit fügt diese
  // Funktion keinen zusätzlichen useState/useEffect/useMemo-Call hinzu und
  // verletzt die Rules of Hooks nicht (Rules of Hooks: alle Hooks vor dem
  // early return weiter unten).
  const renderHauptGreyoutBanner = (): React.ReactElement | null => {
    if (!subEventsOnlyMode) return null;
    const termPlural = (childTermPlural || (isDe ? 'Sub-Events' : 'sub-events')).trim() || (isDe ? 'Sub-Events' : 'sub-events');
    return (
      <div style={{
        display: 'flex', alignItems: 'flex-start', gap: 10,
        padding: '12px 14px', marginBottom: 16,
        background: 'rgba(237,139,0,0.08)',
        border: '1px solid var(--dex-orange, #ed8b00)',
        borderRadius: 'var(--dex-radius, 12px)',
        fontSize: '0.85rem', color: 'var(--dex-gray-700)',
        lineHeight: 1.5,
      }}>
        <Icon iconName="Info" style={{ fontSize: 18, color: 'var(--dex-orange, #ed8b00)', flexShrink: 0, marginTop: 2 }} />
        <div>
          {isDe ? (
            <>
              <strong>Hauptevent ist im aktuellen Modus nicht buchbar</strong> — diese Einstellungen werden nicht verwendet. Wechsle auf einen {termPlural}-Tab, um dort die Konfiguration zu pflegen.
            </>
          ) : (
            <>
              <strong>The main event is not bookable in the current mode</strong> — these settings are not used. Switch to a {termPlural} tab to maintain the configuration there.
            </>
          )}
        </div>
      </div>
    );
  };

  // v15.6: Style-Helfer für den ausgegrauten Hauptevent-Tab-Inhalt. Bei
  // subEventsOnlyMode wird der gesamte Hauptevent-Tab-Inhalt mit Opacity 0.55
  // und pointer-events:none umhüllt, damit der Organizer optisch sofort sieht,
  // dass dieser Bereich aktuell wirkungslos ist. Der Hinweis-Banner steht
  // außerhalb der Hülle, bleibt also klar lesbar.
  const hauptGreyoutWrapperStyle = (): React.CSSProperties => (
    subEventsOnlyMode
      ? { opacity: 0.55, pointerEvents: 'none', userSelect: 'none' }
      : {}
  );

  // v17.3: Unsaved-Changes-Tracking + Navigation-Guard. Wir snapshotten
  // beim Mount die wichtigsten Form-Felder und prüfen bei Bedarf, ob sich
  // etwas geändert hat. Beim Zurück-Klick fragt ein Modal nach.
  const [unsavedConfirmOpen, setUnsavedConfirmOpen] = React.useState<null | { resolve: (_ok: boolean) => void }>(null);
  const initialFormSnapshotRef = React.useRef<string>('');
  const computeFormSnapshot = React.useCallback((): string => {
    return JSON.stringify({
      title, description, location,
      addrStreet, addrHouseNo, addrZip, addrCity,
      organizer, organizerEmails: organizerEmails.join(';'),
      startDate, endDate, registrationDeadline, lastDeregisterDate,
      maxParticipants, waitlistEnabled,
      audience, locationFilter, filterMode,
      contactName, contactEmail, contactInfo,
      eventImageUrl,
      teamRegistrationEnabled, teamSize, askTeamName, teamPartialAllowed,
      teamOpenSlotsVisible, teamJoinRequiresApproval,
      askSalutation, requireSubEventSelection,
      // Custom-Fields nur via Anzahl + Labels — JSON.stringify auf das
      // gesamte Array waere instabil bei id-Aenderungen.
      customFieldsHash: (customFields || []).map(f => `${f.id}:${f.label}:${f.type}:${f.required}`).join('|'),
      agendaLen: (agenda || []).length,
      docsLen: (documents || []).length,
      subEventsLen: (subEvents || []).length,
      outlookBody, outlookHeading, outlookSubheading, outlookSubject,
      disableEmails, disableRegistrationEmail, disableCancellationEmail, disableOutlook,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      emailTemplateOverridesHash: JSON.stringify(emailTemplateOverrides || {}),
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    title, description, location, addrStreet, addrHouseNo, addrZip, addrCity,
    organizer, organizerEmails, startDate, endDate, registrationDeadline, lastDeregisterDate,
    maxParticipants, waitlistEnabled, audience, locationFilter, filterMode,
    contactName, contactEmail, contactInfo, eventImageUrl,
    teamRegistrationEnabled, teamSize, askTeamName, teamPartialAllowed,
    teamOpenSlotsVisible, teamJoinRequiresApproval, askSalutation, requireSubEventSelection,
    customFields, agenda, documents, subEvents,
    outlookBody, outlookHeading, outlookSubheading, outlookSubject, disableEmails, disableRegistrationEmail, disableCancellationEmail, disableOutlook,
    emailTemplateOverrides,
  ]);
  React.useEffect(() => {
    // Initial-Snapshot ein paar Ticks nach dem ersten Render setzen, damit
    // alle initialen useEffect-Loads (z.B. editEvent → State-Hydration) durch sind.
    const t = setTimeout(() => { initialFormSnapshotRef.current = computeFormSnapshot(); }, 100);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editEvent?.id, isEditMode]);
  React.useEffect(() => {
    if (submitted) {
      setNavigationGuard(null);
      return;
    }
    const guard = async (): Promise<boolean> => {
      const isDirty = initialFormSnapshotRef.current !== '' && computeFormSnapshot() !== initialFormSnapshotRef.current;
      if (!isDirty) return true;
      return new Promise<boolean>(resolve => { setUnsavedConfirmOpen({ resolve }); });
    };
    setNavigationGuard(guard);
    return () => setNavigationGuard(null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [submitted, computeFormSnapshot]);

  if (submitted) {
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '64px 32px' }}>
          <h2>{isEditMode ? 'Event erfolgreich aktualisiert!' : 'Event erfolgreich erstellt!'}</h2>
          <p className="mt-8" style={{ color: 'var(--dex-gray-600)' }}>
            &bdquo;{title}&ldquo; wurde {isEditMode ? 'aktualisiert' : 'angelegt'}.
          </p>
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center', flexWrap: 'wrap' }}>
            <button
              className="btn btn-primary"
              onClick={() => {
                // v9.43: Hard-Reload mit Deep-Link statt Soft-Navigate. Grund:
                // nach Event-Erstellung braucht SharePoint ein paar Sekunden, bis
                // die neue Subsite + Listen API-konsistent sind, und der
                // Permission-Cache des Browsers hat für die frische Subsite noch
                // keine gültigen Tokens. Ein Soft-Navigate führt deshalb in einen
                // Render-Crash mit weißem Screen (React #300).
                //
                // Der Hard-Reload räumt den Permission-Cache auf. Damit der User
                // nicht auf der Landing-Seite landet und manuell zur Eventliste
                // klicken muss, hängen wir einen Deep-Link ?action=event-created
                // dran. Der Bootstrap der App liest diesen Parameter, navigiert
                // automatisch zur Eventliste und zeigt eine grüne Erfolgs-Banner-
                // Meldung mit dem Event-Titel.
                const action = isEditMode ? 'event-updated' : 'event-created';
                const targetEventId = selectedEventId || '';
                const url = window.location.pathname + '?action=' + action + (targetEventId ? '&event=' + encodeURIComponent(targetEventId) : '');
                window.location.href = url;
              }}
            >
              Zur Übersicht
            </button>
            <button className="btn btn-secondary" onClick={() => { setSubmitted(false); setTitle(''); }}>Weiteres Event erstellen</button>
          </div>
          <p style={{ marginTop: 20, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
            <em>Hinweis: Beim Klick auf {'„Zur Übersicht“'} wird die Seite einmal neu geladen, damit SharePoint die frisch erstellte Subsite überall sauber einbindet — du landest direkt in der Eventliste mit einer Erfolgs-Meldung.</em>
          </p>
        </div>
      </div>
    );
  }

  // Hilfsfunktion fuer die Vorschau
  const formatPreviewDate = (val: string): string => {
    if (!val) return '--';
    const d = new Date(val);
    return d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
      ' ' + d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  };

  // Vorschau-Sektion rendern
  const renderPreviewSection = (sectionId: string): React.ReactElement | null => {
    switch (sectionId) {
      case 'event':
        return (
          <div className="registration-event" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
            <div className="section-header section-header--red">Selected Event</div>
            <div className="registration-event__card">
              <div className="registration-event__image" style={{
                background: eventImageUrl
                  ? `url(${eventImageUrl}) center/cover`
                  : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
              }}>
                <div className="registration-event__overlay">
                  <h4>{title || 'Event Titel'}</h4>
                  <p>{formatPreviewDate(startDate)} until<br />{formatPreviewDate(endDate)}</p>
                </div>
              </div>
            </div>
          </div>
        );
      case 'personal':
        return (
          <div className="registration-form" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
            <div className="section-header">Personal Information</div>
            <div style={{ padding: '16px 20px' }}>
              <div className="form-group"><label className="form-label"><span className="required">*</span> Salutation</label><select className="form-select" disabled><option>Please select</option></select></div>
              <div className="form-group"><label className="form-label"><span className="required">*</span> First Name</label><input className="form-input" disabled placeholder="First Name" /></div>
              <div className="form-group"><label className="form-label"><span className="required">*</span> Surname</label><input className="form-input" disabled placeholder="Surname" /></div>
              <div className="form-group"><label className="form-label"><span className="required">*</span> E-Mail</label><input className="form-input" disabled placeholder="email@deloitte.de" /></div>
            </div>
          </div>
        );
      case 'specific':
        return (
          <div className="registration-specific" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
            <div className="section-header">Event specific Information</div>
            <div style={{ padding: '16px 20px' }}>
              {customFields.filter(f => f.label).length === 0 ? (
                <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', fontSize: '0.9rem' }}>No additional information required.</p>
              ) : (
                customFields.filter(f => f.label).map(field => (
                  <div className="form-group" key={field.id}>
                    <label className="form-label">{field.required && <span className="required">*</span>}{field.label}</label>
                    {field.type === 'select' && field.multi ? (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 4, padding: 8, border: '1px solid var(--dex-gray-200)', borderRadius: 6, background: '#fff' }}>
                        {field.options.map(o => o.trim()).filter(Boolean).map(opt => (
                          <label key={opt} style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                            <input type="checkbox" disabled />
                            <span>{opt}</span>
                          </label>
                        ))}
                        <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-400)', marginTop: 2 }}>Mehrere Auswahl möglich</span>
                      </div>
                    ) : field.type === 'select' ? (
                      <select className="form-select" disabled><option>Please select</option>{field.options.map(o => o.trim()).filter(Boolean).map(opt => <option key={opt}>{opt}</option>)}</select>
                    ) : field.type === 'checkbox' ? (
                      <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.9rem' }}><input type="checkbox" disabled /> {field.label}</label>
                    ) : (
                      <input className="form-input" disabled placeholder={field.label} type={field.type === 'number' ? 'number' : 'text'} />
                    )}
                  </div>
                ))
              )}
            </div>
          </div>
        );
      case 'actions':
        return (
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 16 }}>
            <button className="btn btn-danger" disabled style={{ opacity: 0.5 }}><Trash2 size={16} /> Delete</button>
            <button className="btn btn-primary" disabled style={{ opacity: 0.5 }}><Send size={16} /> Register</button>
          </div>
        );
      default:
        return null;
    }
  };


  // Hint-Bullets pro Step. Werden ueber das i-Icon in der Progress-Bar
  // (Mouseover) als Tooltip eingeblendet — vorher wurden sie als
  // dauerhafte gruene Hinweis-Box am Anfang jedes Steps angezeigt
  // (renderStepIntro), das war fuer geuebte Organizer zu viel Rauschen.
  const STEP_HINTS_DE: string[][] = [
    [
      'Event-Titel und Beschreibung — werden auf der Eventliste und der Registrierungsseite angezeigt',
      'Event-Bild hochladen (wird oben auf der Detailseite und in den Mails verwendet)',
      'Als Entwurf speichern — taucht dann nur für Admins, Organizer und Test-Team auf',
      'Organizer auswählen — bekommen alle Organizer-Mails (Cancel-/Roommate- etc.) und sehen das Event im Admin Center',
      'Optional: QR-Code-Scanner-User für Check-In am Event-Tag (ohne weitere Bearbeitungs-Rechte)',
      'Standort-Filter und Audience festlegen — wer das Event in der Liste sieht',
    ],
    [
      // v15 Step 2: Sub-Events
      'Optional Sub-Events (Workshops, Sessions, Programmpunkte) anlegen — jeder bekommt im Anmeldeformular eine eigene Anmelde-Checkbox',
      'Bezeichnung wählen: Sub-Events / Workshops / Sessions / Programmpunkte / Event-Sections / eigene Bezeichnung',
      'Anmelde-Modus: zusätzlich zum Hauptevent ODER nur für Sub-Events (Hauptevent-Anmeldung dann ausgeblendet)',
      'Ort, Kapazität, Anmeldefrist und Felder pro Sub-Event werden in den jeweiligen Folge-Schritten (3, 4, 5) per Tab gepflegt',
    ],
    [
      // v15 Step 3: Ort & Programm (mit Tabs pro Sub-Event)
      'Veranstaltungsort und Adresse erfassen — pro Sub-Event optional eigener Ort (per Tab)',
      'Start- und End-Datum (mit Uhrzeit) festlegen',
      'Agenda mehrtägig pflegen (Drag-Reihenfolge pro Tag)',
      'Transferzeiten — Bus / Shuttle / Bahn von/zum Veranstaltungsort',
    ],
    [
      // v15 Step 4: Kapazität & Sichtbarkeit (mit Tabs pro Sub-Event)
      'Maximale Teilnehmerzahl festlegen (oder Unbegrenzt) — pro Sub-Event eigene Kapazität per Tab (Default: vom Hauptevent übernehmen)',
      'Anmeldefrist setzen — pro Sub-Event eigene Deadline möglich (leer = Hauptevent-Deadline gilt)',
      'Optional: Letzte Abmeldemöglichkeit — danach können sich Teilnehmer nicht mehr selbst abmelden',
      'Warteliste aktivieren — voll besetzte Events nehmen weitere Anmeldungen auf, bis ein Platz frei wird',
      'Optional: Geteilte Kapazität — zwei frei benannte Gruppen mit eigener Platzzahl + eigener oder gemeinsamer Warteliste',
    ],
    [
      // v15 Step 5: Felder (mit Tabs pro Sub-Event)
      'Feldtyp wählen: Text, Zahl, Dropdown, Checkbox, Personen-Suche oder Roommate (Doppelzimmer)',
      'Mehrfachauswahl bei Dropdowns (z.B. mehrere Allergien anhaken)',
      'Pflichtfeld setzen (rotes Sternchen, Anmeldung blockiert wenn leer)',
      'Beschreibung pro Feld — landet als „i"-Tooltip neben dem Feld-Label',
      'Sichtbarkeitsbedingung: Feld nur dann anzeigen wenn eine andere Frage einen bestimmten Wert hat (z.B. „Zimmerart nur fragen wenn Hotel = ja")',
      'Pro Sub-Event eigene Felder per Tab (Default: vom Hauptevent übernehmen)',
    ],
    [
      // v15 Step 6: Kommunikation
      'E-Mail-Sprache (DE/EN) für die automatischen Mails an die Teilnehmer wählen',
      'Pro Mail-Template (Anmeldung, Storno, Warteliste, Erinnerung, QR-Code…) den Subject/Heading/Body anpassen — mit Live-Vorschau',
      'Eigenes Logo / Header-Bild für Mail und Outlook-Termin hochladen',
      'Outlook-Termin-Body individuell gestalten (Live-Vorschau zeigt wie das Outlook-Element später aussieht)',
      'Benachrichtigungen optional komplett deaktivieren — z.B. für interne Entwurfs-Events',
      'Pro Sub-Event eigene Mail-Texte + Outlook-Body + Disable-Toggles per Tab',
    ],
    [
      // v15 Step 7: Team-Anmeldung
      'Team-Anmeldung erlauben — ein Teilnehmer kann sich für sich + sein Team gleichzeitig anmelden',
      'Team-Größe festlegen (2-N Personen)',
      'Optional Team-Namen abfragen — z.B. für Quiz- oder Lauf-Teams',
      'Beitritts-Modus: nur komplette Teams ODER auch Teil-Teams erlaubt',
      'Optional: offene Slots öffentlich sichtbar — andere Teilnehmer können beitreten (ggf. mit Lead-Approval)',
    ],
    [
      // v15 Step 8: Dokumente
      'Dokumente hochladen (PDF) — Teilnehmer sehen sie auf MyEvents als Inline-Vorschau oder Download',
    ],
    [
      // v15 Step 9: Fun-Zone
      'Quiz-Fragen für das Event anlegen — Multiple-Choice mit beliebig vielen Antwortoptionen',
      'Pro Frage optional ein Bild hochladen (Logo, Foto-Quiz, etc.)',
      'Mehrere richtige Antworten möglich (Mehrfachauswahl) — werden alle für volle Punktzahl gebraucht',
      'Cluster-Größe steuern: wie viele Fragen pro „Spielblock" angezeigt werden — Teilnehmer kann zwischenspeichern und später weitermachen',
      'Live-Highscore + Statistik im Admin Center sehen (welche Fragen am häufigsten falsch beantwortet werden)',
    ],
  ];
  const STEP_HINTS_EN: string[][] = [
    [
      'Event title and description — shown on the event list and registration page',
      'Upload an event image (used at the top of the detail page and in emails)',
      'Save as draft — only visible to admins, organizers and the test team',
      'Pick the organizers — they receive all organizer emails (cancellation/roommate etc.) and see the event in the admin center',
      'Optional: QR scanner users for check-in on event day (no further editing rights)',
      'Set location filter and audience — who sees the event in the list',
    ],
    [
      // v15 Step 2: Sub-Events
      'Optionally add sub-events (workshops, sessions, program items) — each gets its own registration checkbox in the attendee form',
      'Pick the term: Sub-Events / Workshops / Sessions / Programmpunkte / Event-Sections / custom',
      'Registration mode: in addition to the main event OR sub-events only (main-event registration hidden)',
      'Location, capacity, deadline and fields per sub-event are configured in the following steps (3, 4, 5) via tabs',
    ],
    [
      // v15 Step 3: Location & Programme (with tabs per sub-event)
      'Set event location and address — per sub-event an own location is possible (via tab)',
      'Set start and end date (incl. time)',
      'Maintain the event programme / agenda (multi-day supported, drag-reorder per day)',
      'Transfer times — bus / shuttle / train to and from the venue',
    ],
    [
      // v15 Step 4: Capacity & Visibility (with tabs per sub-event)
      'Set the maximum number of attendees (or Unlimited) — per sub-event own capacity via tab (default: inherit from main event)',
      'Set the registration deadline — per sub-event own deadline possible (empty = main-event deadline applies)',
      'Optional: last self-cancel date — after that self-cancel is locked (late cancel)',
      'Enable waitlist — full events accept new registrations and promote them once a spot frees up',
      'Optional: split capacity — two freely-named groups with own seat count + own or shared waitlist',
    ],
    [
      // v15 Step 5: Fields (with tabs per sub-event)
      'Pick a field type: text, number, dropdown, checkbox, people search or roommate (double room)',
      'Multi-select for dropdowns (e.g. tick multiple allergies)',
      'Mark required (red asterisk, blocks submit when empty)',
      'Description per field — appears as „i" tooltip next to the field label',
      'Visibility condition: only show this field when another question has a specific value (e.g. „Only ask room type if Hotel = yes")',
      'Per sub-event own fields via tab (default: inherit from main event)',
    ],
    [
      // v15 Step 6: Communication
      'Pick the email language (DE/EN) for automatic emails to attendees',
      'Edit subject / heading / body per email template (registration, cancellation, waitlist, reminder, QR code…) — with live preview',
      'Upload a custom logo / header image for the email and Outlook event',
      'Customise the Outlook event body (live preview shows how the Outlook item will appear)',
      'Optionally disable notifications entirely — e.g. for internal draft events',
      'Per sub-event own mail texts + Outlook body + disable toggles via tab',
    ],
    [
      // v15 Step 7: Team Registration
      'Allow team registration — an attendee can register themselves + their team at once',
      'Set team size (2-N people)',
      'Optionally ask for a team name — e.g. quiz or running teams',
      'Join mode: complete teams only OR partial teams allowed',
      'Optional: open slots publicly visible — other attendees can join (with optional lead approval)',
    ],
    [
      // v15 Step 8: Documents
      'Upload documents (PDF) — attendees see them on MyEvents as inline preview or download',
    ],
    [
      // v15 Step 9: Fun-Zone
      'Create quiz questions for the event — multiple choice with any number of answer options',
      'Optionally upload an image per question (logo, photo quiz, etc.)',
      'Multiple correct answers are supported — all of them must be picked for full points',
      'Control cluster size: how many questions per „play block" — attendees can save progress and continue later',
      'See live highscore + statistics in the admin center (which questions are most often answered incorrectly)',
    ],
  ];

  const steps = [
    { label: t('create.step.basics'), icon: '1' },
    // v15.0: Sub-Events kommen vor „Ort & Programm". Hintergrund:
    // Steps 3-5 (Ort, Kapazität, Felder) zeigen pro-Sub-Event-Tabs,
    // damit der Organizer pro Sub-Event eigenes Ort / eigene Kapazität /
    // eigene Felder pflegen kann — dafür müssen die Sub-Events schon
    // angelegt sein, deshalb ist Step 2 der Sub-Events-Step.
    { label: t('create.step.subevents'), icon: '2' },
    { label: t('create.step.datetime'), icon: '3' },
    { label: t('create.step.capacity'), icon: '4' },
    { label: t('create.step.fields'), icon: '5' },
    { label: t('create.step.communication'), icon: '6' },
    // v15.0: Team-Anmeldung kommt jetzt nach Kommunikation (vorher nach
    // Kapazität). Reihenfolge spiegelt den realen Setup-Workflow besser
    // wider: erst die Komm-Texte stehen, dann entscheidet der Organizer
    // ob Team-Anmeldung relevant ist.
    { label: t('create.step.team'), icon: '7' },
    { label: t('create.step.documents'), icon: '8' },
    { label: t('create.step.funzone'), icon: '9' },
  ];

  // Tooltip-State: welcher Step zeigt gerade seinen Hint-Tooltip an?
  const [hintStepIdx, setHintStepIdx] = React.useState<number | null>(null);

  const getStepErrors = (): string[] => {
    const errors: string[] = [];
    switch (currentStep) {
      case 0:
        // v11.8 BUG-FIX: startDate / endDate / endBeforeStart-Check
        // ist von case 1 nach case 0 gewandert, weil die Felder
        // schon in Schritt 1 (Grundlagen) eingegeben werden — wenn
        // die Validierung erst beim Verlassen von Schritt 2 (Ort &
        // Programm) ausgewertet wurde, hatte der Weiter-Button keinen
        // sichtbaren Effekt: die DatePicker-Felder sind in Step 2 gar
        // nicht sichtbar, der User sah nicht warum's nicht weiterging.
        if (!title) errors.push('title');
        if (!organizer) errors.push('organizer');
        if (!startDate) errors.push('startDate');
        if (!endDate) errors.push('endDate');
        if (startDate && endDate && new Date(endDate) <= new Date(startDate)) errors.push('endBeforeStart');
        // v9.14: description ist optional — kein Pflichtfeld mehr
        break;
      case 1:
        // Schritt 2 (Sub-Events) ist ohne Pflicht-Validierung — der
        // Organizer kann den Schritt auch komplett leer lassen.
        // (v15.0: Sub-Events kommen jetzt VOR Ort & Programm.)
        // v18.36: Aber WENN ein Sub-Event Datum hat, darf das Ende nicht vor
        // dem Start liegen — sonst failt der Outlook-Create-Flow mit HTTP 400.
        if (subEvents.some(s => s.title && s.title.trim() && s.startDate && s.endDate && new Date(s.endDate) <= new Date(s.startDate))) {
          errors.push('subEventEndBeforeStart');
        }
        break;
      case 2:
        // Schritt 3 (Ort & Programm) ist ohne Pflicht-Validierung —
        // Adresse / Agenda / Transferzeiten sind alle optional. Datum-
        // Checks laufen in case 0 (Grundlagen).
        break;
      case 3:
        if (registrationDeadline && startDate && new Date(registrationDeadline) > new Date(startDate)) errors.push('deadlineAfterStart');
        if (lastDeregisterDate && startDate && new Date(lastDeregisterDate) > new Date(startDate)) errors.push('deregAfterStart');
        if (!unlimitedParticipants && (maxParticipants === '' || isNaN(Number(maxParticipants)) || Number(maxParticipants) < 0)) errors.push('maxParticipants');
        break;
    }
    return errors;
  };

  const canProceed = (): boolean => getStepErrors().length === 0;

  const fieldHasError = (fieldName: string): boolean => triedNext && getStepErrors().indexOf(fieldName) >= 0;

  const errorBorderStyle = (fieldName: string): React.CSSProperties =>
    fieldHasError(fieldName) ? { borderColor: 'var(--dex-red)', boxShadow: '0 0 0 2px rgba(218,41,28,0.15)' } : {};

  // v7.23: Intro-Hilfsbox pro Wizard-Step. Zeigt eine Liste was der User in
  // diesem Schritt einstellen kann + Verweis aufs Handbuch. DE/EN bilingual.
  // v7.25: pastell-gruener Hintergrund (statt grau), Feature-Items als kompakte
  // Zeilen mit gruenem Check-Icon (statt klassischer Disc-Bullets).
  // v7.26: Items in einem auto-fit-Grid (bis zu 3 Spalten ab Wide-Screen),
  // damit die Box nicht extrem lang wird wenn viele Items drin sind.
  // No-Op seit v7.36: die Hint-Box wird nicht mehr inline am Step-Anfang
  // gerendert. Stattdessen liegen die Hints in STEP_HINTS_DE/EN und werden
  // ueber das i-Icon in der Progress-Bar (Mouseover) angezeigt. Funktion
  // bleibt aus Kompatibilitaetsgruenden mit den 7 bestehenden Call-Sites.
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const renderStepIntro = (_bulletsDe: string[], _bulletsEn: string[]): React.ReactElement | null => null;

  // v15.0: kleine TabStrip-Komponente fuer pro-Sub-Event-Tabs in den
  // Schritten 3 (Ort), 4 (Kapazitaet) und 6 (Felder). Visuell konsistent
  // mit dem Komm-Tab-Pattern in Schritt 7 (gruene Unterstreichung des
  // aktiven Tabs, leichte Hover/Active-Styles). Tab 0 ist immer das
  // Haupt-Event, Tabs 1..N entsprechen `subEvents[0..N-1]`.
  const renderPerEventTabStrip = (
    activeIdx: number,
    onChange: (idx: number) => void,
    mainLabel: string,
    ariaLabel: string,
  ): React.ReactElement | null => {
    if (subEvents.length === 0) return null;
    const tabs: Array<{ label: string; isMain: boolean }> = [
      { label: mainLabel, isMain: true },
      ...subEvents.map(s => ({
        label: (s.title || (isDe ? 'Sub-Event ohne Titel' : 'Untitled sub-event')).trim(),
        isMain: false,
      })),
    ];
    return (
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 16 }}>
        <div
          role="tablist"
          aria-label={ariaLabel}
          style={{
            display: 'flex', flexWrap: 'wrap', gap: 6, flex: 1,
            borderBottom: '1px solid var(--dex-gray-200)',
            paddingBottom: 0,
          }}
        >
          {tabs.map((tab, tabIdx) => {
            const active = tabIdx === activeIdx;
            return (
              <button
                key={tabIdx}
                type="button"
                role="tab"
                aria-selected={active}
                onClick={() => onChange(tabIdx)}
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
                title={tab.label}
              >
                {tab.isMain && (
                  <span style={{ fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: 0.4, color: active ? 'var(--dex-green-dark)' : 'var(--dex-gray-400)' }}>
                    {isDe ? 'Haupt' : 'Main'}
                  </span>
                )}
                <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>{tab.label}</span>
              </button>
            );
          })}
        </div>
      </div>
    );
  };

  // Nutzungsbedingungen-Modal: zeigt sich beim ersten Aufruf der
  // Create-Event-Seite. Nach Akzeptieren wird die Maske weggeklappt; bei
  // Bearbeiten bestehender Events (isEditMode) wird sie nicht angezeigt.
  const showTermsModal = !isEditMode && !tcAccepted;

  // v10.23: Zebra-Hintergrund fuer Schritt-3-Bloecke (Kapazitaet & Sichtbarkeit).
  // Counter wird pro Render zurueckgesetzt; conditional Bloecke verschieben den
  // Index nur wenn sie tatsaechlich rendern, sodass die Alternation auch dann
  // sauber bleibt, wenn Filterverknuepfung oder Sichtbarkeit-pruefen-Buttons
  // ausgeblendet sind. Wird bewusst NUR in Schritt 3 verwendet — andere Steps
  // bekommen das in einer spaeteren Iteration nachgezogen.
  let _zebraS3Idx = 0;
  const zebraS3Bg = (): string => {
    const c = _zebraS3Idx % 2 === 0 ? 'var(--dex-gray-50, #fafafa)' : '#ffffff';
    _zebraS3Idx++;
    return c;
  };

  return (
    <div className="page-container" style={{ maxWidth: 1100, marginLeft: 'auto', marginRight: 'auto' }}>
      {showTermsModal && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
        >
          <div
            className="card"
            style={{
              width: '100%', maxWidth: 720, maxHeight: '90vh', overflow: 'auto',
              padding: 28, borderRadius: 16, background: '#fff',
              boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
            }}
          >
            <h2 style={{ margin: '0 0 4px', fontSize: '1.3rem' }}>
              {isDe
                ? 'Deloitte Event Experience Platform — Nutzungsbedingungen (Deutschland)'
                : 'Deloitte Event Experience Platform — Terms of Use (Germany)'}
            </h2>
            <p style={{ margin: '0 0 16px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
              {isDe ? 'Letzte Überarbeitung: 28.04.2026' : 'Last revised: 28 April 2026'}
            </p>

            {/* Eingeklappte Kurzfassung — die volle Fassung kann der Nutzer
                ueber den Toggle ausklappen. Die Checkbox-Bestaetigung ist
                trotzdem Pflicht (siehe weiter unten). */}
            <div
              style={{
                background: 'var(--dex-gray-50, #f8f9fa)',
                border: '1px solid var(--dex-gray-200, #e5e7eb)',
                borderRadius: 10,
                padding: '12px 14px',
                fontSize: '0.88rem',
                lineHeight: 1.5,
                color: 'var(--dex-gray-700)',
              }}
            >
              {isDe
                ? <>Bitte gehe sorgfältig mit personenbezogenen Daten der Teilnehmer um, sammle nur das absolut Nötige, nutze die Daten ausschließlich für den vereinbarten Event-Zweck und beachte die Datenschutzregeln von Deloitte Deutschland. Volltext über den Button unten einsehen.</>
                : <>Please handle attendees&apos; personal data with care, collect only what is absolutely necessary, use the data exclusively for the agreed event purpose, and follow Deloitte Germany&apos;s data-protection rules. Use the button below to read the full text.</>}
            </div>

            <button
              type="button"
              onClick={() => setTcExpanded(v => !v)}
              style={{
                marginTop: 10,
                background: 'none',
                border: 'none',
                color: 'var(--dex-green-dark, #4a7c1f)',
                fontWeight: 600,
                fontSize: '0.85rem',
                cursor: 'pointer',
                padding: '4px 0',
                display: 'inline-flex',
                alignItems: 'center',
                gap: 6,
              }}
            >
              {tcExpanded
                ? (isDe ? '▲ Vollständige Bedingungen einklappen' : '▲ Hide full terms')
                : (isDe ? '▼ Vollständige Bedingungen anzeigen' : '▼ Show full terms')}
            </button>

            {tcExpanded && (
              <div style={{ fontSize: '0.88rem', lineHeight: 1.55, color: 'var(--dex-gray-800)', marginTop: 12 }}>
                {isDe ? (
                  <>
                    <p>
                      Der Zugang zur Event Experience Platform wird dir als Mitarbeiter von Deloitte Deutschland gewährt,
                      damit du das Teilnehmermanagement für Veranstaltungen, Events, Workshops oder andere Termine
                      organisieren kannst.
                    </p>

                    <p style={{ marginBottom: 6 }}>Die Plattform dient zur Koordination von:</p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Internen Deloitte Veranstaltungen</li>
                      <li>Externen Veranstaltungen, bei denen das Teilnehmermanagement für Deloitte-Mitarbeiter organisiert wird (bspw. Laufveranstaltungen wie B2Run, oder JPMorgan)</li>
                    </ul>

                    <p>
                      <strong>Wichtiger Hinweis:</strong> Externe Nicht-Deloitte-Mitarbeiter werden über dieses Tool
                      nicht koordiniert und erhalten keinen Zugang zur Plattform.
                    </p>

                    <p>Jedes Event, das du erstellst, muss den nachfolgenden Richtlinien folgen.</p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Wichtige Datenschutzhinweise</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Die Teilnahme an Events ist immer freiwillig und darf nicht erzwungen werden.</li>
                      <li>Vermeide die Sammlung personenbezogener Daten so weit wie möglich.</li>
                      <li>Sammle nur die Daten, die du unbedingt benötigst, um den Zweck des Events zu erreichen.</li>
                      <li>Reduziere Freitextfelder auf das absolute Minimum, um individuelle Informationen zur Identifizierung von Personen zu vermeiden.</li>
                      <li>Verwende gesammelte Daten ausschließlich für den definierten und genehmigten Zweck. Falls Abweichungen notwendig sind, wende dich im Voraus an das Datenschutz-Team.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Berechtigungen und Datenzugriff</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>Als Event-Ersteller / Administrator:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Du erhältst Admin-Funktionalitäten für dein spezifisches Event.</li>
                      <li>Du kannst auf die gesamte Teilnehmerliste deines Events zugreifen.</li>
                      <li>Diese Berechtigung gilt ausschließlich für das von dir erstellte Event.</li>
                      <li>Du darfst Teilnehmerinformationen nicht mit anderen teilen oder für andere Zwecke verwenden.</li>
                    </ul>

                    <p style={{ marginBottom: 6 }}><strong>Als Event-Teilnehmer:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Du kannst dich für Events an- oder abmelden.</li>
                      <li>Du erhältst Informationen zum jeweiligen Event.</li>
                      <li>Du hast keinen Zugriff auf die Teilnehmerliste oder Informationen über andere Teilnehmer.</li>
                      <li>Du siehst nur deine eigenen Event-Anmeldungen und -Daten.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Datenschutzbestimmungen im Detail</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>Beschränkung der Sammlung personenbezogener und vertraulicher Daten:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Nur was unbedingt erforderlich ist, um den beabsichtigten Zweck zu erreichen.</li>
                      <li>Offene Fragen auf das Minimum reduzieren (um die Sammlung unnötiger oder nicht autorisierter Daten zu vermeiden).</li>
                    </ul>

                    <p>
                      <strong>Sammle keine sensiblen personenbezogenen Daten</strong> — das heißt: keine Daten bezüglich
                      Rasse oder ethnischer Herkunft, religiöser oder philosophischer Überzeugungen,
                      Gewerkschaftsmitgliedschaft, politischer Meinungen, medizinischer oder gesundheitlicher Zustände
                      oder Informationen über das Sexualleben oder die sexuelle Orientierung einer Person. Falls sensible
                      personenbezogene Daten gesammelt werden müssen, kontaktiere zuerst das Team unter
                      {' '}<a href="mailto:privacy@deloitte.de">privacy@deloitte.de</a>.
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Besondere Bestimmungen für das Teilnehmermanagement</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Teilnehmerdaten dürfen nur für das spezifische Event verwendet werden, für das sie gesammelt wurden.</li>
                      <li>Die Weitergabe von Teilnehmerlisten an Dritte ist untersagt.</li>
                      <li>Teilnehmerdaten anderer Events sind nicht einsehbar.</li>
                      <li>Nach Abschluss des Events sind Teilnehmerdaten gemäß den Deloitte-Richtlinien zu behandeln.</li>
                    </ul>

                    <p>
                      Ermögliche anonyme Antworten, wann immer möglich. Verwende personenbezogene und vertrauliche Daten,
                      die in einem Event gesammelt wurden, nicht für andere Zwecke als den ursprünglich angegebenen.
                      Sprich dich mit dem Datenschutz-Team ab, falls eine andere Nutzung der Daten beabsichtigt ist
                      (du benötigst die vorherige schriftliche Einwilligung der betroffenen Personen / Teilnehmer
                      unter Verwendung einer entsprechenden Vorlage).
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Kontaktinformationen</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Datenschutz-Fragen: <a href="mailto:privacy@deloitte.de">privacy@deloitte.de</a></li>
                    </ul>

                    <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                      Diese Richtlinien gelten für alle Arten von Events, einschließlich Workshops, Seminare,
                      Webinare, Konferenzen und andere Veranstaltungen, deren Teilnehmermanagement für
                      Deloitte-Mitarbeiter über die Event Experience Platform organisiert wird.
                    </p>
                  </>
                ) : (
                  <>
                    <p>
                      Access to the Event Experience Platform is granted to you as an employee of Deloitte Germany so
                      that you can organise attendee management for events, workshops or other appointments.
                    </p>

                    <p style={{ marginBottom: 6 }}>The platform is used to coordinate:</p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Internal Deloitte events</li>
                      <li>External events for which attendee management is organised on behalf of Deloitte employees (e.g. running events such as B2Run or JPMorgan)</li>
                    </ul>

                    <p>
                      <strong>Important note:</strong> External non-Deloitte employees are not coordinated through this
                      tool and will not be granted access to the platform.
                    </p>

                    <p>Every event you create must follow the guidelines below.</p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Key data-protection guidance</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Attending events is always voluntary and must never be enforced.</li>
                      <li>Avoid collecting personal data wherever possible.</li>
                      <li>Only collect data that is strictly necessary to achieve the event&apos;s purpose.</li>
                      <li>Keep free-text fields to an absolute minimum to avoid collecting individual information that could identify people.</li>
                      <li>Use collected data exclusively for the defined and approved purpose. If you need to deviate, contact the data-protection team in advance.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Permissions and data access</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>As event creator / administrator:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>You receive admin functionality for your specific event.</li>
                      <li>You can access the entire attendee list of your event.</li>
                      <li>This permission is limited to the event you created.</li>
                      <li>You may not share attendee information with others or use it for other purposes.</li>
                    </ul>

                    <p style={{ marginBottom: 6 }}><strong>As event attendee:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>You can register for or unregister from events.</li>
                      <li>You receive information about the relevant event.</li>
                      <li>You have no access to the attendee list or information about other attendees.</li>
                      <li>You only see your own event registrations and data.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Data-protection rules in detail</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>Restricting the collection of personal and confidential data:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Only what is strictly necessary to achieve the intended purpose.</li>
                      <li>Reduce open-ended questions to a minimum (to avoid collecting unnecessary or unauthorised data).</li>
                    </ul>

                    <p>
                      <strong>Do not collect sensitive personal data</strong> — that is, no data on race or ethnic origin,
                      religious or philosophical beliefs, trade-union membership, political opinions, medical or health
                      conditions, or information about a person&apos;s sex life or sexual orientation. If sensitive personal
                      data must be collected, contact the team first at
                      {' '}<a href="mailto:privacy@deloitte.de">privacy@deloitte.de</a>.
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Specific rules for attendee management</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Attendee data may only be used for the specific event for which it was collected.</li>
                      <li>Sharing attendee lists with third parties is prohibited.</li>
                      <li>Attendee data of other events is not accessible.</li>
                      <li>After the event, attendee data must be handled in line with Deloitte policy.</li>
                    </ul>

                    <p>
                      Allow anonymous responses wherever possible. Do not use personal or confidential data collected for
                      one event for purposes other than the originally stated one. Coordinate with the data-protection
                      team if you intend to use the data differently (you will need prior written consent from the
                      affected individuals / attendees, using an appropriate template).
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Contact</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Data-protection questions: <a href="mailto:privacy@deloitte.de">privacy@deloitte.de</a></li>
                    </ul>

                    <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                      These guidelines apply to all types of events including workshops, seminars, webinars, conferences
                      and any other events whose attendee management for Deloitte employees is organised through the
                      Event Experience Platform.
                    </p>
                  </>
                )}
              </div>
            )}

            <label
              style={{
                display: 'flex', alignItems: 'flex-start', gap: 10,
                marginTop: 20, padding: 14,
                background: tcCheckbox ? 'rgba(134,188,37,0.08)' : 'var(--dex-gray-50, #f8f9fa)',
                border: `1px solid ${tcCheckbox ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200, #e5e7eb)'}`,
                borderRadius: 10, cursor: 'pointer',
                transition: 'all 0.15s ease',
              }}
            >
              <input
                type="checkbox"
                checked={tcCheckbox}
                onChange={e => setTcCheckbox(e.target.checked)}
                style={{ marginTop: 2, width: 18, height: 18, accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
              />
              <span style={{ fontSize: '0.9rem', lineHeight: 1.4 }}>
                {isDe
                  ? 'Ich habe die Nutzungs- und Datenschutzbedingungen gelesen und akzeptiere sie. Ich bestätige, dass ich mich beim Anlegen und Verwalten dieses Events an die Datenschutzbestimmungen halten werde.'
                  : 'I have read and accept the terms of use and data-protection rules. I confirm that I will follow the data-protection rules when creating and managing this event.'}
              </span>
            </label>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 12, marginTop: 20 }}>
              <button
                type="button"
                className="btn btn-outline"
                onClick={() => goBack()}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                disabled={!tcCheckbox}
                onClick={() => setTcAccepted(true)}
                style={{ opacity: tcCheckbox ? 1 : 0.5, cursor: tcCheckbox ? 'pointer' : 'not-allowed' }}
              >
                <Check size={16} /> {isDe ? 'Akzeptieren & weiter' : 'Accept & continue'}
              </button>
            </div>
          </div>
        </div>
      )}
      <div>
        {/* ===== Step Progress Bar =====
            v14.8: drei Layout-Fixes für das 9-Schritt-Layout:
            (1) Linie endet exakt auf der Mittelachse des ersten/letzten
                Kreises — vorher fix `left/right: 10%`, was zufällig nur
                für n=8 stimmte; jetzt dynamisch über `100 / (steps.length * 2)`.
            (2) Linie etwas dicker (5 statt 3 px) + abgerundet — sonst
                verschwindet sie bei 9 Schritten optisch.
            (3) `alignItems: flex-start` statt `center` — sonst rutschen
                Kreise nach unten, wenn ein Label (z.B. „Kapazität &
                Sichtbarkeit") umbricht.
            Die Linie sitzt bei top=17, height=5 (Mitte bei 19.5 px) —
            das deckt sich exakt mit der Mitte der 40-px-Kreise. */}
        <div style={{ marginBottom: 32 }}>
          {(() => {
            const sidePct = 100 / (steps.length * 2);
            const spanPct = 100 - 2 * sidePct;
            return (
          <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', position: 'relative' }}>
            {/* Verbindungslinie */}
            <div style={{ position: 'absolute', top: 17, left: `${sidePct}%`, right: `${sidePct}%`, height: 5, background: 'var(--dex-gray-200)', borderRadius: 3, zIndex: 0 }} />
            <div style={{ position: 'absolute', top: 17, left: `${sidePct}%`, height: 5, background: 'var(--dex-green)', borderRadius: 3, zIndex: 1, width: `${(currentStep / Math.max(1, steps.length - 1)) * spanPct}%`, transition: 'width 0.4s ease' }} />
            {steps.map((step, idx) => (
              <div
                key={idx}
                onClick={() => { if (idx <= currentStep || canProceed()) setCurrentStep(idx); }}
                style={{
                  display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8,
                  zIndex: 2, cursor: idx <= currentStep ? 'pointer' : 'default',
                  flex: 1,
                }}
              >
                <div style={{
                  width: 40, height: 40, borderRadius: '50%',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  fontWeight: 700, fontSize: '1rem',
                  background: idx <= currentStep ? 'var(--dex-green)' : '#fff',
                  color: idx <= currentStep ? '#fff' : 'var(--dex-gray-400)',
                  border: idx <= currentStep ? '3px solid var(--dex-green)' : '3px solid var(--dex-gray-200)',
                  transition: 'all 0.3s ease',
                  boxShadow: idx === currentStep ? '0 0 0 4px rgba(134,188,37,0.2)' : 'none',
                }}>
                  {idx < currentStep ? '✓' : step.icon}
                </div>
                <span style={{
                  fontSize: '0.75rem', fontWeight: idx === currentStep ? 700 : 500,
                  color: idx <= currentStep ? 'var(--dex-green)' : 'var(--dex-gray-400)',
                  transition: 'color 0.3s ease',
                  textAlign: 'center',
                }}>
                  {step.label}
                </span>
                {/* v9.27/v9.37: i-Icon UNTER dem Step-Label (vorher inline rechts daneben).
                    Hover zeigt die Hints fuer diesen Step.
                    v9.37: Styling identisch zur InfoTooltip-Komponente (serif, 20x20,
                    1.5px-Border) — sonst wirkt das wizard-i im Vergleich klobig. */}
                <span
                  role="button"
                  tabIndex={0}
                  aria-label={isDe ? 'Hinweise zu diesem Schritt' : 'Hints for this step'}
                  onMouseEnter={() => setHintStepIdx(idx)}
                  onMouseLeave={() => setHintStepIdx(null)}
                  onFocus={() => setHintStepIdx(idx)}
                  onBlur={() => setHintStepIdx(null)}
                  onClick={e => { e.stopPropagation(); setHintStepIdx(prev => prev === idx ? null : idx); }}
                  style={{
                    position: 'relative',
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    width: 20, height: 20, borderRadius: '50%',
                    background: hintStepIdx === idx ? 'var(--dex-gray-100, #f0f0f0)' : 'transparent',
                    color: 'var(--dex-gray-700, #555)',
                    border: `1.5px solid ${hintStepIdx === idx ? 'var(--dex-gray-700, #555)' : 'var(--dex-gray-500, #888)'}`,
                    fontSize: '0.7rem', fontWeight: 700, fontFamily: 'serif',
                    cursor: 'help',
                    marginTop: 4,
                    userSelect: 'none',
                    transition: 'background 0.15s, border-color 0.15s',
                  }}
                >
                  i
                  {hintStepIdx === idx && (
                    <div
                      role="tooltip"
                      style={{
                        // v9.40: Styling 1:1 wie InfoTooltip (siehe InfoTooltip.tsx),
                        // damit die zwei Tooltip-Varianten optisch konsistent wirken.
                        position: 'absolute',
                        top: 'calc(100% + 8px)',
                        left: '50%',
                        transform: 'translateX(-50%)',
                        width: 'max-content',
                        maxWidth: 480,
                        minWidth: 280,
                        background: 'rgba(40,40,40,0.96)',
                        color: '#fff',
                        padding: '12px 16px',
                        borderRadius: 8,
                        boxShadow: '0 6px 18px rgba(0,0,0,0.28)',
                        // v15: explizit Sans-Serif — vorher 'inherit', was
                        // den serif-Font des parent „i"-Icons (s. unten
                        // fontFamily:'serif' fuer das i-Glyph) übernommen
                        // hat und den ganzen Tooltip Times-artig erscheinen
                        // ließ. Jetzt 1:1 wie InfoTooltip.
                        fontFamily: 'Aptos, "Open Sans", "Segoe UI", Arial, Helvetica, sans-serif',
                        fontSize: '0.82rem',
                        lineHeight: 1.55,
                        fontWeight: 400,
                        fontStyle: 'normal',
                        textAlign: 'left',
                        whiteSpace: 'normal',
                        zIndex: 1500,
                        pointerEvents: 'none',
                      }}
                    >
                      <div style={{ fontWeight: 700, marginBottom: 8, color: 'rgba(255,255,255,0.92)' }}>
                        {isDe ? 'Was ich hier einstellen kann' : 'What I can configure here'}
                      </div>
                      <ul style={{ margin: 0, paddingLeft: 18 }}>
                        {(isDe ? STEP_HINTS_DE : STEP_HINTS_EN)[idx]?.map((b, bi) => (
                          <li key={bi} style={{ marginBottom: 4 }}>{b}</li>
                        ))}
                      </ul>
                    </div>
                  )}
                </span>
              </div>
            ))}
          </div>
            );
          })()}
        </div>

        {/* ===== Formular ===== */}
        <div>
          <div className="card" style={{ borderRadius: 16, boxShadow: '0 2px 12px rgba(0,0,0,0.06)' }}>
            <div className="creation-form">
              {error && (
                <div style={{ padding: '10px 16px', background: '#fce4ec', color: '#c62828', borderRadius: 8, marginBottom: 16, fontSize: '0.85rem' }}>
                  {error}
                </div>
              )}

              {!isEditMode && currentStep === 0 && (
                <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 16 }}>
                  <button
                    className="btn btn-outline"
                    onClick={() => setShowDemoVariantModal(true)}
                    style={{ fontSize: '0.8rem', padding: '4px 12px' }}
                    title={isDe ? 'Demo-Vorlage auswählen' : 'Choose demo template'}
                  >
                    {isDe ? 'Demo' : 'Demo'}
                  </button>
                </div>
              )}

              {/* ===== Schritt 1: Grundlagen =====
                  v9.32: 1-basierte UI-Nummerierung (in der Logik bleibt
                  currentStep 0-basiert) — siehe CLAUDE.md. */}
              <div style={{ display: currentStep === 0 ? 'block' : 'none' }}>
              <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                {isDe ? 'Schritt 1 — Grundlagen' : 'Step 1 — Basics'}
              </h2>
              <p style={{ margin: '0 0 20px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                {isDe
                  ? 'Hier definierst du das Fundament des Events: Titel, Datum, Beschreibung, Bild und die Personen, die das Event verantworten oder testen.'
                  : 'Here you define the foundation of the event: title, date, description, image and the people who run or test it.'}
              </p>

              {/* v11.57 / v11.63: Hinweisbox bei ausstehendem Outlook-Sync.
                  Sichtbar bei editEvent, wenn OutlookDirty=true auf dem
                  Hauptevent ODER auf mindestens einem Sub-Event gesetzt ist
                  (und Outlook fuer das jeweilige Event nicht deaktiviert
                  wurde). Auf neuen Events nie. */}
              {(() => {
                if (!editEvent) return null;
                // v18.51: im „Nur für Sub-Events"-Modus kein Hauptevent-Dirty-Hinweis.
                const topDirty = editEvent.outlookDirty === true && editEvent.disableOutlook !== true && !editEvent.subEventsOnlyMode;
                const dirtySubs = childEventsOf(editEvent.id).filter(k => k.outlookDirty === true && k.disableOutlook !== true);
                if (!topDirty && dirtySubs.length === 0) return null;
                const subCount = dirtySubs.length;
                let bodyDe: string;
                let bodyEn: string;
                if (topDirty && subCount > 0) {
                  bodyDe = `Outlook-Synchronisation steht aus: für das Hauptevent UND ${subCount} Sub-Event${subCount === 1 ? '' : 's'}. Beim nächsten Speichern kannst du im Dialog pro Termin entscheiden, ob die Teilnehmer eine „Aktualisierter Termin"-Benachrichtigung bekommen.`;
                  bodyEn = `Outlook sync is pending: for the main event AND ${subCount} sub-event${subCount === 1 ? '' : 's'}. On the next save you can decide per invite whether the attendees should receive an “updated meeting” notification.`;
                } else if (topDirty) {
                  bodyDe = 'Outlook-Synchronisation steht aus: für das Hauptevent. Beim nächsten Speichern kannst du im Dialog pro Termin entscheiden, ob die Teilnehmer eine „Aktualisierter Termin"-Benachrichtigung bekommen.';
                  bodyEn = 'Outlook sync is pending: for the main event. On the next save you can decide per invite whether the attendees should receive an “updated meeting” notification.';
                } else {
                  bodyDe = `Outlook-Synchronisation steht aus: für ${subCount} Sub-Event${subCount === 1 ? '' : 's'}. Beim nächsten Speichern kannst du im Dialog pro Termin entscheiden, ob die Teilnehmer eine „Aktualisierter Termin"-Benachrichtigung bekommen.`;
                  bodyEn = `Outlook sync is pending: for ${subCount} sub-event${subCount === 1 ? '' : 's'}. On the next save you can decide per invite whether the attendees should receive an “updated meeting” notification.`;
                }
                return (
                  <div style={{
                    display: 'flex', alignItems: 'flex-start', gap: 12,
                    border: '2px solid #e0a800', background: '#fff8e1',
                    color: '#5a3e00', padding: '12px 16px', fontSize: 14,
                    borderRadius: 8, marginBottom: 16, lineHeight: 1.55,
                  }}>
                    <Icon iconName="Warning" style={{ fontSize: 22, color: '#e0a800', flexShrink: 0, marginTop: 2 }} />
                    <div>
                      <div style={{ fontWeight: 700, marginBottom: 4 }}>
                        {isDe ? 'Outlook-Synchronisation steht aus' : 'Outlook sync is pending'}
                      </div>
                      <div>{isDe ? bodyDe : bodyEn}</div>
                    </div>
                  </div>
                );
              })()}

              {renderStepIntro(
                [
                  '1. Als Entwurf speichern — Event nur für Admins, Organizer und Test-Team sichtbar; optional Aktiv-Ab-Datum für automatisches Go-Live',
                  '2. Event-Titel',
                  '3. Datum (Start &amp; Ende) — füllt die Anmelde- und Storno-Deadlines automatisch vor',
                  '4. Beschreibung (optional, HTML-Editor)',
                  '5. Event-Bild hochladen — oben auf der Detailseite und in den Mails verwendet',
                  '6. Organizer auswählen — bekommen alle Organizer-Mails',
                  '7. Test-Team — sieht das Event schon im Entwurfsmodus',
                  '8. Check-In Team — darf nur das QR-/Check-In-Tool nutzen',
                ],
                [
                  '1. Save as draft — visible only to admins, organizers, and the test team; optional active-from date for automatic go-live',
                  '2. Event title',
                  '3. Date (start &amp; end) — pre-fills the registration and cancellation deadlines',
                  '4. Description (optional, HTML editor)',
                  '5. Upload an event image — shown at the top of the detail page and in emails',
                  '6. Pick organizers — they receive all organizer emails',
                  '7. Test team — can see the event already in draft mode',
                  '8. Check-in team — may only use the QR / check-in tool',
                ]
              )}

              {/* v9.21: Entwurf-Flag als erster Schritt — vor Title.
                  Default ist on, der Organizer kann die Test-Strecke
                  in Ruhe aufbauen, das Test-Team durchspielen lassen,
                  und ohne Aengste sein Event entwickeln. */}
              <div className="form-group" style={{ marginTop: 0, marginBottom: 20, paddingBottom: 20, borderBottom: '1px solid var(--dex-gray-100)', maxWidth: 720 }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', padding: 14, background: isFictive ? 'rgba(237,139,0,0.06)' : 'var(--dex-gray-50, #f8f9fa)', borderRadius: 'var(--dex-radius, 12px)', border: `1px solid ${isFictive ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-200)'}` }}>
                  <StepBadge n={1} />
                  <input
                    type="checkbox"
                    checked={isFictive}
                    onChange={e => setIsFictive(e.target.checked)}
                    style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 3 }}
                  />
                  <span style={{ fontSize: '0.9rem' }}>
                    <strong>{t('create.fictive')}</strong>
                    <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-500)', lineHeight: 1.5, marginTop: 4 }}>
                      {t('create.fictive.hint')}
                    </span>
                  </span>
                </label>
                {/* v9.21: ActiveFrom direkt unter dem Entwurfs-Toggle — wenn
                    der Organizer ein Live-Datum setzt, geht das Event ab dann
                    auch wenn das Entwurf-Haekchen noch on ist. Optional. */}
                <div style={{ marginTop: 12, paddingLeft: 4 }}>
                  <label style={{ display: 'block', fontSize: '0.85rem', color: 'var(--dex-gray-700)', marginBottom: 4, fontWeight: 500 }}>
                    Aktiv ab (optional)
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Was du hier einstellst:</strong> einen Zeitpunkt, ab dem das Event automatisch live geht — auch wenn der <strong>Entwurf-Haken</strong> noch gesetzt ist.<br /><br />
                        <strong>Anzeige in der App:</strong> bis zu diesem Zeitpunkt sehen <strong>nur Admins, Organizer und Test-Team</strong> das Event. Ab dem gesetzten Datum prüft die App bei jedem Aufruf, ob die Zeit schon erreicht ist; falls ja, wird das Event in der allgemeinen Eventliste eingeblendet.<br /><br />
                        <strong>Auswirkung für Teilnehmer:</strong> bis zum Aktiv-ab-Zeitpunkt taucht das Event nicht in der Liste auf, kann nicht aufgerufen werden und bekommt keine Mails. Ab dem Stichtag ist es ganz normal anmeldbar.<br /><br />
                        <strong>Leer lassen</strong> = kein Auto-Go-Live. Du musst dann manuell den Entwurf-Haken entfernen oder im Admin Center auf <strong>Event aktivieren</strong> klicken.
                      </>
                    ) : (
                      <>
                        <strong>What you set here:</strong> a date/time at which the event automatically goes live — even if the <strong>draft toggle</strong> is still on.<br /><br />
                        <strong>Shown in the app:</strong> until that point, only <strong>admins, organizers and the test team</strong> see the event. Once the timestamp is reached, the app reveals the event in the general event list.<br /><br />
                        <strong>Effect for attendees:</strong> until the active-from date the event is not listed, not openable, and produces no mails. After the timestamp it behaves like any other published event.<br /><br />
                        <strong>Leave empty</strong> = no auto-go-live. Publish manually by clearing the draft toggle or by clicking <strong>Activate event</strong> in the admin center.
                      </>
                    )} />
                  </label>
                  <DatePicker
                    selected={activeFrom ? new Date(activeFrom) : null}
                    onChange={(date: Date | null) => setActiveFrom(date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}` : '')}
                    showTimeSelect
                    timeFormat="HH:mm"
                    timeIntervals={15}
                    timeCaption="Uhrzeit"
                    dateFormat="dd.MM.yyyy, HH:mm"
                    locale="de"
                    placeholderText="Datum und Uhrzeit wählen"
                    className="form-input"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    isClearable
                    autoComplete="off"
                  />
                </div>
              </div>

              <div className="form-group" style={{ paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={2} />
                  <span className="required">*</span> {t('create.eventtitle')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> den offiziellen Namen des Events, z.B. <em>Sommerfest 2026</em> oder <em>JPMorgan Lauf 2026</em>.<br /><br />
                      <strong>Anzeige in der App:</strong> der Titel erscheint in der <strong>Eventliste</strong>, im <strong>Header der Detailseite</strong>, in <strong>Meine Events</strong> und im <strong>Admin Center</strong>. Wird auch für die Subsite-URL und für interne Verweise herangezogen.<br /><br />
                      <strong>Automatismen:</strong> der Titel wird 1:1 in den <strong>Betreff aller automatischen Mails</strong> übernommen (Anmelde-Bestätigung, Storno, Warteliste, Nachrück-Mail, QR-Code) sowie in den <strong>Outlook-Termin-Titel</strong> der Teilnehmer.<br /><br />
                      <strong>Empfehlung:</strong> sprechend wählen — der Titel ist das Erste, was Teilnehmer sehen, und identifiziert das Event in ihrem Outlook-Kalender.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the official name of the event, e.g. <em>Summer Party 2026</em> or <em>JPMorgan Run 2026</em>.<br /><br />
                      <strong>Shown in the app:</strong> shown in the <strong>event list</strong>, the <strong>detail page header</strong>, in <strong>My Events</strong> and the <strong>admin center</strong>. Also feeds the subsite URL and internal references.<br /><br />
                      <strong>Automation:</strong> the title is used 1:1 as the <strong>subject of every automated mail</strong> (registration, cancellation, waitlist, promotion, QR-code) and as the <strong>Outlook event title</strong> in attendees{'’'} calendars.<br /><br />
                      <strong>Tip:</strong> pick something descriptive — it is the first thing attendees see and identifies the event in their Outlook calendar.
                    </>
                  )} />
                </label>
                <input className="form-input" value={title} onChange={e => setTitle(e.target.value)} placeholder="z.B. Sommerfest 2026" style={errorBorderStyle('title')} />
                {fieldHasError('title') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
              </div>

              {/* v9.24: Event-Datum direkt nach Title — auto-fillt die Deadlines.
                  Vorher in Step 1, jetzt in Step 0 weil das fundamentale Info ist. */}
              <div className="form-group" style={{ paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={3} />
                  Datum (Start &amp; Ende)
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> Start- und Endzeit des Events (Datum + Uhrzeit, jeweils Berliner Zeit).<br /><br />
                      <strong>Anzeige in der App:</strong> die Werte erscheinen <strong>oben auf der Anmelde-Seite</strong>, in der <strong>Eventliste</strong>, in <strong>Meine Events</strong> und im Admin Center. Die App nutzt sie auch für <strong>Sortierung</strong> (nächste Events zuerst) und für die Logik <strong>Event ist vorbei</strong> (danach werden manche Aktionen wie Anmeldung gesperrt).<br /><br />
                      <strong>Automatismen:</strong> Datum + Uhrzeit landen 1:1 im <strong>Outlook-Termin der Teilnehmer</strong> — der Termin blockt damit den richtigen Slot im Kalender. Außerdem werden die <strong>Anmelde-Deadline</strong> (7 Tage vor Start) und die <strong>Letzte Abmeldemöglichkeit</strong> (3 Tage vor Start) automatisch vorgeschlagen — beide kannst du im Schritt <strong>Kapazität & Sichtbarkeit</strong> jederzeit überschreiben.<br /><br />
                      <strong>Auswirkung für Teilnehmer:</strong> sehen das Datum sofort in der Liste, bekommen es in jeder Bestätigungs-Mail und als Outlook-Eintrag.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> start and end (date + time, Berlin local time).<br /><br />
                      <strong>Shown in the app:</strong> shown <strong>at the top of the registration page</strong>, in the <strong>event list</strong>, in <strong>My Events</strong> and the admin center. Also drives <strong>sort order</strong> (upcoming first) and the <strong>event is over</strong> logic (some actions get locked after that).<br /><br />
                      <strong>Automation:</strong> date + time go 1:1 into the <strong>attendee{'’'}s Outlook event</strong> so the right slot gets blocked. The <strong>registration deadline</strong> (7 days before start) and <strong>last cancellation date</strong> (3 days before start) are auto-suggested — both can be overridden in step <strong>Capacity & Visibility</strong>.<br /><br />
                      <strong>Effect for attendees:</strong> they see the date in the list, in every confirmation email and as an Outlook entry.
                    </>
                  )} />
                </label>
              <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                <div className="form-group" style={{ marginBottom: 0 }}>
                  <label className="form-label">
                    <span className="required">*</span> {t('create.startdate')}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Startzeitpunkt</strong> — Datum + Uhrzeit, ab wann das Event läuft. Wandert 1:1 in den <strong>Outlook-Termin</strong> jedes Teilnehmers (blockt den Kalender-Slot) und in <strong>jede Bestätigungs-Mail</strong>. Bestimmt außerdem die Standard-Vorschläge für <strong>Anmelde-Deadline</strong> (7 Tage vor Start) und <strong>Letzte Abmeldemöglichkeit</strong> (3 Tage vor Start).
                      </>
                    ) : (
                      <>
                        <strong>Start time</strong> — date + time when the event begins. Goes 1:1 into every attendee Outlook event (blocks the calendar slot) and into every <strong>confirmation email</strong>. Also drives the auto-suggestions for <strong>registration deadline</strong> (7 days before start) and <strong>last cancellation date</strong> (3 days before start).
                      </>
                    )} />
                  </label>
                  <DatePicker
                    selected={startDate ? new Date(startDate) : null}
                    onChange={(date: Date | null) => setStartDate(date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}` : '')}
                    showTimeSelect
                    timeFormat="HH:mm"
                    timeIntervals={15}
                    timeCaption="Uhrzeit"
                    dateFormat="dd.MM.yyyy, HH:mm"
                    locale="de"
                    placeholderText="Datum und Uhrzeit wählen"
                    className="form-input"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    isClearable
                    autoComplete="off"
                  />
                  {fieldHasError('startDate') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
                </div>
                <div className="form-group" style={{ marginBottom: 0 }}>
                  <label className="form-label">
                    <span className="required">*</span> {t('create.enddate')}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Endzeitpunkt</strong> — Datum + Uhrzeit, wann das Event vorbei ist. Wandert 1:1 in den <strong>Outlook-Termin</strong> der Teilnehmer (sonst läuft der Termin endlos). Wichtig auch für interne Logik: nach diesem Zeitpunkt zählt das Event als <strong>vorbei</strong> — Anmeldungen werden gesperrt, das Event rutscht in der Liste nach unten und manche automatische Benachrichtigungen (z.B. Late-Cancel-Hinweise) reagieren darauf.
                      </>
                    ) : (
                      <>
                        <strong>End time</strong> — date + time when the event finishes. Goes 1:1 into the attendee Outlook event (otherwise it would never end). Also feeds internal logic: past this point the event counts as <strong>over</strong> — registrations get locked, it drops down the list, and some automated notifications (e.g. late-cancel alerts) react to it.
                      </>
                    )} />
                  </label>
                  <DatePicker
                    selected={endDate ? new Date(endDate) : null}
                    onChange={(date: Date | null) => setEndDate(date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}` : '')}
                    showTimeSelect
                    timeFormat="HH:mm"
                    timeIntervals={15}
                    timeCaption="Uhrzeit"
                    dateFormat="dd.MM.yyyy, HH:mm"
                    locale="de"
                    placeholderText="Datum und Uhrzeit wählen"
                    className="form-input"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    minDate={startDate ? new Date(startDate) : undefined}
                    isClearable
                    autoComplete="off"
                  />
                  {fieldHasError('endDate') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
                </div>
              </div>
              {fieldHasError('endBeforeStart') && <p style={{ color: 'var(--dex-red)', fontSize: '0.8rem', marginTop: -4, marginBottom: 8 }}>{t('create.error.endBeforeStart')}</p>}
              <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginTop: 8, marginBottom: 0 }}>
                Die Uhrzeit wird für den Outlook-Kalendereintrag der Teilnehmer verwendet.
              </p>
              </div>

              <div className="form-group" style={{ paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={4} />
                  {t('create.description')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> die <strong>Hauptbeschreibung</strong> des Events — was findet statt, an wen richtet es sich, was sollten Teilnehmer wissen.<br /><br />
                      <strong>Anzeige in der App:</strong> wird oben auf der <strong>Anmelde-Seite</strong> und unter <strong>Meine Events</strong> angezeigt. Du kannst <strong>HTML-Formatierung</strong> nutzen (Fettdruck, Listen, Links, Bilder) — die App rendert sie 1:1.<br /><br />
                      <strong>Automatismen:</strong> wenn der Outlook-Termin-Body leer ist, wird die Beschreibung als <strong>Fallback in den Outlook-Kalendereintrag</strong> übernommen.<br /><br />
                      <strong>Auswirkung für Teilnehmer:</strong> erste inhaltliche Information vor der Anmeldung. Optional — leer ist erlaubt, aber bei externen Empfängerkreisen empfehlenswert.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>main event description</strong> — what is happening, who it is for, what attendees should know.<br /><br />
                      <strong>Shown in the app:</strong> shown at the top of the <strong>registration page</strong> and under <strong>My Events</strong>. You can use <strong>HTML formatting</strong> (bold, lists, links, images) — the app renders it 1:1.<br /><br />
                      <strong>Automation:</strong> if the Outlook event body is empty, the description is used as a <strong>fallback in the Outlook calendar entry</strong>.<br /><br />
                      <strong>Effect for attendees:</strong> first piece of substantive information before registering. Optional — leaving it blank is fine, but recommended for broader audiences.
                    </>
                  )} />
                </label>
                {/* v9.39: Beschreibung als HTML-Editor (vorher plain textarea).
                    Live-Vorschau im HtmlEditorModal — wird auf der Anmelde-Seite
                    1:1 als HTML gerendert. */}
                <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    onClick={() => { setHtmlEditorMode('description'); setHtmlEditorOpen(true); }}
                    style={{ fontSize: '0.85rem' }}
                  >
                    {isDe ? 'Bearbeiten & Vorschau' : 'Edit & Preview'}
                  </button>
                  <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', flex: 1, minWidth: 200 }}>
                    {description
                      ? `${description.replace(/<[^>]+>/g, '').substring(0, 120)}${description.length > 120 ? '…' : ''}`
                      : (isDe ? 'Keine Beschreibung gesetzt — klicke „Bearbeiten" zum Hinzufügen.' : 'No description set — click „Edit" to add one.')}
                  </span>
                </div>
                {/* v18.73: Hinweis, wenn Name/Datum/Ort des Events redundant in
                    der Beschreibung stehen — die werden bereits separat auf der
                    Anmelde-Seite angezeigt. Mit klickbarem Beispieltext. */}
                {(() => {
                  const plain = (description || '')
                    .replace(/<[^>]+>/g, ' ')
                    .replace(/&nbsp;/gi, ' ')
                    .replace(/\s+/g, ' ')
                    .toLowerCase();
                  if (plain.trim().length < 8) return null;
                  const hits: string[] = [];
                  const tl = title.trim().toLowerCase();
                  if (tl.length >= 4 && plain.indexOf(tl) >= 0) hits.push(isDe ? 'der Event-Name' : 'the event name');
                  const locl = location.trim().toLowerCase();
                  if (locl.length >= 4 && plain.indexOf(locl) >= 0) hits.push(isDe ? 'der Ort' : 'the location');
                  if (startDate) {
                    const d = new Date(startDate);
                    if (!isNaN(d.getTime())) {
                      const dd = String(d.getDate()).padStart(2, '0');
                      const mm = String(d.getMonth() + 1).padStart(2, '0');
                      const yyyy = String(d.getFullYear());
                      const monthsDe = ['januar', 'februar', 'märz', 'april', 'mai', 'juni', 'juli', 'august', 'september', 'oktober', 'november', 'dezember'];
                      const monthsEn = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
                      const mn = (isDe ? monthsDe : monthsEn)[d.getMonth()];
                      const pats = [`${dd}.${mm}.${yyyy}`, `${dd}.${mm}.`, `${d.getDate()}. ${mn}`, `${d.getDate()}.${mn}`, `${d.getDate()} ${mn}`];
                      if (pats.some(p => plain.indexOf(p) >= 0)) hits.push(isDe ? 'das Datum' : 'the date');
                    }
                  }
                  if (hits.length === 0) return null;
                  const joined = hits.length === 1
                    ? hits[0]
                    : hits.slice(0, -1).join(', ') + (isDe ? ' und ' : ' and ') + hits[hits.length - 1];
                  return (
                    <div style={{ marginTop: 12, padding: '12px 14px', background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 8, fontSize: '0.82rem', color: 'var(--dex-gray-800)', lineHeight: 1.5 }}>
                      <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8 }}>
                        <Icon iconName="Info" style={{ fontSize: 15, color: 'var(--dex-orange, #ed8b00)', marginTop: 2 }} />
                        <div style={{ flex: 1 }}>
                          {isDe
                            ? <>In der Beschreibung steht offenbar <strong>{joined}</strong>. <strong>Name, Datum und Ort</strong> des Events werden bereits <strong>separat</strong> auf der Anmelde-Seite angezeigt — du musst sie hier nicht wiederholen. Nutze die Beschreibung lieber für einen einladenden, inhaltlichen Text. Über <strong>„Bearbeiten &amp; Vorschau“</strong> kannst du einen Beispieltext übernehmen.</>
                            : <>Your description appears to contain <strong>{joined}</strong>. The event&rsquo;s <strong>name, date and location</strong> are already shown <strong>separately</strong> on the registration page — no need to repeat them here. Use the description for an inviting, substantive text instead. Via <strong>“Edit &amp; Preview”</strong> you can load an example text.</>}
                        </div>
                      </div>
                    </div>
                  );
                })()}
              </div>

              <div className="form-group">
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={5} />
                  {t('create.eventimage')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> ein <strong>Hauptbild fürs Event</strong> (Foto vom Veranstaltungsort, Eventlogo, Stimmungsbild). Wird zentral als Item-Attachment am Event gespeichert.<br /><br />
                      <strong>Anzeige in der App:</strong> erscheint <strong>oben auf der Anmelde-Seite</strong>, <strong>als Kachel-Hintergrund</strong> in der Eventliste und <strong>in Meine Events</strong>. Macht Events visuell unterscheidbar in einer langen Liste.<br /><br />
                      <strong>Empfehlung:</strong> Querformat (z.B. 16:9), gute Auflösung (mind. 1200px breit). Hochformat funktioniert auch — die App erkennt das automatisch und legt das Bild dann links neben den Detail-Rows ab statt als Banner.<br /><br />
                      <strong>Auswirkung für Teilnehmer:</strong> visueller Wiedererkennungswert in ihrer Eventliste und auf der Anmelde-Seite — beeinflusst nicht die Mails (dort gibt es ein separates Logo).
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>main event image</strong> (venue photo, event logo, mood shot). Stored centrally as an item attachment.<br /><br />
                      <strong>Shown in the app:</strong> shown <strong>at the top of the registration page</strong>, <strong>as the tile background</strong> in the event list and <strong>under My Events</strong>. Makes events visually distinguishable in a long list.<br /><br />
                      <strong>Tip:</strong> landscape (e.g. 16:9), high resolution (at least 1200px wide). Portrait works too — the app detects orientation and places the image to the left of the detail rows instead of as a banner.<br /><br />
                      <strong>Effect for attendees:</strong> visual recognition in their list and on the registration page — does not affect emails (there is a separate logo for those).
                    </>
                  )} />
                </label>
                {imagePreview && (
                  <div style={{ position: 'relative', marginBottom: 8, display: 'block', width: 'fit-content', maxWidth: '100%' }}>
                    <img
                      src={imagePreview}
                      alt="Vorschau"
                      style={{
                        // Korrekte Auflösung beibehalten, nur in der Hoehe begrenzen + max-Breite zur Sicherheit
                        display: 'block',
                        maxHeight: 220,
                        maxWidth: '100%',
                        width: 'auto',
                        height: 'auto',
                        objectFit: 'contain',
                        borderRadius: 'var(--dex-radius)',
                        background: 'var(--dex-gray-100)',
                      }}
                    />
                    <button
                      type="button"
                      onClick={() => { setImageFile(null); setImagePreview(''); setEventImageUrl(''); }}
                      style={{
                        position: 'absolute', top: 8, right: 8, background: 'rgba(0,0,0,0.6)',
                        color: '#fff', border: 'none', borderRadius: '50%', width: 28, height: 28,
                        cursor: 'pointer', fontSize: '1rem', display: 'flex', alignItems: 'center', justifyContent: 'center',
                      }}
                    >
                      <X size={14} />
                    </button>
                  </div>
                )}
                <label style={{
                  display: 'inline-flex', alignItems: 'center', gap: 8,
                  padding: '8px 16px', borderRadius: 'var(--dex-radius)',
                  border: '2px dashed var(--dex-gray-300)', cursor: 'pointer',
                  fontSize: '0.85rem', color: 'var(--dex-gray-600)',
                  transition: 'border-color 0.2s, background 0.2s',
                }}>
                  <Plus size={16} />
                  {imageFile ? imageFile.name : 'Bild auswählen'}
                  <input
                    type="file"
                    accept="image/*"
                    style={{ display: 'none' }}
                    onChange={e => {
                      const file = e.target.files && e.target.files[0];
                      if (file) {
                        setImageUploadError('');
                        setImageFile(file);
                        const reader = new FileReader();
                        reader.onload = ev => setImagePreview(ev.target?.result as string || '');
                        reader.readAsDataURL(file);
                      }
                    }}
                  />
                </label>
                {imageUploadError && (
                  <p style={{ color: 'var(--dex-red, #c00)', fontSize: '0.8rem', marginTop: 4 }}>{imageUploadError}</p>
                )}
              </div>


              <div className="form-group" style={{ position: 'relative', paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={6} />
                  <span className="required">*</span> {t('create.organizer')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> die <strong>verantwortlichen Personen</strong> für dieses Event — beliebige Deloitte-User per Graph-Suche. Du selbst bist standardmäßig vorbefüllt, kannst aber Co-Organizer hinzunehmen oder dich selbst rauslöschen.<br /><br />
                      <strong>Anzeige in der App:</strong> Organizer dürfen das Event <strong>bearbeiten, deaktivieren, löschen</strong>, die <strong>Teilnehmerliste</strong> einsehen, <strong>QR-Codes versenden</strong> und <strong>Massenmails</strong> verschicken. Sie tauchen auf der Anmelde-Seite und in Meine Events als <strong>Ansprechpartner</strong> mit Foto + Mail-Adresse auf.<br /><br />
                      <strong>Automatismen:</strong> Organizer bekommen je nach Einstellung in <strong>Schritt 7 (Kommunikation)</strong> eine BCC-Kopie der Anmelde-/Abmelde-Mails. Late-Cancel- und Roommate-Mails gehen ebenfalls an alle Organizer. Wenn ein Teilnehmer die Outlook-Einladung weiterleitet und der Empfänger nicht angemeldet ist, bekommen die Organizer eine FYI-Mail.<br /><br />
                      <strong>Reihenfolge zählt:</strong> der erste Organizer ist der Haupt-Organizer und wird in Mails als Absender-Name verwendet.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>responsible people</strong> for this event — any Deloitte user via Graph search. You are pre-filled by default, but you can add co-organizers or remove yourself.<br /><br />
                      <strong>Shown in the app:</strong> organizers can <strong>edit, deactivate, delete</strong> the event, see the <strong>attendee list</strong>, <strong>send QR codes</strong> and <strong>mass emails</strong>. They appear on the registration page and in My Events as <strong>contacts</strong> with photo + email.<br /><br />
                      <strong>Automation:</strong> depending on the setting in <strong>step 5 (Communication)</strong>, organizers receive BCC copies of registration / cancellation mails. Late-cancel and roommate mails go to all organizers. If an attendee forwards the Outlook invite to someone unregistered, organizers receive an FYI mail.<br /><br />
                      <strong>Order matters:</strong> the first organizer is the main organizer and is used as the sender name in mails.
                    </>
                  )} />
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ marginLeft: 'auto', fontSize: '0.75rem', padding: '4px 10px', whiteSpace: 'nowrap' }}
                    onClick={() => setBulkOrganizerOpen(true)}
                  >
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                      <Users size={12} /> Massenimport
                    </span>
                  </button>
                </label>
                {/* Mismatch-Warning: bei Legacy-Korruption aus v10.0–v10.2-Closure-Bug
                    haben Events mehr Namen als Emails (oder umgekehrt). Auto-Heal padded
                    inzwischen statt zu truncaten — die Chips zeigen alle Namen, aber bei
                    welchen die Email fehlt, kann der User es hier nachpflegen. */}
                {(() => {
                  const orgList = organizer.split(';').map(s => s.trim()).filter(Boolean);
                  const missingEmailCount = orgList.reduce((acc, _, i) => {
                    return acc + ((organizerEmails[i] || '').trim() === '' ? 1 : 0);
                  }, 0);
                  if (missingEmailCount === 0) return null;
                  return (
                    <div
                      style={{
                        background: '#fff8e1',
                        border: '1px solid #f0c419',
                        borderRadius: 'var(--dex-radius)',
                        padding: '10px 12px',
                        marginBottom: 10,
                        fontSize: '0.82rem',
                        lineHeight: 1.5,
                        color: 'var(--dex-gray-800)',
                      }}
                    >
                      <strong>⚠️ {missingEmailCount} Organizer ohne hinterlegte Email-Adresse.</strong> Bei diesen Personen fehlen die Mails fürs <strong>BCC</strong> der Anmelde-/Abmelde-Mails, die <strong>Outlook-Einladung</strong> und die <strong>Decline-/Forward-Notifications</strong>. Bitte entferne die betroffenen Chips (X) und füge sie über den Picker oder Massenimport neu ein. <em>(Ursache: Legacy-Daten aus einer früheren App-Version — wird beim nächsten Save geheilt.)</em>
                    </div>
                  );
                })()}
                {/* Organizer-Chips (immer sichtbar wenn 1+ Organizer) */}
                {(() => {
                  const orgList = organizer.split(';').map(s => s.trim()).filter(Boolean);
                  if (orgList.length === 0) return null;
                  const move = (idx: number, dir: -1 | 1): void => {
                    const nextNames = [...orgList];
                    const target = idx + dir;
                    if (target < 0 || target >= nextNames.length) return;
                    [nextNames[idx], nextNames[target]] = [nextNames[target], nextNames[idx]];
                    setOrganizer(nextNames.join('; '));
                    setOrganizerEmails(prev => {
                      if (idx >= prev.length || target >= prev.length) return prev;
                      const nextEmails = [...prev];
                      [nextEmails[idx], nextEmails[target]] = [nextEmails[target], nextEmails[idx]];
                      return nextEmails;
                    });
                  };
                  const remove = (idx: number): void => {
                    // Email-aware Remove: bei State-Korruption (z.B. Events aus
                    // v10.0–v10.2 wo der Closure-Bug emails ohne Namen schrieb)
                    // kann die gleiche Email mehrfach in organizerEmails stehen,
                    // während orgList nur einen Eintrag hat. Ein reiner Index-Filter
                    // würde dann nur EINEN Email-Eintrag killen, der Rest bleibt
                    // drin → die Person bleibt für den Picker „bekannt" und ist
                    // ausgegraut. Deshalb: emailToRemove ermitteln und ALLE
                    // Vorkommen aus organizerEmails entfernen.
                    const emailToRemove = (organizerEmails[idx] || '').toLowerCase();
                    const nextNames = orgList.filter((_, i) => i !== idx);
                    setOrganizer(nextNames.join('; '));
                    setOrganizerEmails(prev => {
                      if (!emailToRemove) return prev.filter((_, i) => i !== idx);
                      return prev.filter(e => (e || '').toLowerCase() !== emailToRemove);
                    });
                  };
                  return (
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 10 }}>
                      {orgList.map((name, i) => {
                        const email = organizerEmails[i] || '';
                        return (
                        <span
                          key={`${name}-${i}`}
                          style={{
                            display: 'inline-flex', alignItems: 'center', gap: 6,
                            padding: '3px 6px 3px 4px',
                            background: 'var(--dex-green)', color: '#fff',
                            borderRadius: 999, fontSize: '0.85rem', fontWeight: 500,
                          }}
                        >
                          {email ? (
                            <img
                              src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=S`}
                              alt={name}
                              onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
                              style={{ width: 24, height: 24, borderRadius: '50%', objectFit: 'cover', background: 'rgba(255,255,255,0.25)' }}
                            />
                          ) : null}
                          <span>{name}</span>
                          {orgList.length > 1 && i > 0 && (
                            <button
                              type="button"
                              onClick={() => move(i, -1)}
                              style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.75rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
                              title="Nach vorne"
                            >◀</button>
                          )}
                          {orgList.length > 1 && i < orgList.length - 1 && (
                            <button
                              type="button"
                              onClick={() => move(i, 1)}
                              style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.75rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
                              title="Nach hinten"
                            >▶</button>
                          )}
                          <button
                            type="button"
                            onClick={() => remove(i)}
                            style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.9rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
                            title="Entfernen"
                          >×</button>
                        </span>
                      );
                      })}
                    </div>
                  );
                })()}
                <input
                  className="form-input"
                  value={organizerSearch}
                  onChange={e => {
                    const val = e.target.value;
                    setOrganizerSearch(val);
                    if (organizerTimerRef.current) clearTimeout(organizerTimerRef.current);
                    const q = val.trim();
                    if (!q) { setOrganizerResults([]); return; }
                    // v9.20: Graph-Search statt Role-Filter — jeder Deloitte-User
                    // kann als Organizer hinzugefuegt werden. Damit gibt es nur
                    // einen Picker (kein extra Co-Organizer); der erste in der
                    // Liste ist der "Hauptorganizer", weitere sind gleichwertig.
                    organizerTimerRef.current = setTimeout(async () => {
                      try {
                        const results = await searchUsers(q, organizerIncludeIntl);
                        setOrganizerResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setOrganizerResults([]); }
                    }, 350);
                  }}
                  onBlur={() => {
                    setTimeout(() => { setOrganizerSearch(''); setOrganizerResults([]); }, 150);
                  }}
                  placeholder={t('create.organizer.placeholder')}
                  style={errorBorderStyle('organizer')}
                />
                <InternationalSearchToggle
                  checked={organizerIncludeIntl}
                  onChange={async next => {
                    setOrganizerIncludeIntl(next);
                    const q = organizerSearch.trim();
                    if (q.length >= 1) {
                      try {
                        const results = await searchUsers(q, next);
                        setOrganizerResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setOrganizerResults([]); }
                    }
                  }}
                  isDe={isDe}
                />
                {isSearchingOrganizer && (
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>Suche...</div>
                )}
                {organizerResults.length > 0 && (
                  <div style={{
                    position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
                    background: '#fff', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                    maxHeight: 280, overflowY: 'auto',
                  }}>
                    {organizerResults.map(u => {
                      const alreadyAdded = organizerEmails.indexOf(u.email) >= 0;
                      return (
                        <div
                          key={u.email}
                          style={{
                            padding: '8px 12px', cursor: alreadyAdded ? 'not-allowed' : 'pointer', fontSize: '0.85rem',
                            borderBottom: '1px solid var(--dex-gray-100)',
                            opacity: alreadyAdded ? 0.45 : 1,
                            display: 'flex', alignItems: 'center', gap: 10,
                          }}
                          onMouseDown={() => {
                            if (alreadyAdded) return;
                            const existing = organizer.split(';').map(s => s.trim()).filter(Boolean);
                            const nextNames = [...existing, u.displayName];
                            setOrganizer(nextNames.join('; '));
                            setOrganizerEmails(prev => [...prev, u.email]);
                            setOrganizerSearch('');
                            setOrganizerResults([]);
                          }}
                        >
                          <img
                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(u.email)}&size=S`}
                            alt={u.displayName}
                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                            style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                          />
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontWeight: 600 }}>{u.displayName}</div>
                            <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                              {u.email}{u.location ? ` · ${u.location}` : ''}
                            </div>
                          </div>
                          {alreadyAdded && <span style={{ color: 'var(--dex-green)', fontSize: '0.85rem', flexShrink: 0 }}>✓</span>}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* v18.9: Organizer-Anzeige ausblenden. Rein visuell — die
                  Organizer behalten alle Rechte + Mail-Benachrichtigungen,
                  werden aber auf der Anmelde-Seite und in „Meine Events"
                  nicht als Ansprechpartner-Chips gezeigt. */}
              <div className="form-group" style={{ paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={hideOrganizer}
                    onChange={e => setHideOrganizer(e.target.checked)}
                    style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 2 }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Organizer ausblenden' : 'Hide organizer'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> ob die <strong>Organizer-Kacheln</strong> (Name + Foto + Mail) auf der <strong>Anmelde-Seite</strong> und in <strong>&bdquo;Meine Events&ldquo;</strong> angezeigt werden.<br /><br />
                          <strong>Anzeige in der App:</strong> wenn aktiviert, sehen Teilnehmer <strong>keine Organizer-Ansprechpartner</strong> mehr bei diesem Event. Der optionale Ansprechpartner unten bleibt davon unberührt.<br /><br />
                          <strong>Wichtig:</strong> das ist rein optisch — die Organizer behalten alle <strong>Rechte</strong> (bearbeiten, Teilnehmer verwalten) und ihre <strong>Mail-Benachrichtigungen</strong>.
                        </>
                      : <>
                          <strong>What this controls:</strong> whether the <strong>organizer chips</strong> (name + photo + email) are shown on the <strong>registration page</strong> and in <strong>&bdquo;My Events&ldquo;</strong>.<br /><br />
                          <strong>Where you see it:</strong> when enabled, attendees no longer see organizer contacts for this event. The optional contact person below is not affected.<br /><br />
                          <strong>Note:</strong> this is purely visual — organizers keep all <strong>permissions</strong> (edit, manage attendees) and their <strong>email notifications</strong>.
                        </>
                    } />
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? 'Default: aus — Organizer werden als Ansprechpartner angezeigt.'
                        : 'Default: off — organizers are shown as contacts.'}
                    </span>
                  </span>
                </label>
              </div>

              {/* v10.16: Optionaler Ansprechpartner. Reines Anzeige-Feld
                  (kein Login, keine SP-Permissions) — z.B. die Person vor Ort
                  oder eine Hotline-Mail die Teilnehmer bei Fragen anschreiben
                  sollen. Wird auf Register-/MyEvents-Page zusätzlich zu den
                  Organizern gezeigt. Alles drei optional, Freitext. */}
              <div className="form-group" style={{ paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <span style={{
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    width: 22, height: 22, borderRadius: '50%',
                    background: 'var(--dex-gray-300)', color: '#fff',
                    fontSize: '0.75rem', fontWeight: 700,
                  }}>+</span>
                  {isDe ? 'Ansprechpartner (optional)' : 'Contact person (optional)'}
                </label>
                <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', margin: '0 0 10px', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Zusätzliche Kontaktperson für Rückfragen — z.B. Person vor Ort, externe Agentur, Hotline-Mailbox. Erscheint auf der Anmelde-Seite und in „Meine Events" zusätzlich zu den Organizern. Hat KEINE App-Berechtigung, ist nur ein Anzeige-Feld.'
                    : 'Additional contact for questions — e.g. on-site contact, external agency, hotline mailbox. Appears on the registration page and in „My Events" in addition to the organizers. Has NO app permissions, display-only.'}
                </p>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 10 }}>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                      {isDe ? 'Name' : 'Name'}
                    </label>
                    <input
                      type="text"
                      className="form-input"
                      value={contactName}
                      onChange={e => setContactName(e.target.value)}
                      placeholder={isDe ? 'z.B. Anna Schmitt' : 'e.g. Anna Schmitt'}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                      {isDe ? 'E-Mail' : 'Email'}
                    </label>
                    <input
                      type="email"
                      className="form-input"
                      value={contactEmail}
                      onChange={e => setContactEmail(e.target.value)}
                      placeholder={isDe ? 'z.B. event-helpdesk@example.de' : 'e.g. event-helpdesk@example.com'}
                    />
                  </div>
                </div>
                <div>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                    {isDe ? 'Zusatz-Info / Erreichbarkeit (Freitext)' : 'Additional info / availability (free text)'}
                  </label>
                  <textarea
                    className="form-input"
                    value={contactInfo}
                    onChange={e => setContactInfo(e.target.value)}
                    rows={3}
                    placeholder={isDe
                      ? 'z.B. „Vor Ort am Eventtag ab 7:30 Uhr, mobil unter +49 151 123 456" oder „Bei Fragen vor dem Event direkt per Mail."'
                      : 'e.g. „On-site from 7:30 am on event day, mobile +49 151 123 456" or „For questions before the event, email directly."'}
                  />
                </div>
              </div>

              {/* v9.21: Test-Team pro Event — sieht das Event im Entwurfsmodus
                  und kann sich anmelden, ohne globale Organizer-Rolle. Picker
                  via Graph-Search, beliebige Deloitte-User. */}
              <div className="form-group" style={{ position: 'relative', paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={7} />
                  Test-Team
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> eine kleine Gruppe von Personen, die das Event <strong>schon im Entwurfsmodus</strong> sieht und sich testweise anmelden darf — bevor du es für die echte Zielgruppe freigibst.<br /><br />
                      <strong>Anzeige in der App:</strong> Test-Team-Mitglieder sehen das Event in ihrer Liste, können auf die Anmelde-Seite, sich registrieren, abmelden, eigene Daten ändern. Sie haben <strong>keine Admin-Rechte</strong> — kein Bearbeiten, keine Teilnehmerliste, keine Massenmails. Reguläre User sehen das Event weiterhin nicht, solange der Entwurf-Haken gesetzt ist.<br /><br />
                      <strong>Automatismen:</strong> Test-Anmeldungen lösen ganz normal <strong>Bestätigungs-Mails</strong> und <strong>Outlook-Termine</strong> aus — perfekt um den kompletten Anmelde-Ablauf zu testen. Bei externen Mails (nicht @deloitte.de) greift die normale Umleitung an dich als Organizer.<br /><br />
                      <strong>Empfehlung:</strong> 1–3 Personen reichen typischerweise — ein Co-Organizer und ein naiver Tester, der noch nichts vom Event weiß.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> a small group of people who can see the event <strong>already in draft mode</strong> and register as a test — before you publish it to the real audience.<br /><br />
                      <strong>Shown in the app:</strong> test-team members see the event in their list, can open the registration page, register, cancel, edit their own data. They have <strong>no admin rights</strong> — no edit, no attendee list, no mass mails. Regular users still do not see the event while the draft toggle is on.<br /><br />
                      <strong>Automation:</strong> test registrations trigger normal <strong>confirmation mails</strong> and <strong>Outlook events</strong> — perfect for testing the full flow. External mails (non-@deloitte.de) follow the standard organizer-redirect rule.<br /><br />
                      <strong>Tip:</strong> 1–3 people are usually enough — a co-organizer and one naive tester who has not seen the event yet.
                    </>
                  )} />
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ marginLeft: 'auto', fontSize: '0.75rem', padding: '4px 10px', whiteSpace: 'nowrap' }}
                    onClick={() => setBulkTestTeamOpen(true)}
                  >
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                      <Users size={12} /> Massenimport
                    </span>
                  </button>
                </label>
                {testTeamNames.length > 0 && (() => {
                  const remove = (idx: number): void => {
                    setTestTeamNames(testTeamNames.filter((_, i) => i !== idx));
                    setTestTeamEmails(testTeamEmails.filter((_, i) => i !== idx));
                  };
                  return (
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 10 }}>
                      {testTeamNames.map((name, i) => {
                        const email = testTeamEmails[i] || '';
                        return (
                          <span key={`${email}-${i}`} style={{
                            display: 'inline-flex', alignItems: 'center', gap: 6,
                            padding: '3px 6px 3px 4px',
                            background: '#0ea5e9', color: '#fff',
                            borderRadius: 999, fontSize: '0.85rem', fontWeight: 500,
                          }}>
                            {email && (
                              <img
                                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=S`}
                                alt={name}
                                onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
                                style={{ width: 24, height: 24, borderRadius: '50%', objectFit: 'cover', background: 'rgba(255,255,255,0.25)' }}
                              />
                            )}
                            <span>{name}</span>
                            <button type="button" onClick={() => remove(i)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.9rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }} title="Entfernen">×</button>
                          </span>
                        );
                      })}
                    </div>
                  );
                })()}
                <input
                  className="form-input"
                  value={testTeamSearch}
                  onChange={e => {
                    const val = e.target.value;
                    setTestTeamSearch(val);
                    if (testTeamTimerRef.current) clearTimeout(testTeamTimerRef.current);
                    const q = val.trim();
                    if (!q) { setTestTeamResults([]); return; }
                    testTeamTimerRef.current = setTimeout(async () => {
                      try {
                        const results = await searchUsers(q, testTeamIncludeIntl);
                        setTestTeamResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setTestTeamResults([]); }
                    }, 350);
                  }}
                  onBlur={() => {
                    setTimeout(() => { setTestTeamSearch(''); setTestTeamResults([]); }, 150);
                  }}
                  placeholder="Name oder E-Mail eingeben (alle Deloitte-User)"
                />
                <InternationalSearchToggle
                  checked={testTeamIncludeIntl}
                  onChange={async next => {
                    setTestTeamIncludeIntl(next);
                    const q = testTeamSearch.trim();
                    if (q.length >= 1) {
                      try {
                        const results = await searchUsers(q, next);
                        setTestTeamResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setTestTeamResults([]); }
                    }
                  }}
                  isDe={isDe}
                />
                {testTeamResults.length > 0 && (
                  <div style={{
                    position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
                    background: '#fff', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                    maxHeight: 280, overflowY: 'auto',
                  }}>
                    {testTeamResults.map(u => {
                      const alreadyAdded = testTeamEmails.indexOf(u.email) >= 0;
                      return (
                        <div key={u.email} style={{
                          padding: '8px 12px', cursor: alreadyAdded ? 'not-allowed' : 'pointer', fontSize: '0.85rem',
                          borderBottom: '1px solid var(--dex-gray-100)',
                          opacity: alreadyAdded ? 0.45 : 1,
                          display: 'flex', alignItems: 'center', gap: 10,
                        }} onMouseDown={() => {
                          if (alreadyAdded || !u.email) return;
                          setTestTeamNames(prev => [...prev, u.displayName]);
                          setTestTeamEmails(prev => [...prev, u.email]);
                          setTestTeamSearch('');
                          setTestTeamResults([]);
                        }}>
                          <img
                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(u.email)}&size=S`}
                            alt={u.displayName}
                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                            style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                          />
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontWeight: 600 }}>{u.displayName}</div>
                            <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                              {u.email}{u.location ? ` · ${u.location}` : ''}
                            </div>
                          </div>
                          {alreadyAdded && <span style={{ color: 'var(--dex-green)', fontSize: '0.85rem', flexShrink: 0 }}>✓</span>}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* v6.19: QR-Code-Scanner pro Event. Separater Picker im gleichen Stil wie
                  Organizer, aber via Graph-Search (jeder Deloitte-User kann Scanner sein).
                  QR-Scanner haben eingeschränkten Admin-Zugriff (nur QR-Tool + KPIs),
                  erscheinen NICHT in Organizer-Listen auf MyEvents/RegistrationPage und
                  bekommen KEINE Organizer-Mails. */}
              <div className="form-group" style={{ position: 'relative', paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={8} />
                  {t('create.qrscanners') || 'QR-Code-Scanner'}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> Personen, die am Event-Tag <strong>nur das Check-In-Tool</strong> bedienen dürfen — z.B. Helfer am Empfangstresen oder am Stadioneingang. Beliebige Deloitte-User per Graph-Suche.<br /><br />
                      <strong>Anzeige in der App:</strong> Check-In-Team-Mitglieder sehen oben im Header das <strong>QR-Scanner-Icon</strong> und können den <strong>Check-In-Modus</strong> öffnen — QR-Codes scannen, Teilnehmer manuell ein-/auschecken, Check-In-KPIs sehen. Sie haben <strong>keine weiteren Rechte</strong>: kein Edit, keine Teilnehmerliste, keine Mails.<br /><br />
                      <strong>Automatismen:</strong> Check-In-Team taucht <strong>nicht in der Organizer-Liste</strong> auf der Anmelde-Seite auf und bekommt <strong>keine Organizer-Mails</strong> (BCC, Late-Cancel etc.).<br /><br />
                      <strong>Empfehlung:</strong> für jedes Event genau die Personen eintragen, die am Veranstaltungstag wirklich am Empfang stehen.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> people who may operate <strong>only the check-in tool</strong> on the event day — e.g. helpers at the welcome desk or stadium entrance. Any Deloitte user via Graph search.<br /><br />
                      <strong>Shown in the app:</strong> check-in team members see the <strong>QR scanner icon</strong> in the header and can open the <strong>check-in mode</strong> — scan QR codes, manually check attendees in/out, view check-in KPIs. They have <strong>no further rights</strong>: no edit, no attendee list, no emails.<br /><br />
                      <strong>Automation:</strong> check-in team does <strong>not appear in the organizer list</strong> on the registration page and does <strong>not receive organizer emails</strong> (BCC, late-cancel etc.).<br /><br />
                      <strong>Tip:</strong> for each event, list exactly the people who will actually staff the welcome desk.
                    </>
                  )} />
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ marginLeft: 'auto', fontSize: '0.75rem', padding: '4px 10px', whiteSpace: 'nowrap' }}
                    onClick={() => setBulkQrScannerOpen(true)}
                  >
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                      <Users size={12} /> Massenimport
                    </span>
                  </button>
                </label>
                {qrScannerNames.length > 0 && (() => {
                  const move = (idx: number, dir: -1 | 1): void => {
                    const target = idx + dir;
                    if (target < 0 || target >= qrScannerNames.length) return;
                    const nextNames = [...qrScannerNames];
                    const nextEmails = [...qrScannerEmails];
                    [nextNames[idx], nextNames[target]] = [nextNames[target], nextNames[idx]];
                    [nextEmails[idx], nextEmails[target]] = [nextEmails[target], nextEmails[idx]];
                    setQrScannerNames(nextNames);
                    setQrScannerEmails(nextEmails);
                  };
                  const remove = (idx: number): void => {
                    setQrScannerNames(qrScannerNames.filter((_, i) => i !== idx));
                    setQrScannerEmails(qrScannerEmails.filter((_, i) => i !== idx));
                  };
                  return (
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 10 }}>
                      {qrScannerNames.map((name, i) => {
                        const email = qrScannerEmails[i] || '';
                        return (
                          <span
                            key={`${email}-${i}`}
                            style={{
                              display: 'inline-flex', alignItems: 'center', gap: 6,
                              padding: '3px 6px 3px 4px',
                              background: 'var(--dex-orange, #ed8b00)', color: '#fff',
                              borderRadius: 999, fontSize: '0.85rem', fontWeight: 500,
                            }}
                          >
                            {email ? (
                              <img
                                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=S`}
                                alt={name}
                                onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
                                style={{ width: 24, height: 24, borderRadius: '50%', objectFit: 'cover', background: 'rgba(255,255,255,0.25)' }}
                              />
                            ) : null}
                            <span>{name}</span>
                            {qrScannerNames.length > 1 && i > 0 && (
                              <button type="button" onClick={() => move(i, -1)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.75rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }} title="Nach vorne">◀</button>
                            )}
                            {qrScannerNames.length > 1 && i < qrScannerNames.length - 1 && (
                              <button type="button" onClick={() => move(i, 1)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.75rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }} title="Nach hinten">▶</button>
                            )}
                            <button type="button" onClick={() => remove(i)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.9rem', lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }} title="Entfernen">×</button>
                          </span>
                        );
                      })}
                    </div>
                  );
                })()}
                <input
                  className="form-input"
                  value={qrScannerSearch}
                  onChange={e => {
                    const val = e.target.value;
                    setQrScannerSearch(val);
                    if (qrScannerTimerRef.current) clearTimeout(qrScannerTimerRef.current);
                    const q = val.trim();
                    if (!q) { setQrScannerResults([]); return; }
                    // v9.18: Graph-Search statt Role-Filter — jeder Deloitte-User
                    // kann QR-Scanner sein. Debounce 350ms.
                    qrScannerTimerRef.current = setTimeout(async () => {
                      try {
                        const results = await searchUsers(q, qrScannerIncludeIntl);
                        setQrScannerResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setQrScannerResults([]); }
                    }, 350);
                  }}
                  onBlur={() => {
                    setTimeout(() => { setQrScannerSearch(''); setQrScannerResults([]); }, 150);
                  }}
                  placeholder={t('create.qrscanners.placeholder') || 'Name oder E-Mail eingeben (alle Deloitte-User)'}
                />
                <InternationalSearchToggle
                  checked={qrScannerIncludeIntl}
                  onChange={async next => {
                    setQrScannerIncludeIntl(next);
                    const q = qrScannerSearch.trim();
                    if (q.length >= 1) {
                      try {
                        const results = await searchUsers(q, next);
                        setQrScannerResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setQrScannerResults([]); }
                    }
                  }}
                  isDe={isDe}
                />
                {qrScannerResults.length > 0 && (
                  <div style={{
                    position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
                    background: '#fff', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                    maxHeight: 280, overflowY: 'auto',
                  }}>
                    {qrScannerResults.map(u => {
                      const alreadyAdded = qrScannerEmails.indexOf(u.email) >= 0;
                      return (
                        <div
                          key={u.email}
                          style={{
                            padding: '8px 12px', cursor: alreadyAdded ? 'not-allowed' : 'pointer', fontSize: '0.85rem',
                            borderBottom: '1px solid var(--dex-gray-100)',
                            opacity: alreadyAdded ? 0.45 : 1,
                            display: 'flex', alignItems: 'center', gap: 10,
                          }}
                          onMouseDown={() => {
                            if (alreadyAdded || !u.email) return;
                            setQrScannerNames(prev => [...prev, u.displayName]);
                            setQrScannerEmails(prev => [...prev, u.email]);
                            setQrScannerSearch('');
                            setQrScannerResults([]);
                          }}
                        >
                          <img
                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(u.email)}&size=S`}
                            alt={u.displayName}
                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                            style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                          />
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontWeight: 600 }}>{u.displayName}</div>
                            <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                              {u.email}{u.location ? ` · ${u.location}` : ''}
                            </div>
                          </div>
                          {alreadyAdded && <span style={{ color: 'var(--dex-green)', fontSize: '0.85rem', flexShrink: 0 }}>✓</span>}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* Duplikat-Hinweis: gleiche Person in mehreren Team-Listen.
                  Co-Organizer haben automatisch Check-In- und Test-Team-Rechte
                  (Header.canCheckIn-Logik, Drafts-Sichtbarkeit für Organizer),
                  doppelte Einträge sind redundant. Test-Team und Check-In allein
                  sind orthogonale Rollen — die warnen wir nicht.

                  Pattern: nach jedem Add (Massenimport oder Einzel-Pick) updated
                  sich die Memo automatisch und die Warnung erscheint inline. Pro
                  Eintrag ein Ein-Klick-Button um die Person aus der überflüssigen
                  Liste zu entfernen. */}
              {(() => {
                const orgSet = new Set(organizerEmails.map(e => (e || '').toLowerCase()));
                const ttSet = new Set(testTeamEmails.map(e => (e || '').toLowerCase()));
                const qrSet = new Set(qrScannerEmails.map(e => (e || '').toLowerCase()));
                const allEmails = new Set<string>();
                organizerEmails.forEach(e => allEmails.add((e || '').toLowerCase()));
                testTeamEmails.forEach(e => allEmails.add((e || '').toLowerCase()));
                qrScannerEmails.forEach(e => allEmails.add((e || '').toLowerCase()));

                const orgNames = organizer.split(';').map(s => s.trim()).filter(Boolean);
                const dups: Array<{ email: string; name: string; inOrg: boolean; inTt: boolean; inQr: boolean }> = [];
                // Array.from() statt `for ... of Set` — TS-Target ES5 erlaubt kein
                // direktes Set-Iterieren ohne --downlevelIteration.
                for (const e of Array.from(allEmails)) {
                  if (!e) continue;
                  const inOrg = orgSet.has(e);
                  const inTt = ttSet.has(e);
                  const inQr = qrSet.has(e);
                  // Nur Co-Organizer + (Test|Check-In) ist redundant. Test+Check-In
                  // ohne Co-Organizer sind unterschiedliche Funktionen → nicht warnen.
                  if (!inOrg) continue;
                  if (!inTt && !inQr) continue;
                  // Display-Name aus dem ersten Treffer ziehen (Org bevorzugt).
                  let name = e;
                  const idxOrg = organizerEmails.findIndex(x => (x || '').toLowerCase() === e);
                  if (idxOrg >= 0 && orgNames[idxOrg]) name = orgNames[idxOrg];
                  else {
                    const idxTt = testTeamEmails.findIndex(x => (x || '').toLowerCase() === e);
                    if (idxTt >= 0 && testTeamNames[idxTt]) name = testTeamNames[idxTt];
                    else {
                      const idxQr = qrScannerEmails.findIndex(x => (x || '').toLowerCase() === e);
                      if (idxQr >= 0 && qrScannerNames[idxQr]) name = qrScannerNames[idxQr];
                    }
                  }
                  dups.push({ email: e, name, inOrg, inTt, inQr });
                }
                if (dups.length === 0) return null;

                const removeFromTestTeam = (emailLc: string): void => {
                  const idx = testTeamEmails.findIndex(x => (x || '').toLowerCase() === emailLc);
                  if (idx < 0) return;
                  setTestTeamNames(testTeamNames.filter((_, i) => i !== idx));
                  setTestTeamEmails(testTeamEmails.filter((_, i) => i !== idx));
                };
                const removeFromQr = (emailLc: string): void => {
                  const idx = qrScannerEmails.findIndex(x => (x || '').toLowerCase() === emailLc);
                  if (idx < 0) return;
                  setQrScannerNames(qrScannerNames.filter((_, i) => i !== idx));
                  setQrScannerEmails(qrScannerEmails.filter((_, i) => i !== idx));
                };
                const removeAllOverlap = (emailLc: string): void => {
                  removeFromTestTeam(emailLc);
                  removeFromQr(emailLc);
                };

                return (
                  <div
                    className="form-group"
                    style={{
                      background: '#fff8e1',
                      border: '1px solid #f0c419',
                      borderRadius: 'var(--dex-radius)',
                      padding: '12px 14px',
                      marginBottom: 20,
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                      <span style={{ fontSize: '1.1rem' }}>⚠️</span>
                      <strong style={{ fontSize: '0.95rem' }}>
                        {dups.length === 1
                          ? '1 Person ist mehrfach gelistet'
                          : `${dups.length} Personen sind mehrfach gelistet`}
                      </strong>
                    </div>
                    <p style={{ margin: '0 0 10px', fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                      Co-Organizer dürfen automatisch das <strong>Check-In-Tool</strong> nutzen und sehen
                      Events auch im <strong>Entwurfsmodus</strong> — ein zusätzlicher Eintrag im Test-Team oder
                      Check-In-Team ist daher nicht nötig. Du kannst die überflüssigen Einträge hier auf
                      einen Klick entfernen.
                    </p>
                    <ul style={{ margin: 0, padding: 0, listStyle: 'none' }}>
                      {dups.map(d => {
                        const teamsLabel: string[] = [];
                        if (d.inOrg) teamsLabel.push('Co-Organizer');
                        if (d.inTt) teamsLabel.push('Test-Team');
                        if (d.inQr) teamsLabel.push('Check-In-Team');
                        return (
                          <li
                            key={d.email}
                            style={{
                              padding: '8px 0',
                              borderTop: '1px solid #f0c419',
                              display: 'flex',
                              alignItems: 'center',
                              gap: 10,
                              flexWrap: 'wrap',
                            }}
                          >
                            <div style={{ flex: '1 1 200px', minWidth: 0 }}>
                              <strong>{d.name}</strong>{' '}
                              <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>{d.email}</span>
                              <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
                                Aktuell in: {teamsLabel.join(', ')}
                              </div>
                            </div>
                            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                              {d.inTt && (
                                <button
                                  type="button"
                                  className="btn btn-secondary"
                                  style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                  onClick={() => removeFromTestTeam(d.email)}
                                >
                                  Aus Test-Team entfernen
                                </button>
                              )}
                              {d.inQr && (
                                <button
                                  type="button"
                                  className="btn btn-secondary"
                                  style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                  onClick={() => removeFromQr(d.email)}
                                >
                                  Aus Check-In-Team entfernen
                                </button>
                              )}
                              {d.inTt && d.inQr && (
                                <button
                                  type="button"
                                  className="btn btn-primary"
                                  style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                  onClick={() => removeAllOverlap(d.email)}
                                >
                                  Aus beiden entfernen
                                </button>
                              )}
                            </div>
                          </li>
                        );
                      })}
                    </ul>
                  </div>
                );
              })()}

              </div>

              {/* ===== Step 3 (v15.0: vormals Step 2): Ort & Programm ===== */}
              <div style={{ display: currentStep === 2 ? 'block' : 'none' }}>
              <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                {isDe ? 'Schritt 3 — Ort & Programm' : 'Step 3 — Location & Programme'}
              </h2>
              <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                {isDe
                  ? 'Hier sagst du, wo das Event stattfindet, wie der Tagesablauf aussieht und wie Teilnehmer hinkommen.'
                  : 'Here you say where the event takes place, what the schedule looks like and how attendees get there.'}
              </p>
              {renderStepIntro(
                [
                  'Veranstaltungsort und Adresse erfassen',
                  'Agenda pflegen — Tagesablauf für die Teilnehmer',
                  'Optional: Transferzeiten (Bus/Bahn/Treffpunkt) hinterlegen',
                ],
                [
                  'Set event location and address',
                  'Maintain the agenda — schedule shown to participants',
                  'Optional: add transfer times (bus/train/meeting point)',
                ]
              )}
              {/* v9.28: Banner — alle Infos in diesem Schritt landen direkt
                  beim Teilnehmer (Anmelde-Seite + Meine Events). */}
              <div style={{
                display: 'flex', alignItems: 'flex-start', gap: 10,
                padding: '12px 14px', marginBottom: 20,
                background: 'rgba(0,118,168,0.06)',
                border: '1px solid #0076a8',
                borderRadius: 'var(--dex-radius, 12px)',
                fontSize: '0.85rem', color: 'var(--dex-gray-700)',
                lineHeight: 1.5,
              }}>
                <Icon iconName="Info" style={{ fontSize: 18, color: '#0076a8', flexShrink: 0, marginTop: 2 }} />
                <div>
                  {isDe
                    ? <>Alle Eingaben in diesem Schritt — <strong>Veranstaltungsort, Adresse, Agenda und Transferzeiten</strong> — werden den Teilnehmern direkt auf der <strong>Anmelde-Seite</strong> und später unter <strong>{'„Meine Events“'}</strong> angezeigt.</>
                    : <>All inputs in this step — <strong>venue, address, agenda and transfer times</strong> — are shown to attendees directly on the <strong>registration page</strong> and later under <strong>{'„My Events“'}</strong>.</>}
                </div>
              </div>

              {/* v15.3: pro-Sub-Event-Tabs fuer den Ort. Tab 0 = Haupt-Event
                  (komplette Ort/Adresse/Agenda/Transferzeiten-UI bleibt
                  unveraendert). Tabs N>0 = vollwertige Ort/Adresse/Agenda/
                  Transferzeiten-UI pro Sub-Event — kein Inheritance-Toggle
                  mehr, jedes Sub-Event hat eigene Werte. Per
                  „Vom Hauptevent kopieren"-Button kann der Organizer die
                  Hauptevent-Werte als Startpunkt uebernehmen. */}
              {renderPerEventTabStrip(
                activeLocationTabIdx,
                setActiveLocationTabIdx,
                `${isDe ? 'Haupt-Event' : 'Main event'}: ${title || (isDe ? 'Ohne Titel' : 'Untitled')}`,
                isDe ? 'Event-Tab wechseln (Ort & Programm)' : 'Switch event tab (location & programme)'
              )}

              {activeLocationTabIdx > 0 && (() => {
                const seIdx = activeLocationTabIdx - 1;
                const se = subEvents[seIdx];
                if (!se) return null;
                const seAddr = se.locationAddress || { street: '', houseNo: '', zip: '', city: '' };
                const seAgenda = se.agenda || [];
                const seTransfers = se.transferTimes || [];
                const updateSub = (patch: Partial<SubEventDraft>): void => {
                  setSubEvents(prev => prev.map((x, i) => i === seIdx ? { ...x, ...patch } : x));
                };
                const updateSubAgendaItem = (id: string, patch: Partial<AgendaItem>): void => {
                  updateSub({ agenda: seAgenda.map(a => a.id === id ? { ...a, ...patch } : a) });
                };
                return (
                  <div>
                    {/* v15.3: „Vom Hauptevent kopieren"-Button. Uebernimmt
                        Ort, Adresse, Agenda und Transferzeiten vom Hauptevent
                        als Startwerte fuer dieses Sub-Event. */}
                    <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }}>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.8rem', padding: '6px 14px' }}
                        onClick={() => updateSub({
                          location: location,
                          locationAddress: { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity },
                          agenda: agenda.map(a => ({ ...a, id: `ag-${Date.now()}-${Math.random().toString(36).slice(2, 6)}` })),
                          transferTimes: transferTimes.map(tt => ({ ...tt, id: `tr-${Date.now()}-${Math.random().toString(36).slice(2, 6)}` })),
                        })}
                        title={isDe
                          ? 'Uebernimmt Ort, Adresse, Agenda und Transferzeiten vom Hauptevent als Startwerte'
                          : 'Copies location, address, agenda and transfer times from the main event as starting values'}
                      >
                        {isDe ? 'Vom Hauptevent kopieren' : 'Copy from main event'}
                      </button>
                    </div>
                    <div className="form-group">
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={9} />
                        {t('create.location')}
                      </label>
                      <input
                        className="form-input"
                        value={se.location || ''}
                        onChange={e => updateSub({ location: e.target.value })}
                        placeholder={isDe ? 'z.B. RheinEnergieStadion, Köln' : 'e.g. RheinEnergieStadion, Cologne'}
                      />
                    </div>
                    <div className="form-group">
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={10} />
                        {isDe ? 'Adresse' : 'Address'}
                      </label>
                      <div style={{ display: 'grid', gridTemplateColumns: '3fr 1fr', gap: 8, marginBottom: 8 }}>
                        <input className="form-input" value={seAddr.street} onChange={e => updateSub({ locationAddress: { ...seAddr, street: e.target.value } })} placeholder="Straße" />
                        <input className="form-input" value={seAddr.houseNo} onChange={e => updateSub({ locationAddress: { ...seAddr, houseNo: e.target.value } })} placeholder="Hausnr." />
                      </div>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 3fr', gap: 8 }}>
                        <input className="form-input" value={seAddr.zip} onChange={e => updateSub({ locationAddress: { ...seAddr, zip: e.target.value } })} placeholder="PLZ" />
                        <input className="form-input" value={seAddr.city} onChange={e => updateSub({ locationAddress: { ...seAddr, city: e.target.value } })} placeholder="Ort" />
                      </div>
                    </div>
                    {/* v18.44: Outlook-Ort pro Sub-Event überschreibbar (auch hier, nicht nur im Outlook-Editor). */}
                    <div className="form-group" style={{ marginTop: 16 }}>
                      <label className="form-label">
                        {isDe ? 'Ort im Outlook-Termin' : 'Location in the Outlook event'}
                      </label>
                      <input
                        className="form-input"
                        value={se.outlookLocation || ''}
                        onChange={e => updateSub({ outlookLocation: e.target.value })}
                        placeholder={buildOutlookLocation(se.location, seAddr) || (isDe ? 'z.B. Mezzomar, Harffstraße 110a, Düsseldorf' : 'e.g. Mezzomar, Harffstraße 110a, Düsseldorf')}
                      />
                      <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                        {isDe
                          ? 'Leer lassen = automatisch aus Veranstaltungsort + Adresse dieses Sub-Events.'
                          : 'Leave empty = automatic from this sub-event\'s venue + address.'}
                      </span>
                    </div>
                    <div className="form-group" style={{ marginTop: 24 }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '1rem', fontWeight: 700 }}>
                        <StepBadge n={11} />
                        {t('create.agenda')}
                      </label>
                      {seAgenda
                        .slice()
                        .sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time))
                        .map(item => (
                        <div key={item.id} style={{
                          display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'flex-start',
                          padding: '10px 12px', marginBottom: 8,
                          background: 'var(--dex-gray-50, #fafafa)', borderRadius: 'var(--dex-radius)',
                          border: '1px solid var(--dex-gray-200)',
                        }}>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 120 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.date')}</label>
                            <input type="date" className="form-input" value={item.date} onChange={e => updateSubAgendaItem(item.id, { date: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 80 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.time')}</label>
                            <input type="time" className="form-input" value={item.time} onChange={e => updateSubAgendaItem(item.id, { time: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 80 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.endtime')}</label>
                            <input type="time" className="form-input" value={item.endTime || ''} onChange={e => updateSubAgendaItem(item.id, { endTime: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2, flex: 1, minWidth: 150 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.title')}</label>
                            <input type="text" className="form-input" value={item.title} onChange={e => updateSubAgendaItem(item.id, { title: e.target.value })} placeholder={t('create.agenda.title')} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2, flex: 1, minWidth: 150 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.desc')}</label>
                            <input type="text" className="form-input" value={item.description || ''} onChange={e => updateSubAgendaItem(item.id, { description: e.target.value })} placeholder={t('create.agenda.desc')} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                          </div>
                          <div style={{ display: 'flex', alignItems: 'flex-end', paddingBottom: 2 }}>
                            <button type="button" onClick={() => updateSub({ agenda: seAgenda.filter(a => a.id !== item.id) })} style={{
                              background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)',
                              fontSize: '1.1rem', padding: '4px', lineHeight: 1,
                            }} title={t('general.delete')}>
                              <X size={16} />
                            </button>
                          </div>
                        </div>
                      ))}
                      <button type="button" className="btn btn-outline" onClick={() => updateSub({
                        agenda: [...seAgenda, {
                          id: `ag-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
                          date: se.startDate ? se.startDate.slice(0, 10) : '',
                          time: '',
                          endTime: '',
                          icon: 'Calendar',
                          title: '',
                          description: '',
                        }],
                      })} style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4 }}>
                        <Plus size={14} /> {t('create.agenda.add')}
                      </button>
                    </div>
                    <div className="form-group" style={{ marginTop: 24 }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '1rem', fontWeight: 700 }}>
                        <StepBadge n={12} />
                        {t('create.transfers')}
                      </label>
                      {seTransfers.map(tt => (
                        <div key={tt.id} style={{
                          padding: '12px 14px', marginBottom: 8,
                          background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                          border: '1px solid var(--dex-gray-200)',
                        }}>
                          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr auto', gap: 8, marginBottom: 8 }}>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.location')}</label>
                              <input type="text" className="form-input" value={tt.location} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, location: e.target.value } : x) })} placeholder="Stadt..." style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.meetingpoint')}</label>
                              <input type="text" className="form-input" value={tt.meetingPoint || ''} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, meetingPoint: e.target.value } : x) })} placeholder="z.B. Hbf..." style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.address')}</label>
                              <input type="text" className="form-input" value={tt.address || ''} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, address: e.target.value } : x) })} placeholder="Straße, PLZ Ort" style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div style={{ display: 'flex', alignItems: 'flex-end', paddingBottom: 4 }}>
                              <button type="button" onClick={() => updateSub({ transferTimes: seTransfers.filter(x => x.id !== tt.id) })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)', padding: '4px', lineHeight: 1 }} title={t('general.delete')}>
                                <X size={16} />
                              </button>
                            </div>
                          </div>
                          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 2fr', gap: 8 }}>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.date')}</label>
                              <input type="date" className="form-input" value={tt.date} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, date: e.target.value } : x) })} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.departure')}</label>
                              <input type="time" className="form-input" value={tt.departureTime} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, departureTime: e.target.value } : x) })} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.arrival')}</label>
                              <input type="time" className="form-input" value={tt.arrivalTime} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, arrivalTime: e.target.value } : x) })} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.desc')}</label>
                              <input type="text" className="form-input" value={tt.description || ''} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, description: e.target.value } : x) })} placeholder={t('create.transfers.desc')} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                          </div>
                        </div>
                      ))}
                      <button type="button" className="btn btn-outline" onClick={() => updateSub({
                        transferTimes: [...seTransfers, {
                          id: `tr-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
                          location: '',
                          meetingPoint: '',
                          address: '',
                          date: se.startDate ? se.startDate.slice(0, 10) : '',
                          departureTime: '',
                          arrivalTime: '',
                          description: '',
                        }],
                      })} style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4 }}>
                        <Plus size={14} /> {t('create.transfers.add')}
                      </button>
                    </div>
                  </div>
                );
              })()}

              <div style={{ display: activeLocationTabIdx === 0 ? 'block' : 'none' }}>
              {/* v15.7: Step 3 NICHT mehr ausgrauen wenn subEventsOnlyMode —
                  Ort/Adresse/Agenda/Transferzeiten sind übergreifende
                  Event-Infos (gehen auf die Event-Detail-Seite, in „Meine
                  Events"-Card usw.). Sie bleiben immer relevant, auch wenn
                  das Hauptevent nicht anmeldbar ist. */}
              <div>
              <div className="form-group">
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={9} />
                  {t('create.location')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> den <strong>Namen des Veranstaltungsortes</strong> (z.B. RheinEnergieStadion, Köln oder Deloitte Düsseldorf, Schwannstraße 6).<br /><br />
                      <strong>Anzeige in der App:</strong> erscheint auf der <strong>Anmelde-Seite</strong>, in der <strong>Eventliste</strong>, in <strong>Meine Events</strong> als Ort-Information.<br /><br />
                      <strong>Automatismen:</strong> wandert in den <strong>Outlook-Termin der Teilnehmer</strong> als Ort-Feld — so sehen sie auf einen Blick, wo sie hin müssen. Falls Bing Maps den Ort findet, blendet Outlook automatisch eine Karte ein.<br /><br />
                      <strong>Empfehlung:</strong> sprechender Name + Stadt — verwende die Adresse für die strukturierte Detail-Adresse darunter, hier reicht der Veranstaltungsort.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>name of the venue</strong> (e.g. RheinEnergieStadion, Cologne or Deloitte Düsseldorf, Schwannstraße 6).<br /><br />
                      <strong>Shown in the app:</strong> appears on the <strong>registration page</strong>, in the <strong>event list</strong> and in <strong>My Events</strong> as the location.<br /><br />
                      <strong>Automation:</strong> goes into the attendee Outlook event as the location field — so they immediately see where to go. If Bing Maps recognises it, Outlook auto-inserts a map.<br /><br />
                      <strong>Tip:</strong> descriptive name + city — use the structured address below for full details, here just the venue name.
                    </>
                  )} />
                </label>
                <input className="form-input" value={location} onChange={e => setLocation(e.target.value)} placeholder="z.B. RheinEnergieStadion, Köln" />
              </div>
              <div className="form-group">
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={10} />
                  Adresse
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> die <strong>strukturierte Adresse</strong> (Straße, Hausnr., PLZ, Ort) — getrennt eingegeben, damit die Adresse einheitlich aussieht.<br /><br />
                      <strong>Anzeige in der App:</strong> wird auf der <strong>Anmelde-Seite</strong> und in <strong>Meine Events</strong> sauber formatiert angezeigt — Teilnehmer können auf die Adresse klicken und sie z.B. in Google Maps öffnen.<br /><br />
                      <strong>Automatismen:</strong> wird in <strong>Bestätigungs-Mails</strong> und im <strong>Outlook-Termin-Body</strong> als Klartext-Adresse mitgegeben — z.B. zum Kopieren ins Navi.<br /><br />
                      <strong>Optional:</strong> wenn der Veranstaltungsort oben schon eindeutig genug ist, kannst du die strukturierte Adresse leer lassen.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>structured address</strong> (street, number, ZIP, city) — entered field by field for consistent formatting.<br /><br />
                      <strong>Shown in the app:</strong> shown nicely on the <strong>registration page</strong> and under <strong>My Events</strong> — attendees can click it and e.g. open it in Google Maps.<br /><br />
                      <strong>Automation:</strong> goes into the <strong>confirmation mails</strong> and the <strong>Outlook event body</strong> as plain text — handy to copy into a navigation device.<br /><br />
                      <strong>Optional:</strong> if the venue name above is already explicit enough, you can leave the structured address empty.
                    </>
                  )} />
                </label>
                <div style={{ display: 'grid', gridTemplateColumns: '3fr 1fr', gap: 8, marginBottom: 8 }}>
                  <input className="form-input" value={addrStreet} onChange={e => setAddrStreet(e.target.value)} placeholder="Straße" />
                  <input className="form-input" value={addrHouseNo} onChange={e => setAddrHouseNo(e.target.value)} placeholder="Hausnr." />
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 3fr', gap: 8 }}>
                  <input className="form-input" value={addrZip} onChange={e => setAddrZip(e.target.value)} placeholder="PLZ" />
                  <input className="form-input" value={addrCity} onChange={e => setAddrCity(e.target.value)} placeholder="Ort" />
                </div>
              </div>

              {/* v18.40: Ort im Outlook-Termin (überschreibbar) */}
              <div className="form-group" style={{ marginTop: 16 }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  {isDe ? 'Ort im Outlook-Termin' : 'Location in the Outlook event'}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> den Text, der im <strong>&bdquo;Ort&ldquo;-Feld des Outlook-Termins</strong> der Teilnehmer steht.<br /><br />
                      <strong>Standard:</strong> wird automatisch aus <strong>Veranstaltungsort + Adresse</strong> oben zusammengebaut (siehe Platzhalter). Lässt du das Feld <strong>leer</strong>, wird immer dieser aktuelle Standard verwendet.<br /><br />
                      <strong>Überschreiben:</strong> Trägst du hier etwas ein, wird genau dieser Text als Termin-Ort genommen — z.&nbsp;B. ein abweichender Raum, ein Online-Link oder ein Kurzname.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the text shown in the <strong>&bdquo;Location&ldquo; field of attendees&apos; Outlook event</strong>.<br /><br />
                      <strong>Default:</strong> built automatically from <strong>venue + address</strong> above (see placeholder). Leave it <strong>empty</strong> to always use that current default.<br /><br />
                      <strong>Override:</strong> type something here to use exactly that text as the event location — e.g. a different room, an online link or a short name.
                    </>
                  )} />
                </label>
                <input
                  className="form-input"
                  value={outlookLocationOverride}
                  onChange={e => setOutlookLocationOverride(e.target.value)}
                  placeholder={buildOutlookLocation(location, { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity }) || (isDe ? 'z.B. Mezzomar, Harffstraße 110a, Düsseldorf' : 'e.g. Mezzomar, Harffstraße 110a, Düsseldorf')}
                />
                <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                  {isDe
                    ? 'Leer lassen = automatisch aus Veranstaltungsort + Adresse. Eingabe überschreibt den Termin-Ort.'
                    : 'Leave empty = automatic from venue + address. Any input overrides the event location.'}
                </span>
              </div>

              {/* ===== Agenda Editor ===== */}
              <div className="form-group" style={{ marginTop: 24 }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '1rem', fontWeight: 700 }}>
                  <StepBadge n={11} />
                  {t('create.agenda')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> den <strong>Programmablauf des Events</strong> als Liste — pro Punkt: Datum, Start- und Endzeit, Titel, optionale Beschreibung und ein Icon (z.B. Kaffee, Vortrag, Pause).<br /><br />
                      <strong>Anzeige in der App:</strong> erscheint als <strong>schöner Timeline-Block</strong> auf der Anmelde-Seite und in Meine Events — Punkte werden automatisch nach Datum + Uhrzeit sortiert. Mehrtägige Events werden tageweise gruppiert.<br /><br />
                      <strong>Automatismen:</strong> die Agenda landet <strong>nicht</strong> automatisch im Outlook-Termin-Body (dafür gibt es das eigene Feld <strong>Text im Outlook-Termin</strong> in Schritt 7).<br /><br />
                      <strong>Empfehlung:</strong> hilft Teilnehmern, sich auf den Tag einzustellen — bei Tagungen oder Auswärtsterminen sehr empfohlen, bei kurzen Office-Events optional.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>event programme</strong> as a list — per item: date, start/end time, title, optional description and an icon (e.g. coffee, talk, break).<br /><br />
                      <strong>Shown in the app:</strong> shown as a <strong>nice timeline block</strong> on the registration page and in My Events — items are auto-sorted by date + time. Multi-day events are grouped per day.<br /><br />
                      <strong>Automation:</strong> the agenda is <strong>not</strong> automatically pulled into the Outlook event body (there is a dedicated field <strong>Text in the Outlook event</strong> in step 7 for that).<br /><br />
                      <strong>Tip:</strong> helps attendees plan their day — strongly recommended for off-site events / conferences, optional for short office events.
                    </>
                  )} />
                </label>
                {agenda
                  .slice()
                  .sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time))
                  .map((item) => (
                  <div key={item.id} style={{
                    display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'flex-start',
                    padding: '10px 12px', marginBottom: 8,
                    background: 'var(--dex-gray-50, #fafafa)', borderRadius: 'var(--dex-radius)',
                    border: '1px solid var(--dex-gray-200)',
                  }}>
                    {/* Icon Picker - Grüner Kreis mit weißem Icon */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, position: 'relative' }}>
                      <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.icon')}</label>
                      <button
                        type="button"
                        onClick={() => { setIconPickerOpen(iconPickerOpen === item.id ? null : item.id); setIconSearch(''); setShowAllIcons(false); }}
                        style={{
                          width: 38, height: 38, display: 'flex', alignItems: 'center', justifyContent: 'center',
                          border: 'none', borderRadius: '50%',
                          background: 'var(--dex-green-dark, #6b9a1e)', cursor: 'pointer',
                        }}
                        title={item.icon || 'Calendar'}
                      >
                        <Icon iconName={item.icon || 'Calendar'} style={{ fontSize: 18, color: '#fff' }} />
                      </button>
                      {iconPickerOpen === item.id && (
                        <div style={{
                          position: 'absolute', top: '100%', left: 0, zIndex: 100,
                          background: '#fff', border: '1px solid var(--dex-gray-300)', borderRadius: 12,
                          boxShadow: '0 6px 20px rgba(0,0,0,0.15)', padding: 10, width: 300,
                        }}>
                          <input
                            type="text"
                            className="form-input"
                            placeholder="Icon suchen..."
                            value={iconSearch}
                            onChange={e => setIconSearch(e.target.value)}
                            style={{ fontSize: '0.82rem', padding: '6px 10px', marginBottom: 8, borderRadius: 8 }}
                            autoFocus
                          />
                          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(6, 1fr)', gap: 6, maxHeight: 240, overflowY: 'auto', padding: '2px 0' }}>
                            {(showAllIcons ? [...AGENDA_ICONS, ...EXTENDED_ICONS] : AGENDA_ICONS)
                              .filter(ic => !iconSearch || ic.label.toLowerCase().includes(iconSearch.toLowerCase()) || ic.name.toLowerCase().includes(iconSearch.toLowerCase()) || ic.category.includes(iconSearch.toLowerCase()))
                              .map(ic => (
                                <button
                                  key={ic.name}
                                  type="button"
                                  title={ic.label}
                                  onClick={() => { updateAgendaItem(item.id, { icon: ic.name }); setIconPickerOpen(null); }}
                                  style={{
                                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                                    width: 42, height: 42, border: 'none', borderRadius: '50%', cursor: 'pointer',
                                    background: item.icon === ic.name ? 'var(--dex-green-dark, #6b9a1e)' : 'var(--dex-gray-100, #f3f3f3)',
                                    transition: 'background 0.15s, transform 0.1s',
                                  }}
                                >
                                  <Icon iconName={ic.name} style={{ fontSize: 18, color: item.icon === ic.name ? '#fff' : 'var(--dex-gray-700)' }} />
                                </button>
                              ))
                            }
                          </div>
                          {!showAllIcons && !iconSearch && (
                            <button
                              type="button"
                              onClick={() => setShowAllIcons(true)}
                              style={{
                                width: '100%', marginTop: 8, padding: '6px 0', fontSize: '0.78rem',
                                background: 'none', border: '1px dashed var(--dex-gray-300)', borderRadius: 8,
                                color: 'var(--dex-green, #86bc25)', cursor: 'pointer', fontWeight: 600,
                              }}
                            >
                              + {t('create.agenda.icon') === 'Icon' ? 'Show all icons' : 'Alle Icons anzeigen'}
                            </button>
                          )}
                        </div>
                      )}
                    </div>

                    {/* Date */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 120 }}>
                      <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.date')}</label>
                      <input type="date" className="form-input" value={item.date} onChange={e => updateAgendaItem(item.id, { date: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                    </div>

                    {/* Start Time */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 80 }}>
                      <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.time')}</label>
                      <input type="time" className="form-input" value={item.time} onChange={e => updateAgendaItem(item.id, { time: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                    </div>

                    {/* End Time */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 80 }}>
                      <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.endtime')}</label>
                      <input type="time" className="form-input" value={item.endTime || ''} onChange={e => updateAgendaItem(item.id, { endTime: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                    </div>

                    {/* Title */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, flex: 1, minWidth: 150 }}>
                      <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.title')}</label>
                      <input type="text" className="form-input" value={item.title} onChange={e => updateAgendaItem(item.id, { title: e.target.value })} placeholder={t('create.agenda.title')} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                    </div>

                    {/* Description */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, flex: 1, minWidth: 150 }}>
                      <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.desc')}</label>
                      <input type="text" className="form-input" value={item.description || ''} onChange={e => updateAgendaItem(item.id, { description: e.target.value })} placeholder={t('create.agenda.desc')} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                    </div>

                    {/* Delete */}
                    <div style={{ display: 'flex', alignItems: 'flex-end', paddingBottom: 2 }}>
                      <button type="button" onClick={() => removeAgendaItem(item.id)} style={{
                        background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)',
                        fontSize: '1.1rem', padding: '4px', lineHeight: 1,
                      }} title={t('general.delete')}>
                        <X size={16} />
                      </button>
                    </div>
                  </div>
                ))}
                <button type="button" className="btn btn-outline" onClick={addAgendaItem} style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4 }}>
                  <Plus size={14} /> {t('create.agenda.add')}
                </button>
              </div>

              {/* ===== Transferzeiten Editor ===== */}
              <div className="form-group" style={{ marginTop: 24 }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '1rem', fontWeight: 700 }}>
                  <StepBadge n={12} />
                  {t('create.transfers')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> <strong>An- und Abreise-Infos</strong> für Teilnehmer — z.B. Bus-/Shuttle-/Bahn-Treffpunkte mit Datum, Abfahrt, Ankunft und optionaler Zusatzinfo (Bus-Kennzeichen, Treffpunkt-Schild, Wagen-Nr.). Pro Stadt ein eigener Eintrag möglich.<br /><br />
                      <strong>Anzeige in der App:</strong> erscheint als <strong>eigener Block</strong> auf der Anmelde-Seite und in Meine Events mit allen Details auf einen Blick.<br /><br />
                      <strong>Automatismen:</strong> Transferzeiten gehen <strong>nicht</strong> in den Outlook-Termin (sonst würde der Termin Bus als Konkurrenz-Termin im Kalender blocken). Sie sind nur in der App sichtbar.<br /><br />
                      <strong>Empfehlung:</strong> bei Auswärtsterminen mit organisierter Anreise sehr empfohlen — bei rein lokalen Office-Events nicht nötig.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> <strong>arrival and departure info</strong> for attendees — e.g. bus/shuttle/train pickups with date, departure, arrival and an optional note (bus number, meeting-point sign, carriage no.). One entry per city.<br /><br />
                      <strong>Shown in the app:</strong> shown as a <strong>dedicated block</strong> on the registration page and in My Events with all details at a glance.<br /><br />
                      <strong>Automation:</strong> transfer times do <strong>not</strong> end up in the Outlook event (otherwise the bus trip would clash with the actual event in the calendar). They live only in the app.<br /><br />
                      <strong>Tip:</strong> strongly recommended for off-site events with organised travel — not needed for local office events.
                    </>
                  )} />
                </label>
                {transferTimes.map((tt) => (
                  <div key={tt.id} style={{
                    padding: '12px 14px', marginBottom: 8,
                    background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                    border: '1px solid var(--dex-gray-200)',
                  }}>
                    {/* Zeile 1: Stadt + Treffpunkt + Adresse + Löschen */}
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr auto', gap: 8, marginBottom: 8 }}>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.location')}</label>
                        <input type="text" className="form-input" list={`transfer-locations-${tt.id}`} value={tt.location} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, location: e.target.value } : x))} placeholder="Stadt eingeben..." style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                        <datalist id={`transfer-locations-${tt.id}`}>
                          {locationOptions.filter(o => o !== 'All').map(opt => (
                            <option key={opt} value={opt} />
                          ))}
                        </datalist>
                      </div>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.meetingpoint')}</label>
                        <input type="text" className="form-input" value={tt.meetingPoint || ''} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, meetingPoint: e.target.value } : x))} placeholder="z.B. Flughafen, Hbf..." style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.address')}</label>
                        <input type="text" className="form-input" value={tt.address || ''} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, address: e.target.value } : x))} placeholder="Straße, PLZ Ort" style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                      <div style={{ display: 'flex', alignItems: 'flex-end', paddingBottom: 4 }}>
                        <button type="button" onClick={() => setTransferTimes(transferTimes.filter(x => x.id !== tt.id))} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)', padding: '4px', lineHeight: 1 }} title={t('general.delete')}>
                          <X size={16} />
                        </button>
                      </div>
                    </div>
                    {/* Zeile 2: Datum + Abfahrt + Ankunft + Beschreibung */}
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 2fr', gap: 8 }}>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.date')}</label>
                        <input type="date" className="form-input" value={tt.date} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, date: e.target.value } : x))} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.departure')}</label>
                        <input type="time" className="form-input" value={tt.departureTime} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, departureTime: e.target.value } : x))} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.arrival')}</label>
                        <input type="time" className="form-input" value={tt.arrivalTime} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, arrivalTime: e.target.value } : x))} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.desc')}</label>
                        <input type="text" className="form-input" value={tt.description || ''} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, description: e.target.value } : x))} placeholder={t('create.transfers.desc')} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                    </div>
                  </div>
                ))}
                <button type="button" className="btn btn-outline" onClick={() => setTransferTimes([...transferTimes, { id: `tr-${Date.now()}`, location: '', meetingPoint: '', address: '', date: startDate ? startDate.slice(0, 10) : '', departureTime: '', arrivalTime: '', description: '' }])} style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4 }}>
                  <Plus size={14} /> {t('create.transfers.add')}
                </button>
              </div>

              </div>{/* v15.6: close hauptGreyoutWrapperStyle div (Step 3) */}
              </div>{/* v15.0: close activeLocationTabIdx===0 wrapper (Top-Level Ort/Adresse/Agenda/Transfer) */}

              </div>{/* close Step 3 (Ort & Programm) */}

              {/* ===== Step 2 (v15.0, vormals Step 3): Sub-Events =====
                  Sub-Events (Workshops / Sessions / Programmpunkte), plus
                  Bezeichnungs-Dropdown und Anmelde-Modus (Hauptevent +
                  Sub-Events vs. nur Sub-Events).
                  v15.0: vorgezogen vor „Ort & Programm", damit die folgenden
                  Steps pro-Sub-Event-Tabs anbieten koennen. */}
              <div style={{ display: currentStep === 1 ? 'block' : 'none' }}>
              <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                {isDe ? 'Schritt 2 — Sub-Events' : 'Step 2 — Sub-events'}
              </h2>
              <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                {isDe
                  ? <><strong>Optional</strong> — lege zusätzliche Sessions, Workshops oder Programmpunkte zum Hauptevent an. Hier legst du auch fest, wie diese Bausteine in der App heißen und ob sich Teilnehmer nur für Sub-Events oder zusätzlich auch fürs Hauptevent anmelden können.</>
                  : <><strong>Optional</strong> — add additional sessions, workshops or program items to the main event. Here you also configure what these building blocks are called in the app and whether attendees register only for sub-events or for the main event as well.</>}
              </p>

              {/* Bezeichnungs-Dropdown */}
              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '14px 16px', marginBottom: 12,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <label className="form-label" style={{ fontSize: '0.95rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8 }}>
                  {isDe ? 'Bezeichnung der Sub-Events' : 'Sub-event naming'}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> wie die untergeordneten Bausteine in der App genannt werden. Die <strong>Standardbezeichnung</strong> ist &bdquo;Sub-Event&ldquo;. Du kannst aber z.B. &bdquo;Workshop&ldquo;, &bdquo;Session&ldquo;, &bdquo;Programmpunkt&ldquo; oder &bdquo;Event-Section&ldquo; auswählen — oder eine eigene Bezeichnung (Singular + Plural) eintippen.<br /><br />
                      <strong>Anzeige in der App:</strong> der gewählte Begriff erscheint überall dort, wo bisher &bdquo;Sub-Event(s)&ldquo; stand — z.B. im Anmeldeformular der Teilnehmer (&bdquo;Verfügbare Workshops&ldquo;), in &bdquo;Meine Events&ldquo; und im Admin Center.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> how the child building blocks are named throughout the app. The <strong>default</strong> is &bdquo;Sub-event&ldquo;. You can pick e.g. &bdquo;Workshop&ldquo;, &bdquo;Session&ldquo;, &bdquo;Programmpunkt&ldquo; or &bdquo;Event section&ldquo; — or type your own singular + plural.<br /><br />
                      <strong>Shown in the app:</strong> the chosen term replaces &bdquo;Sub-event(s)&ldquo; everywhere — e.g. in the attendee registration form (&bdquo;Available workshops&ldquo;), in &bdquo;My events&ldquo; and in the admin center.
                    </>
                  )} />
                </label>
                {(() => {
                  const presets = [
                    { key: 'subevent',     singular: isDe ? 'Sub-Event' : 'Sub-event',         plural: isDe ? 'Sub-Events' : 'Sub-events' },
                    { key: 'workshop',     singular: 'Workshop',                                plural: 'Workshops' },
                    { key: 'session',      singular: 'Session',                                 plural: 'Sessions' },
                    { key: 'programmpunkt', singular: isDe ? 'Programmpunkt' : 'Program item', plural: isDe ? 'Programmpunkte' : 'Program items' },
                    { key: 'section',      singular: isDe ? 'Event-Section' : 'Event section', plural: isDe ? 'Event-Sections' : 'Event sections' },
                  ];
                  const matchKey = (() => {
                    // v15.9: customTermMode hat Priorität — wer in den
                    // Custom-Modus geklickt hat bleibt dort, auch wenn
                    // beide Inputs noch leer sind.
                    if (customTermMode) return 'custom';
                    const s = (childTermSingular || '').trim();
                    const p = (childTermPlural || '').trim();
                    if (!s && !p) return 'subevent';
                    const hit = presets.find(x => x.singular === s && x.plural === p);
                    return hit ? hit.key : 'custom';
                  })();
                  return (
                    <>
                      <select
                        className="form-input"
                        value={matchKey}
                        onChange={e => {
                          const k = e.target.value;
                          if (k === 'custom') {
                            // v15.9: Custom-Modus sticky machen, auch ohne
                            // initiale Werte — sonst kippt der Dropdown
                            // sofort zurück auf 'subevent'.
                            setCustomTermMode(true);
                            return;
                          }
                          // Preset gewählt → Custom-Modus aufheben + Werte
                          // aus dem Preset übernehmen.
                          setCustomTermMode(false);
                          const preset = presets.find(x => x.key === k);
                          if (preset) {
                            setChildTermSingular(preset.singular);
                            setChildTermPlural(preset.plural);
                          }
                        }}
                        style={{ marginTop: 6, maxWidth: 360 }}
                      >
                        {presets.map(p => (
                          <option key={p.key} value={p.key}>
                            {p.plural}
                          </option>
                        ))}
                        <option value="custom">{isDe ? 'Eigene Bezeichnung…' : 'Custom term…'}</option>
                      </select>
                      {matchKey === 'custom' && (
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginTop: 10, maxWidth: 480 }}>
                          <div>
                            <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>{isDe ? 'Singular' : 'Singular'}</label>
                            <input
                              type="text"
                              className="form-input"
                              value={childTermSingular}
                              onChange={e => setChildTermSingular(e.target.value)}
                              placeholder={isDe ? 'z.B. Modul' : 'e.g. Module'}
                              style={{ padding: '6px 10px', fontSize: '0.9rem' }}
                            />
                          </div>
                          <div>
                            <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>{isDe ? 'Plural' : 'Plural'}</label>
                            <input
                              type="text"
                              className="form-input"
                              value={childTermPlural}
                              onChange={e => setChildTermPlural(e.target.value)}
                              placeholder={isDe ? 'z.B. Module' : 'e.g. Modules'}
                              style={{ padding: '6px 10px', fontSize: '0.9rem' }}
                            />
                          </div>
                        </div>
                      )}
                    </>
                  );
                })()}
              </div>

              {/* Anmelde-Modus */}
              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '14px 16px', marginBottom: 16,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <label className="form-label" style={{ fontSize: '0.95rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8 }}>
                  {isDe ? 'Anmelde-Modus' : 'Registration mode'}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> ob sich Teilnehmer zusätzlich zum Hauptevent für Sub-Events anmelden können (Standard) oder ob es <strong>überhaupt kein Hauptevent-Anmelden</strong> mehr gibt und die Anmeldung ausschließlich über einzelne Sub-Events läuft.<br /><br />
                      <strong>Anzeige in der App:</strong> im Modus &bdquo;Nur Sub-Events&ldquo; ist die Anmelde-Checkbox für das Hauptevent im Teilnehmerformular ausgeblendet — der Teilnehmer muss zwingend mindestens einen Sub-Event auswählen. Im Schritt 7 (Kommunikation) wird der Haupt-Event-Tab ausgegraut, weil die Hauptevent-Kommunikation in diesem Modus nicht greift.<br /><br />
                      <strong>Empfehlung:</strong> nutze &bdquo;Nur Sub-Events&ldquo; für Mehrtages-Programme, in denen jeder Teilnehmer aus einem Pool von Slots wählt und ein Hauptevent-Slot keinen Sinn ergibt.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> whether attendees can additionally register for sub-events alongside the main event (default) or whether there is <strong>no main-event registration at all</strong> and registration runs exclusively via individual sub-events.<br /><br />
                      <strong>Shown in the app:</strong> in &bdquo;Sub-events only&ldquo; mode the main-event registration checkbox in the attendee form is hidden — the attendee must pick at least one sub-event. In step 7 (Communication) the main-event tab is greyed out, because main-event communication does not apply in this mode.<br /><br />
                      <strong>Tip:</strong> use &bdquo;Sub-events only&ldquo; for multi-day programmes where every attendee picks from a pool of slots and a main-event slot makes no sense.
                    </>
                  )} />
                </label>
                {/* v15.3.1: Custom-styled Radio-Cards in echtem Deloitte-Grün —
                    weil Edge mit accentColor uneinheitlich rendert (mal solid,
                    mal hollow), hier explizite Visual-Komposition: nativer
                    Input visually-hidden + eigener grüner Outer-Border-Kreis
                    mit grünem Inner-Dot wenn ausgewählt. Funktioniert 1:1 in
                    allen Browsern + matched die Optik der Registration-Page. */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 8 }}>
                  {[false, true].map(modeVal => {
                    const selected = !!subEventsOnlyMode === modeVal;
                    return (
                      <label
                        key={String(modeVal)}
                        style={{
                          display: 'flex', alignItems: 'flex-start', gap: 10,
                          padding: '10px 14px', borderRadius: 8,
                          border: `1px solid ${selected ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                          background: selected ? 'rgba(134,188,37,0.06)' : '#fff',
                          cursor: 'pointer',
                          transition: 'border-color 0.15s, background 0.15s',
                        }}
                      >
                        {/* Hidden native radio for accessibility + form-state */}
                        <input
                          type="radio"
                          name="subEventsMode"
                          checked={selected}
                          onChange={() => setSubEventsOnlyMode(modeVal)}
                          style={{
                            position: 'absolute', opacity: 0, pointerEvents: 'none',
                            width: 1, height: 1, margin: -1, padding: 0,
                            border: 0, overflow: 'hidden', clip: 'rect(0 0 0 0)',
                          }}
                        />
                        {/* Visual custom radio */}
                        <span
                          aria-hidden="true"
                          style={{
                            display: 'inline-block',
                            width: 18, height: 18, borderRadius: '50%',
                            border: `2px solid ${selected ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-400, #9aa0a6)'}`,
                            background: '#fff',
                            position: 'relative', flexShrink: 0,
                            marginTop: 2,
                            transition: 'border-color 0.15s',
                          }}
                        >
                          {selected && (
                            <span style={{
                              position: 'absolute', inset: 3,
                              borderRadius: '50%',
                              background: 'var(--dex-green, #86bc25)',
                            }} />
                          )}
                        </span>
                        <span style={{ fontSize: '0.88rem', flex: 1 }}>
                          {modeVal === false
                            ? (isDe
                                ? <>Anmeldung für <strong>Hauptevent + {(childTermPlural || 'Sub-Events').trim() || 'Sub-Events'}</strong> (Standard)</>
                                : <>Registration for <strong>main event + {(childTermPlural || 'sub-events').trim() || 'sub-events'}</strong> (default)</>)
                            : (isDe
                                ? <>Nur für <strong>{(childTermPlural || 'Sub-Events').trim() || 'Sub-Events'}</strong> (kein Hauptevent-Anmelden)</>
                                : <>Only for <strong>{(childTermPlural || 'sub-events').trim() || 'sub-events'}</strong> (no main-event registration)</>)}
                        </span>
                      </label>
                    );
                  })}
                </div>
              </div>

                {/* ===== Sub-Events (z.B. Workshop-Tage, Networking-Dinner, Kick-off-Sessions) ===== */}
                <div className="form-group" style={{ marginTop: 0 }}>
                  <label className="form-label" style={{ fontSize: '1rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8 }}>
                    <StepBadge n={13} />
                    {/* v15.5: dynamische Bezeichnung — verwendet den oben
                        gewählten Plural-Term statt fix „Sub-Events". */}
                    {(childTermPlural || (isDe ? 'Sub-Events' : 'Sub-events'))} {isDe ? '(optional)' : '(optional)'}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Was du hier einstellst:</strong> <strong>zusätzliche Sessions</strong> zum Hauptevent — z.B. eine Trainingsreihe, optionale Workshops, Side-Events am Vortag. Pro Session ein eigener Eintrag mit Titel, Ort, Start/Ende, Anmeldeschluss und Kapazität.<br /><br />
                        <strong>Anzeige in der App:</strong> Teilnehmer sehen die Sessions auf der Anmelde-Seite als <strong>eigene Anmelde-Bereiche</strong> — Haupt-Event und Sessions können <strong>unabhängig voneinander</strong> an- und abgewählt werden. Niemand muss zwingend das Haupt-Event mitbuchen, um sich für eine Session anzumelden.<br /><br />
                        <strong>Automatismen:</strong> pro Session-An-/Abmeldung gibt es eine <strong>eigene Bestätigungs-Mail</strong> und einen <strong>eigenen Outlook-Termin</strong> (im Deloitte-Layout). Pro Session optional ein- oder ausschaltbar (Mails / Outlook).<br /><br />
                        <strong>Empfehlung:</strong> nutze Sub-Events für mehrtägige Trainings, optionale Add-on-Workshops, oder regelmäßige Side-Sessions zu einem Haupt-Event. Bei einfachen Single-Day-Events nicht nötig.
                      </>
                    ) : (
                      <>
                        <strong>What you set here:</strong> <strong>additional sessions</strong> attached to the main event — e.g. a training series, optional workshops, side events on the day before. One entry per session with title, location, start/end, registration cutoff and capacity.<br /><br />
                        <strong>Shown in the app:</strong> attendees see the sessions on the registration page as <strong>independent registration blocks</strong> — main event and sessions can be picked or skipped <strong>independently</strong>. Nobody has to book the main event to sign up for a session.<br /><br />
                        <strong>Automation:</strong> each session registration / cancellation triggers its <strong>own confirmation mail</strong> and its <strong>own Outlook event</strong> (in the Deloitte template). Per session optionally toggleable (mails / Outlook).<br /><br />
                        <strong>Tip:</strong> use sub-events for multi-day trainings, optional add-on workshops, or recurring side sessions of a main event. Not needed for simple single-day events.
                      </>
                    )} />
                  </label>
                  {subEvents.length === 0 && (
                    <div style={{
                      padding: 10, border: '1px dashed var(--dex-gray-300)', borderRadius: 'var(--dex-radius)',
                      color: 'var(--dex-gray-500)', fontSize: '0.82rem', marginBottom: 8, marginTop: 4,
                      textAlign: 'center', background: 'var(--dex-gray-50, #fafafa)',
                    }}>
                      {t('create.subevents.empty')}
                    </div>
                  )}
                  {subEvents.map((se, idx) => {
                    // SubEvent-Daten werden intern als UTC-ISO gespeichert, fuer die
                    // react-datepicker-Komponenten brauchen wir Date-Objekte mit den
                    // richtigen Berlin-Lokalzeiten. Wir parsen via isoToLocal, was den
                    // Berlin-Wert als "YYYY-MM-DDTHH:MM" liefert, und bauen daraus ein
                    // JavaScript-Date-Objekt mit lokalen Werten.
                    const startDateObj = se.startDate ? (() => {
                      const local = isoToLocal(se.startDate); // "YYYY-MM-DDTHH:MM" in Berlin
                      if (!local) return null;
                      const [dp, tp] = local.split('T');
                      const [y, mo, da] = dp.split('-').map(n => parseInt(n, 10));
                      const [h, mi] = (tp || '00:00').split(':').map(n => parseInt(n, 10));
                      return new Date(y, mo - 1, da, h, mi, 0, 0);
                    })() : null;
                    const endDateObj = se.endDate ? (() => {
                      const local = isoToLocal(se.endDate);
                      if (!local) return null;
                      const [dp, tp] = local.split('T');
                      const [y, mo, da] = dp.split('-').map(n => parseInt(n, 10));
                      const [h, mi] = (tp || '00:00').split(':').map(n => parseInt(n, 10));
                      return new Date(y, mo - 1, da, h, mi, 0, 0);
                    })() : null;
                    // v15: deadlineObj entfernt — der Anmeldeschluss-Editor
                    // wandert nach Schritt 4 (Kapazität) in den Sub-Event-Tab.
                    const dateToBerlinIso = (d: Date | null): string => {
                      if (!d) return '';
                      const local = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}T${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
                      return berlinLocalToUtcIso(local);
                    };
                    return (
                      <div key={se.id} style={{
                        padding: '14px 16px', marginBottom: 10, marginTop: 4,
                        background: 'var(--dex-gray-50, #fafafa)', borderRadius: 'var(--dex-radius)',
                        border: '1px solid var(--dex-gray-200)', borderLeft: '3px solid var(--dex-green, #86bc25)',
                      }}>
                        {/* Header-Zeile: Titel (prominent) + Loeschen */}
                        <div style={{ display: 'flex', gap: 10, alignItems: 'flex-start', marginBottom: 10 }}>
                          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 2 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.subevents.title')}</label>
                            <input
                              type="text"
                              className="form-input"
                              value={se.title}
                              placeholder={t('create.subevents.title.placeholder')}
                              onChange={e => {
                                const v = e.target.value;
                                setSubEvents(subEvents.map((x, i) => i === idx ? { ...x, title: v } : x));
                              }}
                              style={{ padding: '6px 10px', fontSize: '0.9rem', fontWeight: 600 }}
                            />
                          </div>
                          <button
                            type="button"
                            onClick={() => setSubEvents(subEvents.filter(x => x.id !== se.id))}
                            style={{
                              background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)',
                              fontSize: '1.1rem', padding: '4px', lineHeight: 1, marginTop: 18,
                            }}
                            title={t('create.subevents.remove')}
                          >
                            <X size={16} />
                          </button>
                        </div>

                        {/* Zeit-Zeile: Start + Ende + Anmeldeschluss.
                            v15.0: „Max. Teilnehmer" entfaellt aus dieser
                            Karte — Kapazitaet wird jetzt in Schritt 4
                            (Kapazitaet) pro Sub-Event-Tab gepflegt. */}
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10, marginBottom: 10 }}>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.subevents.start')}</label>
                            <DatePicker
                              selected={startDateObj}
                              onChange={(date: Date | null) => {
                                const iso = dateToBerlinIso(date);
                                setSubEvents(subEvents.map((x, i) => i === idx ? { ...x, startDate: iso } : x));
                              }}
                              showTimeSelect
                              timeFormat="HH:mm"
                              timeIntervals={15}
                              timeCaption="Uhrzeit"
                              dateFormat="dd.MM.yyyy, HH:mm"
                              locale="de"
                              placeholderText="Datum und Uhrzeit"
                              className="form-input"
                              wrapperClassName="dex-datepicker-wrapper"
                              calendarClassName="dex-datepicker-calendar"
                              popperPlacement="bottom-start"
                              maxDate={endDateObj || undefined}
                              isClearable
                              autoComplete="off"
                            />
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.subevents.end')}</label>
                            <DatePicker
                              selected={endDateObj}
                              onChange={(date: Date | null) => {
                                const iso = dateToBerlinIso(date);
                                setSubEvents(subEvents.map((x, i) => i === idx ? { ...x, endDate: iso } : x));
                              }}
                              showTimeSelect
                              timeFormat="HH:mm"
                              timeIntervals={15}
                              timeCaption="Uhrzeit"
                              dateFormat="dd.MM.yyyy, HH:mm"
                              locale="de"
                              placeholderText="Datum und Uhrzeit"
                              className="form-input"
                              wrapperClassName="dex-datepicker-wrapper"
                              calendarClassName="dex-datepicker-calendar"
                              popperPlacement="bottom-start"
                              minDate={startDateObj || undefined}
                              isClearable
                              autoComplete="off"
                            />
                          </div>
                          {/* v15: Anmeldeschluss raus aus der Sub-Event-Card —
                              wird jetzt in Schritt 4 (Kapazität & Sichtbarkeit)
                              pro Sub-Event-Tab gepflegt (analog zur Kapazität
                              mit „vom Hauptevent übernehmen"-Toggle). */}
                        </div>
                        {/* v18.36: Ende-vor-Start-Hinweis pro Sub-Event. */}
                        {startDateObj && endDateObj && endDateObj <= startDateObj && (
                          <p style={{ color: 'var(--dex-red, #c00)', fontSize: '0.8rem', margin: '-4px 0 8px' }}>
                            {isDe
                              ? 'Das Enddatum dieses Sub-Events liegt vor dem Startdatum — bitte korrigieren.'
                              : 'The end date of this sub-event is before the start date — please correct it.'}
                          </p>
                        )}

                        {/* Beschreibung. v15.0: „Ort" entfaellt aus dieser
                            Karte — wird in Schritt 3 (Ort & Programm) pro
                            Sub-Event-Tab gepflegt. */}
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 10, marginBottom: 10 }}>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.subevents.description')}</label>
                            <input
                              type="text"
                              className="form-input"
                              value={se.description || ''}
                              onChange={e => {
                                const v = e.target.value;
                                setSubEvents(subEvents.map((x, i) => i === idx ? { ...x, description: v } : x));
                              }}
                              style={{ padding: '6px 10px', fontSize: '0.85rem' }}
                            />
                          </div>
                        </div>

                        {/* v15: Mail- und Outlook-Toggles raus aus der Sub-Event-
                            Card — sie leben jetzt ausschliesslich in Schritt 6
                            (Kommunikation) pro Sub-Event-Tab. Hier bleibt nur
                            das absolut Notwendige zur Anlage des Sub-Events. */}
                      </div>
                    );
                  })}
                  <button
                    type="button"
                    className="btn btn-outline"
                    style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4 }}
                    onClick={() => {
                      const newSub: SubEventDraft = {
                        id: (typeof crypto !== 'undefined' && crypto.randomUUID) ? crypto.randomUUID() : `se_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
                        title: '',
                        description: '',
                        location: '',
                        startDate: '',
                        endDate: '',
                        maxParticipants: 0,
                        disableEmails: false,
                        disableOutlook: false,
                        // v15.3: neue Sub-Events starten leer und vollwertig.
                        // Default-Werte koennen per „Vom Hauptevent kopieren"-
                        // Button in den Steps 3/4/5 einmalig vorbelegt werden.
                        locationAddress: { street: '', houseNo: '', zip: '', city: '' },
                        agenda: [],
                        transferTimes: [],
                        lastDeregisterDate: '',
                        locationFilter: '',
                        audience: '',
                        filterMode: 'OR',
                        waitlistEnabled: true,
                        askSalutation: false,
                      };
                      setSubEvents([...subEvents, newSub]);
                    }}
                  >
                    <Plus size={14} /> {t('create.subevents.add')}
                  </button>
                </div>
              </div>{/* close Step 2 (Sub-Events) */}

              {/* ===== Step 4 (v14.8: vormals Step 3): Kapazität, Fristen & Sichtbarkeit ===== */}
              <div style={{ display: currentStep === 3 ? 'block' : 'none' }}>
              <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                {isDe ? 'Schritt 4 — Kapazität & Sichtbarkeit' : 'Step 4 — Capacity & Visibility'}
              </h2>
              <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                {isDe
                  ? 'Hier legst du fest, wer das Event überhaupt sieht, wie viele Plätze es gibt und bis wann sich Teilnehmer an- bzw. abmelden können.'
                  : 'Here you decide who can see the event in the first place, how many spots there are, and the deadlines for registration and cancellation.'}
              </p>
              {renderStepIntro(
                [
                  'Sichtbarkeit: Standort-Filter und Mailverteiler/User festlegen — wer das Event in der Liste sieht',
                  'Anmelde-Deadline + letzte Abmeldemöglichkeit (vorbefüllt anhand des Event-Datums, jederzeit überschreibbar)',
                  'Maximale Teilnehmerzahl festlegen (oder Unbegrenzt)',
                  'Warteliste aktivieren — voll besetzte Events nehmen weitere Anmeldungen auf, bis ein Platz frei wird',
                  'Optional: Geteilte Kapazität — zwei frei benannte Gruppen mit eigener Platzzahl + eigener oder gemeinsamer Warteliste',
                ],
                [
                  'Visibility: configure location filter + mailing lists/individual users — who sees the event in the list',
                  'Registration deadline + last cancellation date (pre-filled from the event date, always overridable)',
                  'Set the maximum number of attendees (or Unlimited)',
                  'Enable waitlist — full events accept new registrations and promote them once a spot frees up',
                  'Optional: split capacity — two freely-named groups with own seat count + own or shared waitlist',
                ]
              )}

              {/* v15.0: pro-Sub-Event-Tabs fuer Kapazitaet. Tab 0 = Haupt-
                  Event (komplette Sichtbarkeit/Deadlines/MaxParticipants/
                  Split-UI). Tabs N>0 = schlanke MaxParticipants-only-UI pro
                  Sub-Event mit Inheritance-Toggle. Sichtbarkeit, Filter,
                  Deadlines, Split-Capacity bleiben Top-Level — pro Sub-Event
                  ist nur die Platzzahl relevant. */}
              {renderPerEventTabStrip(
                activeCapacityTabIdx,
                setActiveCapacityTabIdx,
                `${isDe ? 'Haupt-Event' : 'Main event'}: ${title || (isDe ? 'Ohne Titel' : 'Untitled')}`,
                isDe ? 'Event-Tab wechseln (Kapazität)' : 'Switch event tab (capacity)'
              )}

              {activeCapacityTabIdx > 0 && (() => {
                const seIdx = activeCapacityTabIdx - 1;
                const se = subEvents[seIdx];
                if (!se) return null;
                const updateSub = (patch: Partial<SubEventDraft>): void => {
                  setSubEvents(prev => prev.map((x, i) => i === seIdx ? { ...x, ...patch } : x));
                };
                const seLocationFilterList = (se.locationFilter || '').split(',').map(s => s.trim()).filter(Boolean);
                return (
                  <div>
                    {/* v15.3: „Vom Hauptevent kopieren"-Button. Uebernimmt
                        Kapazitaets-/Sichtbarkeits-/Deadline-/Filter-Werte
                        vom Hauptevent als Startwerte fuer dieses Sub-Event. */}
                    <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }}>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.8rem', padding: '6px 14px' }}
                        onClick={() => updateSub({
                          maxParticipants: parseInt(maxParticipants, 10) || 0,
                          registrationDeadline: registrationDeadline ? berlinLocalToUtcIso(registrationDeadline) : '',
                          lastDeregisterDate: lastDeregisterDate ? berlinLocalToUtcIso(lastDeregisterDate) : '',
                          locationFilter: locationFilter,
                          audience: audience,
                          filterMode: filterMode,
                          waitlistEnabled: waitlistEnabled,
                        })}
                        title={isDe
                          ? 'Uebernimmt Teilnehmerzahl, Deadlines, Sichtbarkeit und Warteliste vom Hauptevent als Startwerte'
                          : 'Copies capacity, deadlines, visibility and waitlist from the main event as starting values'}
                      >
                        {isDe ? 'Vom Hauptevent kopieren' : 'Copy from main event'}
                      </button>
                    </div>

                    {/* v15.6: Sichtbarkeits-Sektion analog Hauptevent —
                        Header mit Auge-Icon plus erklärender Lead-Text. */}
                    <div style={{
                      paddingBottom: 12, marginBottom: 16,
                      borderBottom: '2px solid var(--dex-gray-100)',
                    }}>
                      <h3 style={{ margin: '0 0 6px', fontSize: '1.05rem', display: 'flex', alignItems: 'center', gap: 8 }}>
                        <Icon iconName="Hide3" style={{ fontSize: 18, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                        {isDe ? 'Sichtbarkeit' : 'Visibility'}
                      </h3>
                      <p style={{ margin: 0, fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                        {isDe
                          ? <>Auch dieses Sub-Event ist standardmäßig für <strong>alle Mitarbeiter von Deloitte Deutschland</strong> sichtbar. Über Standortfilter und Mailverteiler kannst du den Empfängerkreis gezielt einschränken — unabhängig vom Hauptevent.</>
                          : <>By default this sub-event is visible for <strong>all Deloitte Germany employees</strong>. Use the location filter and mailing lists below to restrict the audience — independent from the main event.</>}
                      </p>
                    </div>

                    <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={13} />
                        {isDe ? 'Standortfilter' : 'Location filter'}
                      </label>
                      <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                        {isDe
                          ? <>Wählst du hier einen oder mehrere Standorte aus, sehen <strong>nur Mitarbeiter mit diesen Standorten</strong> dieses Sub-Event. Leer = für alle sichtbar.</>
                          : <>If you pick one or more locations here, <strong>only employees from those locations</strong> will see this sub-event. Empty = visible to everyone.</>}
                      </p>
                      <LocationMultiSelect
                        options={locationOptions}
                        selected={seLocationFilterList}
                        onChange={list => updateSub({ locationFilter: list.join(', ') })}
                        isDe={isDe}
                      />
                    </div>

                    <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={14} />
                        {isDe ? 'Mailverteiler / einzelne User' : 'Mailing lists / individual users'}
                      </label>
                      <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                        {isDe
                          ? <>Mail-Adressen oder Gruppen, kommagetrennt. Leer = keine zusätzliche Einschränkung. Du kannst auch direkt eine Verteiler-Mail eintippen (z.B. <code>SAPAlliance@deloitte.com</code>) oder Sondergruppen wie <code>DEALL</code>, <code>DEKOELN</code>.</>
                          : <>Email addresses or groups, comma-separated. Empty = no extra restriction. You can also type a distribution list directly (e.g. <code>SAPAlliance@deloitte.com</code>) or special groups like <code>DEALL</code>, <code>DEKOELN</code>.</>}
                      </p>
                      <textarea
                        className="form-input"
                        value={se.audience || ''}
                        onChange={e => updateSub({ audience: e.target.value })}
                        placeholder={isDe ? 'z.B. max@deloitte.de, DEKOELN' : 'e.g. max@deloitte.de, DEKOELN'}
                        rows={2}
                        style={{ resize: 'vertical' }}
                      />
                    </div>

                    {(seLocationFilterList.length > 0 && (se.audience || '').trim().length > 0) && (
                      <div className="form-group" style={{ padding: '16px 20px 16px 30px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                        <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                          <StepBadge n={15} />
                          {isDe ? 'Filterverknüpfung' : 'Filter combination'}
                        </label>
                        <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginTop: -4, marginBottom: 12, lineHeight: 1.55 }}>
                          {isDe
                            ? <>Beide Filter sind gesetzt — bestimmt, ob für eine Person <strong>einer</strong> der Filter (ODER) oder <strong>beide</strong> (UND) zutreffen müssen, damit das Sub-Event in ihrer Liste auftaucht.</>
                            : <>Both filters are set — defines whether a person needs to match <strong>either</strong> filter (OR) or <strong>both</strong> filters (AND) for the sub-event to appear in their list.</>}
                        </p>
                        <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
                          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.9rem', cursor: 'pointer' }}>
                            <input
                              type="radio"
                              name={`subFilterMode-${se.id}`}
                              checked={(se.filterMode || 'OR') === 'OR'}
                              onChange={() => updateSub({ filterMode: 'OR' })}
                            />
                            <strong>{isDe ? 'ODER' : 'OR'}</strong>
                            <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>– {isDe ? 'Einer der Filter reicht' : 'one filter is enough'}</span>
                          </label>
                          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.9rem', cursor: 'pointer' }}>
                            <input
                              type="radio"
                              name={`subFilterMode-${se.id}`}
                              checked={se.filterMode === 'AND'}
                              onChange={() => updateSub({ filterMode: 'AND' })}
                            />
                            <strong>{isDe ? 'UND' : 'AND'}</strong>
                            <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>– {isDe ? 'Beides muss zutreffen' : 'both must match'}</span>
                          </label>
                        </div>
                      </div>
                    )}

                    {/* Deadlines: zwei DatePicker nebeneinander, gleicher Look
                        wie im Hauptevent. */}
                    <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={(seLocationFilterList.length > 0 && (se.audience || '').trim().length > 0) ? 16 : 15} />
                        {isDe ? 'Anmelde- und Abmeldefristen' : 'Registration & cancellation deadlines'}
                      </label>
                      <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                        {isDe
                          ? <>Frei pro Sub-Event setzbar. Leer lassen → die Fristen des Hauptevents gelten.</>
                          : <>Settable per sub-event. Leave empty → the main event’s deadlines apply.</>}
                      </p>
                      <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                        <div className="form-group" style={{ marginBottom: 0 }}>
                          <label className="form-label">{t('create.deadline')}</label>
                          <DatePicker
                            selected={se.registrationDeadline ? new Date(se.registrationDeadline) : null}
                            onChange={(date: Date | null) => {
                              if (!date) { updateSub({ registrationDeadline: '' }); return; }
                              const local = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
                              updateSub({ registrationDeadline: berlinLocalToUtcIso(local) });
                            }}
                            showTimeSelect
                            timeFormat="HH:mm"
                            timeIntervals={15}
                            timeCaption="Uhrzeit"
                            dateFormat="dd.MM.yyyy, HH:mm"
                            locale="de"
                            placeholderText={isDe ? 'Anmelde-Deadline' : 'Registration deadline'}
                            className="form-input"
                            wrapperClassName="dex-datepicker-wrapper"
                            calendarClassName="dex-datepicker-calendar"
                            popperPlacement="bottom-start"
                            isClearable
                            autoComplete="off"
                          />
                        </div>
                        <div className="form-group" style={{ marginBottom: 0 }}>
                          <label className="form-label">{t('create.lastcancel')}</label>
                          <DatePicker
                            selected={se.lastDeregisterDate ? new Date(se.lastDeregisterDate) : null}
                            onChange={(date: Date | null) => {
                              if (!date) { updateSub({ lastDeregisterDate: '' }); return; }
                              const local = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
                              updateSub({ lastDeregisterDate: berlinLocalToUtcIso(local) });
                            }}
                            showTimeSelect
                            timeFormat="HH:mm"
                            timeIntervals={15}
                            timeCaption="Uhrzeit"
                            dateFormat="dd.MM.yyyy, HH:mm"
                            locale="de"
                            placeholderText={isDe ? 'Abmeldefrist' : 'Last cancellation'}
                            className="form-input"
                            wrapperClassName="dex-datepicker-wrapper"
                            calendarClassName="dex-datepicker-calendar"
                            popperPlacement="bottom-start"
                            isClearable
                            autoComplete="off"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Teilnehmerzahl & Warteliste — analog Hauptevent als
                        eine kombinierte Card mit „Unbegrenzt"-Toggle und
                        Warteliste-Toggle. Split-Capacity bleibt Hauptevent-only
                        (Scope-Eingrenzung, siehe v15.6 Refactor-Plan). */}
                    <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={(seLocationFilterList.length > 0 && (se.audience || '').trim().length > 0) ? 17 : 16} />
                        {isDe ? 'Teilnehmerzahl & Warteliste' : 'Capacity & waitlist'}
                      </label>
                      <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                        <div className="form-group" style={{ marginBottom: 0 }}>
                          <label className="form-label">
                            {isDe ? 'Max. Teilnehmer (0 = unbegrenzt)' : 'Max. attendees (0 = unlimited)'}
                          </label>
                          <div className="toggle-wrapper" style={{ marginTop: 4, marginBottom: 8 }}>
                            <label className="toggle">
                              <input
                                type="checkbox"
                                checked={(se.maxParticipants || 0) === 0}
                                onChange={e => {
                                  if (e.target.checked) {
                                    updateSub({ maxParticipants: 0, waitlistEnabled: false });
                                  } else {
                                    updateSub({ maxParticipants: 50 });
                                  }
                                }}
                              />
                              <span className="toggle-slider" />
                            </label>
                            <span style={{ fontSize: '0.9rem' }}>
                              {(se.maxParticipants || 0) === 0
                                ? (isDe ? 'Unbegrenzt' : 'Unlimited')
                                : (`${se.maxParticipants} ${isDe ? 'Plätze' : 'seats'}`)}
                            </span>
                          </div>
                          {(se.maxParticipants || 0) > 0 && (
                            <input
                              type="number"
                              min={0}
                              className="form-input"
                              value={se.maxParticipants || 0}
                              onChange={e => {
                                const v = parseInt(e.target.value, 10) || 0;
                                updateSub({ maxParticipants: v });
                              }}
                              placeholder={isDe ? 'Anzahl' : 'Count'}
                            />
                          )}
                        </div>
                        {(se.maxParticipants || 0) > 0 && (
                          <div className="form-group" style={{ marginBottom: 0 }}>
                            <label className="form-label">{t('create.waitlist')}</label>
                            <div className="toggle-wrapper" style={{ marginTop: 8 }}>
                              <label className="toggle">
                                <input
                                  type="checkbox"
                                  checked={typeof se.waitlistEnabled === 'boolean' ? se.waitlistEnabled : true}
                                  onChange={e => updateSub({ waitlistEnabled: e.target.checked })}
                                />
                                <span className="toggle-slider" />
                              </label>
                              <span style={{ fontSize: '0.9rem' }}>
                                {(typeof se.waitlistEnabled === 'boolean' ? se.waitlistEnabled : true) ? t('create.enabled') : t('create.disabled')}
                              </span>
                            </div>
                          </div>
                        )}
                      </div>
                      <p style={{ margin: '12px 0 0', fontSize: '0.75rem', color: 'var(--dex-gray-500)', lineHeight: 1.5 }}>
                        {isDe
                          ? <em>Hinweis: <strong>Geteilte Kapazität</strong> (zwei Gruppen mit eigener Platzzahl) ist aktuell nur auf Hauptevent-Ebene möglich — Sub-Events nutzen die einfache Gesamtkapazität.</em>
                          : <em>Note: <strong>Split capacity</strong> (two groups with separate seat counts) is currently main-event-only — sub-events use the simple total capacity.</em>}
                      </p>
                    </div>
                  </div>
                );
              })()}

              <div style={{ display: activeCapacityTabIdx === 0 ? 'block' : 'none' }}>
              {renderHauptGreyoutBanner()}
              <div style={hauptGreyoutWrapperStyle()}>

              {/* v9.24: Sichtbarkeits-Steuerungen aus Step 0 hierher verschoben.
                  Die Frage 'wer darf das Event sehen' passt logisch zu Kapazitaet/Fristen
                  als 'Wer-Wann-Wieviel' und entlastet Step 0 (Grundlagen). */}
              {/* Zwischenüberschrift: alle Sichtbarkeits-Steuerungen
                  (Standortfilter + Mailverteiler/einzelne User) gruppieren,
                  damit der Organizer auf einen Blick versteht, dass es hier
                  um die Frage geht: wer darf das Event ueberhaupt sehen. */}
              <div style={{
                paddingBottom: 12, marginBottom: 16,
                borderBottom: '2px solid var(--dex-gray-100)',
              }}>
                <h3 style={{ margin: '0 0 6px', fontSize: '1.05rem', display: 'flex', alignItems: 'center', gap: 8 }}>
                  <Icon iconName="Hide3" style={{ fontSize: 18, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                  {isDe ? 'Sichtbarkeit' : 'Visibility'}
                </h3>
                <p style={{ margin: 0, fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                  {isDe ? (
                    <>
                      Dieses Event ist standardmäßig für <strong>alle Mitarbeiter von Deloitte Deutschland</strong>{' '}
                      sichtbar und buchbar. Über die folgenden beiden Bereiche — Standortfilter sowie
                      Mailverteiler / einzelne User — kannst du den Empfängerkreis gezielt einschränken.
                      Mitarbeiter außerhalb der definierten Auswahl sehen das Event nicht in ihrer
                      Übersicht und können sich entsprechend nicht anmelden.
                    </>
                  ) : (
                    <>
                      By default, this event is visible and bookable for <strong>all Deloitte Germany employees</strong>.
                      You can narrow down the audience using the two sections below — the location filter
                      and mailing lists / individual users. Employees outside the selected scope will
                      not see the event in their overview and cannot register for it.
                    </>
                  )}
                </p>
              </div>

              <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={13} />
                  {isDe ? 'Standortfilter' : 'Location filter'}
                </label>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                  {isDe ? (
                    <>
                      Wählst du hier einen oder mehrere Standorte aus, sehen <strong>nur Mitarbeiter mit diesen Standorten</strong> das Event.<br />
                      <em>Beispiel: &bdquo;Köln&ldquo; und &bdquo;Düsseldorf&ldquo; → Nur Mitarbeiter mit diesen Standorten sehen das Event. Alle anderen sehen es nicht.</em>
                    </>
                  ) : (
                    <>
                      If you pick one or more locations here, <strong>only employees from those locations</strong> will see the event.<br />
                      <em>Example: &bdquo;Cologne&ldquo; and &bdquo;Düsseldorf&ldquo; → only employees with one of these locations will see the event. Everyone else will not.</em>
                    </>
                  )}
                </p>
                {/* Multi-Select-Dropdown — kompakter als die alten Pillen,
                    erlaubt Suche + Mehrfachauswahl. Aktuelle Auswahl wird
                    direkt im Trigger-Button als Chip-Liste angezeigt. */}
                <LocationMultiSelect
                  options={locationOptions}
                  selected={locationFilter.split(',').map(s => s.trim()).filter(Boolean)}
                  onChange={list => setLocationFilter(list.join(', '))}
                  isDe={isDe}
                />
                {!locationFilter && (
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-green)', marginTop: 8 }}>
                    {isDe
                      ? 'Kein Standort ausgewählt → Event ist für alle sichtbar.'
                      : 'No location selected → event is visible to everyone.'}
                  </p>
                )}
              </div>

              <div className="form-group" style={{ position: 'relative', padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={14} />
                  {isDe ? 'Mailverteiler / einzelne User' : 'Mailing lists / individual users'}
                </label>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                  {isDe
                    ? <>Wähle <strong>einzelne Personen</strong> oder ganze <strong>Mailverteiler bzw. Security-Gruppen aus Entra</strong> aus. Wenn auch ein Standortfilter gesetzt ist, kannst du unten festlegen, ob beide Bedingungen (UND) oder eine davon (ODER) reichen.</>
                    : <>Pick <strong>individual people</strong> or entire <strong>mailing lists / security groups from Entra</strong>. If you also set a location filter, you can decide below whether both conditions (AND) or either of them (OR) is enough.</>}
                </p>
                {/* v16.4: Hinweis fuer den Organizer, dass Mitglieder zum
                    Save-Zeitpunkt eingefroren werden und das Event bei DL-
                    Mitglieder-Aenderungen einmal neu gespeichert werden muss,
                    damit die neuen Mitglieder die Sichtbarkeit bekommen. */}
                <div style={{
                  fontSize: '0.78rem', color: 'var(--dex-orange-dark, #b35a00)',
                  background: 'rgba(237,139,0,0.08)', border: '1px dashed var(--dex-orange, #ed8b00)',
                  borderRadius: 6, padding: '8px 12px', marginBottom: 12, lineHeight: 1.5,
                }}>
                  {isDe
                    ? <><strong>Hinweis:</strong> Die Mitglieder der ausgewählten Mailverteiler werden beim Speichern des Events einmal aufgelöst und gespeichert — das ist der schnelle Pfad für den Sichtbarkeits-Check. Wenn sich später Mitglieder eines Verteilers ändern (z.B. neue Person zur DL hinzugefügt), <strong>speichere das Event einmal neu</strong>, damit die App den frischen Stand bekommt.</>
                    : <><strong>Note:</strong> The members of the selected distribution lists are resolved and cached when the event is saved — this is the fast path for the visibility check. If list members change later (e.g. new person added to a DL), <strong>re-save the event once</strong> to refresh the cache.</>}
                </div>
                {/* Chip-Liste der bereits ausgewaehlten Audience-Eintraege.
                    Bei vielen Eintraegen: Inline-Suche + Pagination (nur 10 sichtbar, 'Mehr anzeigen'-Button). */}
                {audience.trim().length > 0 && (() => {
                  const allEntries = audience.split(',').map(s => s.trim()).filter(Boolean);
                  const chipSearchLc = audienceChipSearch.trim().toLowerCase();
                  const filtered = chipSearchLc
                    ? allEntries.filter(e => e.toLowerCase().indexOf(chipSearchLc) >= 0)
                    : allEntries;
                  const visibleLimit = 10;
                  const visible = audienceShowAll || chipSearchLc ? filtered : filtered.slice(0, visibleLimit);
                  const hiddenCount = filtered.length - visible.length;
                  return (
                    <div style={{ marginBottom: 8 }}>
                      {/* Meta-Zeile mit Anzahl + Such-Input (nur wenn viele Eintraege) */}
                      {allEntries.length > visibleLimit && (
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
                          <span>{allEntries.length} Einträge{chipSearchLc && ` — ${filtered.length} Treffer`}</span>
                          <input
                            type="text"
                            className="form-input"
                            value={audienceChipSearch}
                            onChange={e => setAudienceChipSearch(e.target.value)}
                            placeholder="In Zielgruppe suchen..."
                            style={{ flex: 1, maxWidth: 260, fontSize: '0.75rem', padding: '4px 8px' }}
                          />
                          {audienceChipSearch && (
                            <button
                              type="button"
                              onClick={() => setAudienceChipSearch('')}
                              style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 2, color: 'var(--dex-gray-400)' }}
                              title="Suche löschen"
                            >
                              <X size={14} />
                            </button>
                          )}
                        </div>
                      )}
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                        {visible.map((entry, i) => {
                          const isEmail = entry.indexOf('@') >= 0;
                          return (
                            <span key={`${entry}-${i}`} style={{
                              display: 'inline-flex', alignItems: 'center', gap: 6,
                              padding: '4px 10px', borderRadius: 999,
                              background: isEmail ? 'rgba(0,118,168,0.10)' : 'rgba(134,188,37,0.12)',
                              color: isEmail ? 'var(--dex-blue, #0076a8)' : 'var(--dex-green-dark)',
                              fontSize: '0.8rem', fontWeight: 600,
                            }}>
                              {entry}
                              <button
                                type="button"
                                title="Mitglieder anzeigen"
                                onClick={() => openMembersModal(entry)}
                                style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0, display: isEmail ? 'inline-flex' : 'none', color: 'inherit' }}
                              >
                                <Users size={12} />
                              </button>
                              <button
                                type="button"
                                onClick={() => removeAudienceItem(entry)}
                                title="Entfernen"
                                style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0, color: 'inherit', fontWeight: 700 }}
                              >
                                <X size={12} />
                              </button>
                            </span>
                          );
                        })}
                      </div>
                      {/* Mehr / Weniger Button */}
                      {!chipSearchLc && allEntries.length > visibleLimit && (
                        <button
                          type="button"
                          className="btn btn-secondary"
                          style={{ fontSize: '0.75rem', padding: '4px 10px', marginTop: 6 }}
                          onClick={() => setAudienceShowAll(!audienceShowAll)}
                        >
                          {audienceShowAll ? `Weniger anzeigen (${allEntries.length - visibleLimit} ausblenden)` : `Alle anzeigen (+${hiddenCount} weitere)`}
                        </button>
                      )}
                      {chipSearchLc && filtered.length === 0 && (
                        <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', marginTop: 4, fontStyle: 'italic' }}>
                          Kein Treffer für &bdquo;{audienceChipSearch}&ldquo; in der Zielgruppe.
                        </div>
                      )}
                    </div>
                  );
                })()}
                {/* Such-Input */}
                <input
                  className="form-input"
                  value={audienceSearch}
                  onChange={e => {
                    const val = e.target.value;
                    setAudienceSearch(val);
                    if (audienceTimerRef.current) clearTimeout(audienceTimerRef.current);
                    if (val.trim().length >= 2) {
                      audienceTimerRef.current = setTimeout(async () => {
                        setIsSearchingAudience(true);
                        try {
                          const [users, groups] = await Promise.all([
                            searchUsers(val.trim(), audienceIncludeIntl),
                            searchGroups(val.trim()),
                          ]);
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const u: Array<{ kind: 'user' | 'group'; email: string; displayName: string }> = users.map((x: any) => ({ kind: 'user' as const, email: x.email, displayName: x.displayName }));
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const g: Array<{ kind: 'user' | 'group'; email: string; displayName: string }> = groups.map((x: any) => ({ kind: 'group' as const, email: x.email, displayName: x.displayName }));
                          setAudienceResults([...g, ...u]); // Gruppen zuerst anzeigen
                        } catch { setAudienceResults([]); }
                        setIsSearchingAudience(false);
                      }, 300);
                    } else {
                      setAudienceResults([]);
                    }
                  }}
                  placeholder="Personen oder Gruppen suchen (z.B. SAPAlliance, max@deloitte.de, DEKOELN)"
                />
                <InternationalSearchToggle
                  checked={audienceIncludeIntl}
                  onChange={async next => {
                    setAudienceIncludeIntl(next);
                    const q = audienceSearch.trim();
                    if (q.length >= 2) {
                      setIsSearchingAudience(true);
                      try {
                        const [users, groups] = await Promise.all([
                          searchUsers(q, next),
                          searchGroups(q),
                        ]);
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const u: Array<{ kind: 'user' | 'group'; email: string; displayName: string }> = users.map((x: any) => ({ kind: 'user' as const, email: x.email, displayName: x.displayName }));
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const g: Array<{ kind: 'user' | 'group'; email: string; displayName: string }> = groups.map((x: any) => ({ kind: 'group' as const, email: x.email, displayName: x.displayName }));
                        setAudienceResults([...g, ...u]);
                      } catch { setAudienceResults([]); }
                      setIsSearchingAudience(false);
                    }
                  }}
                  isDe={isDe}
                />
                {isSearchingAudience && (
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>Suche...</div>
                )}
                {audienceResults.length > 0 && (
                  <div style={{
                    position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
                    background: '#fff', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                    maxHeight: 280, overflowY: 'auto',
                  }}>
                    {audienceResults.map((r, i) => (
                      <div
                        key={`${r.kind}-${r.email}-${i}`}
                        style={{
                          padding: '8px 12px', cursor: 'pointer', fontSize: '0.85rem',
                          borderBottom: '1px solid var(--dex-gray-100)',
                          display: 'flex', alignItems: 'center', gap: 8,
                        }}
                        onMouseDown={() => {
                          addAudienceItem(r.email);
                          setAudienceSearch('');
                          setAudienceResults([]);
                        }}
                      >
                        <span style={{
                          fontSize: '0.7rem', fontWeight: 700,
                          padding: '2px 8px', borderRadius: 4,
                          background: r.kind === 'group' ? 'rgba(134,188,37,0.18)' : 'rgba(0,118,168,0.14)',
                          color: r.kind === 'group' ? 'var(--dex-green-dark)' : 'var(--dex-blue, #0076a8)',
                        }}>
                          {r.kind === 'group' ? 'GRUPPE' : 'USER'}
                        </span>
                        <strong>{r.displayName}</strong>
                        <span style={{ color: 'var(--dex-gray-400)', marginLeft: 'auto' }}>{r.email}</span>
                      </div>
                    ))}
                  </div>
                )}
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 6 }}>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ fontSize: '0.75rem', padding: '6px 12px', whiteSpace: 'nowrap' }}
                    onClick={() => setBulkAudienceOpen(true)}
                  >
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}><Users size={12} /> Massenimport (Liste einfügen)</span>
                  </button>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', margin: 0, flex: 1 }}>
                    Klicke einen Treffer an um ihn hinzuzufügen. Bei Gruppen kannst du im Chip per <Users size={11} /> die Mitglieder einsehen.
                    Statt zu suchen kannst du auch direkt die Verteiler-Mail eintippen (z.B. SAPAlliance@deloitte.com) oder Sondergruppen wie <code>DEALL</code>, <code>DEKOELN</code>.
                  </p>
                </div>
              </div>

              {/* Filterverknuepfung: nur sichtbar wenn beide Bereiche
                  (Standortfilter + Mailverteiler) Werte haben — sonst gibt
                  es nichts zu kombinieren. */}
              {locationFilter && audience && (
                <div className="form-group" style={{ padding: '16px 20px 16px 30px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                  <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <StepBadge n={15} />
                    {isDe ? 'Filterverknüpfung' : 'Filter combination'}
                  </label>
                  <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginTop: -4, marginBottom: 12, lineHeight: 1.6 }}>
                    {isDe ? (
                      <>
                        <p style={{ margin: '0 0 8px' }}>
                          Bestimmt, wie der <strong>Standortfilter</strong> und der <strong>Mailverteiler / einzelne User</strong> miteinander kombiniert werden — also welche Bedingungen für eine Person erfüllt sein müssen, damit sie das Event in ihrer Liste sieht.
                        </p>
                        <ul style={{ margin: '0 0 8px 18px', padding: 0 }}>
                          <li style={{ marginBottom: 6 }}>
                            <strong>ODER (Default):</strong> <em>einer der beiden Filter reicht.</em> Beispiel: <strong>Standort = Köln</strong>, <strong>Verteiler = SAPALL</strong> → jede Person, die <strong>in Köln</strong> sitzt <strong>ODER</strong> in <strong>SAPALL</strong> ist, sieht das Event. Praktisch wenn du bewusst einen <strong>breiten Empfängerkreis</strong> willst (z.B. Standort-Mitarbeiter <strong>plus</strong> Fachgruppe).
                          </li>
                          <li>
                            <strong>UND:</strong> <em>beide Filter müssen zutreffen.</em> Beispiel: <strong>Standort = Köln</strong>, <strong>Verteiler = SAPALL</strong> → nur wer <strong>in Köln</strong> sitzt <strong>UND</strong> in <strong>SAPALL</strong> ist, sieht das Event. Praktisch wenn du den Empfängerkreis <strong>strikt eingrenzen</strong> willst (z.B. nur die SAP-Kollegen <strong>am Standort Köln</strong>).
                          </li>
                        </ul>
                        <p style={{ margin: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                          <strong>Hinweis:</strong> diese Auswahl erscheint nur, wenn <strong>beide</strong> Filter gesetzt sind — sonst gibt es nichts zu kombinieren.
                        </p>
                      </>
                    ) : (
                      <>
                        <p style={{ margin: '0 0 8px' }}>
                          Defines how the <strong>location filter</strong> and the <strong>mailing lists / individual users</strong> are combined — i.e. which conditions must be true for a person before the event shows up in their list.
                        </p>
                        <ul style={{ margin: '0 0 8px 18px', padding: 0 }}>
                          <li style={{ marginBottom: 6 }}>
                            <strong>OR (default):</strong> <em>either filter is enough.</em> Example: <strong>Location = Cologne</strong>, <strong>list = SAPALL</strong> → anyone <strong>in Cologne</strong> <strong>OR</strong> in <strong>SAPALL</strong> sees the event. Useful when you intentionally want a <strong>broad audience</strong> (e.g. location staff <strong>plus</strong> a domain group).
                          </li>
                          <li>
                            <strong>AND:</strong> <em>both filters must match.</em> Example: <strong>Location = Cologne</strong>, <strong>list = SAPALL</strong> → only people <strong>in Cologne</strong> <strong>AND</strong> in <strong>SAPALL</strong> see the event. Useful when you want to <strong>strictly narrow</strong> the audience (e.g. only the SAP colleagues <strong>at the Cologne site</strong>).
                          </li>
                        </ul>
                        <p style={{ margin: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                          <strong>Note:</strong> this selector only appears when <strong>both</strong> filters are set — otherwise there is nothing to combine.
                        </p>
                      </>
                    )}
                  </div>
                  <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.9rem', cursor: 'pointer' }}>
                      <input type="radio" name="filterMode" value="OR" checked={filterMode === 'OR'} onChange={() => setFilterMode('OR')} />
                      <strong>{isDe ? 'ODER' : 'OR'}</strong>
                      <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>
                        – {isDe ? 'Einer der Filter reicht' : 'one filter is enough'}
                      </span>
                    </label>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.9rem', cursor: 'pointer' }}>
                      <input type="radio" name="filterMode" value="AND" checked={filterMode === 'AND'} onChange={() => setFilterMode('AND')} />
                      <strong>{isDe ? 'UND' : 'AND'}</strong>
                      <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>
                        – {isDe ? 'Beides muss zutreffen' : 'both must match'}
                      </span>
                    </label>
                  </div>
                </div>
              )}

              {/* Sichtbarkeit-Pruefen-Button: ans Ende der gesamten
                  Sichtbarkeits-Sektion (nach Standortfilter, Mailverteiler
                  und Filterverknuepfung), weil der Organizer typischerweise
                  erst alles befuellt, bevor er die kombinierte Sichtbarkeit
                  in einer Vorschau gegen-prueft. */}
              {(locationFilter || audience) && (
                <div className="form-group" style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)', flexWrap: 'wrap' }}>
                  <button
                    className="btn btn-outline"
                    style={{ fontSize: '0.8rem', padding: '6px 14px', whiteSpace: 'nowrap' }}
                    onClick={async () => {
                      setShowEmailModal(true);
                      // v8.9: Verteiler einmal aufloesen und cachen, damit
                      // jeder Such-Treffer in O(1) gegen die Mailgruppen-
                      // Mitgliedschaft geprueft werden kann.
                      setVisibilityCacheLoading(true);
                      const cache = new Set<string>();
                      const audItems = audience.split(',').map(s => s.trim()).filter(Boolean);
                      for (const item of audItems) {
                        if (item.indexOf('@') < 0) continue;
                        // Direkter User-Eintrag → in den Cache
                        cache.add(item.toLowerCase());
                        // Verteiler/Gruppe → Members aufloesen
                        try {
                          const grp = await getGroupMembers(item).catch(() => null);
                          if (grp && grp.members) {
                            for (const m of grp.members) {
                              if (m.email) cache.add(m.email.toLowerCase());
                            }
                          }
                        } catch { /* */ }
                      }
                      setVisibilityAudienceCache(cache);
                      setVisibilityCacheLoading(false);
                    }}
                    type="button"
                  >
                    <Users size={14} /> {isDe ? 'Sichtbarkeit prüfen' : 'Check visibility'}
                  </button>
                  <button
                    className="btn btn-outline"
                    style={{ fontSize: '0.8rem', padding: '6px 14px', whiteSpace: 'nowrap' }}
                    onClick={async () => {
                      // Resolver: Mailgruppen-Members via Graph aufloesen,
                      // einzelne E-Mails direkt durchreichen, Mitglieds-Quelle
                      // markieren (z.B. 'SAPALL@deloitte.com').
                      setExcludeModalOpen(true);
                      setExcludeResolving(true);
                      const resolved: Array<{ email: string; displayName: string; firstName: string; lastName: string; jobTitle: string; location: string; source: string }> = [];
                      const seen = new Set<string>();
                      // v8.9: Standorte zuerst aufloesen — alle User des
                      // Standorts werden via Graph geholt (officeLocation
                      // exact match, Fallback startsWith).
                      const locItems = locationFilter.split(',').map(s => s.trim()).filter(Boolean);
                      for (const loc of locItems) {
                        try {
                          const users = await searchUsersByLocation(loc).catch(() => []);
                          for (const u of users) {
                            const k = (u.email || '').toLowerCase();
                            if (!k || seen.has(k)) continue;
                            seen.add(k);
                            resolved.push({
                              email: u.email,
                              displayName: u.displayName,
                              firstName: u.firstName || '',
                              lastName: u.lastName || '',
                              jobTitle: u.jobTitle || '',
                              location: u.location || loc,
                              source: loc,
                            });
                          }
                        } catch { /* skip */ }
                      }
                      const audItems = audience.split(',').map(s => s.trim()).filter(Boolean);
                      for (const item of audItems) {
                        try {
                          if (item.indexOf('@') >= 0) {
                            // Wenn sichtbar wie eine Gruppe (z.B. SAPALL@), getGroupMembers; sonst direkt
                            const grp = await getGroupMembers(item).catch(() => null);
                            if (grp && grp.members && grp.members.length > 0) {
                              for (const m of grp.members) {
                                const k = (m.email || '').toLowerCase();
                                if (!k || seen.has(k)) continue;
                                seen.add(k);
                                resolved.push({
                                  email: m.email,
                                  displayName: m.displayName,
                                  firstName: m.firstName || '',
                                  lastName: m.lastName || '',
                                  jobTitle: m.jobTitle || '',
                                  location: m.location || '',
                                  source: item,
                                });
                              }
                            } else {
                              // Direkter User-Eintrag
                              const k = item.toLowerCase();
                              if (!seen.has(k)) {
                                seen.add(k);
                                resolved.push({ email: item, displayName: item, firstName: '', lastName: '', jobTitle: '', location: '', source: 'direkt' });
                              }
                            }
                          }
                        } catch { /* skip */ }
                      }
                      setExcludeResolvedUsers(resolved);
                      setExcludeResolving(false);
                    }}
                    type="button"
                  >
                    <Users size={14} /> {isDe ? 'Personen ausschließen' : 'Exclude users'}
                    {excludedUsers.length > 0 && (
                      <span style={{ marginLeft: 6, padding: '1px 7px', background: 'var(--dex-red, #c00)', color: '#fff', borderRadius: 999, fontSize: '0.7rem', fontWeight: 700 }}>
                        {excludedUsers.length}
                      </span>
                    )}
                  </button>
                  <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', margin: 0, lineHeight: 1.5, flex: 1, minWidth: 200 }}>
                    {isDe
                      ? 'Öffnet eine Vorschau, mit der du anhand einer Testperson verifizieren kannst, ob die kombinierte Sichtbarkeit (Standortfilter + Mailverteiler/User + Verknüpfung) wirklich passt — bevor du das Event speicherst.'
                      : 'Opens a preview where you can use a test person to verify whether the combined visibility (location filter + mailing lists/users + AND/OR combination) really matches — before you save the event.'}
                  </p>
                </div>
              )}

              <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={(locationFilter && audience) ? 16 : 15} />
                  {isDe ? 'Anmelde- und Abmeldefristen' : 'Registration & cancellation deadlines'}
                  <InfoTooltip text={isDe
                    ? 'Bis wann können sich Teilnehmer anmelden bzw. ohne Rückfrage abmelden? Beide Werte werden anhand des Event-Datums automatisch vorgeschlagen, du kannst sie jederzeit überschreiben.'
                    : 'Until when can attendees register / cancel themselves without consequence? Both values are auto-suggested from the event date and can be overridden at any time.'} />
                </label>
              <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                <div className="form-group" style={{ marginBottom: 0 }}>
                  <label className="form-label">
                    {t('create.deadline')}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Anmelde-Deadline</strong> — bis zu diesem Stichtag können sich Teilnehmer selbst registrieren.<br /><br />
                        <strong>Auswirkung für Teilnehmer:</strong> nach dem Stichtag ist der <strong>Anmelden-Button gesperrt</strong> (auch via Direktlink), reguläre User können sich nicht mehr selbst eintragen. <strong>Organizer und Co-Organizer</strong> dürfen weiterhin manuell Teilnehmer anlegen — die Deadline gilt nur für Self-Registration.<br /><br />
                        Vorbefüllt mit <strong>7 Tagen vor Event-Start</strong>, frei überschreibbar.
                      </>
                    ) : (
                      <>
                        <strong>Registration deadline</strong> — until this cutoff attendees can self-register.<br /><br />
                        <strong>Effect for attendees:</strong> past the cutoff the <strong>register button is locked</strong> (also via direct link), regular users can no longer sign themselves up. <strong>Organizers and co-organizers</strong> can still add attendees manually — the deadline only applies to self-registration.<br /><br />
                        Pre-filled with <strong>7 days before event start</strong>, freely overridable.
                      </>
                    )} />
                  </label>
                  <DatePicker
                    selected={registrationDeadline ? new Date(registrationDeadline) : null}
                    onChange={(date: Date | null) => setRegistrationDeadline(date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}` : '')}
                    showTimeSelect
                    timeFormat="HH:mm"
                    timeIntervals={15}
                    timeCaption="Uhrzeit"
                    dateFormat="dd.MM.yyyy, HH:mm"
                    locale="de"
                    placeholderText="Anmelde-Deadline"
                    className="form-input"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    isClearable
                    autoComplete="off"
                  />
                </div>
                <div className="form-group" style={{ marginBottom: 0 }}>
                  <label className="form-label">
                    {t('create.lastcancel')}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Letzte Abmeldemöglichkeit</strong> — bis zu diesem Stichtag können sich Teilnehmer <strong>ohne Rückfrage</strong> selbst abmelden.<br /><br />
                        <strong>Auswirkung für Teilnehmer:</strong> nach dem Stichtag ist der <strong>Abmelden-Button für reguläre User ausgeblendet</strong> — sie müssen aktiv den Organizer kontaktieren, der dann manuell abmeldet. <strong>Organizer und Co-Organizer</strong> können weiterhin jederzeit Teilnehmer abmelden.<br /><br />
                        <strong>Automatismen:</strong> je nach Einstellung in <strong>Schritt 7 (Kommunikation)</strong> bekommen die Organizer eine <strong>Late-Cancel-Mail</strong> mit Name + Mail des Abmeldenden — damit Hotel, Catering oder Bus angepasst werden können.<br /><br />
                        Vorbefüllt mit <strong>3 Tagen vor Event-Start</strong>.
                      </>
                    ) : (
                      <>
                        <strong>Last cancellation date</strong> — until this cutoff attendees can <strong>self-cancel without consequences</strong>.<br /><br />
                        <strong>Effect for attendees:</strong> past the cutoff the <strong>cancel button is hidden for regular users</strong> — they have to actively contact the organizer who then cancels them manually. <strong>Organizers and co-organizers</strong> can still cancel attendees at any time.<br /><br />
                        <strong>Automation:</strong> depending on the setting in <strong>step 5 (Communication)</strong>, organizers receive a <strong>late-cancel mail</strong> with name + email of the person — so hotel, catering or bus can be adjusted.<br /><br />
                        Pre-filled with <strong>3 days before event start</strong>.
                      </>
                    )} />
                  </label>
                  <DatePicker
                    selected={lastDeregisterDate ? new Date(lastDeregisterDate) : null}
                    onChange={(date: Date | null) => setLastDeregisterDate(date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}` : '')}
                    showTimeSelect
                    timeFormat="HH:mm"
                    timeIntervals={15}
                    timeCaption="Uhrzeit"
                    dateFormat="dd.MM.yyyy, HH:mm"
                    locale="de"
                    placeholderText="Abmeldefrist"
                    className="form-input"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    isClearable
                    autoComplete="off"
                  />
                </div>
              </div>
              {fieldHasError('deadlineAfterStart') && <p style={{ color: 'var(--dex-red)', fontSize: '0.8rem', marginTop: 8, marginBottom: 0 }}>{t('create.error.deadlineAfterStart')}</p>}
              {fieldHasError('deregAfterStart') && <p style={{ color: 'var(--dex-red)', fontSize: '0.8rem', marginTop: 8, marginBottom: 0 }}>{t('create.error.deregAfterStart')}</p>}
              </div>

              {/* v9.17: Reihenfolge umgestellt — Standard-Teilnehmerzahl
                  steht oben, Split-Toggle wird unter dem Block subtler
                  angezeigt. Die Mehrheit der Events nutzt nur eine
                  Gesamtkapazitaet; der B2Run-Sonderfall ist Opt-in. */}

              <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={(locationFilter && audience) ? 17 : 16} />
                  {isDe ? 'Teilnehmerzahl & Warteliste' : 'Capacity & waitlist'}
                </label>
              {/* v10.20: Geteilte Kapazität — generisch fuer beliebige Events.
                  Labels werden vom Organizer frei gewaehlt (z.B. "Vormittag /
                  Nachmittag", "VIP / Standard", "Lauf / Walk"). Default-Fallback
                  ist 'Durchstarter' / 'Funstarter' fuer Backward-Compat mit
                  B2Run-Events vor v10.20. */}
              {useSplitCapacities ? (
                <div style={{ padding: 16, background: 'var(--dex-green-light, #f0fdf4)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-green)', marginBottom: 16 }}>
                  <label className="form-label" style={{ marginBottom: 4 }}>
                    {isDe ? 'Geteilte Kapazität' : 'Split capacity'}
                    <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> zwei getrennte Kapazitäten innerhalb des Events. Beispiele: <strong>Vormittag / Nachmittag</strong>, <strong>VIP / Standard</strong>, <strong>Lauf 5 km / Lauf 10 km</strong>. Du legst die <strong>Bezeichnungen frei fest</strong> und vergibst pro Gruppe eine eigene Platzzahl mit eigener Warteliste.<br /><br />
                      <strong>Anzeige in der App:</strong> Teilnehmer sehen auf der Anmelde-Seite <strong>zwei nebeneinanderstehende Boxen</strong> mit deinen Bezeichnungen, jeweils mit &bdquo;X / Y Plätze frei&ldquo;. Die Wahl ist <strong>verpflichtend</strong>, bevor man auf Anmelden klickt.<br /><br />
                      <strong>Automatismen:</strong> ist eine der zwei Gruppen voll, kommen weitere Anmeldungen <strong>nur in die Warteliste dieser Gruppe</strong> — nicht in die andere. Beim Nachrücken bleibt der Typ <strong>erhalten</strong> (eine VIP-Wartelisten-Person rückt nicht in einen Standard-Platz).<br /><br />
                      <strong>Empfehlung:</strong> nur verwenden, wenn die zwei Gruppen <strong>wirklich getrennt</strong> behandelt werden sollen (eigenes Catering, eigener Bus, eigener Slot beim Veranstalter). Bei einer einfachen Gesamtkapazität reicht die Standard-Teilnehmerzahl unten.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> two separate capacities within the event. Examples: <strong>morning / afternoon</strong>, <strong>VIP / standard</strong>, <strong>5 km run / 10 km run</strong>. You <strong>name the two groups freely</strong> and give each its own seat count and waitlist.<br /><br />
                      <strong>Shown in the app:</strong> attendees see <strong>two side-by-side boxes</strong> on the registration page, each labelled with your text and showing &ldquo;X / Y seats free&rdquo;. Picking one is <strong>required</strong> before submitting.<br /><br />
                      <strong>Automation:</strong> when one group is full, further sign-ups land <strong>only on that group&apos;s waitlist</strong> — not the other. When promoting from waitlist, the <strong>group is preserved</strong> (a waitlisted VIP is not auto-promoted into a standard slot).<br /><br />
                      <strong>Tip:</strong> only use this when the two groups really need to be <strong>handled separately</strong> (own catering, own bus, separate slot with the supplier). For a single overall capacity the standard attendee count below is enough.
                    </>
                  )} />
                  </label>
                  <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                    {isDe
                      ? 'Vergib pro Gruppe eine eigene Bezeichnung und Platzzahl. Die Bezeichnungen erscheinen auf der Anmeldeseite als zwei Auswahl-Boxen.'
                      : 'Give each group its own name and seat count. The names appear on the registration page as two selectable boxes.'}
                  </p>
                  {/* v10.20: zwei Text-Inputs fuer die frei waehlbaren Bezeichnungen.
                      Wenn der Organizer nichts eintraegt, faellt die Registration-
                      Seite auf 'Durchstarter' / 'Funstarter' zurueck. */}
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 12 }}>
                    <div>
                      <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                        {isDe ? 'Bezeichnung Gruppe A' : 'Group A label'}
                      </label>
                      <input
                        className="form-input"
                        type="text"
                        value={splitLabelA}
                        onChange={e => setSplitLabelA(e.target.value)}
                        placeholder={isDe ? 'z.B. Vormittag, VIP, Durchstarter' : 'e.g. morning, VIP, starter'}
                        maxLength={40}
                      />
                    </div>
                    <div>
                      <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                        {isDe ? 'Bezeichnung Gruppe B' : 'Group B label'}
                      </label>
                      <input
                        className="form-input"
                        type="text"
                        value={splitLabelB}
                        onChange={e => setSplitLabelB(e.target.value)}
                        placeholder={isDe ? 'z.B. Nachmittag, Standard, Funstarter' : 'e.g. afternoon, standard, fun'}
                        maxLength={40}
                      />
                    </div>
                  </div>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                    <div>
                      <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                        <Icon iconName="People" style={{ fontSize: 14, marginRight: 6, color: 'var(--dex-green-dark, #6b9a1e)' }} />
                        {isDe ? 'Plätze' : 'Seats'} {splitLabelA.trim() || (isDe ? 'Gruppe A' : 'Group A')}
                      </label>
                      <input
                        className="form-input"
                        type="number"
                        min={0}
                        value={durchstarterCapacity}
                        onChange={e => setDurchstarterCapacity(e.target.value)}
                        placeholder="z.B. 10"
                      />
                    </div>
                    <div>
                      <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                        <Icon iconName="People" style={{ fontSize: 14, marginRight: 6, color: 'var(--dex-orange, #ff8c00)' }} />
                        {isDe ? 'Plätze' : 'Seats'} {splitLabelB.trim() || (isDe ? 'Gruppe B' : 'Group B')}
                      </label>
                      <input
                        className="form-input"
                        type="number"
                        min={0}
                        value={funstarterCapacity}
                        onChange={e => setFunstarterCapacity(e.target.value)}
                        placeholder="z.B. 90"
                      />
                    </div>
                  </div>
                  <div style={{ marginTop: 12, padding: '8px 12px', background: '#fff', borderRadius: 8, fontSize: '0.85rem' }}>
                    <strong>{t('create.b2runcap.total')}:</strong> {((parseInt(durchstarterCapacity, 10) || 0) + (parseInt(funstarterCapacity, 10) || 0))} {t('create.b2runcap.seats')}
                  </div>

                  {/* v10.20: Waitlist-Modus bei Split-Capacity. Default
                      'separate' (zwei Wartelisten) — entspricht dem alten
                      B2Run-Verhalten und ist die typ-bewusste Variante.
                      Alternative 'shared' = eine gemeinsame Warteliste, FIFO
                      ueber beide Gruppen. Sinnvoll wenn die Gruppen
                      organisatorisch fluide sind (z.B. Vormittag/Nachmittag
                      bei einem Workshop, wo der naechste freie Slot egal ist
                      welche Gruppe). Nur sichtbar wenn Warteliste aktiviert
                      ist (Toggle weiter unten). */}
                  {waitlistEnabled && (
                    <div style={{ marginTop: 12, padding: '12px 14px', background: '#fff', borderRadius: 8 }}>
                      <label className="form-label" style={{ marginBottom: 6 }}>
                        {isDe ? 'Warteliste-Verhalten' : 'Waitlist behaviour'}
                        <InfoTooltip text={isDe ? (
                          <>
                            <strong>Was du hier einstellst:</strong> ob bei einer Anmeldung über die jeweilige Kapazität hinaus eine <strong>gemeinsame</strong> oder zwei <strong>getrennte</strong> Wartelisten greifen.<br /><br />
                            <strong>Getrennt (Default):</strong> jede Gruppe hat ihre eigene Warteliste. Wer auf der {splitLabelA.trim() || 'Gruppe A'}-Warteliste landet, rückt nur in einen frei werdenden {splitLabelA.trim() || 'Gruppe A'}-Platz nach. Saubere Trennung — sinnvoll wenn die zwei Gruppen wirklich unterschiedliche Slots beim Veranstalter, eigenes Catering oder eigenen Bus haben.<br /><br />
                            <strong>Gemeinsam:</strong> alle Wartelistler stehen in einer einzigen Schlange. Wer am längsten wartet, rückt zuerst nach — egal in welche Gruppe der frei werdende Platz gehört. Sinnvoll wenn die Gruppen-Wahl nur eine UI-Komfort-Sache ist (z.B. Vormittag/Nachmittag-Slot bei einem Workshop) und der Organizer sich nicht um Typen kümmern will.<br /><br />
                            <strong>Auswirkung für Teilnehmer:</strong> bei <strong>getrennt</strong> kann es passieren, dass jemand in der einen Schlange weiter hinten steht, obwohl die andere Gruppe leer ist — dann muss man <strong>aktiv umsteigen</strong> (über den Fallback-Dialog beim nächsten Versuch). Bei <strong>gemeinsam</strong> rutscht jeder hoch sobald irgendwo ein Platz frei wird.
                          </>
                        ) : (
                          <>
                            <strong>What you set here:</strong> whether sign-ups beyond the per-group capacity land on <strong>one shared</strong> or <strong>two separate</strong> waitlists.<br /><br />
                            <strong>Separate (default):</strong> each group has its own waitlist. Someone on the {splitLabelA.trim() || 'group A'} waitlist only moves up into a freed {splitLabelA.trim() || 'group A'} seat. Clean separation — useful when the two groups have genuinely different supplier slots, own catering, own bus.<br /><br />
                            <strong>Shared:</strong> all waitlisters stand in one queue. Whoever has waited longest moves up first — regardless of which group the freed seat belongs to. Useful when the group split is just a UI convenience (e.g. morning / afternoon slot at a workshop) and the organizer does not want to manage types.<br /><br />
                            <strong>Effect for attendees:</strong> with <strong>separate</strong> someone may be further back in their queue while the other group is empty — they then have to <strong>actively switch</strong> (via the fallback dialog at next attempt). With <strong>shared</strong> everyone moves up as soon as a seat opens anywhere.
                          </>
                        )} />
                      </label>
                      <div style={{ display: 'flex', gap: 16, marginTop: 4 }}>
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: '0.85rem' }}>
                          <input
                            type="radio"
                            name="splitWaitlistMode"
                            checked={!splitSharedWaitlist}
                            onChange={() => setSplitSharedWaitlist(false)}
                          />
                          {isDe ? 'Getrennte Wartelisten pro Gruppe' : 'Separate waitlist per group'}
                        </label>
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: '0.85rem' }}>
                          <input
                            type="radio"
                            name="splitWaitlistMode"
                            checked={!!splitSharedWaitlist}
                            onChange={() => setSplitSharedWaitlist(true)}
                          />
                          {isDe ? 'Eine gemeinsame Warteliste' : 'One shared waitlist'}
                        </label>
                      </div>
                    </div>
                  )}

                  {/* v11.25: Display-Reihenfolge der zwei Gruppen-Karten in der
                      Registrierungs-UI umkehren. Reine Anzeige-Toggle —
                      splitLabelA/B, Kapazitäten und die internen StarterType-IDs
                      bleiben unangetastet. Nur bei aktiver Split-Capacity
                      sinnvoll. */}
                  {useSplitCapacities && (
                    <div style={{ marginTop: 12, padding: '12px 14px', background: '#fff', borderRadius: 8 }}>
                      <label style={{ display: 'inline-flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.85rem' }}>
                        <input
                          type="checkbox"
                          checked={splitDisplayOrderReversed}
                          onChange={e => setSplitDisplayOrderReversed(e.target.checked)}
                          style={{ marginTop: 3 }}
                        />
                        <span>
                          <strong>{isDe ? 'Reihenfolge in der Registrierung umkehren' : 'Reverse order on registration page'}</strong>
                          <InfoTooltip text={isDe ? (
                            <>
                              <strong>Was du hier einstellst:</strong> ob die zwei Gruppen-Karten (&bdquo;{splitLabelA.trim() || 'Gruppe A'}&ldquo; und &bdquo;{splitLabelB.trim() || 'Gruppe B'}&ldquo;) in der Registrierungs-Maske in der <strong>aktuellen</strong> oder in <strong>umgekehrter</strong> Reihenfolge angezeigt werden.<br /><br />
                              <strong>Anzeige in der App:</strong> aus = &bdquo;{splitLabelA.trim() || 'Gruppe A'}&ldquo; links, &bdquo;{splitLabelB.trim() || 'Gruppe B'}&ldquo; rechts. An = umgekehrt: &bdquo;{splitLabelB.trim() || 'Gruppe B'}&ldquo; links, &bdquo;{splitLabelA.trim() || 'Gruppe A'}&ldquo; rechts. Gilt für Registrierung und für die Kapazitäts-Übersicht im Admin-Center.<br /><br />
                              <strong>Auswirkung für Teilnehmer:</strong> rein optisch — ändert nichts an den Plätzen, Wartelisten oder bestehenden Anmeldungen. Nur eine andere Reihenfolge der Auswahl-Buttons.<br /><br />
                              <strong>Hinweis:</strong> wenn du tatsächlich &bdquo;Gruppe 2&ldquo; zur prominenteren machen willst, ohne deine Labels und Kapazitäten zu vertauschen, ist <strong>dieses Häkchen</strong> der saubere Weg. Manuelles Vertauschen von Label A ↔ B + Kapazitäten zerschießt die Verbindung zu den existierenden Anmeldungen.
                            </>
                          ) : (
                            <>
                              <strong>What you set here:</strong> whether the two group cards (&bdquo;{splitLabelA.trim() || 'group A'}&ldquo; and &bdquo;{splitLabelB.trim() || 'group B'}&ldquo;) appear in the <strong>current</strong> or <strong>reversed</strong> order on the registration page.<br /><br />
                              <strong>Shown in the app:</strong> off = &bdquo;{splitLabelA.trim() || 'group A'}&ldquo; left, &bdquo;{splitLabelB.trim() || 'group B'}&ldquo; right. On = reversed.<br /><br />
                              <strong>Effect for attendees:</strong> purely visual — does not affect seats, waitlists or existing registrations. Just a different order of selection buttons.<br /><br />
                              <strong>Note:</strong> if you want to make &bdquo;group 2&ldquo; the more prominent one without swapping your labels and capacities, this checkbox is the clean way. Manually swapping label A ↔ B + capacities breaks the link to existing registrations.
                            </>
                          )} />
                        </span>
                      </label>
                    </div>
                  )}

                  {/* v6.15: Starter-Typ → Startblock-Zuordnung (optional).
                      Nur sinnvoll wenn Startblocks definiert sind (Reiter "Event-spezifische Felder").
                      Wenn gesetzt, wird der Startblock bei der Registrierung automatisch
                      anhand des gewählten Starter-Typs gesetzt — der User muss den Block
                      nicht extra auswählen. */}
                  {b2runStartblocks.length > 0 && (
                    <div style={{ marginTop: 12, padding: '10px 12px', background: '#fff', borderRadius: 8 }}>
                      <label className="form-label" style={{ marginBottom: 6 }}>
                        {t('create.b2runcap.starterblock.title') || 'Starter-Typ → Startblock-Zuordnung'}
                        <InfoTooltip text={isDe ? (
                          <>
                            <strong>Was du hier einstellst:</strong> eine <strong>fixe Zuordnung Starter-Typ → Startblock</strong>. Beispiel: Durchstarter immer Block A, Funstarter immer Block C.<br /><br />
                            <strong>Anzeige in der App:</strong> Teilnehmer wählen <strong>nur</strong> den Starter-Typ — der Startblock wird automatisch gesetzt, der Block-Selector verschwindet aus der Anmelde-Maske. Eine Frage weniger für den User.<br /><br />
                            <strong>Automatismen:</strong> die Startblock-Kapazität wird gegen den jeweiligen Starter-Typ gerechnet — freie Plätze pro Typ = freie Plätze im zugeordneten Block.<br /><br />
                            <strong>Leer:</strong> Teilnehmer wählen Starter-Typ <em>und</em> Startblock manuell — der Organizer hat dann gemischte Blöcke und muss bei Engpässen selbst zuteilen.
                          </>
                        ) : (
                          <>
                            <strong>What you set here:</strong> a <strong>fixed mapping starter type → start block</strong>. Example: Durchstarter always block A, Funstarter always block C.<br /><br />
                            <strong>Shown in the app:</strong> attendees only pick the <strong>starter type</strong> — the start block is set automatically, the block selector disappears from the form. One question less for the user.<br /><br />
                            <strong>Automation:</strong> the start-block capacity counts against the matching starter type — free slots per type = free slots in the assigned block.<br /><br />
                            <strong>Empty:</strong> attendees pick starter type <em>and</em> start block manually — the organizer ends up with mixed blocks and has to redistribute manually when slots get tight.
                          </>
                        )} />
                      </label>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                        <div>
                          <label style={{ display: 'block', fontSize: '0.78rem', marginBottom: 4, color: 'var(--dex-gray-600)' }}>Durchstarter →</label>
                          <select
                            className="form-select"
                            value={durchstarterStartblock}
                            onChange={e => setDurchstarterStartblock(e.target.value)}
                          >
                            <option value="">{t('create.b2runcap.starterblock.none') || '— kein automatischer Block —'}</option>
                            {b2runStartblocks.map(b => <option key={b} value={b}>{b}</option>)}
                          </select>
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '0.78rem', marginBottom: 4, color: 'var(--dex-gray-600)' }}>Funstarter →</label>
                          <select
                            className="form-select"
                            value={funstarterStartblock}
                            onChange={e => setFunstarterStartblock(e.target.value)}
                          >
                            <option value="">{t('create.b2runcap.starterblock.none') || '— kein automatischer Block —'}</option>
                            {b2runStartblocks.map(b => <option key={b} value={b}>{b}</option>)}
                          </select>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* v10.24: Leistungsnachweis-Pflicht-Toggle wurde entfernt.
                      Stattdessen kann der Organizer in Schritt 6 (Felder) ein
                      eigenes Pflichtfeld vom Typ Checkbox anlegen und es ueber
                      'Sichtbar fuer Teilnehmergruppe → Nur Gruppe A' gezielt
                      auf eine Split-Gruppe einschraenken. Das ersetzt den
                      hartkodierten B2Run-Leistungsnachweis-Sonderfall durch
                      ein generisches Pro-Gruppe-Feld-Konzept. Hinweis steht
                      hier, damit der Organizer beim Migrieren weiss wo das
                      Feature jetzt liegt. */}
                  <div style={{ marginTop: 12, padding: '10px 12px', background: 'rgba(21,101,192,0.06)', border: '1px solid rgba(21,101,192,0.4)', borderRadius: 8, fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                    <div style={{ fontWeight: 700, color: 'var(--dex-blue, #1565c0)', marginBottom: 4, display: 'flex', alignItems: 'center', gap: 6 }}>
                      <Icon iconName="Info" style={{ fontSize: 16 }} />
                      {isDe ? 'Pflichtfelder pro Gruppe' : 'Required fields per group'}
                    </div>
                    {isDe ? (
                      <>
                        Du möchtest für eine der zwei Gruppen ein zusätzliches Pflichtfeld einblenden — z.B. eine Checkbox &bdquo;Leistungsnachweis vorhanden&ldquo; nur für die Gruppe der schnellen Läufer? Lege das Feld in <strong>Schritt 6 (Felder)</strong> an und stelle dort den Selector <strong>&bdquo;Sichtbar für Teilnehmergruppe&ldquo;</strong> auf <strong>&bdquo;Nur {(splitLabelA || '').trim() || 'Gruppe A'}&ldquo;</strong> bzw. <strong>&bdquo;Nur {(splitLabelB || '').trim() || 'Gruppe B'}&ldquo;</strong>. Das Feld wird dann in der Anmeldung dynamisch ein- oder ausgeblendet, sobald der Teilnehmer eine der zwei Boxen anklickt.
                      </>
                    ) : (
                      <>
                        Want to show an extra required field only for one of the two groups — e.g. a checkbox &ldquo;Performance proof available&rdquo; just for the fast-runner group? Add the field in <strong>step 6 (Fields)</strong> and set the <strong>&ldquo;Visible for attendee group&rdquo;</strong> selector there to <strong>&ldquo;{(splitLabelA || '').trim() || 'Group A'} only&rdquo;</strong> or <strong>&ldquo;{(splitLabelB || '').trim() || 'Group B'} only&rdquo;</strong>. The field will then be shown or hidden dynamically as the attendee picks one of the two boxes.
                      </>
                    )}
                  </div>

                  <div style={{ marginTop: 12 }}>
                    <label className="form-label" style={{ marginBottom: 4 }}>
                      {t('create.waitlist')}
                      <InfoTooltip text={isDe ? (
                        <>
                          <strong>Wartelisten für Split-Kapazitäten</strong> — bei aktivierter Warteliste hat <strong>jeder Starter-Typ seine eigene Liste</strong>. Durchstarter-Anmeldungen über der Durchstarter-Kapazität landen auf der Durchstarter-Warteliste, dasselbe für Funstarter. Beim Nachrücken wird <strong>typ-bewusst</strong> befördert — der älteste Durchstarter-Eintrag rückt in einen frei werdenden Durchstarter-Platz, kein Mix.
                        </>
                      ) : (
                        <>
                          <strong>Waitlists for split capacities</strong> — when the waitlist is on, <strong>each starter type has its own queue</strong>. Durchstarter sign-ups beyond the Durchstarter capacity go on the Durchstarter waitlist, same for Funstarter. Promotion is <strong>type-aware</strong> — the oldest Durchstarter waiting moves into a freed Durchstarter slot, no mixing.
                        </>
                      )} />
                    </label>
                    <div className="toggle-wrapper" style={{ marginTop: 4 }}>
                      <label className="toggle">
                        <input type="checkbox" checked={waitlistEnabled} onChange={e => setWaitlistEnabled(e.target.checked)} />
                        <span className="toggle-slider" />
                      </label>
                      <span style={{ fontSize: '0.9rem' }}>{waitlistEnabled ? t('create.enabled') : t('create.disabled')}</span>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                  <div className="form-group">
                    <label className="form-label">
                      {t('create.maxparticipants')}
                      <InfoTooltip text={isDe ? (
                        <>
                          <strong>Was du hier einstellst:</strong> die <strong>maximale Teilnehmerzahl</strong> oder den Modus <strong>Unbegrenzt</strong>.<br /><br />
                          <strong>Anzeige in der App:</strong> ist die Kapazität voll, sehen Teilnehmer den <strong>roten Banner Alle Plätze sind belegt</strong> auf der Anmelde-Seite. Im Admin Center wird ein Auslastungs-KPI live mitgeführt.<br /><br />
                          <strong>Automatismen:</strong> ist <strong>Warteliste</strong> aktiv und die Kapazität voll, werden neue Anmeldungen automatisch auf <strong>Status Warteliste</strong> gesetzt und bekommen die Wartelisten-Bestätigungs-Mail. Bei einer Abmeldung rückt der älteste Wartelisten-Eintrag automatisch nach und bekommt Nachrück-Mail + Outlook-Termin.<br /><br />
                          <strong>Modus Unbegrenzt:</strong> keine Auslastungs-Anzeige, keine Warteliste — wird typischerweise für interne All-Hands-Mails oder reine Info-Events verwendet.
                        </>
                      ) : (
                        <>
                          <strong>What you set here:</strong> the <strong>maximum attendee count</strong> or the <strong>Unlimited</strong> mode.<br /><br />
                          <strong>Shown in the app:</strong> when full, attendees see the <strong>red banner All spots are taken</strong> on the registration page. The admin center shows live capacity KPIs.<br /><br />
                          <strong>Automation:</strong> if <strong>waitlist</strong> is on and capacity is full, new sign-ups land in <strong>status Waitlist</strong> and get a waitlist confirmation mail. On cancellation, the oldest waitlist entry is auto-promoted and receives a promotion mail + Outlook event.<br /><br />
                          <strong>Unlimited mode:</strong> no capacity indicator, no waitlist — typically used for internal all-hands or info-only events.
                        </>
                      )} />
                    </label>
                    <div className="toggle-wrapper" style={{ marginTop: 4, marginBottom: 8 }}>
                      <label className="toggle">
                        <input
                          type="checkbox"
                          checked={unlimitedParticipants}
                          onChange={e => {
                            const unlimited = e.target.checked;
                            setUnlimitedParticipants(unlimited);
                            if (unlimited) {
                              setMaxParticipants('');
                              setWaitlistEnabled(false);
                            }
                          }}
                        />
                        <span className="toggle-slider" />
                      </label>
                      <span style={{ fontSize: '0.9rem' }}>{unlimitedParticipants ? (t('create.maxparticipants') === 'Max. Participants' ? 'Unlimited' : 'Unbegrenzt') : (maxParticipants ? `${maxParticipants} Plätze` : 'Anzahl eingeben')}</span>
                    </div>
                    {!unlimitedParticipants && (
                      <>
                        <input
                          className="form-input"
                          type="number"
                          min={0}
                          value={maxParticipants}
                          onChange={e => setMaxParticipants(e.target.value.replace(/[^0-9]/g, ''))}
                          placeholder="Anzahl"
                          style={errorBorderStyle('maxParticipants')}
                        />
                        {fieldHasError('maxParticipants') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
                      </>
                    )}
                  </div>
                  {!unlimitedParticipants && (
                    <div className="form-group">
                      <label className="form-label">
                        {t('create.waitlist')}
                        <InfoTooltip text={isDe ? (
                          <>
                            <strong>Was du hier einstellst:</strong> ob bei vollem Event eine <strong>Warteliste</strong> akzeptiert wird oder neue Anmeldungen sofort blockiert werden.<br /><br />
                            <strong>Anzeige in der App:</strong> bei aktiver Warteliste können Teilnehmer sich auch über die Kapazitäts-Grenze hinaus anmelden — bekommen Status <strong>Warteliste</strong> mit Positions-Nummer. Im Admin Center erscheint eine eigene <strong>Warteliste-Kachel</strong>.<br /><br />
                            <strong>Automatismen:</strong> Wartelisten-Anmeldungen bekommen die <strong>Wartelisten-Bestätigungs-Mail</strong>. Sobald jemand absagt, rückt der älteste Wartelisten-Eintrag automatisch nach (<strong>First-In, First-Out</strong>) — bekommt eine <strong>Nachrück-Mail</strong> mit Outlook-Termin und der Status wechselt auf Angemeldet.<br /><br />
                            <strong>Auswirkung für Teilnehmer:</strong> sie sehen ihre Position auf der Warteliste auf der Anmelde-Seite und werden automatisch informiert, wenn ein Platz frei wird.<br /><br />
                            <strong>Aus:</strong> bei vollem Event ist der Anmelde-Button gesperrt — neue Interessenten müssen den Organizer direkt kontaktieren.
                          </>
                        ) : (
                          <>
                            <strong>What you set here:</strong> whether full events accept a <strong>waitlist</strong> or new registrations are blocked immediately.<br /><br />
                            <strong>Shown in the app:</strong> when waitlist is on, attendees can register past the capacity limit — they get status <strong>Waitlist</strong> with a position number. The admin center shows a dedicated <strong>waitlist tile</strong>.<br /><br />
                            <strong>Automation:</strong> waitlist sign-ups receive the <strong>waitlist confirmation mail</strong>. As soon as someone cancels, the oldest entry is auto-promoted (<strong>first-in, first-out</strong>) — they receive a <strong>promotion mail</strong> with Outlook event and their status flips to Registered.<br /><br />
                            <strong>Effect for attendees:</strong> they see their waitlist position on the registration page and are notified automatically when a spot frees up.<br /><br />
                            <strong>Off:</strong> when capacity is full, the register button is locked — new interested people have to contact the organizer directly.
                          </>
                        )} />
                      </label>
                      <div className="toggle-wrapper" style={{ marginTop: 8 }}>
                        <label className="toggle">
                          <input type="checkbox" checked={waitlistEnabled} onChange={e => setWaitlistEnabled(e.target.checked)} />
                          <span className="toggle-slider" />
                        </label>
                        <span style={{ fontSize: '0.9rem' }}>{waitlistEnabled ? t('create.enabled') : t('create.disabled')}</span>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* v9.17: Split-Capacity-Toggle UNTER dem Teilnehmerzahl-Block,
                  bewusst subtil — der Großteil der Events nutzt eine einzige
                  Teilnehmerzahl. Nur wer einen Lauf mit getrennten Starter-
                  Töpfen anlegt, klickt diesen Toggle. */}
              <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: '0.8rem', color: 'var(--dex-gray-600)', cursor: 'pointer', padding: '8px 0', marginTop: 4 }}>
                <input
                  type="checkbox"
                  checked={useSplitCapacities}
                  onChange={e => setUseSplitCapacities(e.target.checked)}
                  style={{ marginTop: 2, cursor: 'pointer' }}
                />
                <span>
                  {t('create.splitcap.label')}
                  <span style={{ display: 'block', fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 1 }}>
                    {t('create.splitcap.hint')}
                  </span>
                </span>
              </label>
              </div>

              </div>{/* v15.6: close hauptGreyoutWrapperStyle div (Step 4) */}
              </div>{/* v15.0: close activeCapacityTabIdx===0 wrapper (Top-Level Sichtbarkeit/Deadlines/Max/Split) */}

              </div>{/* close Step 4 (Kapazitaet & Sichtbarkeit) */}

              {/* ===== Step 5 (v14.8: vormals Step 4): Team-Anmeldung =====
                  Renderblock für den Wizard-Schritt Team-Anmeldung.
                  Konfiguriert Team-Anmeldung-Toggle + Teamgröße +
                  Team-Name-Frage. v15: Index 4 → 6 (Team kommt jetzt nach
                  Kommunikation). */}
              <div style={{ display: currentStep === 6 ? 'block' : 'none' }}>
              <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                {isDe ? 'Schritt 7 — Team-Anmeldung' : 'Step 7 — Team Registration'}
              </h2>
              <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                {isDe
                  ? <><strong>Optional</strong> — erlaube einer Person, ein ganzes Team gleichzeitig anzumelden. Praktisch z.B. für Lauf-Teams, Workshop-Gruppen oder Tische bei einer Abendveranstaltung. Default: aus.</>
                  : <><strong>Optional</strong> — let a single person register an entire team in one go. Handy e.g. for running teams, workshop groups or tables at an evening event. Default: off.</>}
              </p>

              {/* Toggle Team-Anmeldung erlauben */}
              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '14px 16px', marginBottom: 12,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={teamRegistrationEnabled}
                    onChange={e => setTeamRegistrationEnabled(e.target.checked)}
                    style={{ marginTop: 3, cursor: 'pointer' }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Team-Anmeldung erlauben' : 'Allow team registration'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> ob eine Person ein <strong>ganzes Team</strong> über das Anmeldeformular anmelden darf — statt sich nur selbst einzutragen.<br /><br />
                          <strong>Anzeige in der App:</strong> der Team-Lead sieht nach Eingabe seiner eigenen Daten ein zusätzliches Formularfeld pro weiterem Team-Mitglied (Name, E-Mail). Default: aus — dann verhält sich das Event wie gewohnt (eine Person meldet sich selbst an).<br /><br />
                          <strong>Auswirkung für Teilnehmer:</strong> die mit angemeldeten Personen bekommen automatisch eine eigene Bestätigungsmail und (sofern Outlook aktiv ist) eigene Kalender-Einladung — sie müssen sich nicht selber registrieren.
                        </>
                      : <>
                          <strong>What this controls:</strong> whether one person can register an <strong>entire team</strong> via the registration form — instead of only registering themselves.<br /><br />
                          <strong>Where you see it:</strong> after entering their own details, the team lead sees an additional form block per team member (name, email). Default: off — the event behaves as usual (one person registers themselves).<br /><br />
                          <strong>For attendees:</strong> co-registered members automatically receive their own confirmation email and (if Outlook is enabled) their own calendar invite — they do not have to register themselves.
                        </>
                    } />
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? 'Wenn aktiviert, kann eine Person ein ganzes Team anmelden — die anderen Mitglieder bekommen Bestätigungsmail + Outlook-Termin automatisch.'
                        : 'When enabled, one person can register an entire team — the other members automatically receive a confirmation mail + Outlook invite.'}
                    </span>
                  </span>
                </label>
              </div>

              {/* Team-Größe + Team-Name-Frage — ausgegraut wenn Team-Anmeldung aus */}
              <div style={{
                background: teamRegistrationEnabled ? '#ffffff' : 'var(--dex-gray-50, #fafafa)',
                borderRadius: 12, padding: '14px 16px', marginBottom: 12,
                border: '1px solid var(--dex-gray-200)',
                opacity: teamRegistrationEnabled ? 1 : 0.55,
                transition: 'opacity 0.2s ease',
              }}>
                <div style={{ marginBottom: 14 }}>
                  <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <strong>{isDe ? 'Team-Größe' : 'Team size'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> die maximale Anzahl Personen pro Team (inkl. Team-Lead). Min. 2, Max. 20. Default 4.<br /><br />
                          <strong>Anzeige in der App:</strong> der Team-Lead sieht so viele Mitglied-Slots wie hier gesetzt; einzelne Slots können leer bleiben, ein Team ist also nicht zwingend voll.<br /><br />
                          <strong>Auswirkung für Teilnehmer:</strong> ein Team kann maximal so groß werden — versucht der Team-Lead, mehr Mitglieder einzutragen, wird er gestoppt.
                        </>
                      : <>
                          <strong>What this controls:</strong> the maximum number of people per team (including the team lead). Min. 2, max. 20. Default 4.<br /><br />
                          <strong>Where you see it:</strong> the team lead sees as many member slots as configured here; slots can stay empty, so teams are not required to be full.<br /><br />
                          <strong>For attendees:</strong> a team caps at this size — attempting to add more members is blocked.
                        </>
                    } />
                  </label>
                  <input
                    type="number"
                    className="form-input"
                    min={2}
                    max={20}
                    value={teamSize}
                    disabled={!teamRegistrationEnabled}
                    onChange={e => {
                      const v = parseInt(e.target.value, 10);
                      if (isNaN(v)) { setTeamSize(2); return; }
                      setTeamSize(Math.max(2, Math.min(20, v)));
                    }}
                    style={{ maxWidth: 120 }}
                  />
                </div>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}>
                  <input
                    type="checkbox"
                    checked={askTeamName}
                    disabled={!teamRegistrationEnabled}
                    onChange={e => setAskTeamName(e.target.checked)}
                    style={{ marginTop: 3, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Team-Namen abfragen' : 'Ask for team name'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> ob der Team-Lead beim Anmelden zusätzlich einen <strong>frei wählbaren Team-Namen</strong> eingeben muss (z.B. &bdquo;Die schnellen Sieben&ldquo;).<br /><br />
                          <strong>Anzeige in der App:</strong> der Team-Name erscheint auf der Seite &bdquo;Meine Events&ldquo; beim Team-Lead und allen Mitgliedern. Bei offenen Slots (Team noch nicht voll) wird der Team-Name in der Slot-Liste angezeigt, damit andere Teilnehmer bei Interesse beitreten können.<br /><br />
                          <strong>Auswirkung für Teilnehmer:</strong> macht das Team identifizierbar. Bleibt die Option aus, wird das Team nur intern über den Namen des Team-Leads referenziert.
                        </>
                      : <>
                          <strong>What this controls:</strong> whether the team lead has to enter a <strong>freely chosen team name</strong> during registration (e.g. &ldquo;The Fast Seven&rdquo;).<br /><br />
                          <strong>Where you see it:</strong> the team name appears on &ldquo;My Events&rdquo; for the team lead and all members. For open slots (team not full yet), the name is displayed in the slot list so other attendees can join.<br /><br />
                          <strong>For attendees:</strong> makes the team identifiable. If turned off, teams are referenced internally only via the team lead&apos;s name.
                        </>
                    } />
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? 'Wenn aktiv, gibt der Team-Lead bei der Anmeldung einen Team-Namen ein, der dann auf der MyEvents-Seite und in offenen Slots angezeigt wird.'
                        : 'When enabled, the team lead enters a team name during registration which is shown on the MyEvents page and in open slots.'}
                    </span>
                  </span>
                </label>
              </div>

              {/* v11.81: Beitritts-Modus — Sub-Box mit Modus + Sichtbarkeit + Approval */}
              <div style={{
                background: teamRegistrationEnabled ? '#ffffff' : 'var(--dex-gray-50, #fafafa)',
                borderRadius: 12, padding: '14px 16px', marginBottom: 12,
                border: '1px solid var(--dex-gray-200)',
                opacity: teamRegistrationEnabled ? 1 : 0.55,
                transition: 'opacity 0.2s ease',
              }}>
                <div style={{ fontWeight: 700, fontSize: '0.95rem', marginBottom: 10, color: 'var(--dex-gray-800)' }}>
                  {isDe ? 'Beitritts-Modus' : 'Join mode'}
                </div>

                {/* Radio-Group: komplette vs. Teil-Teams */}
                <div style={{ marginBottom: 14 }}>
                  <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginBottom: 8, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}>
                    <input
                      type="radio"
                      name="teamPartialMode"
                      checked={!teamPartialAllowed}
                      disabled={!teamRegistrationEnabled}
                      onChange={() => setTeamPartialAllowed(false)}
                      style={{ marginTop: 3, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}
                    />
                    <span style={{ flex: 1 }}>
                      <strong>{isDe ? 'Nur komplette Teams' : 'Only complete teams'}</strong>
                      <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                        {isDe
                          ? 'Der Team-Lead muss alle N Mitglieder beim Anmelden eintragen. Halbe Teams sind nicht möglich.'
                          : 'The team lead must enter all N members during registration. Partial teams are not possible.'}
                      </span>
                    </span>
                  </label>
                  <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}>
                    <input
                      type="radio"
                      name="teamPartialMode"
                      checked={teamPartialAllowed}
                      disabled={!teamRegistrationEnabled}
                      onChange={() => setTeamPartialAllowed(true)}
                      style={{ marginTop: 3, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}
                    />
                    <span style={{ flex: 1 }}>
                      <strong>{isDe ? 'Auch Teil-Teams erlaubt' : 'Partial teams allowed'}</strong>
                      <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                        {isDe
                          ? 'Der Team-Lead kann z.B. 2 von 4 Mitgliedern anmelden, die restlichen 2 Slots bleiben offen — andere Personen können später beitreten (siehe nächste Option).'
                          : 'The team lead can register e.g. 2 of 4 members; the remaining 2 slots stay open — others can join later (see next option).'}
                      </span>
                    </span>
                  </label>
                </div>

                {/* Checkbox: Sichtbarkeit offener Slots */}
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginBottom: 12, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}>
                  <input
                    type="checkbox"
                    checked={teamOpenSlotsVisible}
                    disabled={!teamRegistrationEnabled}
                    onChange={e => {
                      const v = e.target.checked;
                      setTeamOpenSlotsVisible(v);
                      if (!v) setTeamJoinRequiresApproval(false);
                    }}
                    style={{ marginTop: 3, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Unvollständige Teams öffentlich für Beitritt sichtbar' : 'Open teams publicly visible for joining'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> ob andere Teilnehmer Teams mit offenen Slots in der Anmeldeseite sehen und beitreten können.<br /><br />
                          <strong>Anzeige in der App:</strong> auf der Anmeldeseite erscheint eine Liste &bdquo;Teams mit freien Plätzen&ldquo; — pro Team mit der Anzahl freier Slots und (falls aktiviert) dem Team-Namen, aber <strong>ohne</strong> die Namen der bereits angemeldeten Mitglieder (Privatsphäre).<br /><br />
                          <strong>Auswirkung für Teilnehmer:</strong> wer noch in keinem Team ist, kann mit einem Klick einem offenen Slot beitreten — entweder sofort gültig oder erst nach Bestätigung durch den Team-Lead (siehe nächste Option).
                        </>
                      : <>
                          <strong>What this controls:</strong> whether other attendees see and can join teams with open slots on the registration page.<br /><br />
                          <strong>Where you see it:</strong> the registration page shows a list &ldquo;teams with free seats&rdquo; — per team with the count of free slots and (if enabled) the team name, but <strong>without</strong> the names of already-registered members (privacy).<br /><br />
                          <strong>For attendees:</strong> anyone not yet in a team can join an open slot with one click — either immediately or only after lead approval (see next option).
                        </>
                    } />
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? <>Wenn aktiv: andere Teilnehmer sehen offene Slots in der Registrierungsseite als &bdquo;Team mit X freien Plätzen&ldquo; — <strong>ohne</strong> die Namen der bereits angemeldeten Mitglieder (Privatsphäre).</>
                        : <>When active: other attendees see open slots on the registration page as &ldquo;team with X free seats&rdquo; — <strong>without</strong> the names of already-registered members (privacy).</>}
                    </span>
                  </span>
                </label>

                {/* Checkbox: Approval-Pflicht durch Team-Lead */}
                <label style={{
                  display: 'flex', alignItems: 'flex-start', gap: 10,
                  cursor: (teamRegistrationEnabled && teamOpenSlotsVisible) ? 'pointer' : 'not-allowed',
                  opacity: (teamRegistrationEnabled && teamOpenSlotsVisible) ? 1 : 0.55,
                  transition: 'opacity 0.2s ease',
                }}>
                  <input
                    type="checkbox"
                    checked={teamJoinRequiresApproval}
                    disabled={!teamRegistrationEnabled || !teamOpenSlotsVisible}
                    onChange={e => setTeamJoinRequiresApproval(e.target.checked)}
                    style={{ marginTop: 3, cursor: (teamRegistrationEnabled && teamOpenSlotsVisible) ? 'pointer' : 'not-allowed' }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Beitritt erfordert Bestätigung durch Team-Kapitän' : 'Joining requires team captain approval'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> ob jede Beitrittsanfrage zu einem offenen Team-Slot erst vom Team-Lead bestätigt werden muss.<br /><br />
                          <strong>Anzeige in der App:</strong> der Team-Lead bekommt eine Mail mit <strong>&bdquo;Bestätigen&ldquo;</strong>- und <strong>&bdquo;Ablehnen&ldquo;</strong>-Buttons pro Anfrage. Bis zur Bestätigung steht der Beitretende in einer Approve-Queue und ist noch nicht offiziell im Team.<br /><br />
                          <strong>Auswirkung für Teilnehmer:</strong> wenn aktiv, wird der Beitritt erst nach Bestätigung gültig — und der Beitretende bekommt erst dann seine Bestätigungsmail und (falls Outlook aktiv) den Kalendertermin. Wenn aus: Beitritt ist sofort gültig.
                        </>
                      : <>
                          <strong>What this controls:</strong> whether every join request to an open team slot has to be confirmed by the team lead first.<br /><br />
                          <strong>Where you see it:</strong> the team lead receives an email with <strong>&ldquo;Confirm&rdquo;</strong> and <strong>&ldquo;Reject&rdquo;</strong> buttons per request. Until confirmed, the joiner sits in an approve queue and is not yet officially in the team.<br /><br />
                          <strong>For attendees:</strong> if active, the join only becomes valid after confirmation — and the joiner receives their confirmation mail and (if Outlook is enabled) the calendar invite only at that point. If off: join is immediately valid.
                        </>
                    } />
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? 'Wenn aktiv: jeder Beitritt zu einem offenen Team geht erst in eine Approve-Queue. Der Team-Lead bekommt eine Mail mit „Bestätigen / Ablehnen"-Buttons. Erst nach Bestätigung ist die Person im Team. Wenn aus: Beitritt ist sofort gültig.'
                        : 'When active: every join to an open team enters an approve queue. The team lead gets an email with "Confirm / Reject" buttons. Only after confirmation is the person in the team. When off: joins are immediately valid.'}
                    </span>
                  </span>
                </label>
              </div>

              {/* v15: alter Hinweis „Logik folgt mit v11.82+" entfernt —
                  die komplette Team-Anmelde-Logik (Multi-Person-Form,
                  Mails, Outlook, Slot-Beitritt, Lead-Approval, Admin-Center-
                  Team-Management) ist seit v11.82–v11.86 live. */}

              {/* ===== v18.33: Self-Check-in ===== */}
              <div className="form-section" style={{ marginTop: 24, paddingTop: 20, borderTop: '1px solid var(--dex-gray-200, #eee)' }}>
                <div className="form-group">
                  <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    {isDe ? 'Self-Check-in per QR-Code' : 'Self check-in via QR code'}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Was du hier einstellst:</strong> ob sich Teilnehmer am Veranstaltungstag <strong>selbst einchecken</strong> können, indem sie einen event-spezifischen QR-Code scannen — ohne dass das Check-in-Team jeden einzeln abhaken muss.<br /><br />
                        <strong>Anzeige in der App:</strong> nach dem Speichern findest du im <strong>Admin Center</strong> dieses Events zwei Wege: ein <strong>druckbares QR-PDF</strong> (zum Aushängen, bequem) und eine <strong>rotierende Live-Anzeige</strong> für einen Bildschirm am Eingang (der Code wechselt automatisch, damit ein abfotografierter Code schnell verfällt).<br /><br />
                        <strong>Automatismen:</strong> wer den Code mit der Handy-Kamera scannt, wird automatisch als <strong>anwesend</strong> markiert — sofern er angemeldet und im erlaubten Zeitfenster ist. Jeder checkt immer nur <strong>sich selbst</strong> ein.<br /><br />
                        <strong>Auswirkung für Teilnehmer:</strong> kein Anstehen am Check-in-Schalter — einfach Code scannen, fertig.
                      </>
                    ) : (
                      <>
                        <strong>What you set here:</strong> whether attendees can <strong>check themselves in</strong> on the event day by scanning an event-specific QR code — no need for the check-in team to tick everyone off manually.<br /><br />
                        <strong>Where you see it:</strong> after saving you get two options in this event&apos;s <strong>admin center</strong>: a <strong>printable QR PDF</strong> (convenient, for posting) and a <strong>rotating live display</strong> for a screen at the entrance (the code changes automatically so a photographed code expires quickly).<br /><br />
                        <strong>Automations:</strong> whoever scans the code with their phone camera is automatically marked <strong>present</strong> — provided they are registered and within the allowed time window. Everyone only ever checks in <strong>themselves</strong>.<br /><br />
                        <strong>For attendees:</strong> no queue at the check-in desk — just scan and done.
                      </>
                    )} />
                  </label>
                  <div className="toggle-wrapper" style={{ marginTop: 4 }}>
                    <label className="toggle">
                      <input
                        type="checkbox"
                        checked={selfCheckInEnabled}
                        onChange={e => {
                          const on = e.target.checked;
                          setSelfCheckInEnabled(on);
                          if (on) {
                            if (!selfCheckInToken) setSelfCheckInToken(generateSelfCheckInToken());
                            setShowSelfCheckInModal(true);
                          }
                        }}
                      />
                      <span className="toggle-slider" />
                    </label>
                    <span style={{ fontSize: '0.9rem' }}>{selfCheckInEnabled ? t('create.enabled') : t('create.disabled')}</span>
                  </div>
                </div>

                {selfCheckInEnabled && (
                  <div style={{ marginTop: 14, padding: 16, background: 'var(--dex-gray-50, #f7f7f5)', borderRadius: 10, border: '1px solid var(--dex-gray-200, #eee)' }}>
                    <div style={{ fontSize: '0.9rem', color: 'var(--dex-gray-700, #444)', lineHeight: 1.5, marginBottom: 14 }}>
                      {isDe
                        ? 'Der QR-Code (PDF zum Drucken sowie die rotierende Live-Anzeige) steht nach dem Speichern im Admin Center dieses Events bereit.'
                        : 'The QR code (printable PDF and the rotating live display) is available in this event\'s admin center after saving.'}
                    </div>

                    {/* Optionales Zeitfenster */}
                    <div style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-700, #444)', marginBottom: 8 }}>
                      {isDe ? 'Check-in-Zeitfenster (optional)' : 'Check-in time window (optional)'}
                    </div>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                      <div className="form-group">
                        <label className="form-label" style={{ fontSize: '0.8rem' }}>{isDe ? 'Von' : 'From'}</label>
                        <input
                          type="datetime-local"
                          className="form-input"
                          value={selfCheckInFrom}
                          onChange={e => setSelfCheckInFrom(e.target.value)}
                        />
                      </div>
                      <div className="form-group">
                        <label className="form-label" style={{ fontSize: '0.8rem' }}>{isDe ? 'Bis' : 'Until'}</label>
                        <input
                          type="datetime-local"
                          className="form-input"
                          value={selfCheckInTo}
                          onChange={e => setSelfCheckInTo(e.target.value)}
                        />
                      </div>
                    </div>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500, #888)', marginTop: 6 }}>
                      {isDe
                        ? 'Leer lassen = Check-in ist nur am Veranstaltungstag möglich (vom Start- bis zum End-Datum).'
                        : 'Leave empty = check-in is only possible on the event day (from start to end date).'}
                    </div>

                    <div style={{ display: 'flex', gap: 10, marginTop: 14, flexWrap: 'wrap' }}>
                      <button
                        type="button"
                        onClick={() => setShowSelfCheckInModal(true)}
                        style={{ padding: '8px 14px', borderRadius: 8, border: '1px solid var(--dex-gray-300, #ccc)', background: '#fff', color: 'var(--dex-gray-700, #444)', fontWeight: 600, cursor: 'pointer', fontSize: '0.85rem' }}
                      >
                        <Icon iconName="Info" style={{ marginRight: 6 }} />
                        {isDe ? 'Wie funktioniert das?' : 'How does it work?'}
                      </button>
                      {!!(editEvent?.selfCheckInEnabled && editEvent?.selfCheckInToken) && (
                        <button
                          type="button"
                          onClick={() => downloadSelfCheckInPdf({
                            eventTitle: title || 'Event',
                            token: editEvent.selfCheckInToken as string,
                          })}
                          style={{ padding: '8px 14px', borderRadius: 8, border: '1px solid var(--dex-green, #86bc25)', background: 'var(--dex-green, #86bc25)', color: '#fff', fontWeight: 600, cursor: 'pointer', fontSize: '0.85rem' }}
                        >
                          <Icon iconName="PDF" style={{ marginRight: 6 }} />
                          {isDe ? 'QR-PDF herunterladen' : 'Download QR PDF'}
                        </button>
                      )}
                    </div>
                  </div>
                )}
              </div>

              </div>

              {/* ===== Step 5 (v15: vormals Step 6): Registrierungsfelder ===== */}
              <div style={{ display: currentStep === 4 ? 'block' : 'none' }}>
              <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                {isDe ? 'Schritt 5 — Felder' : 'Step 5 — Fields'}
              </h2>
              <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                {isDe
                  ? <><strong>Optional</strong> — die Standard-Teilnehmerdaten (<strong>Vorname, Nachname, E-Mail</strong>) werden bei jeder Anmeldung automatisch erfasst, dazu kommen aus dem Deloitte-Profil <strong>Job Title, Standort, Department und Telefonnummer</strong>. Hier in Schritt 6 ergänzt du <strong>nur zusätzliche Fragen</strong>, die du speziell für dieses Event brauchst — vom T-Shirt-Größen-Dropdown bis zur Pflicht-Checkbox für AGB / Datenschutz. Optional kannst du oben das <strong>Anrede</strong>-Dropdown einblenden. Wenn dein Event keine Zusatzfragen braucht, kannst du diesen Schritt einfach leer lassen.</>
                  : <><strong>Optional</strong> — the standard attendee data (<strong>first name, last name, email</strong>) is captured automatically for every registration, plus <strong>job title, location, department and phone</strong> are pulled from the Deloitte profile. In step 6 you only add <strong>extra questions</strong> specific to this event — from a T-shirt size dropdown to a privacy / terms required checkbox. Optionally enable the <strong>salutation</strong> dropdown on top. If your event needs no extra questions, you can simply leave this step empty.</>}
              </p>

              {/* v15: Anrede-Toggle ist nach UNTEN den Datenschutz-Hinweis
                  gewandert — Organizer soll erst den Sammle-keine-sensiblen-
                  Daten-Hinweis lesen, dann erst die Optional-Anrede-Checkbox
                  setzen. Siehe weiter unten. */}

              {renderStepIntro(
                [
                  'Feldtyp wählen: Text, Zahl, Dropdown, Checkbox, Personen-Suche oder Roommate (Doppelzimmer)',
                  'Mehrfachauswahl bei Dropdowns (z.B. mehrere Allergien anhaken)',
                  'Pflichtfeld setzen (rotes Sternchen, Anmeldung blockiert wenn leer)',
                  'Beschreibung pro Feld — landet als „i"-Tooltip neben dem Feld-Label',
                  'Sichtbarkeitsbedingung: Feld nur dann anzeigen wenn eine andere Frage einen bestimmten Wert hat (z.B. „Zimmerart nur fragen wenn Hotel = ja")',
                  'Reihenfolge per Drag oder Pfeilen — die Nummerierung passt sich automatisch an',
                ],
                [
                  'Pick a field type: text, number, dropdown, checkbox, people search or roommate (double room)',
                  'Multi-select for dropdowns (e.g. tick multiple allergies)',
                  'Mark required (red asterisk, blocks submit when empty)',
                  'Description per field — appears as „i" tooltip next to the field label',
                  'Visibility condition: only show this field when another question has a specific value (e.g. „Only ask room type if Hotel = yes")',
                  'Reordering via drag or arrows — numbering updates automatically',
                ]
              )}

              {/* Datenschutz-Hinweis ueber der Template-Auswahl — links
                  angeordnet, orangener Akzent damit der Organizer beim Anlegen
                  neuer Felder bewusst entscheidet, was wirklich abgefragt
                  werden muss. Seit v7.35 deckungsgleich mit dem Hinweis aus
                  den Nutzungsbedingungen (Sammeln keiner sensiblen Daten),
                  damit Organizer den selben Wortlaut wie bei der initialen
                  Bestaetigung sehen. */}
              <div style={{
                display: 'flex', alignItems: 'flex-start', gap: 10,
                padding: '12px 14px', marginBottom: 16,
                background: 'rgba(237,139,0,0.06)',
                border: '1px solid var(--dex-orange, #ed8b00)',
                borderRadius: 'var(--dex-radius, 12px)',
                fontSize: '0.82rem', color: 'var(--dex-gray-700)',
                lineHeight: 1.5,
              }}>
                <span style={{
                  flexShrink: 0, fontSize: '1.1rem', lineHeight: 1,
                  color: 'var(--dex-orange, #ed8b00)', fontWeight: 700,
                }}>⚠</span>
                <div>
                  <strong style={{ color: 'var(--dex-orange, #ed8b00)' }}>
                    {isDe ? 'Sammle keine sensiblen personenbezogenen Daten' : 'Do not collect sensitive personal data'}
                  </strong>{' '}
                  {isDe
                    ? <>— das heißt: keine Daten bezüglich Rasse oder ethnischer Herkunft, religiöser oder philosophischer Überzeugungen, Gewerkschaftsmitgliedschaft, politischer Meinungen, medizinischer oder gesundheitlicher Zustände oder Informationen über das Sexualleben oder die sexuelle Orientierung einer Person. Falls sensible personenbezogene Daten gesammelt werden müssen, kontaktiere zuerst das Team unter <a href="mailto:privacy@deloitte.de" style={{ color: 'var(--dex-orange, #ed8b00)', fontWeight: 600 }}>privacy@deloitte.de</a>.</>
                    : <>— that means: no data on race or ethnic origin, religious or philosophical beliefs, trade-union membership, political opinions, medical or health conditions, or information about a person&apos;s sex life or sexual orientation. If sensitive personal data must be collected, contact the team first at <a href="mailto:privacy@deloitte.de" style={{ color: 'var(--dex-orange, #ed8b00)', fontWeight: 600 }}>privacy@deloitte.de</a>.</>}
                </div>
              </div>

              {/* v18.57: Anrede-Abfrage-Toggle nach unten verschoben — sitzt jetzt
                  direkt unter den „Vorgeschlagene Felder"-Buttons. */}

              {/* v18.35: Anmeldesprache vorgeben */}
              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '12px 16px', marginBottom: 14,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 6 }}>
                  {isDe ? 'Sprache des Anmeldeformulars' : 'Registration form language'}
                  <InfoTooltip text={isDe
                    ? <>
                        <strong>Was du hier einstellst:</strong> in welcher Sprache die <strong>komplette Anmeldeseite</strong> (alle Texte, Buttons und der <strong>Datenschutz-Disclaimer</strong>) angezeigt wird.<br /><br />
                        <strong>Anzeige in der App:</strong> bei <strong>Automatisch</strong> folgt die Anmeldeseite der App-Sprache des Teilnehmers. Wählst du <strong>Immer Deutsch</strong> oder <strong>Immer Englisch</strong>, wird die Anmeldeseite <strong>fest in dieser Sprache</strong> angezeigt — auch wenn der Teilnehmer die App z.&nbsp;B. auf Deutsch nutzt. Ein kleiner Hinweis im Kopfbereich zeigt das an.<br /><br />
                        <strong>Auswirkung für Teilnehmer:</strong> bei einem englischsprachigen Event siehst du die Anmeldung samt Disclaimer auf Englisch, egal welche App-Sprache du eingestellt hast.
                      </>
                    : <>
                        <strong>What you set here:</strong> the language in which the <strong>entire registration page</strong> (all texts, buttons and the <strong>privacy disclaimer</strong>) is shown.<br /><br />
                        <strong>Where you see it:</strong> with <strong>Automatic</strong> the page follows the attendee&apos;s app language. Choosing <strong>Always German</strong> or <strong>Always English</strong> forces the registration page into that language — even if the attendee uses the app in another language. A small hint in the header indicates this.<br /><br />
                        <strong>For attendees:</strong> for an English-language event you see the registration and disclaimer in English regardless of your app language.
                      </>
                  } />
                </label>
                <select
                  className="form-input"
                  value={registrationLanguage}
                  onChange={e => setRegistrationLanguage(e.target.value as '' | 'de' | 'en')}
                  style={{ width: '100%', maxWidth: 460 }}
                >
                  <option value="">{isDe ? 'Automatisch (App-Sprache des Teilnehmers)' : 'Automatic (attendee\'s app language)'}</option>
                  <option value="de">{isDe ? 'Immer Deutsch' : 'Always German'}</option>
                  <option value="en">{isDe ? 'Immer Englisch' : 'Always English'}</option>
                </select>
                <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 6 }}>
                  {isDe
                    ? 'Default: Automatisch. Bei fester Sprache wird die Anmeldeseite inkl. Disclaimer immer in dieser Sprache angezeigt.'
                    : 'Default: Automatic. With a fixed language the registration page incl. disclaimer is always shown in that language.'}
                </span>
              </div>

              {/* v18.57: Deutsch/Englisch-Toggle nach unten verschoben — sitzt jetzt
                  direkt unter den „Vorgeschlagene Felder"-Buttons. */}

              {/* v10.21: Template-Dropdown ist entfallen — der Organizer
                  pickt B2Run-Felder einzeln per Suggested-Felder-Modal
                  (eingeklappte Sektion "B2Run-spezifische Felder"). Damit
                  fuehrt kein Weg mehr ueber ein hartes B2Run-Template, das
                  zusaetzlich Logik (Auto-Split-Capacity etc.) ausloeste —
                  saubere Trennung zwischen Feld-Konfiguration und
                  Kapazitaets-Modell. */}

              {/* B2Run Startbloecke - moderne Liste mit + Button. Wird
                  unverändert angezeigt, sobald das b2run_startblock-Feld in
                  customFields steht (ueber das Suggested-Felder-Modal
                  ausgewaehlt oder beim Edit eines Legacy-Events vorhanden). */}
              {customFields.some(f => f.id === 'b2run_startblock') && (
                <div className="form-group" style={{ marginBottom: 24, padding: 16, background: 'var(--dex-green-light, #f0fdf4)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-green)' }}>
                  <label className="form-label" style={{ marginBottom: 4 }}>
                    {t('create.startblocks')}
                    <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> die <strong>Startblöcke</strong>, in denen Teilnehmer ihren Lauf starten — z.B. Block A (schnell), Block B (mittel), Block C (Walking). Pro Block ein eigener Eintrag.<br /><br />
                      <strong>Anzeige in der App:</strong> bei der Anmeldung erscheint ein <strong>Dropdown Startblock</strong>, das diese Liste enthält. Falls du oben in Schritt 3 eine Starter-Typ-Zuordnung gemacht hast, ist das Dropdown automatisch gefüllt und disabled.<br /><br />
                      <strong>Auswirkung für Teilnehmer:</strong> der gewählte Startblock landet in der Bestätigungs-Mail und im Admin Center — wichtig damit ihr beim Veranstalter wisst, in welcher Welle wer startet.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>start blocks</strong> in which attendees begin their run — e.g. block A (fast), block B (medium), block C (walking). One entry per block.<br /><br />
                      <strong>Shown in the app:</strong> at registration, a <strong>start-block dropdown</strong> appears that contains this list. If you set up a starter-type mapping in step 3 above, the dropdown is auto-filled and disabled.<br /><br />
                      <strong>Effect for attendees:</strong> the selected start block ends up in the confirmation mail and in the admin center — important so you know which wave each attendee is in when coordinating with the organiser.
                    </>
                  )} />
                  </label>
                  <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                    {t('create.startblocks.hint')}
                  </p>

                  {/* Bestehende Startbloecke als Liste */}
                  {b2runStartblocks.length > 0 && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 12 }}>
                      {b2runStartblocks.map((block, idx) => (
                        <div
                          key={idx}
                          style={{
                            display: 'flex', alignItems: 'center', gap: 8,
                            padding: '8px 12px', borderRadius: 'var(--dex-radius, 12px)',
                            background: '#fff', border: '1px solid var(--dex-gray-200)',
                          }}
                        >
                          <Icon iconName="Running" style={{ fontSize: 16, color: 'var(--dex-green-dark, #6b9a1e)', flexShrink: 0 }} />
                          <span style={{ flex: 1, fontSize: '0.88rem' }}>{block}</span>
                          <button
                            type="button"
                            onClick={() => removeStartblock(block)}
                            title={t('create.startblocks.remove')}
                            style={{
                              border: 'none', background: 'transparent', cursor: 'pointer',
                              color: 'var(--dex-red, #c00)', padding: 4, borderRadius: 4,
                              display: 'inline-flex', alignItems: 'center',
                            }}
                          >
                            <Trash2 size={14} />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}

                  {/* Neues Startblock hinzufuegen */}
                  <div style={{ display: 'flex', gap: 8 }}>
                    <input
                      type="text"
                      className="form-input"
                      value={newStartblock}
                      onChange={e => setNewStartblock(e.target.value)}
                      onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); addStartblock(); } }}
                      placeholder={t('create.startblocks.placeholder')}
                      style={{ flex: 1 }}
                    />
                    <button
                      type="button"
                      className="btn btn-primary"
                      onClick={addStartblock}
                      disabled={!newStartblock.trim()}
                      style={{ whiteSpace: 'nowrap' }}
                    >
                      <Plus size={14} /> {t('create.startblocks.add')}
                    </button>
                  </div>
                </div>
              )}

              {/* v15.0: pro-Sub-Event-Tabs fuer Felder. Tab 0 = Haupt-Event
                  (komplette Custom-Fields-Liste, B2Run-Startbloecke etc.).
                  Tabs N>0 = schlanke per-Sub-Event-Felder-UI mit Inheritance-
                  Toggle. Im subEventsOnlyMode wird Tab 0 zu „Uebergreifende
                  Felder" / „Cross-cutting fields" — die wirken dann auf alle
                  Sub-Event-Anmeldungen. */}
              {renderPerEventTabStrip(
                activeFieldsTabIdx,
                setActiveFieldsTabIdx,
                subEventsOnlyMode
                  ? (isDe ? 'Übergreifende Felder' : 'Cross-cutting fields')
                  : `${isDe ? 'Haupt-Event' : 'Main event'}: ${title || (isDe ? 'Ohne Titel' : 'Untitled')}`,
                isDe ? 'Event-Tab wechseln (Felder)' : 'Switch event tab (fields)'
              )}

              {activeFieldsTabIdx > 0 && (() => {
                const seIdx = activeFieldsTabIdx - 1;
                const se = subEvents[seIdx];
                if (!se) return null;
                const inherit = false;  // v15.3: inheritance entfernt — Sub-Events haben eigene Felder
                const seFields = se.customFields || [];
                const updateSub = (patch: Partial<SubEventDraft>): void => {
                  setSubEvents(prev => prev.map((x, i) => i === seIdx ? { ...x, ...patch } : x));
                };
                return (
                  <div>
                    {/* v15.6: Lead-paragraph analog Hauptevent-Tab. */}
                    <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                      {isDe
                        ? <><strong>Optional</strong> — die Standard-Teilnehmerdaten (Vorname, Nachname, E-Mail) und Profil-Daten (Job Title, Standort, Department, Telefon) werden automatisch erfasst. Hier ergänzt du <strong>nur Zusatzfragen speziell für dieses Sub-Event</strong>. Wenn das Sub-Event keine eigenen Fragen braucht, kannst du diese Sektion leer lassen.</>
                        : <><strong>Optional</strong> — the standard attendee data (first name, last name, email) and profile data (job title, location, department, phone) are captured automatically. Here you only add <strong>extra questions specific to this sub-event</strong>. If the sub-event needs no extra questions, you can leave this section empty.</>}
                    </p>

                    {/* v18.62: Datenschutz-Hinweis hier ENTFERNT — er steht bereits
                        einmal oben in Schritt 5 (über der Tab-Leiste). Eine
                        Wiederholung pro Sub-Event-Tab ist redundant. */}

                    {/* v15.3: „Anrede abfragen"-Toggle pro Sub-Event */}
                    <div style={{
                      background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                      padding: '14px 16px', marginBottom: 16,
                      border: '1px solid var(--dex-gray-200)',
                      display: 'flex', alignItems: 'center', gap: 10, justifyContent: 'space-between', flexWrap: 'wrap',
                    }}>
                      <label style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={!!se.askSalutation}
                          onChange={e => updateSub({ askSalutation: e.target.checked })}
                          style={{ cursor: 'pointer' }}
                        />
                        <span style={{ fontSize: '0.9rem', fontWeight: 600 }}>
                          {isDe ? 'Anrede für dieses Sub-Event abfragen' : 'Ask for salutation on this sub-event'}
                        </span>
                      </label>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.8rem', padding: '6px 14px' }}
                        onClick={() => updateSub({ askSalutation: askSalutation })}
                        title={isDe
                          ? 'Uebernimmt die Anrede-Abfrage-Einstellung vom Hauptevent'
                          : 'Copies the salutation toggle from the main event'}
                      >
                        {isDe ? 'Vom Hauptevent kopieren' : 'Copy from main event'}
                      </button>
                    </div>
                    <div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12, flexWrap: 'wrap' }}>
                        <strong style={{ fontSize: '0.95rem' }}>
                          {se.title || (isDe ? '(unbenanntes Sub-Event)' : '(unnamed sub-event)')}
                        </strong>
                        <button
                          type="button"
                          className="btn btn-outline"
                          style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                          onClick={() => addSubEventCustomField(se.id)}
                        >
                          <Plus size={12} /> {isDe ? 'Feld hinzufügen' : 'Add field'}
                        </button>
                        {customFields.length > 0 && seFields.length === 0 && (
                          <button
                            type="button"
                            className="btn btn-outline"
                            style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                            onClick={() => copyParentFieldsToSubEvent(se.id)}
                            title={isDe ? `Dupliziert die ${customFields.length} Felder vom Hauptevent als Startpunkt` : 'Duplicates the main-event fields as a starting point'}
                          >
                            {isDe ? `Felder vom Hauptevent kopieren (${customFields.length})` : `Copy fields from main event (${customFields.length})`}
                          </button>
                        )}
                      </div>
                      {seFields.length === 0 ? (
                        <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)', fontStyle: 'italic' }}>
                          {isDe
                            ? 'Keine zusätzlichen Felder definiert. Du kannst Felder hinzufügen oder vom Hauptevent kopieren.'
                            : 'No additional fields defined. You can add fields or copy them from the main event.'}
                        </p>
                      ) : (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                          {seFields.map((field, idx) => (
                            <div
                              key={field.id}
                              style={{
                                background: 'var(--dex-gray-50, #fafafa)',
                                borderRadius: 12,
                                padding: 16,
                                border: '1px solid var(--dex-gray-200)',
                              }}
                            >
                              <div style={{ display: 'flex', gap: 10, alignItems: 'center', marginBottom: 12, flexWrap: 'wrap' }}>
                                <span style={{
                                  flexShrink: 0, width: 26, height: 26, borderRadius: '50%',
                                  background: 'var(--dex-green, #86bc25)', color: '#fff',
                                  display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                  fontWeight: 700, fontSize: '0.78rem',
                                }}>{idx + 1}</span>
                                <input
                                  className="form-input"
                                  value={field.label}
                                  placeholder={isDe ? 'Frage / Feld-Label (z.B. „Welche Strecke?")' : 'Question / field label (e.g. „Which distance?")'}
                                  onChange={e => updateSubEventCustomField(se.id, field.id, { label: e.target.value })}
                                  disabled={inherit}
                                  style={{
                                    flex: '0 1 320px', minWidth: 180, maxWidth: 320,
                                    fontSize: '1rem', fontWeight: 600,
                                    padding: '8px 12px',
                                    color: field.label ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
                                  }}
                                />
                                <select
                                  className="form-select"
                                  value={field.type}
                                  disabled={inherit}
                                  onChange={e => updateSubEventCustomField(se.id, field.id, { type: e.target.value as CustomFieldInput['type'] })}
                                  title={isDe ? 'Art des Feldes' : 'Field type'}
                                  style={{
                                    flex: '0 0 200px', maxWidth: 200,
                                    background: 'rgba(134,188,37,0.08)',
                                    border: '1px solid var(--dex-green, #86bc25)',
                                    color: 'var(--dex-green-dark, #4a7c1f)',
                                    fontWeight: 600,
                                    padding: '8px 10px',
                                  }}
                                >
                                  <option value="text">{isDe ? 'Text (Freitext)' : 'Text (free text)'}</option>
                                  <option value="select">{isDe ? 'Dropdown' : 'Dropdown'}</option>
                                  <option value="number">{isDe ? 'Zahl' : 'Number'}</option>
                                  <option value="checkbox">{isDe ? 'Checkbox' : 'Checkbox'}</option>
                                </select>
                                <label
                                  style={{
                                    display: 'inline-flex', alignItems: 'center', gap: 6,
                                    padding: '6px 12px', borderRadius: 999,
                                    fontSize: '0.78rem', fontWeight: 600, whiteSpace: 'nowrap',
                                    cursor: inherit ? 'not-allowed' : 'pointer', userSelect: 'none',
                                    border: `1px solid ${field.required ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                    background: field.required ? 'rgba(134,188,37,0.10)' : '#fff',
                                    color: field.required ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                                  }}
                                >
                                  <input
                                    type="checkbox"
                                    checked={field.required}
                                    disabled={inherit}
                                    onChange={e => updateSubEventCustomField(se.id, field.id, { required: e.target.checked })}
                                    style={{ display: 'none' }}
                                  />
                                  <span style={{ fontSize: '0.85rem', lineHeight: 1 }}>{field.required ? '✓' : '○'}</span>
                                  {t('create.required')}
                                </label>
                                <button
                                  type="button"
                                  onClick={() => removeSubEventCustomField(se.id, field.id)}
                                  disabled={inherit}
                                  title={isDe ? 'Feld entfernen' : 'Remove field'}
                                  style={{ background: 'none', border: 'none', color: 'var(--dex-red)', padding: 4, cursor: inherit ? 'not-allowed' : 'pointer', flexShrink: 0 }}
                                >
                                  <X size={18} />
                                </button>
                              </div>
                              <div style={{ marginLeft: 32, marginTop: 10 }}>
                                <input
                                  className="form-input"
                                  placeholder={isDe
                                    ? 'Beschreibung (optional, erscheint als „i"-Tooltip neben dem Feld)'
                                    : 'Description (optional, shown as „i" tooltip next to the field)'}
                                  value={field.helpText || ''}
                                  disabled={inherit}
                                  onChange={e => updateSubEventCustomField(se.id, field.id, { helpText: e.target.value })}
                                  style={{ width: '100%', fontSize: '0.82rem', padding: '6px 10px' }}
                                />
                              </div>
                              {field.type === 'select' && (
                                <div style={{
                                  marginTop: 10, marginLeft: 32, padding: '12px 14px',
                                  background: '#fff',
                                  border: '1px solid var(--dex-gray-200)',
                                  borderRadius: 8,
                                }}>
                                  <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-700)', fontWeight: 600, marginBottom: 8 }}>
                                    {isDe ? 'Antwort-Optionen' : 'Answer options'}
                                  </div>
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                                    {field.options.map((opt, oidx) => (
                                      <div key={oidx} style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                                        <span style={{
                                          flexShrink: 0, width: 22, height: 22, borderRadius: '50%',
                                          background: 'var(--dex-gray-200)', color: 'var(--dex-gray-700)',
                                          display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                          fontWeight: 700, fontSize: '0.72rem',
                                        }}>{oidx + 1}</span>
                                        <input
                                          className="form-input"
                                          placeholder={isDe ? `Option ${oidx + 1}` : `Option ${oidx + 1}`}
                                          value={opt}
                                          disabled={inherit}
                                          onChange={e => {
                                            const next = field.options.slice();
                                            next[oidx] = e.target.value;
                                            updateSubEventCustomField(se.id, field.id, { options: next });
                                          }}
                                          style={{ flex: 1, fontSize: '0.85rem', padding: '6px 10px' }}
                                        />
                                        {field.options.length > 1 && (
                                          <button
                                            type="button"
                                            disabled={inherit}
                                            onClick={() => {
                                              const next = field.options.filter((_, i) => i !== oidx);
                                              updateSubEventCustomField(se.id, field.id, { options: next });
                                            }}
                                            title={isDe ? 'Option entfernen' : 'Remove option'}
                                            style={{ background: 'none', border: 'none', color: 'var(--dex-gray-500)', padding: 4, cursor: inherit ? 'not-allowed' : 'pointer' }}
                                          >
                                            <X size={14} />
                                          </button>
                                        )}
                                      </div>
                                    ))}
                                    <button
                                      type="button"
                                      disabled={inherit}
                                      onClick={() => updateSubEventCustomField(se.id, field.id, { options: [...field.options, ''] })}
                                      style={{ alignSelf: 'flex-start', background: 'none', border: '1px dashed var(--dex-gray-300)', padding: '4px 12px', fontSize: '0.78rem', borderRadius: 6, cursor: inherit ? 'not-allowed' : 'pointer', color: 'var(--dex-gray-700)', display: 'inline-flex', alignItems: 'center', gap: 4, marginTop: 4 }}
                                    >
                                      <Plus size={12} /> {isDe ? 'Option hinzufügen' : 'Add option'}
                                    </button>
                                  </div>
                                </div>
                              )}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                );
              })()}

              {/* Tab 0 (Haupt-Event bzw. Uebergreifende Felder): die
                  bestehende Hauptevent-Felder-UI bleibt unveraendert; nur
                  die Sektion „Felder pro Sub-Event" weiter unten wird in
                  v15.0 ausgeblendet, weil pro Sub-Event jetzt einen
                  eigenen Tab. */}
              <div style={{ display: activeFieldsTabIdx === 0 ? 'block' : 'none' }}>
              {/* v16.5: In Step 5 (Felder) ist im subEventsOnlyMode KEIN
                  Greyout — die Felder im ersten Tab sind „uebergreifend"
                  und werden bei JEDER Sub-Event-Anmeldung abgefragt, also
                  in dem Modus besonders relevant. Stattdessen ein info-
                  blauer Hinweis-Banner mit der korrekten Erklaerung. */}
              {subEventsOnlyMode && (() => {
                const termPl = (childTermPlural || (isDe ? 'Sub-Events' : 'sub-events')).trim() || (isDe ? 'Sub-Events' : 'sub-events');
                const termSg = (childTermSingular || (isDe ? 'Sub-Event' : 'sub-event')).trim() || (isDe ? 'Sub-Event' : 'sub-event');
                return (
                  <div style={{
                    display: 'flex', alignItems: 'flex-start', gap: 10,
                    padding: '12px 14px', marginBottom: 16,
                    background: 'rgba(33,150,243,0.06)',
                    border: '1px solid var(--dex-info, #2196f3)',
                    borderRadius: 'var(--dex-radius, 12px)',
                    fontSize: '0.85rem', color: 'var(--dex-gray-700)',
                    lineHeight: 1.5,
                  }}>
                    <Icon iconName="Info" style={{ fontSize: 18, color: 'var(--dex-info, #2196f3)', flexShrink: 0, marginTop: 2 }} />
                    <div>
                      {isDe
                        ? <><strong>Übergreifende Felder</strong> — diese Fragen werden bei der Anmeldung zu jeder einzelnen {termSg} abgefragt. Trag hier ein, was für alle {termPl} gemeinsam gilt; spezifische Felder pro {termSg} pflegst du in den jeweiligen Tabs rechts daneben.</>
                        : <><strong>Cross-cutting fields</strong> — these questions are asked once per {termSg} registration. Configure here what applies across all {termPl}; per-{termSg} specifics go into the individual tabs on the right.</>}
                    </div>
                  </div>
                );
              })()}
              <div>
              {/* Dynamische Felder */}
              <div>
                {/* Bereich-Header: trennt Hauptevent-Felder visuell vom
                    Sub-Event-Block weiter unten (v10.11+).
                    v15.0: im subEventsOnlyMode lautet die Ueberschrift
                    „Uebergreifend fuer alle <childTermPlural>". */}
                <h3 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.15rem', fontWeight: 700 }}>
                  {subEventsOnlyMode
                    ? (isDe
                        ? `Übergreifend für alle ${(childTermPlural || 'Sub-Events').trim() || 'Sub-Events'}`
                        : `Across all ${(childTermPlural || 'sub-events').trim() || 'sub-events'}`)
                    : (isDe ? 'Felder für das Hauptevent' : 'Fields for the main event')}
                </h3>
                <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                  {isDe
                    ? 'Diese Felder werden bei jeder Anmeldung abgefragt — egal ob das Event Sub-Events hat oder nicht. Für Sub-Event-spezifische Fragen wechsle oben auf den jeweiligen Sub-Event-Tab.'
                    : 'These fields are asked at every registration — regardless of whether the event has sub-events. For sub-event-specific questions switch to the respective sub-event tab above.'}
                </p>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12 }}>
                  <StepBadge n={19} />
                  {isDe ? 'Eigene Abfragen / Felder' : 'Custom fields'}
                </label>
                {/* v7.20: "Vorgeschlagene Felder" + "Feld hinzufuegen" stehen
                    nach links (vor dem Custom-Fields-Label), damit alle Action-
                    Buttons konsistent links aligned sind (wie auch der Typ-
                    Selector pro Feld). */}
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 16, flexWrap: 'wrap' }}>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    onClick={openSuggestedModal}
                    style={{ fontSize: '0.85rem', padding: '6px 14px' }}
                    title={isDe ? 'Felder aus einem Katalog waehlen' : 'Pick fields from a catalog'}
                  >
                    {isDe ? 'Vorgeschlagene Felder' : 'Suggested fields'}
                  </button>
                  <button className="btn btn-outline" onClick={addCustomField} style={{ fontSize: '0.85rem', padding: '6px 14px' }}>
                    <Plus size={14} /> {t('create.addfield')}
                  </button>
                  {customFields.length > 1 && (
                    <button
                      type="button"
                      className={reorderMode ? 'btn btn-primary' : 'btn btn-outline'}
                      onClick={() => setReorderMode(prev => !prev)}
                      style={{ fontSize: '0.85rem', padding: '6px 14px' }}
                      title={isDe ? 'Felder per Hoch/Runter-Pfeile sortieren' : 'Reorder fields with up/down arrows'}
                    >
                      {reorderMode
                        ? (isDe ? 'Fertig' : 'Done')
                        : (isDe ? 'Reihenfolge ändern' : 'Reorder')}
                    </button>
                  )}
                </div>

              {/* v18.57: Anrede-Abfrage + Deutsch/Englisch — von oben hierher
                  verschoben, direkt unter die „Vorgeschlagene Felder"-Buttons. */}
              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '12px 16px', marginBottom: 14,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={askSalutation}
                    onChange={e => setAskSalutation(e.target.checked)}
                    style={{ marginTop: 3, cursor: 'pointer' }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Anrede abfragen?' : 'Ask for salutation?'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> ob der Teilnehmer bei der Anmeldung sein <strong>Geschlecht / die Anrede</strong> (Frau, Herr, Divers, Keine Angabe) angeben muss. Default: <strong>nein</strong> — viele Events brauchen die Anrede nicht und ersparen den Teilnehmern das Feld.<br /><br />
                          <strong>Anzeige in der App:</strong> wenn aktiviert, erscheint im Registrierungsformular ein Pflicht-Dropdown <strong>Anrede</strong> direkt über dem Vorname-Feld. Wenn aus, wird das Feld komplett ausgeblendet und die gespeicherte Anrede bleibt leer.<br /><br />
                          <strong>Auswirkung für Teilnehmer:</strong> wenn aktiv, können sie sich erst anmelden, wenn sie die Anrede gewählt haben. Wenn aus, überspringen sie diesen Schritt komplett.
                        </>
                      : <>
                          <strong>What this controls:</strong> whether attendees have to provide their <strong>salutation / gender</strong> (Mrs, Mr, Diverse, Prefer not to say) when registering. Default: <strong>no</strong> — many events do not need it and skip the field for attendees.<br /><br />
                          <strong>Where you see it:</strong> when enabled, a required <strong>salutation</strong> dropdown appears in the registration form directly above the first name field. When disabled, the field is hidden completely and the stored salutation stays empty.<br /><br />
                          <strong>For attendees:</strong> when enabled, they can only submit once they have picked a salutation. When disabled, they skip this step entirely.
                        </>
                    } />
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? 'Default: nein — Wenn aktiviert, sehen Teilnehmer ein Anrede-Dropdown im Registrierungsformular.'
                        : 'Default: no — when enabled, attendees see a salutation dropdown in the registration form.'}
                    </span>
                  </span>
                </label>
              </div>

              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '12px 16px', marginBottom: 14,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={bilingualFields}
                    onChange={e => setBilingualFields(e.target.checked)}
                    style={{ marginTop: 3, cursor: 'pointer' }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Deutsch und Englisch ermöglichen' : 'Offer German and English'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> ob du pro Custom-Field, das du unten anlegst, <strong>eine englische Variante</strong> der Texte hinterlegen kannst — also Feld-Name, Beschreibung (i-Tooltip), Checkbox-Bestätigungs-Text und Dropdown-Optionen jeweils auf Deutsch UND auf Englisch. Default: <strong>aus</strong>.<br /><br />
                          <strong>Anzeige in der App:</strong> wenn aktiviert, blendet jede Feld-Karte einen zweiten Eingabe-Block für die EN-Variante ein. Teilnehmer mit App-Sprache <strong>Englisch</strong> bekommen automatisch die EN-Texte zu sehen. Wer als App-Sprache Deutsch eingestellt hat, sieht weiterhin die DE-Texte. Zusätzlich folgt das Standard-Anmelde-Formular (Platzhalter, Hinweis-Boxen, Sub-Event-Sektion) ab dann der <strong>App-Spracheinstellung des Teilnehmers</strong> statt der Mail-Sprache des Events.<br /><br />
                          <strong>Auswirkung für Teilnehmer:</strong> internationale Kolleg:innen, die kein Deutsch sprechen, sehen das komplette Anmelde-Formular sauber auf Englisch. Wer als Organizer keine EN-Variante einträgt, fällt im EN-Modus still auf den DE-Wert zurück — die App bricht also nichts kaputt, falls du nur einige Felder übersetzt.
                        </>
                      : <>
                          <strong>What this controls:</strong> whether, for each custom field you create below, you can store <strong>an English variant</strong> of the texts — i.e. field name, description (i-tooltip), checkbox confirmation text and dropdown options in both German AND English. Default: <strong>off</strong>.<br /><br />
                          <strong>Where you see it:</strong> when enabled, each field card shows a second input row for the EN variant. Attendees with app language set to <strong>English</strong> automatically see the EN texts. Attendees with German keep seeing the DE texts. In addition, the standard registration form chrome (placeholders, hint boxes, sub-event section) follows the <strong>attendee&apos;s app language</strong> instead of the event&apos;s email language.<br /><br />
                          <strong>For attendees:</strong> international colleagues who do not speak German see the whole registration form cleanly in English. If an organizer leaves the EN variant empty for some field, the app silently falls back to the DE value — nothing breaks if you only translate a subset of fields.
                        </>
                    } />
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? 'Default: aus — wenn aktiviert, kannst du pro Feld eine englische Variante hinterlegen.'
                        : 'Default: off — when enabled, each field gets a second input row for the English variant.'}
                    </span>
                  </span>
                </label>
              </div>

                {/* v7.20: Spalten-Header oberhalb der Feld-Karten — erklaert
                    auf einen Blick welche Spalte was bedeutet. Nur sichtbar
                    wenn es mindestens 1 Feld gibt, sonst overhead. */}
                {/* v11.1: alter Tabellen-Header (Nr / Typ / Frage) entfernt —
                    seit der Card-Restrukturierung in v10.25 sind die Spalten
                    in jeder Card selbsterklärend (Nummern-Bubble + Label-
                    Input mit Placeholder + grüner Typ-Selector + Pflicht-
                    Pill + Lösch-X). Der Header passte mit den alten Spalten-
                    Breiten nicht mehr und stiftete optisch Verwirrung. */}
                {customFields.map((field, idx) => {
                  const isExpanded = !!fieldExpandOverride[field.id];
                  return (
                  <div
                    key={field.id}
                    draggable
                    onDragStart={() => setDragFieldId(field.id)}
                    onDragOver={(e) => { e.preventDefault(); setDragOverFieldId(field.id); }}
                    onDragLeave={() => { if (dragOverFieldId === field.id) setDragOverFieldId(null); }}
                    onDrop={() => {
                      if (dragFieldId && dragFieldId !== field.id) {
                        const fromIdx = customFields.findIndex(f => f.id === dragFieldId);
                        const toIdx = customFields.findIndex(f => f.id === field.id);
                        if (fromIdx >= 0 && toIdx >= 0) {
                          const updated = [...customFields];
                          const [moved] = updated.splice(fromIdx, 1);
                          updated.splice(toIdx, 0, moved);
                          setCustomFields(updated);
                        }
                      }
                      setDragFieldId(null);
                      setDragOverFieldId(null);
                    }}
                    onDragEnd={() => { setDragFieldId(null); setDragOverFieldId(null); }}
                    style={{
                      opacity: dragFieldId === field.id ? 0.4 : 1,
                      borderTop: dragOverFieldId === field.id ? '3px solid var(--dex-green)' : undefined,
                      background: 'var(--dex-gray-50, #fafafa)',
                      borderRadius: 12,
                      padding: 16,
                      marginBottom: 12,
                      border: '1px solid var(--dex-gray-200)',
                    }}
                  >
                    {/* v10.25: Konsolidierter Header — Label-Input prominent
                        als Titel + Typ-Dropdown rechts daneben + Pflicht-Pill
                        + Lösch-X. Reorder-Pfeile nur im Reorder-Modus. Die
                        separate Art/Beschriftung-Box ist entfallen, weil Typ
                        und Label hier direkt sichtbar / editierbar sind. */}
                    <div style={{ display: 'flex', gap: 10, alignItems: 'center', marginBottom: 12, flexWrap: 'wrap' }}>
                      {reorderMode && (
                        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 1, padding: '0 4px' }}>
                          <button
                            type="button"
                            onClick={() => moveCustomField(field.id, 'up')}
                            disabled={idx === 0}
                            style={{ background: 'none', border: 'none', padding: 0, color: idx === 0 ? 'var(--dex-gray-300)' : 'var(--dex-gray-600)', cursor: idx === 0 ? 'default' : 'pointer', fontSize: '0.85rem', lineHeight: 1 }}
                            title={isDe ? 'Nach oben' : 'Move up'}
                          >▲</button>
                          <button
                            type="button"
                            onClick={() => moveCustomField(field.id, 'down')}
                            disabled={idx === customFields.length - 1}
                            style={{ background: 'none', border: 'none', padding: 0, color: idx === customFields.length - 1 ? 'var(--dex-gray-300)' : 'var(--dex-gray-600)', cursor: idx === customFields.length - 1 ? 'default' : 'pointer', fontSize: '0.85rem', lineHeight: 1 }}
                            title={isDe ? 'Nach unten' : 'Move down'}
                          >▼</button>
                        </div>
                      )}
                      <span style={{
                        flexShrink: 0, width: 26, height: 26, borderRadius: '50%',
                        background: 'var(--dex-green, #86bc25)', color: '#fff',
                        display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                        fontWeight: 700, fontSize: '0.78rem',
                      }}>{idx + 1}</span>
                      {/* v18.56: Textarea statt Input — lange Fragen brechen jetzt
                          um statt abgeschnitten zu werden. Auto-Höhe via ref
                          (height = scrollHeight). resize:none + overflow:hidden,
                          damit es wie ein wachsendes Eingabefeld wirkt. */}
                      <textarea
                        className="form-input"
                        value={field.label}
                        rows={1}
                        placeholder={isDe ? 'Feld-Name (z.B. T-Shirt Größe)' : 'Field name (e.g. T-shirt size)'}
                        onChange={e => updateCustomField(field.id, { label: e.target.value })}
                        ref={(el) => { if (el) { el.style.height = 'auto'; el.style.height = `${el.scrollHeight}px`; } }}
                        style={{
                          flex: '1 1 280px', minWidth: 180, maxWidth: 360,
                          fontSize: '1rem', fontWeight: 600,
                          padding: '8px 12px',
                          resize: 'none', overflow: 'hidden', lineHeight: 1.35,
                          fontFamily: 'inherit',
                          color: field.label ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
                        }}
                      />
                      <select
                        className="form-select"
                        value={field.type}
                        onChange={e => updateCustomField(field.id, { type: e.target.value as CustomFieldInput['type'] })}
                        title={isDe ? 'Art des Feldes' : 'Field type'}
                        style={{
                          /* v11.4: feste Breite, damit Label-Input + Typ-
                             Selector + Pflicht-Pill + X immer in einer Zeile
                             passen. Vorher: flex 0 0 auto = intrinsic Width
                             — bei langen Options wie 'Roommate (Doppelzimmer)'
                             wurde der Selector breiter und drückte das X in
                             eine zweite Zeile. */
                          flex: '0 0 200px', maxWidth: 200,
                          background: 'rgba(134,188,37,0.08)',
                          border: '1px solid var(--dex-green, #86bc25)',
                          color: 'var(--dex-green-dark, #4a7c1f)',
                          fontWeight: 600,
                          padding: '8px 10px',
                        }}
                      >
                        <option value="text">{isDe ? 'Text (Freitext)' : 'Text (free text)'}</option>
                        <option value="select">{isDe ? 'Dropdown' : 'Dropdown'}</option>
                        <option value="number">{isDe ? 'Zahl' : 'Number'}</option>
                        <option value="checkbox">{isDe ? 'Checkbox' : 'Checkbox'}</option>
                        <option value="user">{isDe ? 'Person' : 'Person'}</option>
                        <option value="roommate">{isDe ? 'Roommate' : 'Roommate'}</option>
                        <option value="document">{isDe ? 'Dokument (Upload)' : 'Document (upload)'}</option>
                      </select>
                      <label
                        style={{
                          display: 'inline-flex', alignItems: 'center', gap: 6,
                          padding: '6px 12px', borderRadius: 999,
                          fontSize: '0.78rem', fontWeight: 600, whiteSpace: 'nowrap',
                          cursor: 'pointer', userSelect: 'none',
                          border: `1px solid ${field.required ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                          background: field.required ? 'rgba(134,188,37,0.10)' : '#fff',
                          color: field.required ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                          transition: 'all 0.15s ease',
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={field.required}
                          onChange={e => updateCustomField(field.id, { required: e.target.checked })}
                          style={{ display: 'none' }}
                        />
                        <span style={{ fontSize: '0.85rem', lineHeight: 1 }}>{field.required ? '✓' : '○'}</span>
                        {t('create.required')}
                      </label>
                      <button
                        type="button"
                        onClick={() => toggleFieldExpand(field.id, isExpanded)}
                        title={isExpanded ? (isDe ? 'Details einklappen' : 'Collapse details') : (isDe ? 'Details bearbeiten' : 'Edit details')}
                        aria-expanded={isExpanded}
                        style={{ background: 'none', border: 'none', color: 'var(--dex-gray-500)', padding: 4, cursor: 'pointer', flexShrink: 0, marginLeft: 'auto', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                      >
                        {!isExpanded && (
                          <span style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                            {isDe ? 'Details' : 'Details'}
                          </span>
                        )}
                        <Icon iconName={isExpanded ? 'ChevronUp' : 'ChevronDown'} style={{ fontSize: 14 }} />
                      </button>
                      <button
                        onClick={() => removeCustomField(field.id)}
                        title={isDe ? 'Feld löschen' : 'Delete field'}
                        style={{ background: 'none', border: 'none', color: 'var(--dex-red)', padding: 4, cursor: 'pointer', flexShrink: 0 }}
                      >
                        <X size={18} />
                      </button>
                    </div>

                    {isExpanded && (<>
                    {/* v18.41: People-Picker-Feld → ausgewählte Person bei
                        An-/Abmelde-Mail auf CC (nur für user/roommate-Felder). */}
                    {(field.type === 'user' || field.type === 'roommate') && (
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginTop: 8, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={!!field.ccOnEmails}
                          onChange={e => updateCustomField(field.id, { ccOnEmails: e.target.checked })}
                          style={{ marginTop: 2 }}
                        />
                        <span style={{ flex: 1 }}>
                          <span style={{ fontSize: '0.85rem', fontWeight: 600 }}>
                            {isDe ? 'Ausgewählte Person bei An-/Abmelde-Mail auf CC setzen' : 'CC the selected person on registration/cancellation email'}
                          </span>
                          <InfoTooltip text={isDe ? (
                            <>
                              <strong>Was du hier einstellst:</strong> ob die in diesem Feld <strong>ausgewählte Person</strong> (z.&nbsp;B. die Assistenz) die <strong>Anmelde- und Abmelde-Mail</strong> des Teilnehmers <strong>in Kopie (CC)</strong> bekommt.<br /><br />
                              <strong>Anzeige in der App:</strong> ändert nichts an der Anzeige — wirkt nur beim Mail-Versand.<br /><br />
                              <strong>Automatismen:</strong> die im Feld gewählte Person wird automatisch auf CC der Bestätigungs- bzw. Abmelde-Mail gesetzt. <strong>Der Outlook-Termin ist davon nicht betroffen</strong> — die Person wird also NICHT in den Kalendereintrag eingeladen.<br /><br />
                              <strong>Auswirkung für Teilnehmer:</strong> seine Assistenz ist bei An- und Abmeldung automatisch informiert, ohne dass er sie manuell weiterleiten muss.
                            </>
                          ) : (
                            <>
                              <strong>What you set here:</strong> whether the <strong>person selected in this field</strong> (e.g. the assistant) receives the attendee&apos;s <strong>registration and cancellation email</strong> in <strong>CC</strong>.<br /><br />
                              <strong>Where you see it:</strong> no visible change — only affects email sending.<br /><br />
                              <strong>Automations:</strong> the chosen person is automatically added to CC of the confirmation / cancellation mail. <strong>The Outlook event is not affected</strong> — the person is NOT invited to the calendar entry.<br /><br />
                              <strong>For attendees:</strong> their assistant is automatically kept in the loop on registration and cancellation without manual forwarding.
                            </>
                          )} />
                          <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                            {isDe
                              ? 'Nur die Mails — der Outlook-Termin wird nicht an diese Person gesendet.'
                              : 'Emails only — the Outlook event is not sent to this person.'}
                          </span>
                        </span>
                      </label>
                    )}

                    {/* v17.20: EN-Feld-Name — sichtbar wenn der Bilingual-
                        Toggle oben aktiviert wurde. Sitzt direkt unter dem
                        DE-Feld-Namen, damit der Organizer beide Sprachen in
                        einer Linie liest. Flagge + Placeholder machen klar,
                        welche Sprache gemeint ist. */}
                    {bilingualFields && (
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        marginLeft: 64, marginBottom: 6,
                      }}>
                        <span style={{
                          flexShrink: 0, fontSize: '0.7rem',
                          padding: '2px 8px', borderRadius: 8,
                          background: 'rgba(0,90,156,0.10)',
                          color: '#005a9c', fontWeight: 700, letterSpacing: 0.5,
                        }}>EN</span>
                        <input
                          className="form-input"
                          value={field.labelEn || ''}
                          placeholder={isDe
                            ? 'Englischer Feld-Name (optional — leer = fällt auf den deutschen Text zurück)'
                            : 'English field name (optional — empty = falls back to the German text)'}
                          onChange={e => updateCustomField(field.id, { labelEn: e.target.value })}
                          style={{ flex: 1, fontSize: '0.88rem', padding: '6px 10px' }}
                        />
                      </div>
                    )}

                    {/* v10.24: Pro-Gruppe-Sichtbarkeit — nur sichtbar wenn die
                        Split-Capacity in Schritt 3 aktiv ist. Der Organizer
                        kann ein Feld auf Gruppe A oder Gruppe B beschränken
                        (Beispiel: Pflicht-Checkbox „Leistungsnachweis vorhanden"
                        nur für Durchstarter / Gruppe A). 'all' = beide
                        Gruppen sehen das Feld (Default). Bei Gruppe-A/B-only
                        wird das Feld in der RegistrationPage entsprechend nur
                        beim passenden Wunsch-Typ gerendert. */}
                    {useSplitCapacities && (() => {
                      const labelA = (splitLabelA || '').trim() || 'Durchstarter';
                      const labelB = (splitLabelB || '').trim() || 'Funstarter';
                      const current = field.onlyForGroup || 'all';
                      const pillStyle = (active: boolean): React.CSSProperties => ({
                        display: 'inline-flex', alignItems: 'center', gap: 6,
                        padding: '6px 12px', borderRadius: 999,
                        fontSize: '0.78rem', fontWeight: 600, whiteSpace: 'nowrap',
                        cursor: 'pointer', userSelect: 'none',
                        border: `1px solid ${active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                        background: active ? 'rgba(134,188,37,0.10)' : '#fff',
                        color: active ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                        transition: 'all 0.15s ease',
                      });
                      return (
                        <div style={{
                          marginTop: 10, padding: '12px 14px',
                          background: '#fff',
                          border: '1px solid var(--dex-gray-200)',
                          borderRadius: 8,
                        }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8, flexWrap: 'wrap' }}>
                            <label style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
                              {isDe ? 'Sichtbar für Teilnehmergruppe' : 'Visible for attendee group'}
                            </label>
                            <InfoTooltip text={isDe ? (
                              <>
                                <strong>Was du hier einstellst:</strong> ob dieses Feld bei der Anmeldung für <strong>alle Teilnehmer</strong> oder nur für eine der zwei Kapazitäts-Gruppen sichtbar ist.<br /><br />
                                <strong>Beispiel:</strong> bei einem Lauf-Event ist die Pflicht-Checkbox &bdquo;Leistungsnachweis vorhanden&ldquo; nur für die <strong>Durchstarter-Gruppe</strong> sinnvoll, nicht für Funstarter / Walker. Stelle das Feld dann auf <strong>Nur {labelA}</strong> — Funstarter sehen es gar nicht erst.<br /><br />
                                <strong>Auswirkung in der App:</strong> die Anmelde-Seite blendet das Feld dynamisch ein/aus, sobald der Teilnehmer eine der zwei Boxen wählt. Pflichtfeld-Validierung greift natürlich nur wenn das Feld auch sichtbar ist.<br /><br />
                                <strong>Vorraussetzung:</strong> in Schritt 3 (Kapazität &amp; Sichtbarkeit) muss der Toggle &bdquo;Geteilte Kapazität&ldquo; aktiv sein. Sonst gibt&apos;s keine Gruppen — dieser Selector ist dann ausgeblendet.
                              </>
                            ) : (
                              <>
                                <strong>What you set here:</strong> whether this field is visible to <strong>all attendees</strong> or only to one of the two capacity groups during registration.<br /><br />
                                <strong>Example:</strong> on a running event, a required checkbox &ldquo;Performance proof available&rdquo; only makes sense for the <strong>fast-runner group</strong>, not for fun-runners / walkers. Set the field to <strong>{labelA} only</strong> — fun-runners won&apos;t even see it.<br /><br />
                                <strong>Effect in the app:</strong> the registration page dynamically shows / hides the field as the attendee picks one of the two boxes. Required-field validation only fires when the field is actually visible.<br /><br />
                                <strong>Requirement:</strong> the &ldquo;Split capacity&rdquo; toggle in step 3 (Capacity &amp; Visibility) must be active. Otherwise there are no groups — this selector is then hidden.
                              </>
                            )} />
                          </div>
                          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                            {([
                              { v: 'all', text: isDe ? 'Beide Gruppen' : 'Both groups' },
                              { v: 'A', text: isDe ? `Nur ${labelA}` : `${labelA} only` },
                              { v: 'B', text: isDe ? `Nur ${labelB}` : `${labelB} only` },
                            ] as const).map(opt => (
                              <label key={opt.v} style={pillStyle(current === opt.v)}>
                                <input
                                  type="radio"
                                  name={`onlyForGroup-${field.id}`}
                                  checked={current === opt.v}
                                  onChange={() => updateCustomField(field.id, { onlyForGroup: opt.v as 'all' | 'A' | 'B' })}
                                  style={{ display: 'none' }}
                                />
                                <span style={{ fontSize: '0.85rem', lineHeight: 1 }}>{current === opt.v ? '●' : '○'}</span>
                                {opt.text}
                              </label>
                            ))}
                          </div>
                        </div>
                      );
                    })()}

                    {/* v7.20: Beschreibung pro Feld. v18.18: Darstellung
                        wählbar — „i"-Box neben dem Label ODER Erklär-Text
                        unter dem Label. */}
                    <div style={{ marginLeft: 32, marginTop: 10 }}>
                      <input
                        className="form-input"
                        placeholder={isDe
                          ? 'Beschreibung (optional)'
                          : 'Description (optional)'}
                        value={field.helpText || ''}
                        onChange={e => updateCustomField(field.id, { helpText: e.target.value })}
                        style={{ width: '100%', fontSize: '0.82rem', padding: '6px 10px' }}
                      />
                      {field.helpText && field.helpText.trim() && (
                        <div style={{ display: 'flex', gap: 16, marginTop: 6, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                          <span style={{ fontWeight: 600 }}>{isDe ? 'Anzeige:' : 'Display:'}</span>
                          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 5, cursor: 'pointer' }}>
                            <input
                              type="radio"
                              name={`helpStyle-${field.id}`}
                              checked={(field.helpTextStyle || 'tooltip') !== 'inline'}
                              onChange={() => updateCustomField(field.id, { helpTextStyle: 'tooltip' })}
                            />
                            {isDe ? '„i"-Info-Box (Hover)' : '„i" info box (hover)'}
                          </label>
                          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 5, cursor: 'pointer' }}>
                            <input
                              type="radio"
                              name={`helpStyle-${field.id}`}
                              checked={field.helpTextStyle === 'inline'}
                              onChange={() => updateCustomField(field.id, { helpTextStyle: 'inline' })}
                            />
                            {isDe ? 'Text unter dem Feld-Titel' : 'Text below the field title'}
                          </label>
                        </div>
                      )}
                    </div>
                    {/* v17.20: EN-Variante der Beschreibung. */}
                    {bilingualFields && (
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        marginLeft: 64, marginTop: 4,
                      }}>
                        <span style={{
                          flexShrink: 0, fontSize: '0.65rem',
                          padding: '1px 6px', borderRadius: 6,
                          background: 'rgba(0,90,156,0.10)',
                          color: '#005a9c', fontWeight: 700, letterSpacing: 0.5,
                        }}>EN</span>
                        <input
                          className="form-input"
                          value={field.helpTextEn || ''}
                          placeholder={isDe
                            ? 'Englische Beschreibung (optional)'
                            : 'English description (optional)'}
                          onChange={e => updateCustomField(field.id, { helpTextEn: e.target.value })}
                          style={{ flex: 1, fontSize: '0.78rem', padding: '5px 9px' }}
                        />
                      </div>
                    )}

                    {/* v11.94: Bei Checkbox-Feldern kann der Organizer den
                        Text neben der Checkbox individuell setzen — Default
                        ist „Ja, bestätigen" / „Yes, confirm". */}
                    {field.type === 'checkbox' && (
                      <div style={{ marginLeft: 32, marginTop: 6 }}>
                        <input
                          className="form-input"
                          placeholder={isDe
                            ? 'Text neben Checkbox (optional, Default: „Ja, bestätigen")'
                            : 'Text next to checkbox (optional, default: „Yes, confirm")'}
                          value={field.confirmLabel || ''}
                          onChange={e => updateCustomField(field.id, { confirmLabel: e.target.value })}
                          style={{ width: '100%', fontSize: '0.82rem', padding: '6px 10px' }}
                        />
                      </div>
                    )}
                    {/* v17.20: EN-Variante des Checkbox-Bestätigungstexts. */}
                    {field.type === 'checkbox' && bilingualFields && (
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        marginLeft: 64, marginTop: 4,
                      }}>
                        <span style={{
                          flexShrink: 0, fontSize: '0.65rem',
                          padding: '1px 6px', borderRadius: 6,
                          background: 'rgba(0,90,156,0.10)',
                          color: '#005a9c', fontWeight: 700, letterSpacing: 0.5,
                        }}>EN</span>
                        <input
                          className="form-input"
                          value={field.confirmLabelEn || ''}
                          placeholder={isDe
                            ? 'Englischer Text neben Checkbox (optional, Default: „Yes, confirm")'
                            : 'English text next to checkbox (optional, default: „Yes, confirm")'}
                          onChange={e => updateCustomField(field.id, { confirmLabelEn: e.target.value })}
                          style={{ flex: 1, fontSize: '0.78rem', padding: '5px 9px' }}
                        />
                      </div>
                    )}


                    {/* v10.23: Dropdown-Optionen als gelisteter Editor mit
                        eigener Box. Pro Option eine Zeile mit Nummer +
                        Eingabefeld + Minus-Button. Plus-Button ans Ende
                        zum Hinzufügen. Mehrfachauswahl-Toggle direkt in
                        diesem Block (Kontext: betrifft nur die Optionsliste).
                        Nur sichtbar wenn type === 'select'. */}
                    {field.type === 'select' && (
                      <div style={{
                        marginTop: 10, padding: '12px 14px',
                        background: '#fff',
                        border: '1px solid var(--dex-gray-200)',
                        borderRadius: 8,
                      }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8, flexWrap: 'wrap', gap: 8 }}>
                          <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-700)', fontWeight: 600 }}>
                            {isDe ? 'Optionen' : 'Options'}
                          </div>
                          <label
                            style={{
                              display: 'inline-flex', alignItems: 'center', gap: 6,
                              padding: '4px 10px', borderRadius: 999,
                              fontSize: '0.75rem', fontWeight: 600, whiteSpace: 'nowrap',
                              cursor: 'pointer', userSelect: 'none',
                              border: `1px solid ${field.multi ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                              background: field.multi ? 'rgba(134,188,37,0.10)' : '#fff',
                              color: field.multi ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                              transition: 'all 0.15s ease',
                            }}
                            title={isDe
                              ? 'Wenn aktiv, kann der Teilnehmer mehrere Optionen gleichzeitig auswählen (z.B. mehrere Allergien).'
                              : 'When enabled, attendees can select multiple options at the same time (e.g. multiple allergies).'}
                          >
                            <input
                              type="checkbox"
                              checked={!!field.multi}
                              onChange={e => updateCustomField(field.id, { multi: e.target.checked })}
                              style={{ display: 'none' }}
                            />
                            <span style={{ fontSize: '0.8rem', lineHeight: 1 }}>{field.multi ? '✓' : '○'}</span>
                            {isDe ? 'Mehrfachauswahl möglich' : 'Allow multiple selection'}
                          </label>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                          {(field.options || []).map((opt, optIdx) => (
                            <div key={optIdx} style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                <span style={{ flexShrink: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontWeight: 600, width: 24, textAlign: 'right' }}>
                                  {optIdx + 1}.
                                </span>
                                <input
                                  className="form-input"
                                  value={opt}
                                  placeholder={isDe ? `Option ${optIdx + 1}` : `Option ${optIdx + 1}`}
                                  onChange={e => {
                                    const opts = [...(field.options || [])];
                                    opts[optIdx] = e.target.value;
                                    updateCustomField(field.id, { options: opts });
                                  }}
                                  style={{ flex: 1, fontSize: '0.85rem', padding: '6px 10px' }}
                                />
                                <button
                                  type="button"
                                  onClick={() => {
                                    const opts = [...(field.options || [])];
                                    opts.splice(optIdx, 1);
                                    // v17.20: EN-Optionsliste positional mit-zuruecksetzen,
                                    // damit Index-Mapping konsistent bleibt.
                                    const optsEn = [...(field.optionsEn || [])];
                                    if (optsEn.length > optIdx) optsEn.splice(optIdx, 1);
                                    updateCustomField(field.id, { options: opts, optionsEn: optsEn });
                                  }}
                                  title={isDe ? 'Option entfernen' : 'Remove option'}
                                  style={{
                                    flexShrink: 0, width: 28, height: 28, borderRadius: 6,
                                    background: '#fff', border: '1px solid var(--dex-gray-300)',
                                    color: 'var(--dex-red, #c00)', cursor: 'pointer',
                                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                    fontSize: '1rem', lineHeight: 1, fontWeight: 700,
                                  }}
                                >−</button>
                              </div>
                              {/* v17.20: Positional gemappte EN-Option. */}
                              {bilingualFields && (
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginLeft: 32 }}>
                                  <span style={{
                                    flexShrink: 0, fontSize: '0.65rem',
                                    padding: '1px 6px', borderRadius: 6,
                                    background: 'rgba(0,90,156,0.10)',
                                    color: '#005a9c', fontWeight: 700, letterSpacing: 0.5,
                                  }}>EN</span>
                                  <input
                                    className="form-input"
                                    value={(field.optionsEn || [])[optIdx] || ''}
                                    placeholder={isDe ? 'Englische Variante (optional)' : 'English variant (optional)'}
                                    onChange={e => {
                                      const optsEn = [...(field.optionsEn || [])];
                                      while (optsEn.length <= optIdx) optsEn.push('');
                                      optsEn[optIdx] = e.target.value;
                                      updateCustomField(field.id, { optionsEn: optsEn });
                                    }}
                                    style={{ flex: 1, fontSize: '0.78rem', padding: '5px 9px' }}
                                  />
                                </div>
                              )}
                            </div>
                          ))}
                          <button
                            type="button"
                            onClick={() => updateCustomField(field.id, { options: [...(field.options || []), ''] })}
                            style={{
                              alignSelf: 'flex-start', marginTop: 4,
                              display: 'inline-flex', alignItems: 'center', gap: 6,
                              background: 'rgba(134,188,37,0.08)',
                              border: '1px dashed var(--dex-green, #86bc25)',
                              color: 'var(--dex-green-dark, #4a7c1f)',
                              borderRadius: 6, padding: '6px 12px',
                              fontSize: '0.8rem', fontWeight: 600, cursor: 'pointer',
                            }}
                          >
                            <span style={{ fontSize: '1rem', lineHeight: 1, fontWeight: 700 }}>+</span>
                            {isDe ? 'Option hinzufügen' : 'Add option'}
                          </button>
                        </div>
                      </div>
                    )}
                    {/* v10.23: Roommate-Erklärungs-Box — wird IMMER bei
                        type=roommate angezeigt (nicht mehr nur konditional bei
                        fehlendem Zimmerart-Feld). Erklärt dem Organizer, was
                        das Feld bei der Anmeldung tatsächlich tut: Personen-
                        Picker mit Mailbenachrichtigung an den ausgewählten
                        Roommate, Match-Erkennung im Admin Center wenn beide
                        sich gegenseitig wählen. Plus: Tipp wenn noch kein
                        Zimmerart-Feld da ist. */}
                    {field.type === 'roommate' && (() => {
                      const roomKeywords = ['einzelzimmer', 'doppelzimmer', 'single room', 'double room', 'zimmerart', 'room type'];
                      const hasRoomTypeField = customFields.some(other => {
                        if (other.id === field.id) return false;
                        const lbl = (other.label || '').toLowerCase();
                        const opts = (other.options || []).join(' ').toLowerCase();
                        return roomKeywords.some(k => lbl.indexOf(k) >= 0 || opts.indexOf(k) >= 0);
                      });
                      return (
                        <div style={{
                          marginTop: 10, padding: '12px 14px',
                          background: 'rgba(21,101,192,0.06)',
                          border: '1px solid rgba(21,101,192,0.4)',
                          borderRadius: 8, fontSize: '0.82rem', color: 'var(--dex-gray-700)',
                          lineHeight: 1.5,
                        }}>
                          <div style={{ fontWeight: 700, color: 'var(--dex-blue, #1565c0)', marginBottom: 6, display: 'flex', alignItems: 'center', gap: 6 }}>
                            <Icon iconName="Info" style={{ fontSize: 16 }} />
                            {isDe ? 'So funktioniert das Roommate-Feld' : 'How the roommate field works'}
                          </div>
                          {isDe ? (
                            <>
                              <p style={{ margin: '0 0 6px' }}>
                                Bei der Anmeldung sieht der Teilnehmer einen <strong>Personen-Suchfeld</strong> (Live-Suche im Deloitte-Tenant). Er tippt einen Namen ein, wählt die gewünschte Person als Zimmerpartner und schließt die Anmeldung ab.
                              </p>
                              <p style={{ margin: '0 0 6px' }}>
                                <strong>Was direkt passiert:</strong> die ausgewählte Person bekommt automatisch eine <strong>Roommate-Anfrage-Mail</strong> im Deloitte-Layout — mit Hinweis, dass <em>X den Wunsch geäußert hat, mit ihr/ihm das Zimmer zu teilen</em> und dem Link zur Event-Anmeldung. Der Empfänger kann sich dann selbst anmelden und seinerseits den Anfragenden als Wunsch-Roommate eintragen.
                              </p>
                              <p style={{ margin: '0 0 6px' }}>
                                <strong>Match-Erkennung im Admin Center:</strong> wenn zwei Teilnehmer sich <em>gegenseitig</em> als Wunsch-Roommate eingetragen haben, markiert das Admin Center das Paar als bestätigtes Match (grünes Häkchen). Einseitige Wünsche werden grau angezeigt — der Organizer kann dann selbst entscheiden, ob er die Person trotzdem zuteilt.
                              </p>
                              <p style={{ margin: 0 }}>
                                Sinnvoll <strong>kombiniert mit einem zweiten Feld vom Typ Dropdown</strong> namens &bdquo;Zimmerart&ldquo; mit Optionen &bdquo;Einzelzimmer / Doppelzimmer&ldquo;. Über die <strong>Sichtbarkeitsbedingung</strong> kannst du das Roommate-Feld dann ausblenden, wenn jemand &bdquo;Einzelzimmer&ldquo; wählt — sonst geben auch Einzelzimmer-Bucher einen Roommate an.
                              </p>
                              {!hasRoomTypeField && (
                                <p style={{ margin: '8px 0 0', padding: '8px 10px', background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 6, color: 'var(--dex-gray-700)' }}>
                                  <strong style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>Tipp:</strong> aktuell hast du noch kein Zimmerart-Feld angelegt. Ein Dropdown &bdquo;Zimmerart&ldquo; mit &bdquo;Einzelzimmer / Doppelzimmer&ldquo; wäre eine sinnvolle Ergänzung — sonst können auch Teilnehmer ohne Doppelzimmer-Wunsch einen Roommate angeben.
                                </p>
                              )}
                            </>
                          ) : (
                            <>
                              <p style={{ margin: '0 0 6px' }}>
                                On the registration page the attendee sees a <strong>person picker</strong> (live search of the Deloitte tenant). They type a name, pick their preferred roommate and submit the registration.
                              </p>
                              <p style={{ margin: '0 0 6px' }}>
                                <strong>What happens immediately:</strong> the selected person automatically receives a <strong>roommate-request email</strong> in the Deloitte layout — letting them know that <em>X requested to share a room with them</em> and including a link to the event registration. The recipient can then register and pick the requester back as their preferred roommate.
                              </p>
                              <p style={{ margin: '0 0 6px' }}>
                                <strong>Match detection in the admin center:</strong> when two attendees pick <em>each other</em> as preferred roommate, the admin center marks the pair as a confirmed match (green check). One-sided wishes are shown in grey — the organizer can still assign them manually if desired.
                              </p>
                              <p style={{ margin: 0 }}>
                                Best <strong>combined with a separate Dropdown field</strong> called &ldquo;Room type&rdquo; with options &ldquo;Single / Double&rdquo;. Use the <strong>visibility condition</strong> to hide the roommate field when someone picks &ldquo;Single&rdquo; — otherwise even single-room bookers will be asked to name a roommate.
                              </p>
                              {!hasRoomTypeField && (
                                <p style={{ margin: '8px 0 0', padding: '8px 10px', background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 6, color: 'var(--dex-gray-700)' }}>
                                  <strong style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>Tip:</strong> you don&apos;t have a room-type field yet. A dropdown &ldquo;Room type&rdquo; with &ldquo;Single / Double&rdquo; would be a useful addition — otherwise attendees without a double-room wish can still pick a roommate.
                                </p>
                              )}
                            </>
                          )}
                        </div>
                      );
                    })()}
                    {/* v7.21: Sichtbarkeitsbedingung — Feld nur anzeigen wenn
                        eine andere Frage einen bestimmten Wert hat. Quelle
                        kann nur ein Feld VOR diesem hier sein (idx < aktuell)
                        und muss vom Typ select oder checkbox sein. */}
                    {(() => {
                      const candidateSources = customFields.slice(0, idx).filter(other =>
                        (other.type === 'select' || other.type === 'checkbox') && (other.label || '').trim().length > 0
                      );
                      const sourceField = field.showIf
                        ? customFields.find(o => o.id === field.showIf!.fieldId)
                        : null;
                      const removeShowIf = (): void => {
                        // showIf gezielt loeschen: updateCustomField macht ein
                        // shallow-merge, also setzen wir undefined und filtern
                        // beim Save raus.
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        updateCustomField(field.id, { showIf: undefined as any });
                      };
                      return (
                        <div style={{ marginLeft: 32, marginTop: 10, padding: '10px 12px', background: 'rgba(21,101,192,0.04)', border: '1px dashed var(--dex-gray-300)', borderRadius: 8 }}>
                          {!field.showIf ? (
                            <span style={{ display: 'inline-flex', alignItems: 'center' }}>
                              <button
                                type="button"
                                onClick={() => {
                                  if (candidateSources.length === 0) {
                                    alert(isDe
                                      ? 'Es gibt noch kein Dropdown- oder Checkbox-Feld VOR diesem hier, an das die Sichtbarkeit geknüpft werden könnte. Lege zuerst ein passendes Feld weiter oben an.'
                                      : 'There is no dropdown or checkbox field BEFORE this one yet that visibility could depend on. Please add a suitable field above first.');
                                    return;
                                  }
                                  const first = candidateSources[0];
                                  updateCustomField(field.id, {
                                    showIf: {
                                      fieldId: first.id,
                                      values: first.type === 'checkbox' ? ['true'] : (first.options[0] ? [first.options[0]] : []),
                                    },
                                  });
                                }}
                                style={{
                                  background: 'none', border: 'none', padding: 0,
                                  color: 'var(--dex-green-dark, #4a7c1f)',
                                  fontSize: '0.8rem', fontWeight: 600, cursor: 'pointer',
                                }}
                              >
                                + {isDe ? 'Sichtbarkeitsbedingung hinzufügen' : 'Add visibility condition'}
                              </button>
                              <InfoTooltip
                                text={isDe
                                  ? 'Dieses Feld wird nur angezeigt, wenn die Antwort auf eine andere (zuvor angelegte) Frage einem von dir festgelegten Wert entspricht. Beispiel: „Roommate" wird nur gefragt, wenn die Frage „Zimmerart" mit „Doppelzimmer" beantwortet wurde. Andernfalls bleibt das Feld komplett verborgen — und blockiert auch nicht die Pflichtfeld-Validierung.'
                                  : 'This field is shown only when the answer to another (previously added) question matches a value you specify. Example: "Roommate" is only asked when the question "Room type" is answered with "Double room". Otherwise the field stays fully hidden — and does not block the required-field validation either.'}
                              />
                            </span>
                          ) : (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', fontWeight: 600, display: 'flex', alignItems: 'center' }}>
                                {isDe ? 'Diese Frage nur anzeigen wenn:' : 'Only show this question when:'}
                                <InfoTooltip
                                  text={isDe
                                    ? 'Dieses Feld wird nur angezeigt, wenn die Antwort auf die Quell-Frage einem der gewählten Werte entspricht. Bei Mehrfachauswahl-Quellen reicht ein Treffer. Pflichtfeld-Validierung wird übersprungen, solange das Feld verborgen ist.'
                                    : 'This field is shown only when the answer to the source question matches one of the chosen values. With multi-select sources a single match is enough. Required-field validation is skipped as long as the field stays hidden.'}
                                />
                              </div>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                                <select
                                  className="form-select"
                                  value={field.showIf.fieldId}
                                  onChange={e => {
                                    const newSrc = customFields.find(o => o.id === e.target.value);
                                    if (!newSrc) return;
                                    updateCustomField(field.id, {
                                      showIf: {
                                        fieldId: newSrc.id,
                                        values: newSrc.type === 'checkbox' ? ['true'] : (newSrc.options[0] ? [newSrc.options[0]] : []),
                                      },
                                    });
                                  }}
                                  style={{ fontSize: '0.82rem', padding: '4px 8px', minWidth: 180, maxWidth: 320 }}
                                >
                                  {candidateSources.map(o => (
                                    <option key={o.id} value={o.id}>
                                      {customFields.findIndex(c => c.id === o.id) + 1}. {o.label}
                                    </option>
                                  ))}
                                  {/* fallback wenn die ausgewaehlte Quelle hinter dem Feld gelandet
                                      ist (z.B. nach einem Move) — option in der Liste anzeigen,
                                      aber als ungueltig markiert lassen. */}
                                  {sourceField && !candidateSources.find(c => c.id === sourceField.id) && (
                                    <option value={sourceField.id} disabled>
                                      ⚠ {sourceField.label} ({isDe ? 'liegt hinter diesem Feld' : 'is positioned after this field'})
                                    </option>
                                  )}
                                </select>
                                <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                                  {isDe ? '=' : '='}
                                </span>
                                {sourceField && sourceField.type === 'checkbox' ? (
                                  <select
                                    className="form-select"
                                    value={field.showIf.values[0] || 'true'}
                                    onChange={e => updateCustomField(field.id, {
                                      showIf: { fieldId: field.showIf!.fieldId, values: [e.target.value] },
                                    })}
                                    style={{ fontSize: '0.82rem', padding: '4px 8px', minWidth: 130 }}
                                  >
                                    <option value="true">{isDe ? 'angehakt' : 'checked'}</option>
                                    <option value="false">{isDe ? 'nicht angehakt' : 'unchecked'}</option>
                                  </select>
                                ) : sourceField ? (
                                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                                    {(sourceField.options || []).filter(Boolean).map(opt => {
                                      const checked = field.showIf!.values.indexOf(opt) >= 0;
                                      return (
                                        <label
                                          key={opt}
                                          style={{
                                            display: 'inline-flex', alignItems: 'center', gap: 4,
                                            padding: '4px 10px', borderRadius: 999,
                                            fontSize: '0.78rem', cursor: 'pointer',
                                            border: `1px solid ${checked ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                            background: checked ? 'rgba(134,188,37,0.10)' : '#fff',
                                            color: checked ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                                            fontWeight: 600,
                                          }}
                                        >
                                          <input
                                            type="checkbox"
                                            checked={checked}
                                            onChange={() => {
                                              const next = checked
                                                ? field.showIf!.values.filter(v => v !== opt)
                                                : [...field.showIf!.values, opt];
                                              updateCustomField(field.id, {
                                                showIf: { fieldId: field.showIf!.fieldId, values: next },
                                              });
                                            }}
                                            style={{ display: 'none' }}
                                          />
                                          {checked ? '✓' : '○'} {opt}
                                        </label>
                                      );
                                    })}
                                  </div>
                                ) : null}
                                <button
                                  type="button"
                                  onClick={removeShowIf}
                                  title={isDe ? 'Bedingung entfernen' : 'Remove condition'}
                                  style={{
                                    background: 'none', border: 'none', cursor: 'pointer',
                                    color: 'var(--dex-red, #c00)', fontSize: '0.8rem',
                                    padding: '4px 6px', marginLeft: 'auto',
                                  }}
                                >
                                  ✕ {isDe ? 'entfernen' : 'remove'}
                                </button>
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })()}
                    </>)}
                  </div>
                  );
                })}
              </div>

              {/* === Bereich 2: Felder pro Sub-Event (v10.11+) ============
                  Jedes Sub-Event hat eine eigene Custom-Fields-Liste, die bei
                  der Anmeldung auf einem Sub-Event zusätzlich zu den Hauptevent-
                  Feldern gerendert wird.
                  v15.0: dieser Bereich wird per pro-Sub-Event-Tab oben
                  ersetzt — Block bleibt im Code (mit display:none), um die
                  Inline-Helper-Funktionen `addSubEventCustomField` /
                  `updateSubEventCustomField` etc. nicht entfernen zu muessen
                  und sicherzustellen, dass der Tab-N>0-Code identische
                  Verhaltens-Garantien hat. */}
              <div style={{ display: 'none', marginTop: 32, paddingTop: 24, borderTop: '2px solid var(--dex-gray-200)' }}>
                <h3 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.15rem', fontWeight: 700 }}>
                  {isDe ? 'Felder pro Sub-Event' : 'Fields per sub-event'}
                </h3>
                <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                  {isDe
                    ? 'Optional: pro Sub-Event eine eigene Auswahl-Frage (z.B. „Welche Strecke läufst du?" mit Optionen 5/10/Halbmarathon). Diese Felder erscheinen NUR wenn der Teilnehmer das jeweilige Sub-Event wählt — zusätzlich zu den Feldern des Hauptevents oben.'
                    : 'Optional: a per-sub-event question (e.g. „Which distance?" with options 5/10/Halfmarathon). These fields appear ONLY when an attendee picks that sub-event — in addition to the main-event fields above.'}
                </p>

                {subEvents.length === 0 ? (
                  <div style={{
                    padding: '16px 18px',
                    background: 'rgba(237,139,0,0.08)',
                    border: '1px dashed var(--dex-orange, #ed8b00)',
                    borderRadius: 8,
                    fontSize: '0.85rem',
                    color: 'var(--dex-gray-700)',
                  }}>
                    {isDe
                      ? 'Noch keine Sub-Events angelegt. Sub-Events legst du in Schritt 2 (Ort & Programm, ganz unten im Bereich „Sub-Events") an — danach kannst du hier pro Sub-Event eigene Anmelde-Felder definieren.'
                      : 'No sub-events yet. Add sub-events in Step 2 (Location & Programme, at the bottom in the „Sub-events" block) — then come back here to define per-sub-event registration fields.'}
                  </div>
                ) : (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
                    {subEvents.map(se => {
                      const seFields = se.customFields || [];
                      return (
                        <div
                          key={se.id}
                          style={{
                            border: '1px solid var(--dex-gray-200)',
                            borderRadius: 8,
                            padding: '14px 16px',
                            background: '#fafafa',
                          }}
                        >
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 8, marginBottom: 10 }}>
                            <strong style={{ fontSize: '0.95rem' }}>
                              {se.title || (isDe ? '(unbenanntes Sub-Event)' : '(unnamed sub-event)')}
                            </strong>
                            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                              <button
                                type="button"
                                className="btn btn-outline"
                                style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                                onClick={() => addSubEventCustomField(se.id)}
                              >
                                <Plus size={12} /> {isDe ? 'Feld hinzufügen' : 'Add field'}
                              </button>
                              {customFields.length > 0 && seFields.length === 0 && (
                                <button
                                  type="button"
                                  className="btn btn-outline"
                                  style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                                  onClick={() => copyParentFieldsToSubEvent(se.id)}
                                  title={isDe ? `Dupliziert die ${customFields.length} Felder vom Hauptevent als Startpunkt` : 'Duplicates the main-event fields as a starting point'}
                                >
                                  {isDe ? `Vom Hauptevent kopieren (${customFields.length})` : `Copy from main event (${customFields.length})`}
                                </button>
                              )}
                            </div>
                          </div>

                          {seFields.length === 0 ? (
                            <p style={{ margin: 0, fontSize: '0.8rem', color: 'var(--dex-gray-500)', fontStyle: 'italic' }}>
                              {isDe
                                ? 'Keine zusätzlichen Felder — Teilnehmer wählen nur das Sub-Event ohne weitere Frage.'
                                : 'No additional fields — attendees just pick this sub-event with no further question.'}
                            </p>
                          ) : (
                            // v11.96: Sub-Event-Felder-Layout = exakt gleicher
                            // Look wie die Hauptevent-Felder (Step 5 oben):
                            // numerierte grüne Badge + prominentes Label-Input
                            // + grüner Typ-Selector + Pflicht-Pill + Lösch-X +
                            // gleiches Helptext / Confirm-Label / Optionen-
                            // Layout. Vorher kompakte Mini-Variante mit kleinen
                            // Schriften — visuell inkonsistent zum Hauptevent.
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                              {seFields.map((field, idx) => (
                                <div
                                  key={field.id}
                                  style={{
                                    background: 'var(--dex-gray-50, #fafafa)',
                                    borderRadius: 12,
                                    padding: 16,
                                    border: '1px solid var(--dex-gray-200)',
                                  }}
                                >
                                  {/* Header: Badge + Label + Typ + Pflicht + X */}
                                  <div style={{ display: 'flex', gap: 10, alignItems: 'center', marginBottom: 12, flexWrap: 'wrap' }}>
                                    <span style={{
                                      flexShrink: 0, width: 26, height: 26, borderRadius: '50%',
                                      background: 'var(--dex-green, #86bc25)', color: '#fff',
                                      display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                      fontWeight: 700, fontSize: '0.78rem',
                                    }}>{idx + 1}</span>
                                    <input
                                      className="form-input"
                                      value={field.label}
                                      placeholder={isDe ? 'Frage / Feld-Label (z.B. „Welche Strecke?")' : 'Question / field label (e.g. „Which distance?")'}
                                      onChange={e => updateSubEventCustomField(se.id, field.id, { label: e.target.value })}
                                      style={{
                                        flex: '0 1 320px', minWidth: 180, maxWidth: 320,
                                        fontSize: '1rem', fontWeight: 600,
                                        padding: '8px 12px',
                                        color: field.label ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
                                      }}
                                    />
                                    <select
                                      className="form-select"
                                      value={field.type}
                                      onChange={e => updateSubEventCustomField(se.id, field.id, { type: e.target.value as CustomFieldInput['type'] })}
                                      title={isDe ? 'Art des Feldes' : 'Field type'}
                                      style={{
                                        flex: '0 0 200px', maxWidth: 200,
                                        background: 'rgba(134,188,37,0.08)',
                                        border: '1px solid var(--dex-green, #86bc25)',
                                        color: 'var(--dex-green-dark, #4a7c1f)',
                                        fontWeight: 600,
                                        padding: '8px 10px',
                                      }}
                                    >
                                      <option value="text">{isDe ? 'Text (Freitext)' : 'Text (free text)'}</option>
                                      <option value="select">{isDe ? 'Dropdown' : 'Dropdown'}</option>
                                      <option value="number">{isDe ? 'Zahl' : 'Number'}</option>
                                      <option value="checkbox">{isDe ? 'Checkbox' : 'Checkbox'}</option>
                                    </select>
                                    <label
                                      style={{
                                        display: 'inline-flex', alignItems: 'center', gap: 6,
                                        padding: '6px 12px', borderRadius: 999,
                                        fontSize: '0.78rem', fontWeight: 600, whiteSpace: 'nowrap',
                                        cursor: 'pointer', userSelect: 'none',
                                        border: `1px solid ${field.required ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                        background: field.required ? 'rgba(134,188,37,0.10)' : '#fff',
                                        color: field.required ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                                        transition: 'all 0.15s ease',
                                      }}
                                    >
                                      <input
                                        type="checkbox"
                                        checked={field.required}
                                        onChange={e => updateSubEventCustomField(se.id, field.id, { required: e.target.checked })}
                                        style={{ display: 'none' }}
                                      />
                                      <span style={{ fontSize: '0.85rem', lineHeight: 1 }}>{field.required ? '✓' : '○'}</span>
                                      {t('create.required')}
                                    </label>
                                    {field.type === 'select' && (
                                      <label
                                        style={{
                                          display: 'inline-flex', alignItems: 'center', gap: 6,
                                          padding: '6px 12px', borderRadius: 999,
                                          fontSize: '0.78rem', fontWeight: 600, whiteSpace: 'nowrap',
                                          cursor: 'pointer', userSelect: 'none',
                                          border: `1px solid ${field.multi ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                          background: field.multi ? 'rgba(134,188,37,0.10)' : '#fff',
                                          color: field.multi ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                                          transition: 'all 0.15s ease',
                                        }}
                                      >
                                        <input
                                          type="checkbox"
                                          checked={!!field.multi}
                                          onChange={e => updateSubEventCustomField(se.id, field.id, { multi: e.target.checked })}
                                          style={{ display: 'none' }}
                                        />
                                        <span style={{ fontSize: '0.85rem', lineHeight: 1 }}>{field.multi ? '✓' : '○'}</span>
                                        {isDe ? 'Mehrfachauswahl' : 'Multi-select'}
                                      </label>
                                    )}
                                    <button
                                      type="button"
                                      onClick={() => removeSubEventCustomField(se.id, field.id)}
                                      title={isDe ? 'Feld entfernen' : 'Remove field'}
                                      style={{ background: 'none', border: 'none', color: 'var(--dex-red)', padding: 4, cursor: 'pointer', flexShrink: 0 }}
                                    >
                                      <X size={18} />
                                    </button>
                                  </div>

                                  {/* Beschreibung pro Feld */}
                                  <div style={{ marginLeft: 32, marginTop: 10 }}>
                                    <input
                                      className="form-input"
                                      placeholder={isDe
                                        ? 'Beschreibung (optional, erscheint als „i"-Tooltip neben dem Feld)'
                                        : 'Description (optional, shown as „i" tooltip next to the field)'}
                                      value={field.helpText || ''}
                                      onChange={e => updateSubEventCustomField(se.id, field.id, { helpText: e.target.value })}
                                      style={{ width: '100%', fontSize: '0.82rem', padding: '6px 10px' }}
                                    />
                                  </div>

                                  {/* Checkbox-Confirm-Label */}
                                  {field.type === 'checkbox' && (
                                    <div style={{ marginLeft: 32, marginTop: 6 }}>
                                      <input
                                        className="form-input"
                                        placeholder={isDe
                                          ? 'Text neben Checkbox (optional, Default: „Ja, bestätigen")'
                                          : 'Text next to checkbox (optional, default: „Yes, confirm")'}
                                        value={field.confirmLabel || ''}
                                        onChange={e => updateSubEventCustomField(se.id, field.id, { confirmLabel: e.target.value })}
                                        style={{ width: '100%', fontSize: '0.82rem', padding: '6px 10px' }}
                                      />
                                    </div>
                                  )}

                                  {/* Optionen-Editor für Dropdown-Felder */}
                                  {field.type === 'select' && (
                                    <div style={{
                                      marginTop: 10, marginLeft: 32, padding: '12px 14px',
                                      background: '#fff',
                                      border: '1px solid var(--dex-gray-200)',
                                      borderRadius: 8,
                                    }}>
                                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-700)', fontWeight: 600, marginBottom: 8 }}>
                                        {isDe ? 'Antwort-Optionen' : 'Answer options'}
                                      </div>
                                      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                                        {field.options.map((opt, oidx) => (
                                          <div key={oidx} style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                                            <span style={{
                                              flexShrink: 0, width: 22, height: 22, borderRadius: '50%',
                                              background: 'var(--dex-gray-200)', color: 'var(--dex-gray-700)',
                                              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                              fontWeight: 700, fontSize: '0.72rem',
                                            }}>{oidx + 1}</span>
                                            <input
                                              className="form-input"
                                              placeholder={isDe ? `Option ${oidx + 1}` : `Option ${oidx + 1}`}
                                              value={opt}
                                              onChange={e => {
                                                const next = field.options.slice();
                                                next[oidx] = e.target.value;
                                                updateSubEventCustomField(se.id, field.id, { options: next });
                                              }}
                                              style={{ flex: 1, fontSize: '0.85rem', padding: '6px 10px' }}
                                            />
                                            {field.options.length > 1 && (
                                              <button
                                                type="button"
                                                onClick={() => {
                                                  const next = field.options.filter((_, i) => i !== oidx);
                                                  updateSubEventCustomField(se.id, field.id, { options: next });
                                                }}
                                                title={isDe ? 'Option entfernen' : 'Remove option'}
                                                style={{ background: 'none', border: 'none', color: 'var(--dex-gray-500)', padding: 4, cursor: 'pointer' }}
                                              >
                                                <X size={14} />
                                              </button>
                                            )}
                                          </div>
                                        ))}
                                        <button
                                          type="button"
                                          onClick={() => updateSubEventCustomField(se.id, field.id, { options: [...field.options, ''] })}
                                          style={{ alignSelf: 'flex-start', background: 'none', border: '1px dashed var(--dex-gray-300)', padding: '4px 12px', fontSize: '0.78rem', borderRadius: 6, cursor: 'pointer', color: 'var(--dex-gray-700)', display: 'inline-flex', alignItems: 'center', gap: 4, marginTop: 4 }}
                                        >
                                          <Plus size={12} /> {isDe ? 'Option hinzufügen' : 'Add option'}
                                        </button>
                                      </div>
                                    </div>
                                  )}
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              </div>{/* v16.5: close plain wrapper div (Step 5) — kein Greyout mehr */}
              </div>{/* v15.0: close activeFieldsTabIdx===0 wrapper (Top-Level Felder + hidden Bereich 2) */}

              {/* v18.75: Sicherheitshinweis vor dem Absenden — eigene Section
                  ganz unten in Schritt 5 (gilt event-weit, daher außerhalb der
                  Feld-Tabs). */}
              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '12px 16px', marginTop: 18, marginBottom: 4,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={confirmDialogEnabled}
                    onChange={e => setConfirmDialogEnabled(e.target.checked)}
                    style={{ marginTop: 3, cursor: 'pointer' }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Sicherheitshinweis vor dem Absenden anzeigen?' : 'Show a confirmation prompt before submitting?'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> ob nach dem Klick auf <strong>„Anmelden“</strong> noch ein <strong>Bestätigungs-Dialog</strong> erscheint, bevor die Anmeldung wirklich abgeschickt wird. Default: <strong>nein</strong>.<br /><br />
                          <strong>Zwei Varianten:</strong> die <strong>Auswahl-Übersicht</strong> listet Haupt-Event und gewählte Sub-Events auf — der Teilnehmer kann vor dem Absenden einzelne Punkte noch ab- oder zuwählen. Der <strong>eigene Hinweistext</strong> zeigt stattdessen einen frei formulierten Hinweis (z.B. zu Verbindlichkeit oder Storno-Fristen), den der Teilnehmer bestätigen muss.<br /><br />
                          <strong>Auswirkung für Teilnehmer:</strong> ein zusätzlicher, bewusster Bestätigungsschritt — schützt vor versehentlichen Anmeldungen.
                        </>
                      : <>
                          <strong>What this controls:</strong> whether a <strong>confirmation dialog</strong> appears after clicking <strong>“Register”</strong>, before the registration is actually submitted. Default: <strong>no</strong>.<br /><br />
                          <strong>Two variants:</strong> the <strong>selection summary</strong> lists the main event and selected sub-events — the attendee can de-/select items before submitting. The <strong>custom hint text</strong> instead shows a free-text note (e.g. about binding registration or cancellation deadlines) the attendee must acknowledge.<br /><br />
                          <strong>For attendees:</strong> an extra, deliberate confirmation step — protects against accidental registrations.
                        </>
                    } />
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? 'Default: nein — wenn aktiviert, muss der Teilnehmer nach „Anmelden" noch einen Dialog bestätigen.'
                        : 'Default: no — when enabled, the attendee has to confirm a dialog after clicking „Register".'}
                    </span>
                  </span>
                </label>
                {confirmDialogEnabled && (
                  <div style={{ marginTop: 12, paddingLeft: 30 }}>
                    <div style={{ fontSize: '0.82rem', fontWeight: 600, marginBottom: 8, color: 'var(--dex-gray-700)' }}>
                      {isDe ? 'Was soll der Dialog zeigen?' : 'What should the dialog show?'}
                    </div>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', marginBottom: 10 }}>
                      <input type="radio" name="confirmDialogMode" checked={confirmDialogMode !== 'freetext'} onChange={() => setConfirmDialogMode('summary')} style={{ marginTop: 3 }} />
                      <span style={{ fontSize: '0.85rem' }}>
                        <strong>{isDe ? 'Auswahl-Übersicht' : 'Selection summary'}</strong> — {isDe
                          ? 'listet Haupt-Event und gewählte Sub-Events auf; der Teilnehmer kann vor dem Absenden einzelne Punkte noch ab- oder zuwählen.'
                          : 'lists the main event and selected sub-events; the attendee can de-/select items before submitting.'}
                      </span>
                    </label>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer' }}>
                      <input type="radio" name="confirmDialogMode" checked={confirmDialogMode === 'freetext'} onChange={() => setConfirmDialogMode('freetext')} style={{ marginTop: 3 }} />
                      <span style={{ fontSize: '0.85rem' }}>
                        <strong>{isDe ? 'Eigener Hinweistext' : 'Custom hint text'}</strong> — {isDe
                          ? 'zeigt einen frei formulierten Hinweis, den der Teilnehmer bestätigen muss.'
                          : 'shows a free-text note the attendee must acknowledge.'}
                      </span>
                    </label>
                    {confirmDialogMode === 'freetext' && (
                      <textarea
                        className="form-input"
                        value={confirmDialogText}
                        onChange={e => setConfirmDialogText(e.target.value)}
                        rows={3}
                        placeholder={isDe
                          ? 'z.B. „Bitte beachte: Die Anmeldung ist verbindlich. Eine Stornierung ist nur bis 3 Tage vor dem Event möglich."'
                          : 'e.g. „Please note: registration is binding. Cancellation is only possible up to 3 days before the event."'}
                        style={{ marginTop: 10, width: '100%', resize: 'vertical' }}
                      />
                    )}
                  </div>
                )}
              </div>

              </div>{/* close Step 5 (Felder) — v15 index 4 */}

              {/* ===== Step 6 (v15: vormals Step 7): Kommunikation ===== */}
              <div style={{ display: currentStep === 5 ? 'block' : 'none' }}>
                <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                  {isDe ? 'Schritt 6 — Kommunikation' : 'Step 6 — Communication'}
                </h2>
                <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                  {isDe
                    ? 'Hier konfigurierst du alle automatischen E-Mails und Outlook-Einladungen — Sprache, Logos, Vorlagen und Versandregeln pro Aktion.'
                    : 'Here you configure all automated emails and Outlook invites — language, logos, templates and per-action send rules.'}
                </p>
                {renderStepIntro(
                  [
                    'Sprache der automatischen E-Mails wählen (Deutsch oder Englisch)',
                    'An- oder ausschalten, ob Teilnehmer überhaupt E-Mails und Outlook-Termine bekommen — z.B. um intern zu testen, ohne echte Mails zu verschicken',
                    'Festlegen, wann die Organizer eine Kopie der Anmelde-/Abmelde-Mails bekommen sollen (immer, nie oder erst kurz vorm Event)',
                    'Eigenes Bild für die E-Mails und für den Outlook-Termin hochladen — ersetzt das Standard-Logo',
                    'Den Text im Outlook-Termin individuell formulieren (mit Live-Vorschau)',
                    'Jede einzelne E-Mail (Anmelde-Bestätigung, Abmelde-Bestätigung, Wartelisten-Mail, Nachrück-Mail) mit eigenem Betreff und Text anpassen',
                  ],
                  [
                    'Pick the language for automated emails (German or English)',
                    'Switch on/off whether attendees receive emails and Outlook entries at all — e.g. for internal testing without sending real mails',
                    'Decide when organizers get a copy of the registration / cancellation emails (always, never, or only close to the event)',
                    'Upload a custom image for the emails and the Outlook entry — replaces the default logo',
                    'Phrase the text inside the Outlook entry yourself (with live preview)',
                    'Customise each individual email (registration, cancellation, waitlist, promotion) — own subject and body',
                  ]
                )}
                <h3 className="mb-16">{t('create.step.communication')}</h3>

                {/* v11.57: Tab-Leiste fuer Haupt-Event vs. Sub-Events. Jeder Sub-
                    Event kann eigene Mail-/Outlook-Einstellungen haben. Wenn keine
                    Sub-Events existieren, blenden wir die Tabs komplett aus. */}
                {subEvents.length > 0 && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 16 }}>
                    <div
                      role="tablist"
                      aria-label={isDe ? 'Event-Tab wechseln (Kommunikations-Einstellungen)' : 'Switch event tab (communication settings)'}
                      style={{
                        display: 'flex', flexWrap: 'wrap', gap: 6, flex: 1,
                        borderBottom: '1px solid var(--dex-gray-200)',
                        paddingBottom: 0,
                      }}
                    >
                      {[{ label: `${isDe ? 'Haupt-Event' : 'Main event'}: ${title || (isDe ? 'Ohne Titel' : 'Untitled')}`, isMain: true }, ...subEvents.map(s => ({ label: (shortSubEventTitle(s.title, title) || (isDe ? 'Sub-Event ohne Titel' : 'Untitled sub-event')).trim(), isMain: false }))].map((tab, tabIdx) => {
                        const active = tabIdx === activeCommTabIdx;
                        // v14.8: Haupt-Event-Tab visuell deaktivieren, wenn der
                        // „Nur Sub-Events"-Modus aktiv ist — kein Click, keine
                        // Sichtbarkeit als Ziel der Konfiguration.
                        const isDisabledMain = tab.isMain && subEventsOnlyMode;
                        return (
                          <button
                            key={tabIdx}
                            type="button"
                            role="tab"
                            aria-selected={active}
                            aria-disabled={isDisabledMain}
                            onClick={() => { if (!isDisabledMain) switchCommTab(tabIdx); }}
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
                              cursor: isDisabledMain ? 'not-allowed' : 'pointer',
                              opacity: isDisabledMain ? 0.4 : 1,
                              marginBottom: -1,
                              whiteSpace: 'nowrap',
                              maxWidth: 280,
                              overflow: 'hidden',
                              textOverflow: 'ellipsis',
                              transition: 'background 0.15s, color 0.15s, border-color 0.15s',
                            }}
                            title={isDisabledMain
                              ? (isDe ? 'Hauptevent-Kommunikation nicht relevant — „Nur Sub-Events"-Modus aktiv' : 'Main-event communication not relevant — „sub-events only" mode active')
                              : tab.label}
                          >
                            {tab.isMain && (
                              <span style={{ fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: 0.4, color: active ? 'var(--dex-green-dark)' : 'var(--dex-gray-400)' }}>
                                {isDe ? 'Haupt' : 'Main'}
                              </span>
                            )}
                            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>{tab.label}</span>
                            {isDisabledMain && (
                              <span style={{
                                fontSize: '0.65rem', fontWeight: 600,
                                padding: '2px 6px', borderRadius: 8,
                                background: 'var(--dex-gray-200, #e0e0e0)',
                                color: 'var(--dex-gray-600)',
                                marginLeft: 4,
                              }}>
                                {isDe ? 'nicht relevant — nur Sub-Events' : 'not relevant — sub-events only'}
                              </span>
                            )}
                          </button>
                        );
                      })}
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
                            Du hast in Schritt 3 (Sub-Events) den Modus <strong>&bdquo;Nur {(childTermPlural || 'Sub-Events').trim() || 'Sub-Events'}&ldquo;</strong> gewählt — Teilnehmer können sich gar nicht fürs Hauptevent anmelden, deshalb gibt es auch keine Bestätigungs-Mails und keinen Outlook-Termin fürs Hauptevent. Wechsle auf den Tab eines Sub-Events, um dort die Kommunikation zu konfigurieren.
                          </>
                        : <>
                            <strong>Main-event communication is not relevant in this mode.</strong><br />
                            You picked the <strong>&bdquo;{(childTermPlural || 'sub-events').trim() || 'sub-events'} only&ldquo;</strong> mode in step 3 (Sub-events) — attendees cannot register for the main event, so no confirmation emails or Outlook invites are sent for it. Switch to a sub-event tab to configure communication there.
                          </>}
                    </div>
                  </div>
                )}

                {!(subEventsOnlyMode && activeCommTabIdx === 0) && (
                <>
                <div className="form-group">
                  <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <StepBadge n={20} />
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

                {/* Benachrichtigungen abschalten — v9.39: collapsed by default.
                    v14.4: pre-open, wenn wir auf dem Haupt-Event-Tab sind und
                    Sub-Events existieren — der Organizer soll die Toggles
                    sehen können, um das Hauptevent stumm zu stellen während
                    Sub-Events einzeln kommunizieren. */}
                <details
                  className="form-group"
                  open={activeCommTabIdx === 0 && subEvents.length > 0 ? true : undefined}
                  style={{ marginTop: 24, padding: 16, background: 'var(--dex-gray-50, #f8f9fa)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-gray-200)' }}
                >
                  <summary style={{ cursor: 'pointer', listStyle: 'none', display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', fontWeight: 600 }}>
                    <StepBadge n={21} />
                    {t('create.notifications')}
                    <span style={{ marginLeft: 'auto', fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                      {(disableEmails || disableOutlook)
                        ? (isDe ? '⚠ Kommunikation deaktiviert' : '⚠ Communication disabled')
                        : (isDe ? 'Standard – empfohlen, klick zum Anpassen' : 'Default – recommended, click to adjust')}
                    </span>
                  </summary>
                  <div style={{ marginTop: 12 }}>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                    {t('create.notifications.hint')}
                  </p>
                  <label style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer', marginBottom: 8 }}>
                    <input
                      type="checkbox"
                      checked={!disableEmails}
                      onChange={e => setDisableEmails(!e.target.checked)}
                      style={{ width: 18, height: 18, cursor: 'pointer' }}
                    />
                    <span style={{ fontSize: '0.9rem' }}>
                      <strong>{t('create.notifications.email')}</strong>
                      <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
                        {t('create.notifications.email.desc')}
                      </span>
                    </span>
                  </label>
                  {/* v19.21/v19.22: granulare Sub-Schalter — einzeln die Anmelde-
                      bzw. Abmelde-Bestätigung abschalten. Ab v19.22 pro Tab
                      (Hauptevent UND Sub-Events), nur wenn E-Mails grundsätzlich
                      aktiv sind (Master an). Der gebundene State spiegelt je nach
                      aktivem Tab den Haupt- oder Sub-Event-Wert. */}
                  {!disableEmails && (
                    <div style={{ marginLeft: 24, paddingLeft: 12, borderLeft: '3px solid var(--dex-green, #86bc25)', display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 8 }}>
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={!disableRegistrationEmail}
                          onChange={e => setDisableRegistrationEmail(!e.target.checked)}
                          style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 2, flexShrink: 0 }}
                        />
                        <span style={{ fontSize: '0.9rem' }}>
                          <strong>{isDe ? 'Anmelde-Bestätigung schicken' : 'Send registration confirmation'}</strong>
                          <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-500)', lineHeight: 1.4, marginTop: 2 }}>
                            {isDe
                              ? 'Wenn aktiv: Teilnehmer bekommen bei der Anmeldung eine Bestätigungs-Mail (und, falls Warteliste aktiv, die Warteliste-Mail). Haken aus = es geht keine Anmelde-Bestätigung raus — die Abmelde-Mail bleibt davon unberührt.'
                              : 'When active: attendees receive a confirmation email on registration (plus the waitlist email if a waitlist is active). Unchecked = no registration confirmation is sent — the cancellation email is unaffected.'}
                          </span>
                        </span>
                      </label>
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={!disableCancellationEmail}
                          onChange={e => setDisableCancellationEmail(!e.target.checked)}
                          style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 2, flexShrink: 0 }}
                        />
                        <span style={{ fontSize: '0.9rem' }}>
                          <strong>{isDe ? 'Abmelde-Bestätigung schicken' : 'Send cancellation confirmation'}</strong>
                          <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-500)', lineHeight: 1.4, marginTop: 2 }}>
                            {isDe
                              ? 'Wenn aktiv: Teilnehmer bekommen bei einer Abmeldung eine Bestätigungs-Mail. Haken aus = es geht keine Abmelde-Bestätigung raus (z.B. wenn du Teilnehmer still abmeldest) — die Anmelde-Mail bleibt davon unberührt.'
                              : 'When active: attendees receive a confirmation email when cancelled. Unchecked = no cancellation confirmation is sent (e.g. when you remove attendees silently) — the registration email is unaffected.'}
                          </span>
                        </span>
                      </label>
                    </div>
                  )}
                  <label style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer' }}>
                    <input
                      type="checkbox"
                      checked={!disableOutlook}
                      onChange={e => setDisableOutlook(!e.target.checked)}
                      style={{ width: 18, height: 18, cursor: 'pointer' }}
                    />
                    <span style={{ fontSize: '0.9rem' }}>
                      <strong>{t('create.notifications.outlook')}</strong>
                      <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
                        {t('create.notifications.outlook.desc')}
                      </span>
                    </span>
                  </label>
                  {/* Nur im Edit-Modus: explizite Bestaetigung dass Outlook-Termin aktualisiert werden soll */}
                  {isEditMode && !disableOutlook && (
                    <label style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer', marginLeft: 24, paddingLeft: 12, borderLeft: '3px solid var(--dex-orange, #ed8b00)' }}>
                      <input
                        type="checkbox"
                        checked={triggerOutlookUpdate}
                        onChange={e => setTriggerOutlookUpdate(e.target.checked)}
                        style={{ width: 18, height: 18, cursor: 'pointer' }}
                      />
                      <span style={{ fontSize: '0.9rem' }}>
                        <strong>{t('create.notifications.triggerupdate')}</strong>
                        <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-500)', lineHeight: 1.4, marginTop: 2 }}>
                          {t('create.notifications.triggerupdate.desc')}
                        </span>
                      </span>
                    </label>
                  )}
                  {/* v15.3: Toggle „Anmeldung für mindestens ein Sub-Event
                      verpflichtend" wurde entfernt — der gleiche Effekt wird
                      jetzt komplett über den „Nur Sub-Events"-Modus in
                      Schritt 2 (Sub-Events) erzielt. Doppelte Konfiguration
                      an zwei Stellen war verwirrend. Die requireSubEventSelection-
                      State-Variable bleibt aus Backward-Compat erhalten (alte
                      Events haben sie ggf. als Piggyback gesetzt). */}
                  {/* v14.4: Acknowledgement-Pflicht bei deaktivierter
                      Hauptevent-Kommunikation + vorhandenen Sub-Events. */}
                  {activeCommTabIdx === 0 && subEvents.length > 0 && (disableEmails || disableOutlook) && (
                    <div style={{
                      marginTop: 16, padding: 14,
                      background: 'rgba(237,139,0,0.10)',
                      border: '2px solid var(--dex-orange, #ed8b00)',
                      borderRadius: 10,
                      display: 'flex', flexDirection: 'column', gap: 10,
                    }}>
                      <div style={{ fontSize: '0.88rem', fontWeight: 700, color: 'var(--dex-orange-dark, #b35a00)' }}>
                        {isDe ? '⚠ Hinweis: Kommunikation für das Hauptevent ist deaktiviert' : '⚠ Note: communication for the main event is disabled'}
                      </div>
                      <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
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
                        <span style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>
                          {isDe
                            ? 'Ja, mir ist bewusst, dass Teilnehmer sich für mindestens ein Sub-Event anmelden müssen, um Kommunikation zu erhalten.'
                            : 'Yes, I understand attendees need to register for at least one sub-event to receive communication.'}
                        </span>
                      </label>
                    </div>
                  )}
                  </div>
                </details>

                {/* v8.5: Organizer-BCC-Konfiguration (pro Event) — granular
                    fuer An- und Abmeldungen getrennt einstellbar. v9.39: collapsed by default. */}
                <details className="form-group" style={{ marginTop: 24, padding: 16, background: 'var(--dex-gray-50, #f8f9fa)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-gray-200)' }}>
                  <summary style={{ cursor: 'pointer', listStyle: 'none', display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', fontWeight: 600 }}>
                    <StepBadge n={22} />
                    {isDe ? 'Sollen die Organizer bei An- und Abmeldungen mitlesen?' : 'Should organizers be looped in on registrations / cancellations?'}
                    <span style={{ marginLeft: 'auto', fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                      {isDe ? 'Standard – empfohlen, klick zum Anpassen' : 'Default – recommended, click to adjust'}
                    </span>
                  </summary>
                  <div style={{ marginTop: 12 }}>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12, lineHeight: 1.5 }}>
                    {isDe
                      ? 'Wenn aktiv, bekommt der Organizer eine versteckte Kopie der Bestätigungs-Mail, die an den Teilnehmer rausgeht — der Teilnehmer sieht nicht, dass jemand mitliest. Praktisch um zu wissen, wer sich gerade an- oder abmeldet. Bei großen Events willst du das vielleicht nicht für jede einzelne Anmeldung — dann kannst du es hier gezielt einschränken (z.B. nur kurz vorm Event, wenn kurzfristige Änderungen wichtig sind).'
                      : 'When on, the organizer gets a hidden copy of the confirmation email sent to the attendee — the attendee doesn\'t see they were copied. Handy to know who is signing up or off. For large events you might not want this for every single sign-up — you can narrow it down here (e.g. only close to the event when last-minute changes matter).'}
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
                        ? '„Erst nach der letzten Abmeldemöglichkeit" nutzt das in Schritt 3 (Kapazität & Sichtbarkeit) gesetzte Datum „Letzte Abmeldemöglichkeit". Vor diesem Stichtag gelten Abmeldungen als unproblematisch — danach möchtest du als Organizer aber wissen, wer noch abspringt.'
                        : '„Only after the last cancellation date" uses the date set in step 3 (Capacity & Visibility) under „Last cancellation date". Cancellations before that are considered routine — after that, organizers usually want to know about late drop-outs.'}
                    </p>
                  </div>
                  </div>
                </details>

                {/* Custom-Logo fuer E-Mails — v9.39: collapsed by default. v9.40: gleiche graue Box wie 21/22. */}
                <details className="form-group" style={{ marginTop: 24, padding: 16, background: 'var(--dex-gray-50, #f8f9fa)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-gray-200)' }}>
                  <summary style={{ cursor: 'pointer', listStyle: 'none', display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', fontWeight: 600 }}>
                    <StepBadge n={23} />
                    {t('create.eventlogo.mail')}
                    <span style={{ marginLeft: 'auto', fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                      {isDe ? 'Standard – empfohlen, klick zum Anpassen' : 'Default – recommended, click to adjust'}
                    </span>
                  </summary>
                  <div style={{ marginTop: 12 }}>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 8 }}>
                    {t('create.eventlogo.mail.hint')}
                  </p>
                  {emailLogoPreview && (
                    <div style={{ marginBottom: 8, display: 'flex', alignItems: 'center', gap: 12 }}>
                      <img src={emailLogoPreview} alt="Event-Logo fuer Mails" style={{ maxWidth: 220, maxHeight: 140, borderRadius: 4 }} />
                      <button className="btn btn-secondary" style={{ fontSize: '0.7rem', padding: '2px 8px', color: 'var(--dex-red, #c00)' }}
                        onClick={() => setEmailLogoPreview('')}>{t('create.eventlogo.remove')}</button>
                    </div>
                  )}
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
                      // funktionieren nicht zuverlaessig in Mails (siehe
                      // EmailImageBase64-Pipeline). Empfehlung sind die
                      // offiziellen Deloitte Circular Motifs.
                      const ok = window.confirm(t('create.logoupload.warning'));
                      if (!ok) { e.target.value = ''; return; }
                      const compressed = await compressImage(file, 600, 0.9);
                      const reader = new FileReader();
                      reader.onload = (ev) => setEmailLogoPreview(ev.target?.result as string || '');
                      reader.readAsDataURL(compressed);
                    }} />
                  </label>
                  </div>
                </details>

                {/* Custom-Logo fuer Outlook-Termin — v9.39: collapsed by default. v9.40: gleiche graue Box. */}
                <details className="form-group" style={{ marginTop: 24, padding: 16, background: 'var(--dex-gray-50, #f8f9fa)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-gray-200)' }}>
                  <summary style={{ cursor: 'pointer', listStyle: 'none', display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', fontWeight: 600 }}>
                    <StepBadge n={24} />
                    {t('create.outlooklogo')}
                    <span style={{ marginLeft: 'auto', fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                      {isDe ? 'Standard – empfohlen, klick zum Anpassen' : 'Default – recommended, click to adjust'}
                    </span>
                  </summary>
                  <div style={{ marginTop: 12 }}>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 8 }}>
                    {t('create.outlooklogo.hint')}
                  </p>
                  {outlookLogoPreview && (
                    <div style={{ marginBottom: 8, display: 'flex', alignItems: 'center', gap: 12 }}>
                      <img src={outlookLogoPreview} alt="Event-Logo fuer Outlook" style={{ maxWidth: 220, maxHeight: 140, borderRadius: 4 }} />
                      <button className="btn btn-secondary" style={{ fontSize: '0.7rem', padding: '2px 8px', color: 'var(--dex-red, #c00)' }}
                        onClick={() => setOutlookLogoPreview('')}>{t('create.eventlogo.remove')}</button>
                    </div>
                  )}
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
                      const ok = window.confirm(t('create.logoupload.warning'));
                      if (!ok) { e.target.value = ''; return; }
                      const compressed = await compressImage(file, 600, 0.9);
                      const reader = new FileReader();
                      reader.onload = (ev) => setOutlookLogoPreview(ev.target?.result as string || '');
                      reader.readAsDataURL(compressed);
                    }} />
                  </label>
                  </div>
                </details>

                {/* v9.39: collapsed by default. v9.40: gleiche graue Box. */}
                <details className="form-group" style={{ marginTop: 24, padding: 16, background: 'var(--dex-gray-50, #f8f9fa)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-gray-200)' }}>
                  <summary style={{ cursor: 'pointer', listStyle: 'none', display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', fontWeight: 600 }}>
                    <StepBadge n={25} />
                    {t('create.outlookdesc')}
                    <span style={{ marginLeft: 'auto', fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                      {isDe ? 'Standard – empfohlen, klick zum Anpassen' : 'Default – recommended, click to adjust'}
                    </span>
                  </summary>
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
                  </div>
                </details>

                {/* v9.39: E-Mail-Texte-Block collapsed by default. v9.40: gleiche graue Box, gleiche Schriftgröße wie 21-25. */}
                <details className="form-group" style={{ marginTop: 24, padding: 16, background: 'var(--dex-gray-50, #f8f9fa)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-gray-200)' }}>
                  <summary style={{ cursor: 'pointer', listStyle: 'none', display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', fontWeight: 600 }}>
                    <StepBadge n={26} />
                    {t('create.templates.title')} ({emailLanguage})
                    <span style={{ marginLeft: 'auto', fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                      {isDe ? 'Standard – empfohlen, klick zum Anpassen' : 'Default – recommended, click to adjust'}
                    </span>
                  </summary>
                  <div style={{ marginTop: 12 }}>
                <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 12 }}>
                  {t('create.templates.hint')}
                </p>

                {/* TemplateType in DEX_EmailTemplates ist ASCII 'Nachruecken' (Umlaut nicht erlaubt in Choice-Feld).
                    v9.17: Warteliste/Nachruecken-Templates nur anzeigen, wenn das Event eine
                    Warteliste hat — sonst werden sie ohnehin nie genutzt. */}
                {['Anmeldung', 'Warteliste', 'Abmeldung', 'Nachruecken']
                  .filter(tType => {
                    // v9.28: Wartelisten-/Nachrueck-Templates nur zeigen, wenn das Event
                    // tatsaechlich eine Warteliste haben kann — also Warteliste aktiviert
                    // UND nicht unbegrenzt Teilnehmer (sonst gibt's nie eine volle Kapazitaet).
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
                      padding: 12, marginBottom: 12, background: override ? '#f0fdf4' : '#fff',
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
                      {/* Inline-Editor entfaellt — Edit oeffnet jetzt das HtmlEditorModal mit Live-Preview */}
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
                </details>
                </>
                )}{/* end !(subEventsOnlyMode && tab===0) wrapper, v14.8 */}

              </div>{/* close Step 7 (Kommunikation) */}

              {/* ===== Step 8 (v14.8: vormals Step 7): Dokumente ===== */}
              <div style={{ display: currentStep === 7 ? 'block' : 'none' }}>
                <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                  {isDe ? 'Schritt 8 — Dokumente' : 'Step 8 — Documents'}
                </h2>
                <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                  {isDe
                    ? 'Hier lädst du alle Unterlagen hoch, die deine Teilnehmer rund um das Event brauchen — von der Agenda bis zur Anfahrt.'
                    : 'Here you upload all documents attendees might need around the event — from the agenda to the travel directions.'}
                </p>
                {renderStepIntro(
                  [
                    'Programm / Agenda pflegen (mehrtägig möglich, Drag-Reihenfolge pro Tag)',
                    'Transferzeiten — Bus / Shuttle / Bahn von/zum Veranstaltungsort',
                    'Dokumente hochladen (PDF) — Teilnehmer sehen sie auf MyEvents als Inline-Vorschau oder Download',
                  ],
                  [
                    'Maintain the event programme / agenda (multi-day supported, drag-reorder per day)',
                    'Transfer times — bus / shuttle / train to and from the venue',
                    'Upload documents (PDF) — attendees see them on MyEvents as inline preview or download',
                  ]
                )}
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                  <StepBadge n={28} />
                  {isDe ? 'Dokumente hochladen' : 'Upload documents'}
                </label>
                {/* v9.28: Schlagwoerter fett rendern fuer bessere Lesbarkeit. */}
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', marginBottom: 16, lineHeight: 1.6 }}>
                  {isDe ? (
                    <>
                      Lade hier alle Unterlagen hoch, die deine Teilnehmer rund um das Event brauchen — z.B.
                      die <strong>Detail-Agenda als PDF</strong>, eine <strong>Anfahrtsbeschreibung mit Karte</strong>,
                      die <strong>Hausordnung</strong> des Veranstaltungsorts, eine <strong>Packliste</strong>, das <strong>Teilnehmer-Briefing</strong>
                      {' '}oder eine <strong>Datenschutz-/Foto-Einverständniserklärung</strong>. Die Dokumente erscheinen
                      automatisch unter <strong>{'„Meine Events“'}</strong> in der Detail-Ansicht des Teilnehmers — dort sehen sie eine
                      <strong> Inline-Vorschau</strong> (bei PDFs) und können das Dokument einzeln <strong>herunterladen</strong>.
                      Mehrere Dateien gleichzeitig hochladen geht per <strong>Drag &amp; Drop</strong> oder <strong>Mehrfachauswahl</strong>.
                      Du kannst Dokumente auch <strong>nach dem Event-Live-Gang</strong> noch hinzufügen oder austauschen — die
                      Teilnehmer sehen immer die aktuelle Version. <strong>Tipp:</strong> für Dokumente, die nur intern für die
                      Organizer wichtig sind (z.B. Kontaktliste vom Caterer), nutze eine geteilte
                      <strong> Teams-/SharePoint-Ablage außerhalb von DEX</strong>, da hier alle Teilnehmer Lese-Zugriff haben.
                    </>
                  ) : (
                    <>
                      Upload everything attendees might need around the event — e.g.
                      the <strong>detailed agenda as PDF</strong>, <strong>directions with a map</strong>,
                      the venue&apos;s <strong>house rules</strong>, a <strong>packing list</strong>, the <strong>attendee briefing</strong>
                      {' '}or a <strong>privacy/photo consent form</strong>. Documents show up automatically under
                      <strong>{' „My Events“'}</strong> in the attendee detail view — they get an
                      <strong> inline preview</strong> (for PDFs) and can <strong>download</strong> each file individually.
                      Multiple files can be uploaded at once via <strong>drag &amp; drop</strong> or <strong>multi-select</strong>.
                      You can keep adding or replacing documents <strong>after the event has gone live</strong> — attendees always
                      see the latest version. <strong>Tip:</strong> for documents only meant for organizers (e.g. caterer
                      contact list), use a shared <strong>Teams/SharePoint folder outside DEX</strong>, because every attendee has read access here.
                    </>
                  )}
                </p>

                {documents.map((doc, idx) => (
                  <div key={idx} style={{
                    display: 'flex', alignItems: 'center', gap: 8, padding: '8px 12px', marginBottom: 6,
                    background: 'var(--dex-gray-50, #fafafa)', borderRadius: 'var(--dex-radius)',
                    border: '1px solid var(--dex-gray-200)',
                  }}>
                    <Icon iconName="Page" style={{ fontSize: 16, color: 'var(--dex-gray-600)' }} />
                    <span style={{ flex: 1, fontSize: '0.85rem' }}>{doc.name}</span>
                    {doc.size > 0 && <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)' }}>{(doc.size / 1024).toFixed(0)} KB</span>}
                    <button type="button" onClick={() => setDocuments(documents.filter((_, i) => i !== idx))} style={{
                      background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)',
                      fontSize: '1.1rem', padding: '4px', lineHeight: 1,
                    }} title={t('general.delete')}>
                      <X size={16} />
                    </button>
                  </div>
                ))}

                <label className="btn btn-outline" style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  <Plus size={14} /> {t('create.documents.upload')}
                  <input
                    type="file"
                    multiple
                    style={{ display: 'none' }}
                    onChange={(e) => {
                      const files = e.target.files;
                      if (!files) return;
                      const newDocs = Array.from(files).map(f => ({ name: f.name, file: f, url: '', size: f.size }));
                      setDocuments([...documents, ...newDocs]);
                      e.target.value = '';
                    }}
                  />
                </label>

                {/* v11.0: Teilnehmer-Upload-Toggle. Default OFF — wird nur
                    bei expliziter Aktivierung in „Meine Events" als Upload-
                    Bereich für die Anmeldung sichtbar. Anzeigename und
                    Hinweistext sind frei konfigurierbar. */}
                <div style={{ marginTop: 32, paddingTop: 20, borderTop: '2px solid var(--dex-gray-100)' }}>
                  <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                    <StepBadge n={29} />
                    {isDe ? 'Teilnehmer-Upload erlauben' : 'Allow attendee upload'}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Was du hier einstellst:</strong> ob jeder Teilnehmer in &bdquo;Meine Events&ldquo; eine eigene Datei (z.B. PDF) zu seiner Anmeldung hochladen darf.<br /><br />
                        <strong>Beispiele:</strong> Reisekostenbeleg, unterschriebener Datenschutzbogen, Foto-Einverständnis, Zertifikat als Voraussetzung für die Teilnahme.<br /><br />
                        <strong>Ablauf:</strong> Teilnehmer sieht nach der Anmeldung in &bdquo;Meine Events&ldquo; einen Upload-Block mit deinem Anzeigenamen + Hinweistext. Hochgeladene Dateien werden direkt als <strong>Item-Attachment</strong> an die Teilnehmer-Zeile in der SharePoint-Subsite gehängt — nicht in einer Sammeldatei. Der Teilnehmer kann seine Datei jederzeit ersetzen oder löschen.<br /><br />
                        <strong>Admin-Sicht:</strong> du siehst im Admin-Center pro Teilnehmer-Zeile alle hochgeladenen Dateien als Liste mit Download-Link. Du kannst auch fremde Uploads löschen.<br /><br />
                        <strong>Default: aus.</strong> Nur einschalten, wenn du tatsächlich ein Dokument von Teilnehmern brauchst.
                      </>
                    ) : (
                      <>
                        <strong>What you set here:</strong> whether every attendee can upload a file (e.g. PDF) to their registration via &ldquo;My Events&rdquo;.<br /><br />
                        <strong>Examples:</strong> travel-expense receipt, signed privacy form, photo-consent, certificate as a prerequisite to attend.<br /><br />
                        <strong>Flow:</strong> after registering, the attendee sees an upload block in &ldquo;My Events&rdquo; with the display name and hint text you configure. Uploaded files attach directly as <strong>item attachments</strong> on the attendee&apos;s row in the SharePoint subsite — not into a collection file. Attendees can replace or delete their own file any time.<br /><br />
                        <strong>Admin view:</strong> you see every uploaded file per attendee in the admin center with a download link. You can also delete attendee uploads.<br /><br />
                        <strong>Default: off.</strong> Enable only when you actually need a document from attendees.
                      </>
                    )} />
                  </label>
                  <label
                    style={{
                      display: 'inline-flex', alignItems: 'center', gap: 8,
                      padding: '8px 14px', borderRadius: 999,
                      fontSize: '0.85rem', fontWeight: 600, whiteSpace: 'nowrap',
                      cursor: 'pointer', userSelect: 'none',
                      border: `1px solid ${allowAttendeeUpload ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                      background: allowAttendeeUpload ? 'rgba(134,188,37,0.10)' : '#fff',
                      color: allowAttendeeUpload ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                      transition: 'all 0.15s ease',
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={allowAttendeeUpload}
                      onChange={e => setAllowAttendeeUpload(e.target.checked)}
                      style={{ display: 'none' }}
                    />
                    <span style={{ fontSize: '0.9rem', lineHeight: 1 }}>{allowAttendeeUpload ? '✓' : '○'}</span>
                    {isDe ? 'Teilnehmer dürfen Datei hochladen' : 'Attendees may upload a file'}
                  </label>

                  {allowAttendeeUpload && (
                    <div style={{ marginTop: 14, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
                      <div>
                        <label className="form-label" style={{ fontSize: '0.82rem', marginBottom: 4 }}>
                          {isDe ? 'Anzeige-Name (z.B. „Reisekostenbeleg")' : 'Display name (e.g. „Travel-expense receipt")'}
                        </label>
                        <input
                          className="form-input"
                          value={attendeeUploadLabel}
                          onChange={e => setAttendeeUploadLabel(e.target.value)}
                          placeholder={isDe ? 'z.B. Reisekostenbeleg, Datenschutz-Erklärung' : 'e.g. Travel receipt, privacy form'}
                          maxLength={80}
                        />
                      </div>
                      <div>
                        <label className="form-label" style={{ fontSize: '0.82rem', marginBottom: 4 }}>
                          {isDe ? 'Hinweistext für Teilnehmer (optional)' : 'Hint text for attendees (optional)'}
                        </label>
                        <input
                          className="form-input"
                          value={attendeeUploadHint}
                          onChange={e => setAttendeeUploadHint(e.target.value)}
                          placeholder={isDe ? 'z.B. Bitte unterschrieben hochladen' : 'e.g. Please upload signed'}
                          maxLength={240}
                        />
                      </div>
                    </div>
                  )}
                </div>
              </div>{/* close Step 8 (Dokumente) */}

              {/* ===== Step 9 (v14.8: vormals Step 8): Fun-Zone ===== */}
              <div style={{ display: currentStep === 8 ? 'block' : 'none' }}>
                <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                  {isDe ? 'Schritt 9 — Fun-Zone' : 'Step 9 — Fun Zone'}
                </h2>
                <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                  {isDe
                    ? 'Optional: ein Quiz für die Teilnehmer — Multiple-Choice-Fragen mit Live-Highscore. Perfekt für Networking, Tagungs-Pausen oder Foto-Quiz.'
                    : 'Optional: a quiz for attendees — multiple-choice questions with live highscore. Perfect for networking, breaks at conferences, or photo quizzes.'}
                </p>
                {renderStepIntro(
                  [
                    'Quiz-Fragen für das Event anlegen — Multiple-Choice mit beliebig vielen Antwortoptionen',
                    'Pro Frage optional ein Bild hochladen (Logo, Foto-Quiz, etc.)',
                    'Mehrere richtige Antworten möglich (Mehrfachauswahl) — werden alle für volle Punktzahl gebraucht',
                    'Bereiche anlegen und Fragen per Drag & Drop zuordnen — alle Fragen eines Bereichs werden im Quiz zusammen auf einer Seite angezeigt',
                    'Live-Highscore + Statistik im Admin Center sehen (welche Fragen am häufigsten falsch beantwortet werden)',
                  ],
                  [
                    'Create quiz questions for the event — multiple choice with any number of answer options',
                    'Optionally upload an image per question (logo, photo quiz, etc.)',
                    'Multiple correct answers are supported — all of them must be picked for full points',
                    'Create sections and assign questions via drag & drop — all questions in a section are shown together on one page in the quiz',
                    'See live highscore + statistics in the admin center (which questions are most often answered incorrectly)',
                  ]
                )}
                <h3 className="mb-16">{t('create.step.funzone')}</h3>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 16 }}>
                  {t('create.funzone.hint')}
                </p>

                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                  <StepBadge n={29} />
                  {isDe ? 'Quiz-Bereiche' : 'Quiz sections'}
                </label>
                {/* Bereiche: Header + "+ Bereich"-Button. Fragen koennen per Drag&Drop
                    in Bereiche gezogen werden; jeder Bereich wird im Quiz zusammen
                    auf einer Seite angezeigt. */}
                <div style={{
                  display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12, flexWrap: 'wrap',
                }}>
                  <button
                    type="button"
                    className="btn btn-outline"
                    style={{ fontSize: '0.82rem', padding: '6px 14px' }}
                    onClick={() => {
                      setNewSectionName('');
                      setNewSectionError('');
                      setNewSectionModalOpen(true);
                    }}
                  >
                    <Plus size={14} /> Bereich
                  </button>
                  <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                    Fragen per Drag &amp; Drop in einen Bereich ziehen — alle Fragen eines Bereichs werden im Quiz zusammen angezeigt.
                  </span>
                </div>

                {(() => {
                  // Section-Reihenfolge: zuerst die in Fragen verwendeten (nach erster Erwaehnung),
                  // dann die noch leeren pendingSections.
                  const used: string[] = [];
                  for (const q of quiz) {
                    if (q.section && used.indexOf(q.section) < 0) used.push(q.section);
                  }
                  const allSections: string[] = [...used];
                  for (const p of pendingSections) {
                    if (allSections.indexOf(p) < 0) allSections.push(p);
                  }

                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const handleDrop = (ev: any, targetSection: string | undefined): void => {
                    ev.preventDefault();
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const id = ev.dataTransfer?.getData?.('text/plain') as string | undefined;
                    const qid = id || draggedQuestionId;
                    if (!qid) return;
                    updateQuizQuestion(qid, { section: targetSection });
                    setDraggedQuestionId(null);
                  };

                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const renderQuestionCard = (q: any, qi: number): React.ReactElement => (
                    <div
                      key={q.id}
                      draggable={true}
                      onDragStart={ev => {
                        setDraggedQuestionId(q.id);
                        try { ev.dataTransfer.setData('text/plain', q.id); } catch { /* some browsers restrict */ }
                        ev.dataTransfer.effectAllowed = 'move';
                      }}
                      onDragEnd={() => setDraggedQuestionId(null)}
                      style={{
                        padding: 16, marginBottom: 10, background: 'var(--dex-gray-50, #fafafa)',
                        borderRadius: 12, border: '1px solid var(--dex-gray-200)',
                        cursor: 'grab',
                        opacity: draggedQuestionId === q.id ? 0.5 : 1,
                      }}
                    >
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                        <label style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)', display: 'flex', alignItems: 'center', gap: 8 }}>
                          <span style={{ cursor: 'grab', color: 'var(--dex-gray-400)' }} title="Ziehen, um in einen Bereich zu verschieben">⋮⋮</span>
                          {t('create.funzone.question')} {qi + 1}
                        </label>
                        <button type="button" onClick={() => removeQuizQuestion(q.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red)', padding: 4 }}>
                          <X size={16} />
                        </button>
                      </div>
                      <input
                        className="form-input"
                        value={q.question}
                        onChange={e => updateQuizQuestion(q.id, { question: e.target.value })}
                        placeholder={t('create.funzone.questionplaceholder')}
                        style={{ marginBottom: 10 }}
                      />
                      <div style={{ marginBottom: 10, display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                        {q.imageBase64 ? (
                          <>
                            <img
                              src={q.imageBase64}
                              alt="Frage-Bild"
                              style={{ maxHeight: 80, maxWidth: 160, borderRadius: 8, border: '1px solid var(--dex-gray-200)' }}
                            />
                            <button
                              type="button"
                              onClick={() => updateQuizQuestion(q.id, { imageBase64: undefined })}
                              style={{
                                fontSize: '0.72rem', padding: '4px 10px',
                                border: '1px solid var(--dex-gray-300)', borderRadius: 6,
                                background: '#fff', color: 'var(--dex-red)', cursor: 'pointer',
                              }}
                            >
                              Bild entfernen
                            </button>
                          </>
                        ) : (
                          <label style={{
                            fontSize: '0.78rem', padding: '6px 12px',
                            border: '1px dashed var(--dex-gray-300)', borderRadius: 8,
                            cursor: 'pointer', color: 'var(--dex-gray-600)',
                            display: 'inline-flex', alignItems: 'center', gap: 6,
                          }}>
                            Bild hochladen (optional)
                            <input
                              type="file"
                              accept="image/*"
                              style={{ display: 'none' }}
                              onChange={async e => {
                                const file = e.target.files && e.target.files[0];
                                if (!file) return;
                                try {
                                  const dataUrl = await new Promise<string>((resolve, reject) => {
                                    const reader = new FileReader();
                                    reader.onload = () => resolve(String(reader.result || ''));
                                    reader.onerror = reject;
                                    reader.readAsDataURL(file);
                                  });
                                  const img = new Image();
                                  await new Promise<void>((resolve, reject) => {
                                    img.onload = () => resolve();
                                    img.onerror = reject;
                                    img.src = dataUrl;
                                  });
                                  const maxW = 800;
                                  const scale = img.width > maxW ? maxW / img.width : 1;
                                  const w = Math.round(img.width * scale);
                                  const h = Math.round(img.height * scale);
                                  const canvas = document.createElement('canvas');
                                  canvas.width = w;
                                  canvas.height = h;
                                  const ctx = canvas.getContext('2d');
                                  if (!ctx) return;
                                  ctx.drawImage(img, 0, 0, w, h);
                                  const compressed = canvas.toDataURL('image/jpeg', 0.8);
                                  updateQuizQuestion(q.id, { imageBase64: compressed });
                                } catch {
                                  alert('Bild konnte nicht verarbeitet werden.');
                                }
                                e.target.value = '';
                              }}
                            />
                          </label>
                        )}
                      </div>
                      <label style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginBottom: 4, display: 'block' }}>
                        {t('create.funzone.options')}
                      </label>
                      {q.options.map((opt: string, oi: number) => (
                        <div key={oi} style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 6 }}>
                          <input
                            type="checkbox"
                            checked={q.correctIndices?.includes(oi) || false}
                            onChange={() => {
                              const indices = q.correctIndices || [];
                              const newIndices = indices.includes(oi) ? indices.filter((x: number) => x !== oi) : [...indices, oi];
                              updateQuizQuestion(q.id, { correctIndices: newIndices.length > 0 ? newIndices : [0] });
                            }}
                            title={t('create.funzone.correct')}
                            style={{ accentColor: 'var(--dex-green)' }}
                          />
                          <input
                            className="form-input"
                            value={opt}
                            onChange={e => {
                              const newOpts = [...q.options];
                              newOpts[oi] = e.target.value;
                              updateQuizQuestion(q.id, { options: newOpts });
                            }}
                            placeholder={`${t('create.funzone.option')} ${oi + 1}`}
                            style={{ flex: 1, padding: '6px 10px', fontSize: '0.85rem' }}
                          />
                          {q.options.length > 2 && (
                            <button type="button" onClick={() => {
                              const newOpts = q.options.filter((_: string, i: number) => i !== oi);
                              const newCorrect = (q.correctIndices || []).filter((ci: number) => ci !== oi).map((ci: number) => ci > oi ? ci - 1 : ci);
                              updateQuizQuestion(q.id, { options: newOpts, correctIndices: newCorrect.length > 0 ? newCorrect : [0] });
                            }} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-400)', padding: 2 }}>
                              <X size={14} />
                            </button>
                          )}
                        </div>
                      ))}
                      <button type="button" onClick={() => updateQuizQuestion(q.id, { options: [...q.options, ''] })} style={{
                        fontSize: '0.78rem', padding: '4px 12px', border: '1px dashed var(--dex-gray-300)',
                        borderRadius: 8, background: 'none', color: 'var(--dex-green-dark)', cursor: 'pointer', marginTop: 4,
                      }}>
                        + {t('create.funzone.addoption')}
                      </button>
                      <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-400)', marginTop: 6 }}>
                        {t('create.funzone.correcthint')}
                      </div>
                    </div>
                  );

                  const unsortedQuiz = quiz.filter(q => !q.section);
                  const globalIndexOf = (qid: string): number => quiz.findIndex(x => x.id === qid);

                  return (
                    <>
                      {allSections.map(sec => {
                        const inSec = quiz.filter(q => q.section === sec);
                        return (
                          <div
                            key={`sec-${sec}`}
                            onDragOver={e => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; }}
                            onDrop={e => handleDrop(e, sec)}
                            style={{
                              padding: 12, marginBottom: 14, borderRadius: 12,
                              border: '2px dashed var(--dex-green)',
                              background: 'rgba(134,188,37,0.04)',
                            }}
                          >
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10, gap: 8, flexWrap: 'wrap' }}>
                              <h4 style={{ margin: 0, color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1rem' }}>
                                Bereich: {sec} <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>({inSec.length} {inSec.length === 1 ? 'Frage' : 'Fragen'})</span>
                              </h4>
                              <button
                                type="button"
                                onClick={() => {
                                  if (!window.confirm(isDe ? `Bereich "${sec}" entfernen? Die Fragen bleiben erhalten und landen in "Ohne Bereich".` : `Remove section "${sec}"? The questions are kept and move to "No section".`)) return;
                                  for (const qq of quiz) {
                                    if (qq.section === sec) updateQuizQuestion(qq.id, { section: undefined });
                                  }
                                  setPendingSections(prev => prev.filter(p => p !== sec));
                                }}
                                style={{
                                  fontSize: '0.72rem', padding: '4px 10px',
                                  border: '1px solid var(--dex-gray-300)', borderRadius: 6,
                                  background: '#fff', color: 'var(--dex-red)', cursor: 'pointer',
                                }}
                              >
                                Bereich entfernen
                              </button>
                            </div>
                            {inSec.length === 0 ? (
                              <div style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', fontSize: '0.82rem', padding: '12px 8px', textAlign: 'center' }}>
                                Fragen hierher ziehen
                              </div>
                            ) : (
                              inSec.map(q => renderQuestionCard(q, globalIndexOf(q.id)))
                            )}
                          </div>
                        );
                      })}

                      {/* Ohne Bereich */}
                      <div
                        onDragOver={e => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; }}
                        onDrop={e => handleDrop(e, undefined)}
                        style={allSections.length > 0 ? {
                          padding: 12, marginBottom: 14, borderRadius: 12,
                          border: '2px dashed var(--dex-gray-300)',
                          background: 'var(--dex-gray-50, #fafafa)',
                        } : undefined}
                      >
                        {allSections.length > 0 && (
                          <h4 style={{ margin: '0 0 10px', color: 'var(--dex-gray-600)', fontSize: '0.95rem' }}>
                            Ohne Bereich <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>({unsortedQuiz.length} {unsortedQuiz.length === 1 ? 'Frage' : 'Fragen'})</span>
                          </h4>
                        )}
                        {allSections.length > 0 && unsortedQuiz.length === 0 ? (
                          <div style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', fontSize: '0.82rem', padding: '8px', textAlign: 'center' }}>
                            (leer)
                          </div>
                        ) : (
                          unsortedQuiz.map(q => renderQuestionCard(q, globalIndexOf(q.id)))
                        )}
                      </div>
                    </>
                  );
                })()}

                <button type="button" className="btn btn-outline" onClick={addQuizQuestion} style={{ fontSize: '0.85rem', padding: '8px 20px' }}>
                  <Plus size={14} /> {t('create.funzone.addquestion')}
                </button>
              </div>{/* close Step 9 (Fun-Zone) */}

            </div>{/* close creation-form */}
          </div>{/* close card */}

          {/* Fortschrittsanzeige */}
          {isSubmitting && (
            <div className="mt-24" style={{ padding: '20px 0' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
                  {progressLabel}
                </span>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--dex-green)' }}>
                  {progress}%
                </span>
              </div>
              <div style={{
                width: '100%', height: 8, background: 'var(--dex-gray-200)',
                borderRadius: 4, overflow: 'hidden',
              }}>
                <div style={{
                  width: `${progress}%`, height: '100%',
                  background: progress === 100
                    ? 'var(--dex-green)'
                    : 'linear-gradient(90deg, var(--dex-green), #0076a8)',
                  borderRadius: 4,
                  transition: 'width 0.5s ease',
                }} />
              </div>
            </div>
          )}

          {!isSubmitting && (
            <>
            {/* v9.27: Action-Row — Zurück (links), Vorschau (mitte), Weiter (rechts)
                alle auf gleicher Höhe in einer Reihe. */}
            <div style={{
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              gap: 12, marginTop: 24, flexWrap: 'wrap',
            }}>
              <div style={{ display: 'flex', gap: 8 }}>
                {currentStep === 0 ? (
                  <button className="btn btn-danger" onClick={() => goBack()}><Trash2 size={16} /> {t('create.cancel')}</button>
                ) : (
                  <button className="btn btn-secondary" onClick={() => setCurrentStep(currentStep - 1)}>
                    {t('general.back')}
                  </button>
                )}
              </div>

              <button
                type="button"
                onClick={() => setShowRegisterPreview(true)}
                disabled={!title}
                style={{
                  display: 'inline-flex', alignItems: 'center', gap: 10,
                  padding: '12px 28px', borderRadius: 999,
                  border: '2px solid var(--dex-green)',
                  background: 'linear-gradient(135deg, rgba(134,188,37,0.12), rgba(0,118,168,0.08))',
                  color: 'var(--dex-green-dark, #4a7c1f)',
                  fontSize: '0.95rem', fontWeight: 700, letterSpacing: 0.3,
                  cursor: title ? 'pointer' : 'not-allowed',
                  opacity: title ? 1 : 0.5,
                  boxShadow: title ? '0 2px 8px rgba(134,188,37,0.25)' : 'none',
                  transition: 'transform 0.1s ease, box-shadow 0.1s ease',
                }}
                onMouseEnter={e => { if (title) (e.currentTarget as HTMLButtonElement).style.transform = 'translateY(-1px)'; }}
                onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.transform = 'translateY(0)'; }}
                title={title ? 'So sehen Teilnehmer die Registrierungsseite' : 'Event-Titel eingeben, um die Vorschau zu öffnen'}
              >
                👁 {t('create.registerpreview')}
              </button>

              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                {/* v17.5: Im Edit-Modus immer einen Speichern-Button anzeigen,
                    damit man nicht durch alle Steps klicken muss wenn man
                    nur eine Sache aendert. Button-Label ist klarer:
                    „Aenderungen speichern und zurueck zum Event" plus
                    Info-Tooltip mit Erklaerung. */}
                {isEditMode && currentStep < steps.length - 1 && (
                  <button
                    className="btn btn-primary"
                    disabled={!title}
                    onClick={attemptSubmit}
                    style={{ opacity: !title ? 0.5 : 1 }}
                  >
                    <Send size={16} /> {isDe ? 'Änderungen speichern und zurück zum Event' : 'Save changes and return to event'}
                  </button>
                )}

                {currentStep < steps.length - 1 ? (
                  <button
                    className="btn btn-primary"
                    onClick={() => {
                      setTriedNext(true);
                      if (canProceed()) {
                        setTriedNext(false);
                        setCurrentStep(currentStep + 1);
                      }
                    }}
                  >
                    {t('create.next')}
                  </button>
                ) : (
                  <button
                    className="btn btn-primary"
                    disabled={!title}
                    onClick={attemptSubmit}
                    style={{ opacity: !title ? 0.5 : 1 }}
                  >
                    <Send size={16} /> {isEditMode ? t('create.save') : t('create.submit')}
                  </button>
                )}
              </div>
            </div>
            </>
          )}
        </div>
      </div>

      {/* ===== Vollbild-Vorschau Modal ===== */}
      {showPreview && (
        <div className="preview-modal" style={{
          position: 'fixed', inset: 0, background: '#fff', zIndex: 1000,
          display: 'flex', flexDirection: 'column',
        }}>
          <div className="preview-modal-inner" style={{
            background: '#fff', borderRadius: 0, width: '100%', maxWidth: '100%',
            height: '100%', overflow: 'auto', padding: 0,
          }}>
            <div style={{
              padding: '16px 24px', borderBottom: '1px solid var(--dex-gray-200)',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              position: 'sticky', top: 0, background: '#fff', zIndex: 1, borderRadius: '16px 16px 0 0',
            }}>
              <div>
                <h3 style={{ margin: 0 }}>Vorschau: Registrierungsseite</h3>
                <p style={{ margin: '4px 0 0', fontSize: '0.8rem', color: 'var(--dex-gray-400)' }}>
                  Sektionen per Drag &amp; Drop verschieben
                </p>
              </div>
              <button onClick={() => setShowPreview(false)} style={{ background: 'none', border: 'none', fontSize: '1.5rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}>
                <X size={24} />
              </button>
            </div>

            <div style={{ padding: 24, display: 'flex', flexDirection: 'column', gap: 16 }}>
              {previewSections.map(section => (
                <div
                  key={section.id}
                  draggable
                  onDragStart={() => setDragSectionId(section.id)}
                  onDragOver={(e) => { e.preventDefault(); setDragOverSectionId(section.id); }}
                  onDragLeave={() => { if (dragOverSectionId === section.id) setDragOverSectionId(null); }}
                  onDrop={() => {
                    if (dragSectionId && dragSectionId !== section.id) {
                      const fromIdx = previewSections.findIndex(s => s.id === dragSectionId);
                      const toIdx = previewSections.findIndex(s => s.id === section.id);
                      if (fromIdx >= 0 && toIdx >= 0) {
                        const updated = [...previewSections];
                        const [moved] = updated.splice(fromIdx, 1);
                        updated.splice(toIdx, 0, moved);
                        setPreviewSections(updated);
                      }
                    }
                    setDragSectionId(null);
                    setDragOverSectionId(null);
                  }}
                  onDragEnd={() => { setDragSectionId(null); setDragOverSectionId(null); }}
                  style={{
                    opacity: dragSectionId === section.id ? 0.4 : 1,
                    borderTop: dragOverSectionId === section.id ? '3px solid var(--dex-green)' : undefined,
                    cursor: 'grab',
                    position: 'relative',
                  }}
                >
                  <div style={{
                    position: 'absolute', top: 4, right: 8, fontSize: '0.65rem',
                    color: 'var(--dex-gray-300)', fontWeight: 600, userSelect: 'none',
                  }}>
                    ⠿ verschieben
                  </div>
                  {renderPreviewSection(section.id)}
                </div>
              ))}
            </div>

            <div style={{
              padding: '16px 24px', borderTop: '1px solid var(--dex-gray-200)',
              display: 'flex', gap: 12, justifyContent: 'flex-end',
              position: 'sticky', bottom: 0, background: '#fff', borderRadius: '0 0 16px 16px',
            }}>
              <button className="btn btn-secondary" onClick={() => setShowPreview(false)}>
                Zurück zum Formular
              </button>
              <button
                className="btn btn-primary"
                disabled={!title}
                onClick={() => { setShowPreview(false); attemptSubmit(); }}
              >
                <Send size={16} /> {isEditMode ? t('create.save') : t('create.submit')}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* HTML-Editor-Modal mit Live-Preview (Outlook-Termin, E-Mail-Template oder Beschreibung).
          v9.39: Mode 'description' fuer die Event-Beschreibung — wird auf der Anmelde-Seite
          1:1 als HTML gerendert, deshalb hier auch ein Bearbeiten/Vorschau-Modal wie bei den
          Mail-Templates. */}
      {(() => {
        if (!htmlEditorOpen) return null;
        const isOutlook = htmlEditorMode === 'outlook';
        const isDescription = htmlEditorMode === 'description';
        const tType = htmlEditorTemplateType;
        const defaultTpl = (!isOutlook && !isDescription) ? emailTemplates.find(tp => tp.templateType === tType && tp.language === emailLanguage) : undefined;
        const override = (!isOutlook && !isDescription) ? emailTemplateOverrides[tType] : undefined;
        const currentSubject = override?.subject || defaultTpl?.subject || '';
        const currentHeading = override?.heading || defaultTpl?.heading || '';
        // v15.19: Subheading-Override pro Event. Falls override.subheading
        // explizit gesetzt ist (auch leerer String), nutze diesen Wert.
        const currentSubheading = override?.subheading !== undefined ? override.subheading : '';
        const currentBody = isOutlook
          ? outlookBody
          : isDescription
            ? description
            : (override?.bodyHtml || defaultTpl?.bodyHtml || '');
        // v18.19: Überschrift-Farbe + -Größe (Override > Template-Default).
        const currentHeadingColor = (override?.headingColor) || (defaultTpl?.headingColor) || '#86bc25';
        const currentHeadingFontSize = override?.headingFontSize || '26px';
        // v18.22: Überschrift fett/kursiv + Unter-Überschrift-Formatierung.
        const currentHeadingBold = override?.headingBold;
        const currentHeadingItalic = override?.headingItalic;
        const currentSubheadingColor = override?.subheadingColor || '#000000';
        const currentSubheadingFontSize = override?.subheadingFontSize || '20px';
        const currentSubheadingBold = override?.subheadingBold;
        const currentSubheadingItalic = override?.subheadingItalic;
        // v18.22: zentraler Patch-Helper — merged ein Teil-Update in den
        // Override des aktuellen TemplateTypes und BEWAHRT alle übrigen Felder
        // (vorher droppte z.B. ein Heading-Text-Edit die zuvor gesetzte Farbe).
        const patchOverride = (patch: Partial<EmailOverrideEntry>): void => {
          setEmailTemplateOverrides(prev => {
            const cur = prev[tType];
            return {
              ...prev,
              [tType]: {
                subject: cur?.subject ?? currentSubject,
                heading: cur?.heading ?? currentHeading,
                subheading: cur?.subheading !== undefined ? cur.subheading : currentSubheading,
                bodyHtml: cur?.bodyHtml ?? currentBody,
                ...(cur ? {
                  headingColor: cur.headingColor,
                  headingFontSize: cur.headingFontSize,
                  headingBold: cur.headingBold,
                  headingItalic: cur.headingItalic,
                  subheadingColor: cur.subheadingColor,
                  subheadingFontSize: cur.subheadingFontSize,
                  subheadingBold: cur.subheadingBold,
                  subheadingItalic: cur.subheadingItalic,
                } : {}),
                ...patch,
              },
            };
          });
        };
        // v18.42: read-only Termin/Ort-Labels für den Outlook-Editor — je nach
        // aktivem Tab (Hauptevent oder Sub-Event).
        const olActiveSub = activeCommTabIdx > 0 ? subEvents[activeCommTabIdx - 1] : undefined;
        const olStart = olActiveSub ? olActiveSub.startDate : startDate;
        const olEnd = olActiveSub ? olActiveSub.endDate : endDate;
        const olFmt = (d?: string): string => {
          if (!d) return '';
          try { return new Date(d).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); } catch { return ''; }
        };
        // v18.44: Auto-Ort (= „würde übernommen") als Platzhalter.
        const outlookLocationAuto = olActiveSub
          ? (buildOutlookLocation(olActiveSub.location, olActiveSub.locationAddress) || olActiveSub.location || '')
          : (buildOutlookLocation(location, { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity }));
        // v18.44: aktuelle Override-Werte des aktiven Tabs (leer = übernommen).
        const olLocationOverrideVal = olActiveSub ? (olActiveSub.outlookLocation || '') : outlookLocationOverride;
        const olStartOverrideVal = olActiveSub ? (olActiveSub.outlookStart || '') : outlookStartOverride;
        const olEndOverrideVal = olActiveSub ? (olActiveSub.outlookEnd || '') : outlookEndOverride;
        const setOlLocation = (v: string): void => {
          if (olActiveSub) { const fi = activeCommTabIdx - 1; setSubEvents(prev => prev.map((s, i) => i === fi ? { ...s, outlookLocation: v } : s)); }
          else setOutlookLocationOverride(v);
        };
        const setOlStart = (iso: string): void => {
          if (olActiveSub) { const fi = activeCommTabIdx - 1; setSubEvents(prev => prev.map((s, i) => i === fi ? { ...s, outlookStart: iso } : s)); }
          else setOutlookStartOverride(iso);
        };
        const setOlEnd = (iso: string): void => {
          if (olActiveSub) { const fi = activeCommTabIdx - 1; setSubEvents(prev => prev.map((s, i) => i === fi ? { ...s, outlookEnd: iso } : s)); }
          else setOutlookEndOverride(iso);
        };
        const pad2 = (n: number): string => String(n).padStart(2, '0');
        const olIsoToDate = (iso?: string): Date | null => {
          if (!iso) return null;
          const loc = isoToLocal(iso); if (!loc) return null;
          const [dp, tp] = loc.split('T'); const [y, mo, da] = dp.split('-').map(n => parseInt(n, 10)); const [h, mi] = (tp || '00:00').split(':').map(n => parseInt(n, 10));
          return new Date(y, mo - 1, da, h, mi, 0, 0);
        };
        const olDateToIso = (d: Date | null): string => {
          if (!d) return '';
          return berlinLocalToUtcIso(`${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}T${pad2(d.getHours())}:${pad2(d.getMinutes())}`);
        };
        const dpCommon = {
          showTimeSelect: true, timeFormat: 'HH:mm', timeIntervals: 15, timeCaption: 'Uhrzeit',
          dateFormat: 'dd.MM.yyyy, HH:mm', locale: 'de', className: 'form-input',
          wrapperClassName: 'dex-datepicker-wrapper', calendarClassName: 'dex-datepicker-calendar',
          popperPlacement: 'bottom-start' as const, isClearable: true, autoComplete: 'off',
        };
        const outlookDateEditor = (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
              <label style={{ fontSize: '0.68rem', color: 'var(--dex-gray-400)' }}>Start</label>
              <DatePicker {...dpCommon} selected={olIsoToDate(olStartOverrideVal)} onChange={(d: Date | null) => setOlStart(olDateToIso(d))} placeholderText={olStart ? olFmt(olStart) + ' (übernommen)' : 'Start'} />
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
              <label style={{ fontSize: '0.68rem', color: 'var(--dex-gray-400)' }}>Ende</label>
              <DatePicker {...dpCommon} selected={olIsoToDate(olEndOverrideVal)} onChange={(d: Date | null) => setOlEnd(olDateToIso(d))} placeholderText={olEnd ? olFmt(olEnd) + ' (übernommen)' : 'Ende'} />
            </div>
          </div>
        );
        // v18.46: Standard-Body-Vorlage (mit Platzhaltern) für „Standardtext laden"
        // im Outlook-Editor — Sprache folgt der aktiven Mail-Sprache.
        const outlookDefaultBody = (emailLanguage === 'EN')
          ? '<p>You are registered for the event <strong>{{EventTitle}}</strong>.</p>'
            + '<p>If you are unable to attend, please cancel your registration in time via the <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView" style="color:#86bc25;font-weight:600;">DEX App</a> („My Events").</p>'
            + '<p>For organizational questions please contact <strong>{{Organizer}}</strong>.</p>'
          : '<p>Du bist für das Event <strong>{{EventTitle}}</strong> angemeldet.</p>'
            + '<p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView" style="color:#86bc25;font-weight:600;">DEX App</a> („Meine Events") ab.</p>'
            + '<p>Bei organisatorischen Fragen wende dich bitte an <strong>{{Organizer}}</strong>.</p>';
        // v19.2: Einladender Beispieltext für die Beschreibung — über den
        // „Standardtext laden"-Button im Beschreibungs-Editor übernehmbar (statt
        // wie früher als Inline-Box im Wizard).
        const descriptionExampleHtml = isDe
          ? 'Liebe Kolleginnen und Kollegen,<br><br>wir freuen uns sehr, euch herzlich einzuladen! Es erwartet euch ein abwechslungsreiches Programm mit viel Raum für Austausch und Begegnung.<br><br>Wir freuen uns auf einen schönen gemeinsamen Tag mit euch!'
          : 'Dear colleagues,<br><br>we are delighted to invite you! Look forward to a varied programme with plenty of room for exchange and networking.<br><br>We look forward to seeing you there!';
        return (
          <HtmlEditorModal
            open={htmlEditorOpen}
            onClose={() => setHtmlEditorOpen(false)}
            defaultBodyHtml={isOutlook ? outlookDefaultBody : (isDescription ? descriptionExampleHtml : undefined)}
            title={isOutlook ? 'Outlook-Termin: Body bearbeiten' : isDescription ? (isDe ? 'Event-Beschreibung bearbeiten' : 'Edit event description') : `E-Mail-Template: ${tType}`}
            value={currentBody}
            onChange={(html) => {
              if (isOutlook) {
                setOutlookBody(html);
              } else if (isDescription) {
                setDescription(html);
              } else {
                // v18.22: patchOverride bewahrt alle übrigen Override-Felder
                // (Farbe/Größe/fett/kursiv von Über-/Unter-Überschrift).
                patchOverride({ bodyHtml: html });
              }
            }}
            previewMode={isDescription ? 'plain' : (isOutlook ? 'outlook' : 'email')}
            emailSubject={(!isOutlook && !isDescription) ? currentSubject : undefined}
            onEmailSubjectChange={(!isOutlook && !isDescription) ? (s) => patchOverride({ subject: s }) : undefined}
            emailHeading={(!isOutlook && !isDescription) ? currentHeading : undefined}
            onEmailHeadingChange={(!isOutlook && !isDescription) ? (h) => patchOverride({ heading: h }) : undefined}
            emailSubheading={(!isOutlook && !isDescription) ? currentSubheading : undefined}
            onEmailSubheadingChange={(!isOutlook && !isDescription) ? (s) => patchOverride({ subheading: s }) : undefined}
            emailHeadingColor={(!isOutlook && !isDescription) ? currentHeadingColor : undefined}
            emailHeadingFontSize={(!isOutlook && !isDescription) ? currentHeadingFontSize : undefined}
            onEmailHeadingColorChange={(!isOutlook && !isDescription) ? (hex) => patchOverride({ headingColor: hex }) : undefined}
            onEmailHeadingFontSizeChange={(!isOutlook && !isDescription) ? (px) => patchOverride({ headingFontSize: px }) : undefined}
            emailHeadingBold={(!isOutlook && !isDescription) ? currentHeadingBold : undefined}
            emailHeadingItalic={(!isOutlook && !isDescription) ? currentHeadingItalic : undefined}
            onEmailHeadingBoldChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ headingBold: b }) : undefined}
            onEmailHeadingItalicChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ headingItalic: b }) : undefined}
            emailSubheadingColor={(!isOutlook && !isDescription) ? currentSubheadingColor : undefined}
            emailSubheadingFontSize={(!isOutlook && !isDescription) ? currentSubheadingFontSize : undefined}
            emailSubheadingBold={(!isOutlook && !isDescription) ? currentSubheadingBold : undefined}
            emailSubheadingItalic={(!isOutlook && !isDescription) ? currentSubheadingItalic : undefined}
            onEmailSubheadingColorChange={(!isOutlook && !isDescription) ? (hex) => patchOverride({ subheadingColor: hex }) : undefined}
            onEmailSubheadingFontSizeChange={(!isOutlook && !isDescription) ? (px) => patchOverride({ subheadingFontSize: px }) : undefined}
            onEmailSubheadingBoldChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ subheadingBold: b }) : undefined}
            onEmailSubheadingItalicChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ subheadingItalic: b }) : undefined}
            imageWidth={!isDescription ? headerImageLayout.width : undefined}
            imagePaddingV={!isDescription ? headerImageLayout.paddingV : undefined}
            imagePaddingH={!isDescription ? headerImageLayout.paddingH : undefined}
            onImageWidthChange={!isDescription ? (w) => setHeaderImageLayout(p => ({ ...p, width: w })) : undefined}
            onImagePaddingVChange={!isDescription ? (v) => setHeaderImageLayout(p => ({ ...p, paddingV: v })) : undefined}
            onImagePaddingHChange={!isDescription ? (h) => setHeaderImageLayout(p => ({ ...p, paddingH: h })) : undefined}
            outlookHeading={isOutlook ? outlookHeading : undefined}
            onOutlookHeadingChange={isOutlook ? setOutlookHeading : undefined}
            outlookSubheading={isOutlook ? outlookSubheading : undefined}
            onOutlookSubheadingChange={isOutlook ? setOutlookSubheading : undefined}
            outlookSubject={isOutlook ? outlookSubject : undefined}
            onOutlookSubjectChange={isOutlook ? setOutlookSubject : undefined}
            outlookDateEditor={isOutlook ? outlookDateEditor : undefined}
            outlookLocationValue={isOutlook ? olLocationOverrideVal : undefined}
            onOutlookLocationChange={isOutlook ? setOlLocation : undefined}
            outlookLocationAuto={isOutlook ? outlookLocationAuto : undefined}
            previewVars={{
              // v17.5: Im Sub-Event-Kommunikations-Tab den Titel des
              // aktiven Sub-Events einsetzen, sonst den Hauptevent-Titel.
              EventTitle: (() => {
                if (activeCommTabIdx > 0) {
                  const sub = subEvents[activeCommTabIdx - 1];
                  return (sub && sub.title && sub.title.trim()) || title || 'Event Title';
                }
                return title || 'Event Title';
              })(),
              Name: 'Max Mustermann',
              Organizer: organizer || 'Organisator',
              AppUrl: 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView',
              WaitlistPosition: '1',
              Address: [addrStreet, addrHouseNo].filter(Boolean).join(' ') + ((addrZip || addrCity) ? ', ' + [addrZip, addrCity].filter(Boolean).join(' ') : ''),
              Location: location || 'Veranstaltungsort',
              StartDate: startDate ? new Date(startDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
              EndDate: endDate ? new Date(endDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
              EventDate: startDate ? new Date(startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
            }}
            insertableVars={isOutlook ? [
              // v17.16: {{Name}} hier ENTFERNT — der Outlook-Termin geht
              // an alle Teilnehmer gleichzeitig, eine pro-Person-Anrede
              // ist nicht moeglich. Vorher konnte der Organizer {{Name}}
              // einfuegen, was bei allen Empfaengern als unaufgeloester
              // Platzhalter „{{Name}}" stehen blieb.
              { key: '{{EventTitle}}', label: 'Event' },
              { key: '{{Organizer}}', label: 'Organizer' },
              { key: '{{Location}}', label: 'Ort' },
              { key: '{{Address}}', label: 'Adresse' },
              { key: '{{StartDate}}', label: 'Start' },
              { key: '{{EndDate}}', label: 'Ende' },
              { key: '{{AppUrl}}', label: 'App Link' },
            ] : [
              { key: '{{Name}}', label: 'Name' },
              { key: '{{EventTitle}}', label: 'Event' },
              { key: '{{Organizer}}', label: 'Organizer' },
              { key: '{{AppUrl}}', label: 'App Link' },
              { key: '{{WaitlistPosition}}', label: 'Waitlist #' },
            ]}
            imageBase64={(isOutlook ? outlookLogoPreview : emailLogoPreview) || ''}
          />
        );
      })()}

      {/* Register-Page-Preview-Modal (zeigt, was Teilnehmer sehen wuerden) */}
      <RegisterPreviewModal
        open={showRegisterPreview}
        onClose={() => setShowRegisterPreview(false)}
        data={{
          title,
          description,
          location,
          locationAddress: { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity },
          startDate,
          endDate,
          imagePreview,
          organizers: organizer.split(';').map(s => s.trim()).filter(Boolean),
          organizerEmails,
          maxParticipants: Number(maxParticipants) || 0,
          unlimitedParticipants,
          customFields: customFields.map(f => ({
            id: f.id,
            label: f.label,
            type: f.type,
            required: f.required,
            visible: f.visible !== false,
            options: f.type === 'select' ? f.options : undefined,
            // v7.24: helpText, multi und showIf an die Live-Preview weiterreichen,
            // damit die echte RegistrationPage genau das rendert was der
            // Teilnehmer spaeter sieht (i-Tooltip, Multi-Select-Liste,
            // Sichtbarkeitsbedingung).
            helpText: f.helpText,
            helpTextStyle: f.helpTextStyle,
            multi: f.multi,
            showIf: f.showIf,
            // v17.20: EN-Varianten an die Preview weiterreichen — sonst sieht
            // der Organizer in der Vorschau nicht, was englische Teilnehmer
            // bekommen wuerden.
            confirmLabel: f.confirmLabel,
            labelEn: f.labelEn,
            helpTextEn: f.helpTextEn,
            confirmLabelEn: f.confirmLabelEn,
            optionsEn: f.optionsEn,
          })),
          isFictive,
          // v14.10: Sub-Events + Sub-Only-Mode + Bezeichnungs-Term an die
          // Vorschau weiterreichen, damit der Organizer auch die Sub-Event-
          // Auswahl im Anmeldeformular sieht (vorher fehlte sie komplett).
          subEvents: subEvents.map(s => ({
            id: s.id,
            title: s.title,
            location: s.location,
            startDate: s.startDate,
            endDate: s.endDate,
            maxParticipants: s.maxParticipants,
            description: s.description,
            customFields: (s.customFields || []).map(f => ({
              id: f.id,
              label: f.label,
              type: f.type,
              required: f.required,
              visible: f.visible !== false,
              options: f.type === 'select' ? f.options : undefined,
              helpText: f.helpText,
              helpTextStyle: f.helpTextStyle,
              multi: f.multi,
              showIf: f.showIf,
            })),
          })),
          subEventsOnlyMode,
          requireSubEventSelection: requireSubEventSelection || subEventsOnlyMode,
          childEventTermSingular: childTermSingular,
          childEventTermPlural: childTermPlural,
          // v17.22: Bilingual-Flag an die Vorschau — sonst rendert die
          // Preview die EN-Varianten nie (useEnVariants prueft event.bilingualFields).
          bilingualFields,
        }}
      />

      {/* Massenimport-Modale — eine generische Komponente, vier Aufruf-Stellen.
          Audience speichert eine `,`-separierte Email-Liste in einem String,
          Teams speichern parallele Names[] + Emails[]-Arrays. Die onAdd-Callbacks
          übersetzen jeweils zwischen Modal-Output (Email + DisplayName) und der
          jeweiligen State-Form. */}
      <BulkUserImportModal
        open={bulkAudienceOpen}
        onClose={() => setBulkAudienceOpen(false)}
        title="Massenimport — Sichtbarkeit"
        description={(
          <p style={{ marginTop: 0 }}>
            Trag Personen, Verteilergruppen oder Email-Adressen direkt in den
            <strong> Sichtbarkeits-Filter</strong> ein. Externe (kein
            <code style={{ margin: '0 4px' }}>@deloitte.de</code>) werden zwar
            geschrieben, sehen das Event aber NICHT — die Plattform ist DEALL-only.
          </p>
        )}
        existingEmails={audience.split(',').map(s => s.trim()).filter(Boolean)}
        searchUsers={searchUsers}
        onAdd={({ email }) => {
          // Audience ist `,`-separierte String-Liste — Email anhängen wenn
          // noch nicht drin (Doppel-Check zur Sicherheit, das Modal filtert
          // schon vor).
          setAudience(prev => {
            const list = (prev || '').split(',').map(s => s.trim()).filter(Boolean);
            if (list.indexOf(email) < 0) list.push(email);
            return list.join(', ');
          });
        }}
      />
      <BulkUserImportModal
        open={bulkOrganizerOpen}
        onClose={() => setBulkOrganizerOpen(false)}
        title="Massenimport — Co-Organizer"
        description={(
          <p style={{ marginTop: 0 }}>
            Mehrere <strong>Co-Organizer</strong> auf einmal hinzufügen. Reihenfolge
            spielt eine Rolle — der erste Eintrag in der Liste bleibt der Haupt-Organizer.
            Massenimport hängt neue Personen <strong>hinten</strong> an.
          </p>
        )}
        existingEmails={organizerEmails}
        searchUsers={searchUsers}
        onAdd={({ email, displayName }) => {
          // WICHTIG: functional setState für `organizer`-String, sonst sehen
          // schnelle Sequenz-Calls (Massenimport mit 10+ Namen) alle dieselbe
          // closure-stale Version und nur der letzte Name landet, während
          // organizerEmails über `prev => ...` korrekt akkumuliert. Das führte
          // zu out-of-sync orgNames/orgEmails-Arrays mit falscher Namen-Email-
          // Zuordnung im Duplikat-Hinweis.
          setOrganizer(prev => {
            const existingNames = (prev || '').split(';').map(s => s.trim()).filter(Boolean);
            return [...existingNames, displayName].join('; ');
          });
          setOrganizerEmails(prev => [...prev, email]);
        }}
      />
      <BulkUserImportModal
        open={bulkTestTeamOpen}
        onClose={() => setBulkTestTeamOpen(false)}
        title="Massenimport — Test-Team"
        description={(
          <p style={{ marginTop: 0 }}>
            Mehrere <strong>Test-Team-Mitglieder</strong> auf einmal hinzufügen. Test-Team
            sieht das Event schon im Entwurfsmodus und kann sich testweise anmelden.
          </p>
        )}
        existingEmails={testTeamEmails}
        searchUsers={searchUsers}
        onAdd={({ email, displayName }) => {
          setTestTeamNames(prev => [...prev, displayName]);
          setTestTeamEmails(prev => [...prev, email]);
        }}
      />
      <BulkUserImportModal
        open={bulkQrScannerOpen}
        onClose={() => setBulkQrScannerOpen(false)}
        title="Massenimport — Check-In Team"
        description={(
          <p style={{ marginTop: 0 }}>
            Mehrere <strong>Check-In-Team-Mitglieder</strong> auf einmal hinzufügen. Diese
            Personen dürfen am Eventtag den QR-Scanner / Check-In-Tool benutzen, haben aber
            keine weiteren Admin-Rechte.
          </p>
        )}
        existingEmails={qrScannerEmails}
        searchUsers={searchUsers}
        onAdd={({ email, displayName }) => {
          setQrScannerNames(prev => [...prev, displayName]);
          setQrScannerEmails(prev => [...prev, email]);
        }}
      />

      {/* Mitglieder-Modal: zeigt die Members einer Entra-Gruppe an */}
      {memberModalOpen && (
        <div style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 1100,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }} onClick={() => setMemberModalOpen(false)}>
          <div className="card" style={{ width: '90%', maxWidth: 560, maxHeight: '80vh', overflow: 'auto', padding: 24 }} onClick={e => e.stopPropagation()}>
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0 }}>
                <Users size={18} /> Mitglieder von <span style={{ color: 'var(--dex-green-dark)' }}>{memberModalGroupName}</span>
              </h3>
              <button
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-600)' }}
                onClick={() => setMemberModalOpen(false)}
              >
                <X size={20} />
              </button>
            </div>
            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginBottom: 12 }}>
              {memberModalGroupEmail}
            </div>
            {memberModalLoading ? (
              <p style={{ color: 'var(--dex-gray-500)', textAlign: 'center', padding: 20 }}>Mitglieder werden geladen...</p>
            ) : memberModalError ? (
              <div style={{ background: 'rgba(218,41,28,0.08)', border: '1px solid var(--dex-red, #c00)', borderRadius: 6, padding: 12, fontSize: '0.85rem', color: 'var(--dex-red, #c00)' }}>
                {memberModalError}
              </div>
            ) : memberModalMembers.length === 0 ? (
              <p style={{ color: 'var(--dex-gray-500)', textAlign: 'center', padding: 20 }}>Keine Mitglieder gefunden.</p>
            ) : (
              <>
                <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginBottom: 8 }}>
                  <strong>{memberModalMembers.length}</strong> {memberModalMembers.length === 1 ? 'Mitglied' : 'Mitglieder'}:
                </p>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                      <th style={{ textAlign: 'left', padding: 6 }}>Name</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>E-Mail</th>
                    </tr>
                  </thead>
                  <tbody>
                    {memberModalMembers.map(m => (
                      <tr key={m.email} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                        <td style={{ padding: 6 }}>{m.displayName}</td>
                        <td style={{ padding: 6, color: 'var(--dex-gray-600)' }}>{m.email}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            )}
          </div>
        </div>
      )}

      {/* v11.88: Demo-Auswahl-Modal — Ersatz fuer den frueheren
          direkten „Demo"-Klick. Vier Karten-Optionen: Standard,
          Mit Gruppen, Mit Sub-Event, Mit Sub-Event + Team. Klick auf
          eine Karte schliesst das Modal und fuellt das Formular mit
          der jeweiligen Variante. */}
      {/* v18.6: Power-User-Hilfe als „?"-Ball unten rechts (nur auf Wizard-
          Seite 1, nur wenn Power User existieren). Klick auf den Ball klappt
          das Panel mit dem Self-Service-Hinweis + den Power-User-Kontakten
          auf; X (oder erneuter Ball-Klick) klappt es wieder zu. */}
      {currentStep === 0 && powerUsers.length > 0 && (
        // v18.51: Ball in die rechte Lücke NEBEN der zentrierten Kachel
        // (page-container max 1100px) verankern statt in die Viewport-Ecke —
        // dort ging er bei breiten Screens/SPFx-Canvas unten rechts verloren.
        // left = Kachelmitte + halbe Kachelbreite + Abstand, geclamped damit er
        // auf schmalen Screens nicht aus dem Bild läuft. Panel öffnet absolut
        // darüber (right:0 = bündig zur Ball-rechten Kante, wächst nach links).
        <div style={{ position: 'fixed', bottom: 40, left: 'min(calc(50vw + 560px), calc(100vw - 230px))', zIndex: 1400 }}>
          {powerUserHelpOpen && (
            <div style={{
              position: 'absolute', bottom: 68, right: 0,
              width: 'min(360px, calc(100vw - 40px))',
              background: '#fff', borderRadius: 'var(--dex-radius-lg, 16px)',
              boxShadow: '0 12px 40px rgba(0,0,0,0.22)',
              border: '1px solid var(--dex-gray-200)',
              overflow: 'hidden',
              animation: 'dexBannerSlideIn 0.25s ease-out',
            }}>
              <div style={{
                display: 'flex', alignItems: 'center', gap: 8,
                padding: '12px 16px', background: 'rgba(0,118,168,0.08)',
                borderBottom: '1px solid var(--dex-gray-200)',
              }}>
                <Icon iconName="Help" style={{ fontSize: 18, color: 'var(--dex-blue, #0076a8)' }} />
                <strong style={{ flex: 1, fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>
                  {isDe ? 'Brauchst du Hilfe?' : 'Need help?'}
                </strong>
                <button
                  type="button"
                  onClick={() => setPowerUserHelpOpen(false)}
                  aria-label={isDe ? 'Schließen' : 'Close'}
                  style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-500)', padding: 2, display: 'inline-flex' }}
                >
                  <X size={18} />
                </button>
              </div>
              <div style={{ padding: '14px 16px' }}>
                <p style={{ margin: '0 0 12px', fontSize: '0.83rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
                  {isDe
                    ? <>DEX ist als <strong>Self-Service</strong> gedacht und versucht, möglichst selbsterklärend zu sein. Wenn du trotzdem mal nicht weiterkommst oder etwas nicht verstehst, wende dich gern an unsere <strong>Power User</strong>:</>
                    : <>DEX is designed as <strong>self-service</strong> and tries to be as self-explanatory as possible. If you do get stuck or something is unclear, feel free to reach out to our <strong>power users</strong>:</>}
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                  {powerUsers.map(pu => {
                    // Name aus „Nachname, Vorname" → „Vorname Nachname".
                    const parts = (pu.userName || '').split(',').map(s => s.trim());
                    const displayName = parts.length === 2 ? `${parts[1]} ${parts[0]}` : (pu.userName || pu.userEmail);
                    const photoUrl = `${siteUrl}/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(pu.userEmail)}&size=M`;
                    return (
                      <a
                        key={pu.id}
                        href={`mailto:${pu.userEmail}`}
                        style={{
                          display: 'flex', alignItems: 'center', gap: 10, textDecoration: 'none',
                          padding: '6px 8px', borderRadius: 'var(--dex-radius, 12px)',
                          border: '1px solid var(--dex-gray-200)', color: 'var(--dex-gray-800)',
                        }}
                      >
                        <img
                          src={photoUrl}
                          alt={displayName}
                          style={{ width: 36, height: 36, borderRadius: '50%', objectFit: 'cover', flexShrink: 0, background: 'var(--dex-gray-100)' }}
                          onError={e => { (e.target as HTMLImageElement).style.visibility = 'hidden'; }}
                        />
                        <span style={{ minWidth: 0 }}>
                          <span style={{ display: 'block', fontSize: '0.85rem', fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{displayName}</span>
                          <span style={{ display: 'block', fontSize: '0.74rem', color: 'var(--dex-blue, #0076a8)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{pu.userEmail}</span>
                        </span>
                      </a>
                    );
                  })}
                </div>
              </div>
            </div>
          )}
          {/* v18.52: beschrifteter Pill-Button statt nur „?" — verständlicher. */}
          <button
            type="button"
            onClick={() => setPowerUserHelpOpen(o => !o)}
            aria-label={isDe ? 'Hilfe: Power User kontaktieren' : 'Help: contact power users'}
            aria-expanded={powerUserHelpOpen}
            title={isDe ? 'Benötigst du Hilfe? Power User kontaktieren' : 'Need help? Contact power users'}
            style={{
              height: 48, borderRadius: 24, flexShrink: 0,
              padding: '0 18px',
              background: 'var(--dex-blue, #0076a8)', color: '#fff',
              border: 'none', cursor: 'pointer',
              boxShadow: '0 6px 20px rgba(0,0,0,0.25)',
              display: 'inline-flex', alignItems: 'center', gap: 8,
              fontSize: '0.9rem', fontWeight: 600, lineHeight: 1, whiteSpace: 'nowrap',
            }}
          >
            {powerUserHelpOpen
              ? <><X size={18} /> {isDe ? 'Schließen' : 'Close'}</>
              : <><Icon iconName="Help" style={{ fontSize: 18 }} /> {isDe ? 'Benötigst du Hilfe?' : 'Need help?'}</>}
          </button>
        </div>
      )}

      {showDemoVariantModal && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
          onClick={() => setShowDemoVariantModal(false)}
        >
          <div
            className="card"
            style={{
              width: '100%', maxWidth: 760, maxHeight: '90vh', overflow: 'auto',
              padding: 24, borderRadius: 16, background: '#fff',
              boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
            }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                {isDe ? 'Demo-Daten laden' : 'Load demo data'}
              </h3>
              <button
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-600)' }}
                onClick={() => setShowDemoVariantModal(false)}
                aria-label={isDe ? 'Schließen' : 'Close'}
              >
                <X size={20} />
              </button>
            </div>
            <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
              {isDe
                ? 'Wähle eine Vorlage. Die ausgewählte Variante überschreibt deine aktuellen Eingaben — verworfen wird nichts, falls du noch nichts gespeichert hast.'
                : 'Choose a template. The selected variant overrides your current input — nothing is lost if you haven\'t saved yet.'}
            </p>
            {(() => {
              const cards: Array<{ key: keyof typeof DEMO_VARIANTS; titleDe: string; titleEn: string; descDe: string; descEn: string }> = [
                {
                  key: 'standard',
                  titleDe: 'Standard',
                  titleEn: 'Standard',
                  descDe: 'Ein Event, eine Gruppe. Typisches Meeting / Lunch.',
                  descEn: 'One event, one group. Typical meeting or lunch.',
                },
                {
                  key: 'groups',
                  titleDe: 'Mit Gruppen',
                  titleEn: 'With groups',
                  descDe: 'Event mit zwei Teilnehmer-Gruppen (Split Capacity), z.B. Vormittag / Nachmittag.',
                  descEn: 'Event with two participant groups (split capacity), e.g. morning / afternoon.',
                },
                {
                  key: 'subevent',
                  titleDe: 'Mit Sub-Event',
                  titleEn: 'With sub-event',
                  descDe: 'Haupt-Event + 1 Sub-Event, z.B. Conference + Dinner.',
                  descEn: 'Main event + 1 sub-event, e.g. conference + dinner.',
                },
                {
                  key: 'subeventTeam',
                  titleDe: 'Mit Sub-Event + Team',
                  titleEn: 'With sub-event + team',
                  descDe: 'Wie links, aber mit Team-Anmeldung (Teams à 4 Personen).',
                  descEn: 'Same as on the left, but with team registration (teams of 4 people).',
                },
              ];
              return (
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: 12 }}>
                  {cards.map(card => (
                    <button
                      key={card.key}
                      type="button"
                      onClick={() => {
                        DEMO_VARIANTS[card.key]();
                        setShowDemoVariantModal(false);
                      }}
                      style={{
                        textAlign: 'left',
                        padding: 16,
                        borderRadius: 12,
                        border: '1px solid var(--dex-gray-200)',
                        background: '#fff',
                        cursor: 'pointer',
                        transition: 'all 0.15s ease',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: 8,
                        minHeight: 120,
                      }}
                      onMouseEnter={e => {
                        e.currentTarget.style.borderColor = 'var(--dex-green-dark, #4a7c1f)';
                        e.currentTarget.style.boxShadow = '0 6px 16px rgba(74,124,31,0.12)';
                        e.currentTarget.style.transform = 'translateY(-2px)';
                      }}
                      onMouseLeave={e => {
                        e.currentTarget.style.borderColor = 'var(--dex-gray-200)';
                        e.currentTarget.style.boxShadow = 'none';
                        e.currentTarget.style.transform = 'translateY(0)';
                      }}
                    >
                      <div style={{ fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
                        {isDe ? card.titleDe : card.titleEn}
                      </div>
                      <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                        {isDe ? card.descDe : card.descEn}
                      </div>
                    </button>
                  ))}
                </div>
              );
            })()}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 20, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
              <button
                type="button"
                className="btn btn-outline"
                onClick={() => setShowDemoVariantModal(false)}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* v17.21: A4-Zusammenfassungs-Modal nach erfolgreichem Save — fragt
          den Organizer, ob er das gesamte Event als PDF oder Word herunter-
          laden moechte (z.B. zur Durchsicht durch einen Partner). Beim
          Klick auf eine Option laeuft der Export sofort, danach feuert der
          eigentliche „Wizard verlassen"-Dispatch (`dex-event-submit-success`).
          „Nein, danke" springt direkt zum Dispatch. */}
      {showSummaryModal && pendingSuccessDispatch && (() => {
        const closeAndDispatch = (): void => {
          // v17.22: Ref VOR dem Dispatch leeren, damit der Unmount-Cleanup
          // nicht ein zweites Mal feuert (Doppel-Navigation/-Banner).
          const payload = pendingSuccessDispatchRef.current || pendingSuccessDispatch;
          pendingSuccessDispatchRef.current = null;
          setShowSummaryModal(false);
          setPendingSuccessDispatch(null);
          try {
            window.dispatchEvent(new CustomEvent('dex-event-submit-success', {
              detail: payload,
            }));
          } catch { /* */ }
        };
        const buildData = (): SummaryData => {
          // Bild als DataURL (falls noch nicht Base64): unten reicht der
          // bestehende imagePreview, der bei neu hochgeladenen Bildern
          // bereits eine Data-URL ist und bei bestehenden Events die
          // SharePoint-URL. Letztere wird im PDF/Doc-Export im Print-View
          // i.d.R. nicht geladen (CORS) — wir bauen einen Fallback-Text.
          const subEventsForSummary = subEvents.map(se => ({
            title: se.title || '',
            startDate: se.startDate,
            endDate: se.endDate,
            location: se.location,
            description: se.description,
            maxParticipants: typeof se.maxParticipants === 'number' ? se.maxParticipants : undefined,
            waitlistEnabled: !!se.waitlistEnabled,
          }));
          const customFieldsForSummary = customFields
            .filter(f => f.label && f.label.trim().length > 0)
            .map(f => ({
              id: f.id,
              label: f.label,
              type: f.type,
              required: !!f.required,
              helpText: f.helpText,
              helpTextStyle: f.helpTextStyle,
              confirmLabel: f.confirmLabel,
              options: f.type === 'select' ? f.options : undefined,
              multi: !!f.multi,
              onlyForGroup: f.onlyForGroup,
              labelEn: f.labelEn,
              helpTextEn: f.helpTextEn,
              confirmLabelEn: f.confirmLabelEn,
              optionsEn: f.optionsEn,
              showIf: f.showIf,
            }));
          // Transferzeiten + Agenda werden in den Summary-Helper als
          // vereinfachte Spalten gemappt — das Detail-Schema bleibt im
          // Wizard, der Export nimmt die fuer Reviewer relevanten Spalten.
          const transfersForSummary = transferTimes.map(t => ({
            time: [t.date, t.departureTime].filter(Boolean).join(' '),
            description: [t.location, t.description, t.meetingPoint].filter(Boolean).join(' — '),
          }));
          const agendaForSummary = agenda.map(a => ({
            time: [a.date, a.time, a.endTime ? ` – ${a.endTime}` : ''].filter(Boolean).join(' '),
            topic: a.title,
            speaker: a.description,
          }));
          const quizForSummary = quiz.map(q => ({
            question: q.question,
            options: q.options,
            correctIndex: (q.correctIndices && q.correctIndices.length > 0) ? q.correctIndices[0] : undefined,
          }));
          const documentsForSummary = documents.map(doc => ({
            name: doc.name,
            size: doc.size,
          }));
          return {
            title,
            description,
            imageDataUrl: imagePreview || eventImageUrl || undefined,
            startDate,
            endDate,
            organizers: organizer.split(';').map(s => s.trim()).filter(Boolean),
            organizerEmails,
            contactName,
            contactEmail,
            contactInfo,
            testTeam: testTeamNames,
            qrScanners: qrScannerNames,
            isFictive,
            activeFrom,
            location,
            address: { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity },
            agenda: agendaForSummary,
            transfers: transfersForSummary,
            locationFilter: locationFilter ? locationFilter.split(';').map(s => s.trim()).filter(Boolean) : [],
            audience: audience ? audience.split(';').map(s => s.trim()).filter(Boolean) : [],
            filterMode,
            excludedUsers,
            registrationDeadline,
            lastDeregisterDate,
            maxParticipants: Number(maxParticipants) || 0,
            unlimitedParticipants,
            waitlistEnabled,
            durchstarterCapacity: Number(durchstarterCapacity) || 0,
            funstarterCapacity: Number(funstarterCapacity) || 0,
            splitLabelA, splitLabelB,
            splitSharedWaitlist,
            teamRegistrationEnabled,
            teamSize,
            askTeamName,
            teamPartialAllowed,
            teamOpenSlotsVisible,
            teamJoinRequiresApproval,
            askSalutation,
            bilingualFields,
            customFields: customFieldsForSummary,
            allowAttendeeUpload,
            attendeeUploadHint,
            attendeeUploadLabel,
            emailLanguage,
            disableEmails,
            disableOutlook,
            outlookHeading,
            outlookSubheading,
            outlookBody,
            notifyOrgRegisterMode,
            notifyOrgRegisterFromDate,
            notifyOrgCancelMode,
            documents: documentsForSummary,
            funZone: quizForSummary,
            quizClusterSize,
            subEvents: subEventsForSummary,
            childTermSingular,
            childTermPlural,
            subEventsOnlyMode,
            requireSubEventSelection,
            generatedAt: new Date().toISOString(),
            locale: isDe ? 'de' : 'en',
          };
        };
        const onPdf = (): void => {
          try { exportSummaryAsPdf(buildData()); } catch (err) {
            console.warn('[DEX] exportSummaryAsPdf failed:', err);
          }
          closeAndDispatch();
        };
        const onDoc = (): void => {
          try { exportSummaryAsDoc(buildData()); } catch (err) {
            console.warn('[DEX] exportSummaryAsDoc failed:', err);
          }
          closeAndDispatch();
        };
        return (
          <div
            style={{
              position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1300,
              display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
            }}
            onClick={closeAndDispatch}
          >
            <div
              className="card"
              style={{
                width: '100%', maxWidth: 560, padding: 24, borderRadius: 16,
                background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
              }}
              onClick={e => e.stopPropagation()}
            >
              <h3 style={{ margin: 0, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                {isDe ? 'Event-Zusammenfassung herunterladen?' : 'Download event summary?'}
              </h3>
              <p style={{ marginTop: 12, color: 'var(--dex-gray-700)', lineHeight: 1.55, fontSize: '0.95rem' }}>
                {isDe
                  ? <>Das Event wurde gespeichert. Möchtest du jetzt eine <strong>A4-Zusammenfassung</strong> mit allen Sektionen (Foto, Beschreibung, Sichtbarkeit, Felder, Kommunikation, Dokumente, Sub-Events…) herunterladen? Du kannst sie z.B. einem Partner zur Durchsicht weiterleiten.</>
                  : <>The event has been saved. Would you like to download a <strong>one-page A4 summary</strong> with every section (photo, description, visibility, fields, communication, documents, sub-events…)? You can forward it to a partner for review.</>}
              </p>
              <div style={{
                marginTop: 18, padding: '10px 14px', background: 'rgba(0,90,156,0.06)',
                border: '1px solid rgba(0,90,156,0.25)', borderRadius: 8,
                fontSize: '0.82rem', color: 'var(--dex-gray-700)',
              }}>
                {isDe
                  ? <><strong>Hinweis:</strong> Beim PDF-Export öffnet sich der Browser-Druckdialog. Wähle dort <strong>&bdquo;Als PDF speichern&ldquo;</strong> als Ziel. Word-Export lädt direkt eine .doc-Datei herunter.</>
                  : <><strong>Note:</strong> The PDF export opens the browser print dialog — pick <strong>&ldquo;Save as PDF&rdquo;</strong> as the destination. Word export downloads a .doc file directly.</>}
              </div>
              <div style={{ marginTop: 22, display: 'flex', flexWrap: 'wrap', gap: 10, justifyContent: 'flex-end' }}>
                <button className="btn btn-outline" onClick={closeAndDispatch}>
                  {isDe ? 'Nein, danke' : 'No, thanks'}
                </button>
                <button className="btn btn-secondary" onClick={onDoc}>
                  {isDe ? 'Als Word (.doc)' : 'As Word (.doc)'}
                </button>
                <button className="btn btn-primary" onClick={onPdf}>
                  {isDe ? 'Als PDF' : 'As PDF'}
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* v8.6: Exclude-Users-Modal — zeigt resolved Mitglieder der
          Verteiler/User-Liste plus Suchfeld. Per Checkbox kann der
          Organizer einzelne Personen ausschliessen, die NICHT mehr in der
          Sichtbarkeit auftauchen sollen. */}
      {excludeModalOpen && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
          onClick={() => setExcludeModalOpen(false)}
        >
          <div
            className="card"
            style={{
              width: '100%', maxWidth: 1100, maxHeight: '90vh', overflow: 'auto',
              padding: 24, borderRadius: 16, background: '#fff',
              boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
            }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0 }}>
                <Users size={18} /> {isDe ? 'Personen ausschließen' : 'Exclude users'}
              </h3>
              <button
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-600)' }}
                onClick={() => setExcludeModalOpen(false)}
              >
                <X size={20} />
              </button>
            </div>

            <p style={{ margin: '0 0 12px', fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
              {isDe
                ? 'Hier kannst du einzelne Personen explizit ausschließen — sie sehen das Event NICHT, auch wenn sie über Standortfilter oder Mailverteiler eigentlich Sichtbarkeit hätten. Die Mitglieder der oben gewählten Mailverteiler werden automatisch aufgelistet (per Microsoft Graph). Standortfilter-User sind nicht direkt aus der App auflistbar — die kannst du über die Suche unten gezielt finden und ausschließen.'
                : 'Here you can explicitly exclude individuals — they will NOT see the event, even if they would otherwise have visibility via location filter or mailing list. Members of the mailing lists chosen above are listed automatically (via Microsoft Graph). Users matched only by location filter cannot be listed directly — use the search below to find and exclude them.'}
            </p>

            {/* Suchfeld — filtert die Tabelle global ueber Email/Vor-/
                Nachname/Position, und ergaenzt bei Bedarf neue User via
                Directory-Suche (z.B. wenn der Gesuchte nicht im Verteiler
                ist, aber explizit ausgeschlossen werden soll). */}
            <div style={{ marginBottom: 12 }}>
              <input
                type="text"
                value={excludeSearch}
                onChange={async e => {
                  const v = e.target.value;
                  setExcludeSearch(v);
                  // Bei Such-Eingabe immer auf Seite 0 zuruecksetzen,
                  // damit der Treffer auch sichtbar ist.
                  setExcludePage(0);
                  if (v.trim().length < 2) return;
                  try {
                    const found = await searchUsers(v.trim(), excludeIncludeIntl);
                    // Nur User, die noch nicht in der resolved-Liste stecken,
                    // anhaengen — sonst Duplikate. seen-Set baut sich durch
                    // resolved + bereits in der Suche gefundene auf.
                    setExcludeResolvedUsers(prev => {
                      const seen = new Set(prev.map(u => u.email.toLowerCase()));
                      const next = [...prev];
                      for (const u of found) {
                        const k = (u.email || '').toLowerCase();
                        if (k && !seen.has(k)) {
                          seen.add(k);
                          // displayName splitten zu first/last falls noetig (Format
                          // 'Nachname, Vorname' oder 'Vorname Nachname').
                          let fn = '';
                          let ln = '';
                          const dn = (u.displayName || '').trim();
                          if (dn.indexOf(',') >= 0) {
                            const parts = dn.split(',').map(s => s.trim());
                            ln = parts[0] || '';
                            fn = parts[1] || '';
                          } else {
                            const parts = dn.split(/\s+/);
                            fn = parts[0] || '';
                            ln = parts.slice(1).join(' ');
                          }
                          next.push({
                            email: u.email,
                            displayName: u.displayName,
                            firstName: fn,
                            lastName: ln,
                            jobTitle: u.jobTitle || '',
                            location: u.location || '',
                            source: isDe ? 'Suche' : 'search',
                          });
                        }
                      }
                      return next;
                    });
                  } catch { /* */ }
                }}
                placeholder={isDe ? 'Person suchen (Name oder E-Mail)' : 'Search person (name or email)'}
                className="form-input"
                style={{ width: '100%' }}
              />
              <InternationalSearchToggle
                checked={excludeIncludeIntl}
                onChange={async next => {
                  setExcludeIncludeIntl(next);
                  const v = excludeSearch.trim();
                  if (v.length < 2) return;
                  try {
                    const found = await searchUsers(v, next);
                    setExcludeResolvedUsers(prev => {
                      const seen = new Set(prev.map(u => u.email.toLowerCase()));
                      const acc = [...prev];
                      for (const u of found) {
                        const k = (u.email || '').toLowerCase();
                        if (k && !seen.has(k)) {
                          seen.add(k);
                          let fn = '';
                          let ln = '';
                          const dn = (u.displayName || '').trim();
                          if (dn.indexOf(',') >= 0) {
                            const parts = dn.split(',').map(s => s.trim());
                            ln = parts[0] || '';
                            fn = parts[1] || '';
                          } else {
                            const parts = dn.split(/\s+/);
                            fn = parts[0] || '';
                            ln = parts.slice(1).join(' ');
                          }
                          acc.push({
                            email: u.email,
                            displayName: u.displayName,
                            firstName: fn,
                            lastName: ln,
                            jobTitle: u.jobTitle || '',
                            location: u.location || '',
                            source: isDe ? 'Suche' : 'search',
                          });
                        }
                      }
                      return acc;
                    });
                  } catch { /* */ }
                }}
                isDe={isDe}
              />
            </div>

            {excludeResolving && (
              <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', textAlign: 'center', padding: 16 }}>
                {isDe ? 'Verteiler werden aufgelöst…' : 'Resolving distribution lists…'}
              </p>
            )}

            {!excludeResolving && excludeResolvedUsers.length === 0 && (
              <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', padding: 16, textAlign: 'center' }}>
                {isDe
                  ? 'Keine Personen aufgelöst. Nutze die Suche, um einzelne Personen hinzuzufügen.'
                  : 'No people resolved. Use the search to add individuals.'}
              </p>
            )}

            {/* v8.8: Tabelle mit Spalten Email / Nachname / Vorname / Position
                / Standort + Filter-Inputs in der Header-Zeile + Sortier-Klick.
                Damit kann der Organizer z.B. nach jobTitle 'Partner' filtern
                und alle nicht-Partner per 'Alle ausschliessen'-Aktion
                excluden. */}
            {(() => {
              const f = excludeFilters;
              // v8.12: Globaler Such-Filter (excludeSearch) wirkt
              // zusaetzlich zur Spalten-Filter-Logik. Match ueber Email,
              // Vor-/Nachname und Position. Damit landet ein Such-Treffer
              // (z.B. 'brenneisen') sofort als einziger sichtbarer Eintrag
              // in der Tabelle.
              const gs = excludeSearch.trim().toLowerCase();
              const filtered = excludeResolvedUsers.filter(u =>
                (!f.email || u.email.toLowerCase().indexOf(f.email.toLowerCase()) >= 0) &&
                (!f.lastName || u.lastName.toLowerCase().indexOf(f.lastName.toLowerCase()) >= 0) &&
                (!f.firstName || u.firstName.toLowerCase().indexOf(f.firstName.toLowerCase()) >= 0) &&
                (!f.jobTitle || u.jobTitle.toLowerCase().indexOf(f.jobTitle.toLowerCase()) >= 0) &&
                (!f.location || u.location.toLowerCase().indexOf(f.location.toLowerCase()) >= 0) &&
                (!gs ||
                  u.email.toLowerCase().indexOf(gs) >= 0 ||
                  u.lastName.toLowerCase().indexOf(gs) >= 0 ||
                  u.firstName.toLowerCase().indexOf(gs) >= 0 ||
                  u.jobTitle.toLowerCase().indexOf(gs) >= 0 ||
                  u.displayName.toLowerCase().indexOf(gs) >= 0)
              );
              const sorted = [...filtered].sort((a, b) => {
                const av = (a[excludeSortBy] || '').toLowerCase();
                const bv = (b[excludeSortBy] || '').toLowerCase();
                if (av < bv) return excludeSortDir === 'asc' ? -1 : 1;
                if (av > bv) return excludeSortDir === 'asc' ? 1 : -1;
                return 0;
              });
              const headerSort = (col: typeof excludeSortBy): void => {
                if (excludeSortBy === col) setExcludeSortDir(d => d === 'asc' ? 'desc' : 'asc');
                else { setExcludeSortBy(col); setExcludeSortDir('asc'); }
              };
              const sortIcon = (col: typeof excludeSortBy): string => excludeSortBy === col ? (excludeSortDir === 'asc' ? ' ▲' : ' ▼') : '';
              const headerStyle: React.CSSProperties = { padding: '8px 6px', textAlign: 'left', fontSize: '0.78rem', fontWeight: 700, color: 'var(--dex-gray-700)', cursor: 'pointer', userSelect: 'none', whiteSpace: 'nowrap' };
              const cellStyle: React.CSSProperties = { padding: '6px', fontSize: '0.82rem', borderBottom: '1px solid var(--dex-gray-100)' };
              const filterInputStyle: React.CSSProperties = { width: '100%', padding: '4px 6px', border: '1px solid var(--dex-gray-200)', borderRadius: 4, fontSize: '0.75rem' };
              const filterActive = !!(f.email || f.lastName || f.firstName || f.jobTitle || f.location);
              const totalPages = Math.max(1, Math.ceil(sorted.length / EXCLUDE_PAGE_SIZE));
              const safePage = Math.min(excludePage, totalPages - 1);
              const pageStart = safePage * EXCLUDE_PAGE_SIZE;
              const pageEnd = Math.min(pageStart + EXCLUDE_PAGE_SIZE, sorted.length);
              const pageItems = sorted.slice(pageStart, pageEnd);
              return (
                <>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6, flexWrap: 'wrap', gap: 8 }}>
                    <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                      {filterActive
                        ? (isDe
                          ? <><strong>{filtered.length}</strong> von <strong>{excludeResolvedUsers.length}</strong> Personen passen zum Filter</>
                          : <><strong>{filtered.length}</strong> of <strong>{excludeResolvedUsers.length}</strong> people match the filter</>)
                        : (isDe
                          ? <><strong>{excludeResolvedUsers.length}</strong> Personen aufgelöst</>
                          : <><strong>{excludeResolvedUsers.length}</strong> people resolved</>)
                      }
                    </span>
                    {filtered.length > 0 && (
                      <div style={{ display: 'flex', gap: 6 }}>
                        <button
                          type="button"
                          onClick={() => {
                            setExcludedUsers(prev => {
                              const seen = new Set(prev.map(e => e.toLowerCase()));
                              const next = [...prev];
                              for (const u of filtered) {
                                const k = u.email.toLowerCase();
                                if (!seen.has(k)) {
                                  seen.add(k);
                                  next.push(u.email);
                                }
                              }
                              return next;
                            });
                          }}
                          style={{ fontSize: '0.72rem', padding: '4px 10px', background: 'rgba(218,41,28,0.08)', border: '1px solid var(--dex-red, #c00)', color: 'var(--dex-red, #c00)', borderRadius: 6, cursor: 'pointer', fontWeight: 600 }}
                        >
                          {isDe ? 'Alle gefilterten ausschließen' : 'Exclude all filtered'}
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            const filterEmails = new Set(filtered.map(u => u.email.toLowerCase()));
                            setExcludedUsers(prev => prev.filter(e => !filterEmails.has(e.toLowerCase())));
                          }}
                          style={{ fontSize: '0.72rem', padding: '4px 10px', background: '#fff', border: '1px solid var(--dex-gray-300)', color: 'var(--dex-gray-700)', borderRadius: 6, cursor: 'pointer' }}
                        >
                          {isDe ? 'Alle wieder einschließen' : 'Include all again'}
                        </button>
                      </div>
                    )}
                  </div>
                  <div style={{ maxHeight: 420, overflowY: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 6 }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                      <thead style={{ position: 'sticky', top: 0, background: 'var(--dex-gray-50, #fafafa)', zIndex: 1 }}>
                        <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                          <th style={{ ...headerStyle, width: 36, cursor: 'default' }} />
                          <th style={headerStyle} onClick={() => headerSort('email')}>E-Mail{sortIcon('email')}</th>
                          <th style={headerStyle} onClick={() => headerSort('lastName')}>{isDe ? 'Nachname' : 'Last name'}{sortIcon('lastName')}</th>
                          <th style={headerStyle} onClick={() => headerSort('firstName')}>{isDe ? 'Vorname' : 'First name'}{sortIcon('firstName')}</th>
                          <th style={headerStyle} onClick={() => headerSort('jobTitle')}>{isDe ? 'Position' : 'Position'}{sortIcon('jobTitle')}</th>
                          <th style={headerStyle} onClick={() => headerSort('location')}>{isDe ? 'Standort' : 'Location'}{sortIcon('location')}</th>
                        </tr>
                        <tr style={{ borderBottom: '1px solid var(--dex-gray-200)', background: '#fff' }}>
                          <th style={{ padding: 4 }} />
                          <th style={{ padding: 4 }}><input style={filterInputStyle} value={f.email} onChange={e => { setExcludeFilters(p => ({ ...p, email: e.target.value })); setExcludePage(0); }} placeholder={isDe ? 'filtern…' : 'filter…'} /></th>
                          <th style={{ padding: 4 }}><input style={filterInputStyle} value={f.lastName} onChange={e => { setExcludeFilters(p => ({ ...p, lastName: e.target.value })); setExcludePage(0); }} placeholder={isDe ? 'filtern…' : 'filter…'} /></th>
                          <th style={{ padding: 4 }}><input style={filterInputStyle} value={f.firstName} onChange={e => { setExcludeFilters(p => ({ ...p, firstName: e.target.value })); setExcludePage(0); }} placeholder={isDe ? 'filtern…' : 'filter…'} /></th>
                          <th style={{ padding: 4 }}><input style={filterInputStyle} value={f.jobTitle} onChange={e => { setExcludeFilters(p => ({ ...p, jobTitle: e.target.value })); setExcludePage(0); }} placeholder={isDe ? 'z.B. Partner' : 'e.g. Partner'} /></th>
                          <th style={{ padding: 4 }}><input style={filterInputStyle} value={f.location} onChange={e => { setExcludeFilters(p => ({ ...p, location: e.target.value })); setExcludePage(0); }} placeholder={isDe ? 'z.B. Köln' : 'e.g. Cologne'} /></th>
                        </tr>
                      </thead>
                      <tbody>
                        {pageItems.map(u => {
                          const emailLc = u.email.toLowerCase();
                          const isExcluded = excludedUsers.some(e => e.toLowerCase() === emailLc);
                          const toggle = (): void => {
                            if (isExcluded) setExcludedUsers(prev => prev.filter(e => e.toLowerCase() !== emailLc));
                            else setExcludedUsers(prev => [...prev, u.email]);
                          };
                          return (
                            <tr
                              key={u.email}
                              onClick={toggle}
                              style={{
                                cursor: 'pointer',
                                background: isExcluded ? 'rgba(218,41,28,0.06)' : 'transparent',
                              }}
                            >
                              <td style={{ ...cellStyle, textAlign: 'center' }}>
                                <input
                                  type="checkbox"
                                  checked={!isExcluded}
                                  onChange={toggle}
                                  onClick={e => e.stopPropagation()}
                                  style={{ accentColor: 'var(--dex-green)', width: 14, height: 14, cursor: 'pointer' }}
                                />
                              </td>
                              <td style={{ ...cellStyle, color: isExcluded ? 'var(--dex-red, #c00)' : 'var(--dex-gray-700)', fontFamily: 'monospace', fontSize: '0.78rem' }}>{u.email}</td>
                              <td style={{ ...cellStyle, fontWeight: 500, color: isExcluded ? 'var(--dex-red, #c00)' : 'var(--dex-gray-800)' }}>{u.lastName || '—'}</td>
                              <td style={{ ...cellStyle, color: isExcluded ? 'var(--dex-red, #c00)' : 'var(--dex-gray-700)' }}>{u.firstName || '—'}</td>
                              <td style={{ ...cellStyle, color: 'var(--dex-gray-600)', fontSize: '0.78rem' }}>{u.jobTitle || '—'}</td>
                              <td style={{ ...cellStyle, color: 'var(--dex-gray-600)', fontSize: '0.78rem' }}>{u.location || '—'}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                  {/* v8.9: Pagination — 200 pro Seite. Wird nur angezeigt
                      wenn es mehr als eine Seite gibt. */}
                  {totalPages > 1 && (
                    <div style={{
                      display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                      marginTop: 8, padding: '6px 4px', flexWrap: 'wrap', gap: 8,
                    }}>
                      <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                        {isDe
                          ? <>Zeige <strong>{pageStart + 1}</strong>–<strong>{pageEnd}</strong> von <strong>{sorted.length}</strong></>
                          : <>Showing <strong>{pageStart + 1}</strong>–<strong>{pageEnd}</strong> of <strong>{sorted.length}</strong></>}
                      </span>
                      <div style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                        <button type="button" disabled={safePage === 0} onClick={() => setExcludePage(0)} style={{ padding: '4px 8px', fontSize: '0.78rem', border: '1px solid var(--dex-gray-300)', background: '#fff', borderRadius: 4, cursor: safePage === 0 ? 'not-allowed' : 'pointer', opacity: safePage === 0 ? 0.4 : 1 }}>«</button>
                        <button type="button" disabled={safePage === 0} onClick={() => setExcludePage(p => Math.max(0, p - 1))} style={{ padding: '4px 8px', fontSize: '0.78rem', border: '1px solid var(--dex-gray-300)', background: '#fff', borderRadius: 4, cursor: safePage === 0 ? 'not-allowed' : 'pointer', opacity: safePage === 0 ? 0.4 : 1 }}>‹ {isDe ? 'Zurück' : 'Prev'}</button>
                        <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-700)', padding: '0 8px' }}>
                          {isDe ? 'Seite' : 'Page'} <strong>{safePage + 1}</strong> / {totalPages}
                        </span>
                        <button type="button" disabled={safePage >= totalPages - 1} onClick={() => setExcludePage(p => Math.min(totalPages - 1, p + 1))} style={{ padding: '4px 8px', fontSize: '0.78rem', border: '1px solid var(--dex-gray-300)', background: '#fff', borderRadius: 4, cursor: safePage >= totalPages - 1 ? 'not-allowed' : 'pointer', opacity: safePage >= totalPages - 1 ? 0.4 : 1 }}>{isDe ? 'Weiter' : 'Next'} ›</button>
                        <button type="button" disabled={safePage >= totalPages - 1} onClick={() => setExcludePage(totalPages - 1)} style={{ padding: '4px 8px', fontSize: '0.78rem', border: '1px solid var(--dex-gray-300)', background: '#fff', borderRadius: 4, cursor: safePage >= totalPages - 1 ? 'not-allowed' : 'pointer', opacity: safePage >= totalPages - 1 ? 0.4 : 1 }}>»</button>
                      </div>
                    </div>
                  )}
                </>
              );
            })()}

            {/* Bereits ausgeschlossene User die NICHT in der resolved-Liste sind
                (z.B. weil sie nur ueber Standortfilter sichtbar waeren und
                ueber die Suche manuell ausgeschlossen wurden in einer
                frueheren Session) — separat darstellen. */}
            {excludedUsers.filter(e => !excludeResolvedUsers.some(u => u.email.toLowerCase() === e.toLowerCase())).length > 0 && (
              <div style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
                <h4 style={{ margin: '0 0 8px', fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                  {isDe ? 'Weitere ausgeschlossene Personen' : 'Other excluded users'}
                </h4>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                  {excludedUsers
                    .filter(e => !excludeResolvedUsers.some(u => u.email.toLowerCase() === e.toLowerCase()))
                    .map(e => (
                      <span key={e} style={{
                        display: 'inline-flex', alignItems: 'center', gap: 6,
                        padding: '3px 4px 3px 10px',
                        background: 'rgba(218,41,28,0.08)',
                        border: '1px solid var(--dex-red, #c00)',
                        color: 'var(--dex-red, #c00)',
                        borderRadius: 999, fontSize: '0.78rem',
                      }}>
                        {e}
                        <button
                          type="button"
                          onClick={() => setExcludedUsers(prev => prev.filter(x => x.toLowerCase() !== e.toLowerCase()))}
                          style={{
                            width: 18, height: 18, borderRadius: '50%',
                            border: 'none', background: 'rgba(218,41,28,0.2)',
                            color: 'var(--dex-red, #c00)', cursor: 'pointer',
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            fontSize: '0.85rem', lineHeight: 1,
                          }}
                          title={isDe ? 'Ausschluss aufheben' : 'Remove exclusion'}
                        >×</button>
                      </span>
                    ))}
                </div>
              </div>
            )}

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 20, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => setExcludeModalOpen(false)}
              >
                {isDe ? 'Fertig' : 'Done'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Email-Verteiler Modal */}
      {showEmailModal && (
        <div style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 1000,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }} onClick={() => setShowEmailModal(false)}>
          <div
            className="card"
            style={{ width: '90%', maxWidth: 600, maxHeight: '80vh', overflow: 'auto', padding: 24 }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0 }}>
                <Users size={18} /> {isDe ? 'Sichtbarkeit prüfen' : 'Check visibility'}
              </h3>
              <button
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-600)' }}
                onClick={() => setShowEmailModal(false)}
              >
                <X size={20} />
              </button>
            </div>

            <p style={{ margin: '0 0 12px', fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
              {isDe
                ? 'Hier kannst du verifizieren, ob die kombinierte Sichtbarkeit (Standortfilter + Mailverteiler / einzelne User + Verknüpfung) wirklich zu der Person passt, die das Event sehen soll. Tippe Name oder E-Mail einer Testperson ein und klick „Suchen" — die Tabelle darunter zeigt, ob sie das Event in ihrer Übersicht sieht und woher die Sichtbarkeit kommt (Standort-Match oder Mitgliedschaft in einem Verteiler).'
                : 'Use this to verify whether the combined visibility (location filter + mailing lists / individual users + the AND/OR mode) actually matches the person you want to reach. Type a test person\'s name or email and click "Search" — the table below shows whether they can see the event and where the visibility comes from (location match or membership in a list).'}
            </p>

            <div style={{ marginBottom: 16, padding: '10px 14px', background: 'var(--dex-gray-100)', borderRadius: 'var(--dex-radius)', fontSize: '0.85rem' }}>
              <div style={{ marginBottom: 6 }}>
                <strong>{isDe ? 'Standortfilter:' : 'Location filter:'}</strong>{' '}
                {locationFilter ? locationFilter.split(',').map(s => s.trim()).map(s => (
                  <span key={s} className="badge badge-green" style={{ marginRight: 6 }}>{s}</span>
                )) : <span style={{ color: 'var(--dex-gray-400)' }}>{isDe ? 'Keine' : 'None'}</span>}
              </div>
              <div style={{ marginBottom: 6 }}>
                <strong>{isDe ? 'Mailverteiler / einzelne User:' : 'Mailing lists / individual users:'}</strong>{' '}
                {audience ? audience.split(',').map(s => s.trim()).map(s => (
                  <span key={s} className="badge badge-orange" style={{ marginRight: 6 }}>{s}</span>
                )) : <span style={{ color: 'var(--dex-gray-400)' }}>{isDe ? 'Keine' : 'None'}</span>}
              </div>
              {locationFilter && audience && (
                <div>
                  <strong>{isDe ? 'Verknüpfung:' : 'Combination:'}</strong>{' '}
                  <span className={`badge ${filterMode === 'AND' ? 'badge-red' : 'badge-green'}`}>
                    {filterMode === 'AND'
                      ? (isDe ? 'UND (beide müssen zutreffen)' : 'AND (both must match)')
                      : (isDe ? 'ODER (eines reicht)' : 'OR (one is enough)')}
                  </span>
                </div>
              )}
            </div>

            <div className="form-group">
              <label className="form-label">User suchen (Name oder E-Mail)</label>
              <div style={{ display: 'flex', gap: 8 }}>
                <input
                  className="form-input"
                  value={emailSearch}
                  onChange={e => setEmailSearch(e.target.value)}
                  placeholder="z.B. Max Mustermann oder mmustermann@"
                  onKeyDown={async e => {
                    if (e.key === 'Enter' && emailSearch.length >= 2) {
                      setIsSearchingEmails(true);
                      const results = await searchUsers(emailSearch, emailSearchIncludeIntl);
                      setEmailSearchResults(results);
                      setIsSearchingEmails(false);
                    }
                  }}
                />
                <button
                  className="btn btn-primary"
                  style={{ whiteSpace: 'nowrap' }}
                  disabled={emailSearch.length < 2 || isSearchingEmails}
                  onClick={async () => {
                    setIsSearchingEmails(true);
                    const results = await searchUsers(emailSearch, emailSearchIncludeIntl);
                    setEmailSearchResults(results);
                    setIsSearchingEmails(false);
                  }}
                >
                  {isSearchingEmails ? '...' : 'Suchen'}
                </button>
              </div>
              <InternationalSearchToggle
                checked={emailSearchIncludeIntl}
                onChange={async next => {
                  setEmailSearchIncludeIntl(next);
                  if (emailSearch.length >= 2) {
                    setIsSearchingEmails(true);
                    try {
                      const results = await searchUsers(emailSearch, next);
                      setEmailSearchResults(results);
                    } catch { /* */ }
                    setIsSearchingEmails(false);
                  }
                }}
                isDe={isDe}
              />
            </div>

            {emailSearchResults.length > 0 && (
              <div style={{ marginTop: 12 }}>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', marginBottom: 8 }}>
                  {emailSearchResults.length} Ergebnis{emailSearchResults.length !== 1 ? 'se' : ''}:
                </p>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                      <th style={{ textAlign: 'left', padding: 6 }}>Name</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>E-Mail</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>Standort</th>
                      <th style={{ textAlign: 'center', padding: 6 }}>{isDe ? 'Sichtbar?' : 'Visible?'}</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Begründung' : 'Reason'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {visibilityCacheLoading && (
                      <tr><td colSpan={5} style={{ padding: 12, textAlign: 'center', fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>
                        {isDe ? 'Verteiler werden aufgelöst…' : 'Resolving distribution lists…'}
                      </td></tr>
                    )}
                    {!visibilityCacheLoading && emailSearchResults.map(u => {
                      // v8.9: Volle isEventVisibleForUser-Logik nachbauen
                      // — Exclude > Standort + Audience + UND/ODER.
                      const emailLc = (u.email || '').toLowerCase();
                      const loc = (u.location || '').toLowerCase();
                      const locationFilters = locationFilter ? locationFilter.split(',').map(s => s.trim().toLowerCase()).filter(Boolean) : [];
                      const hasLocFilter = locationFilters.length > 0 && locationFilters.indexOf('all') < 0;
                      const hasAudFilter = visibilityAudienceCache.size > 0;
                      const isExcluded = excludedUsers.some(e => e.toLowerCase() === emailLc);
                      const matchedLoc = hasLocFilter
                        ? locationFilters.find(f => {
                            const norm = f.replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ä/g, 'ae');
                            return loc.indexOf(f) >= 0 || loc.indexOf(norm) >= 0;
                          })
                        : null;
                      const locMatch = !hasLocFilter || !!matchedLoc;
                      const audMatch = !hasAudFilter || visibilityAudienceCache.has(emailLc);
                      let visible: boolean;
                      let reasonParts: string[] = [];
                      if (!hasLocFilter && !hasAudFilter) {
                        visible = true;
                        reasonParts.push(isDe ? 'Keine Filter gesetzt — für alle sichtbar' : 'No filters set — visible to everyone');
                      } else if (filterMode === 'OR') {
                        visible = (hasLocFilter && locMatch) || (hasAudFilter && audMatch);
                        if (locMatch && hasLocFilter) reasonParts.push(isDe ? `Standort-Match (${matchedLoc})` : `location match (${matchedLoc})`);
                        if (audMatch && hasAudFilter) reasonParts.push(isDe ? 'in Mailverteiler/User-Liste' : 'in mailing list / user');
                        if (!visible) reasonParts.push(isDe ? 'kein Filter passt' : 'no filter matches');
                      } else {
                        // AND
                        visible = (!hasLocFilter || locMatch) && (!hasAudFilter || audMatch);
                        if (hasLocFilter) reasonParts.push(locMatch ? (isDe ? `Standort ✓ (${matchedLoc})` : `location ✓ (${matchedLoc})`) : (isDe ? 'Standort ✗' : 'location ✗'));
                        if (hasAudFilter) reasonParts.push(audMatch ? (isDe ? 'Verteiler ✓' : 'mailing list ✓') : (isDe ? 'Verteiler ✗' : 'mailing list ✗'));
                      }
                      // Exclude hat Vorrang
                      if (isExcluded && visible) {
                        visible = false;
                        reasonParts = [(isDe ? `wäre sichtbar (${reasonParts.join(', ')}), aber explizit ausgeschlossen` : `would be visible (${reasonParts.join(', ')}), but explicitly excluded`)];
                      } else if (isExcluded) {
                        reasonParts.push(isDe ? '+ ausgeschlossen' : '+ excluded');
                      }
                      return (
                        <tr key={u.email} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                          <td style={{ padding: 6 }}>{u.displayName}</td>
                          <td style={{ padding: 6, color: 'var(--dex-gray-600)' }}>{u.email}</td>
                          <td style={{ padding: 6, color: 'var(--dex-gray-500)' }}>{u.location || '-'}</td>
                          <td style={{ padding: 6, textAlign: 'center' }}>
                            {visible
                              ? <span style={{ color: '#22c55e', fontWeight: 700, fontSize: '1.1rem' }}>&#10003;</span>
                              : <span style={{ color: '#ef4444', fontWeight: 700, fontSize: '1.1rem' }}>&#10007;</span>}
                          </td>
                          <td style={{ padding: 6, color: 'var(--dex-gray-600)', fontSize: '0.78rem' }}>
                            {reasonParts.join(', ')}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      )}

      {/* v18.33: Self-Check-in Erklär-Modal */}
      <Modal
        open={showSelfCheckInModal}
        onClose={() => setShowSelfCheckInModal(false)}
        maxWidth={560}
        ariaLabel={isDe ? 'Self-Check-in erklärt' : 'Self check-in explained'}
      >
        {showSelfCheckInModal && (
          <>
            <h3 style={{ marginTop: 0, marginBottom: 6, fontSize: '1.25rem', display: 'flex', alignItems: 'center', gap: 8 }}>
              <Icon iconName="QRCode" style={{ color: 'var(--dex-green, #86bc25)' }} />
              {isDe ? 'Self-Check-in per QR-Code' : 'Self check-in via QR code'}
            </h3>
            <p style={{ margin: '0 0 16px', fontSize: '0.9rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
              {isDe
                ? 'Teilnehmer checken sich am Veranstaltungstag selbst ein, indem sie einen QR-Code scannen — ganz ohne Anstehen am Check-in-Schalter. Du hast zwei Wege, den Code bereitzustellen:'
                : 'Attendees check themselves in on the event day by scanning a QR code — no queue at the check-in desk. You have two ways to provide the code:'}
            </p>

            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              <div style={{ padding: 14, borderRadius: 10, border: '1px solid var(--dex-gray-200, #eee)', background: 'var(--dex-gray-50, #f7f7f5)' }}>
                <div style={{ fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-gray-800, #333)', marginBottom: 4, display: 'flex', alignItems: 'center', gap: 6 }}>
                  <Icon iconName="PDF" style={{ color: 'var(--dex-green, #86bc25)' }} />
                  {isDe ? 'Druckbares QR-PDF (statisch)' : 'Printable QR PDF (static)'}
                </div>
                <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Bequem zum Aushängen oder Auslegen. Hinweis: ein abfotografierter Code lässt sich theoretisch weitergeben — deshalb am besten mit dem Zeitfenster „nur am Event-Tag" kombinieren. Jeder checkt ohnehin nur sich selbst ein.'
                    : 'Convenient for posting. Note: a photographed code could in theory be shared — best combined with the "event day only" time window. Everyone only checks themselves in anyway.'}
                </div>
              </div>
              <div style={{ padding: 14, borderRadius: 10, border: '1px solid var(--dex-gray-200, #eee)', background: 'var(--dex-gray-50, #f7f7f5)' }}>
                <div style={{ fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-gray-800, #333)', marginBottom: 4, display: 'flex', alignItems: 'center', gap: 6 }}>
                  <Icon iconName="Refresh" style={{ color: 'var(--dex-green, #86bc25)' }} />
                  {isDe ? 'Rotierende Live-Anzeige (foto-sicher)' : 'Rotating live display (photo-safe)'}
                </div>
                <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Läuft auf einem Bildschirm am Eingang (Laptop, Tablet, Beamer). Der QR-Code wechselt automatisch alle paar Sekunden — ein abfotografierter Code ist dadurch sofort wertlos. Die beste Wahl, wenn du sicher sein willst, dass nur Anwesende einchecken.'
                    : 'Runs on a screen at the entrance (laptop, tablet, projector). The QR code changes automatically every few seconds — a photographed code is instantly worthless. Best choice if you want to ensure only people on site check in.'}
                </div>
              </div>
            </div>

            <div style={{ marginTop: 16, padding: 12, borderRadius: 10, background: 'rgba(134,188,37,0.08)', border: '1px solid rgba(134,188,37,0.3)', fontSize: '0.85rem', color: 'var(--dex-gray-700, #444)', lineHeight: 1.5 }}>
              <Icon iconName="Camera" style={{ marginRight: 6, color: 'var(--dex-green-dark, #4a7c1f)' }} />
              {isDe
                ? 'Teilnehmer scannen mit der normalen Handy-Kamera (kein In-App-Scanner nötig) — das funktioniert zuverlässig auch in der SharePoint-App.'
                : 'Attendees scan with their normal phone camera (no in-app scanner needed) — this works reliably even inside the SharePoint app.'}
            </div>

            <div style={{ marginTop: 16, fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              {isDe
                ? 'QR-PDF und Live-Anzeige findest du nach dem Speichern im Admin Center dieses Events.'
                : 'You will find the QR PDF and live display in this event\'s admin center after saving.'}
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 20 }}>
              <button
                type="button"
                onClick={() => setShowSelfCheckInModal(false)}
                style={{ padding: '10px 20px', borderRadius: 10, border: 'none', background: 'var(--dex-green, #86bc25)', color: '#fff', fontWeight: 600, cursor: 'pointer' }}
              >
                {isDe ? 'Verstanden' : 'Got it'}
              </button>
            </div>
          </>
        )}
      </Modal>

      {/* v9.28/v13.4: Modal — neuer Quiz-Bereich anlegen, jetzt über <Modal>-Wrapper. */}
      <Modal
        open={newSectionModalOpen}
        onClose={() => setNewSectionModalOpen(false)}
        maxWidth={460}
        ariaLabel="Neuen Quiz-Bereich anlegen"
      >
        {newSectionModalOpen && (
          <>
            <h3 style={{ marginTop: 0, marginBottom: 8, fontSize: '1.15rem' }}>
              Neuen Bereich anlegen
            </h3>
            <p style={{ margin: '0 0 14px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              Bereiche bündeln Quiz-Fragen auf einer gemeinsamen Seite. Vergib einen
              kurzen, sprechenden Namen — z.B. <em>Orte</em>, <em>Geschichte</em> oder <em>Foto-Quiz</em>.
            </p>
            <input
              type="text"
              className="form-input"
              autoFocus
              value={newSectionName}
              placeholder='z.B. "Orte"'
              onChange={e => { setNewSectionName(e.target.value); setNewSectionError(''); }}
              onKeyDown={e => {
                if (e.key === 'Enter') {
                  const name = newSectionName.trim();
                  if (!name) { setNewSectionError('Bitte einen Namen eingeben.'); return; }
                  const existing = new Set<string>();
                  for (const q of quiz) if (q.section) existing.add(q.section);
                  for (const p of pendingSections) existing.add(p);
                  if (existing.has(name)) { setNewSectionError('Ein Bereich mit diesem Namen existiert bereits.'); return; }
                  setPendingSections([...pendingSections, name]);
                  setNewSectionModalOpen(false);
                } else if (e.key === 'Escape') {
                  setNewSectionModalOpen(false);
                }
              }}
              style={{ fontSize: '0.95rem', marginBottom: 8 }}
            />
            {newSectionError && (
              <div style={{ color: 'var(--dex-red, #c00)', fontSize: '0.78rem', marginBottom: 10 }}>{newSectionError}</div>
            )}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 12 }}>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => setNewSectionModalOpen(false)}
              >
                Abbrechen
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => {
                  const name = newSectionName.trim();
                  if (!name) { setNewSectionError('Bitte einen Namen eingeben.'); return; }
                  const existing = new Set<string>();
                  for (const q of quiz) if (q.section) existing.add(q.section);
                  for (const p of pendingSections) existing.add(p);
                  if (existing.has(name)) { setNewSectionError('Ein Bereich mit diesem Namen existiert bereits.'); return; }
                  setPendingSections([...pendingSections, name]);
                  setNewSectionModalOpen(false);
                }}
              >
                <Plus size={14} /> Bereich anlegen
              </button>
            </div>
          </>
        )}
      </Modal>

      {/* Modal: Vorgeschlagene Felder auswaehlen (Multi-Select) — v13.4 über <Modal>. */}
      <Modal
        open={showSuggestedModal}
        onClose={() => setShowSuggestedModal(false)}
        maxWidth={540}
        ariaLabel="Vorgeschlagene Felder auswählen"
      >
        {showSuggestedModal && (
          <div
            style={{
              display: 'flex', flexDirection: 'column', gap: 14,
            }}
          >
            <h2 style={{ margin: 0, fontSize: '1.25rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Vorgeschlagene Felder' : 'Suggested fields'}
            </h2>
            <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
              {isDe
                ? 'Wähle aus dem Katalog, welche Felder dem Event hinzugefügt werden sollen. Du kannst die Felder danach weiter anpassen.'
                : 'Pick the fields you want to add to the event. You can still tweak them afterwards.'}
            </p>
            {/* v10.21: Catalog gruppiert nach Kategorie. Allgemeine Felder
                immer ausgeklappt, B2Run-Felder default eingeklappt mit
                Toggle. Jeder Eintrag bekommt ein Badge mit der Kategorie. */}
            {(() => {
              const generalEntries = SUGGESTED_FIELDS_CATALOG.filter(s => s.category === 'general');
              const b2runEntries = SUGGESTED_FIELDS_CATALOG.filter(s => s.category === 'b2run');
              const renderEntry = (s: SuggestedEntry): React.ReactElement => (
                <label
                  key={s.key}
                  style={{
                    display: 'flex', gap: 10, alignItems: 'flex-start',
                    padding: '10px 12px', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 8, cursor: 'pointer',
                    background: suggestedSelection[s.key] ? 'var(--dex-gray-50, #fafafa)' : '#fff',
                  }}
                >
                  <input
                    type="checkbox"
                    checked={!!suggestedSelection[s.key]}
                    onChange={e => setSuggestedSelection({ ...suggestedSelection, [s.key]: e.target.checked })}
                    style={{ marginTop: 3, flexShrink: 0 }}
                  />
                  {/* v10.23: passendes Fluent-UI-Icon links neben dem Label,
                      damit die Auswahl auf einen Blick visuell wiedererkennbar
                      ist. Farbe analog zur Kategorie (gruen=Allgemein,
                      orange=B2Run). */}
                  <Icon
                    iconName={s.icon}
                    style={{
                      fontSize: 20, flexShrink: 0, marginTop: 2,
                      color: s.category === 'b2run' ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)',
                    }}
                  />
                  <span style={{ flex: 1 }}>
                    <span style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                      <strong style={{ fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>{s.label}</strong>
                      <span style={{
                        fontSize: '0.65rem', fontWeight: 600,
                        padding: '2px 8px', borderRadius: 999,
                        textTransform: 'uppercase', letterSpacing: 0.5,
                        background: s.category === 'b2run' ? 'rgba(237,139,0,0.12)' : 'rgba(134,188,37,0.12)',
                        color: s.category === 'b2run' ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)',
                      }}>
                        {s.category === 'b2run' ? 'B2Run' : (isDe ? 'Allgemein' : 'General')}
                      </span>
                      {/* v10.23: i-Tooltip mit ausfuehrlichem Hinweis was das
                          Feld in der App tut — verhindert Klick-und-Probier-
                          Modus, weil der Organizer schon vor Auswahl sieht
                          welche Frage-Form (Dropdown / Freitext / Pflicht-
                          Checkbox) und welcher Effekt (Anzeige im Admin-Center,
                          Excel-Export, etc.) entsteht. Klick auf das Label
                          (das `<label>`-Wrapping) wuerde die Checkbox togglen
                          — das `onClick`-stopPropagation des InfoTooltip
                          verhindert das. */}
                      <span onClick={e => e.preventDefault()} style={{ display: 'inline-flex' }}>
                        <InfoTooltip text={s.tooltip || s.description} />
                      </span>
                    </span>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>{s.description}</div>
                  </span>
                </label>
              );
              return (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 14, marginTop: 4 }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                    {generalEntries.map(renderEntry)}
                  </div>
                  <div style={{ borderTop: '1px solid var(--dex-gray-200)', paddingTop: 14 }}>
                    <button
                      type="button"
                      onClick={() => setShowB2runSuggested(v => !v)}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        background: 'none', border: 'none', padding: 0,
                        fontSize: '0.85rem', fontWeight: 600, cursor: 'pointer',
                        color: 'var(--dex-gray-700)',
                      }}
                    >
                      <span style={{ display: 'inline-flex', transform: showB2runSuggested ? 'rotate(90deg)' : 'rotate(0deg)', transition: 'transform 0.15s' }}>▶</span>
                      {isDe ? 'B2Run-spezifische Felder' : 'B2Run-specific fields'}
                      <span style={{
                        fontSize: '0.65rem', fontWeight: 600,
                        padding: '2px 8px', borderRadius: 999,
                        background: 'rgba(237,139,0,0.12)',
                        color: 'var(--dex-orange-dark, #b35a00)',
                      }}>
                        B2Run · {b2runEntries.length}
                      </span>
                    </button>
                    {showB2runSuggested && (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 10 }}>
                        <p style={{ margin: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)', lineHeight: 1.45 }}>
                          {isDe
                            ? 'Diese Felder sind speziell für B2Run-Lauf-Events vorgesehen (Startblock, Altersklasse, Datenschutz-Checkbox mit b2run.de-Links etc.). Bei normalen Events brauchst du sie nicht.'
                            : 'These fields are intended for B2Run running events (start block, age group, B2Run-specific privacy checkbox etc.). Skip them for standard events.'}
                        </p>
                        {b2runEntries.map(renderEntry)}
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}
            <div style={{ display: 'flex', gap: 8, justifyContent: 'space-between', marginTop: 6 }}>
              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  type="button"
                  className="btn btn-outline"
                  style={{ fontSize: '0.8rem', padding: '4px 12px' }}
                  onClick={() => {
                    const all: Record<string, boolean> = {};
                    for (const s of SUGGESTED_FIELDS_CATALOG) all[s.key] = true;
                    setSuggestedSelection(all);
                  }}
                >
                  {isDe ? 'Alle' : 'All'}
                </button>
                <button
                  type="button"
                  className="btn btn-outline"
                  style={{ fontSize: '0.8rem', padding: '4px 12px' }}
                  onClick={() => setSuggestedSelection({})}
                >
                  {isDe ? 'Keine' : 'None'}
                </button>
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setShowSuggestedModal(false)}
                >
                  {isDe ? 'Abbrechen' : 'Cancel'}
                </button>
                <button
                  type="button"
                  className="btn btn-primary"
                  onClick={addSelectedSuggestedFields}
                  disabled={!Object.values(suggestedSelection).some(Boolean)}
                >
                  {isDe ? 'Hinzufügen' : 'Add'}
                </button>
              </div>
            </div>
          </div>
        )}
      </Modal>

      {/* v11.57 / v11.63 / v13.4: Outlook-Update-Confirm-Modal über <Modal>-Wrapper.
          dismissable=false, da Schließen nur über Cancel-Button erlaubt. */}
      <Modal
        open={outlookConfirmOpen}
        onClose={cancelOutlookSave}
        maxWidth={620}
        dismissable={false}
        ariaLabel="Outlook-Update bestätigen"
      >
        {outlookConfirmOpen && (
          <div>
            <h2 id="outlook-confirm-title" style={{
              margin: '0 0 10px', fontSize: '1.15rem', fontWeight: 700,
              color: 'var(--dex-green-dark, #4a7c1f)',
            }}>
              {isDe ? 'Outlook-Termin der Teilnehmer aktualisieren?' : 'Update Outlook invite for attendees?'}
            </h2>
            <p style={{ margin: '0 0 14px', fontSize: '0.9rem', color: 'var(--dex-gray-700)', lineHeight: 1.55 }}>
              {isDe
                ? 'Du hast Felder geändert, die für die Teilnehmer-Outlook-Termine relevant sind. Wähle aus, welche Termine du jetzt neu rausschicken willst — der Rest wird gespeichert, aber Outlook bleibt unangetastet (du kannst das später jederzeit nachholen).'
                : 'You changed fields that are relevant to the attendees’ Outlook invites. Pick which invites you want to resend now — everything else is saved, but Outlook is left alone (you can resend later at any time).'}
            </p>
            <div style={{
              border: '1px solid var(--dex-gray-200)',
              borderRadius: 8,
              marginBottom: 14,
              background: 'var(--dex-gray-50, #f8f9fa)',
            }}>
              {outlookConfirmItems.map((it, idx) => {
                const isLast = idx === outlookConfirmItems.length - 1;
                const fieldLabelMap: Record<'title'|'startDate'|'endDate'|'outlookBody'|'location'|'subject'|'layout', { de: string; en: string }> = {
                  title: { de: 'Titel', en: 'Title' },
                  startDate: { de: 'Startzeit', en: 'Start time' },
                  endDate: { de: 'Endzeit', en: 'End time' },
                  outlookBody: { de: 'Termin-Text', en: 'Calendar body' },
                  location: { de: 'Ort', en: 'Location' },
                  subject: { de: 'Betreff', en: 'Subject' },
                  layout: { de: 'Kopfbild (Größe/Abstand)', en: 'Header image (size/spacing)' },
                };
                const changedLabels = it.changedFields.map(f => isDe ? fieldLabelMap[f].de : fieldLabelMap[f].en).join(', ');
                const checked = !!outlookConfirmChecks[it.eventId];
                // v11.69: noOutlookYet-Items bekommen wieder eine Checkbox.
                // Default UNCHECKED. Beim Anhaken wird das Sub-Event in der
                // Eventverwaltung komplett neu angelegt (DEX_Events-Item
                // delete + create mit `existingSubsiteUrl`), damit der
                // Outlook-Termin entsteht. Die bestehende Teilnehmerliste
                // mit allen Anmeldungen bleibt unangetastet.
                if (it.noOutlookYet) {
                  return (
                    <label
                      key={it.eventId}
                      style={{
                        display: 'flex', alignItems: 'flex-start', gap: 12,
                        padding: '12px 14px',
                        borderBottom: isLast ? 'none' : '1px solid var(--dex-gray-200)',
                        cursor: 'pointer',
                        background: '#fffaf0',
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={e => {
                          const next = e.target.checked;
                          setOutlookConfirmChecks(prev => ({ ...prev, [it.eventId]: next }));
                        }}
                        style={{ width: 18, height: 18, marginTop: 2, cursor: 'pointer', flexShrink: 0 }}
                      />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: '0.92rem', fontWeight: 600, color: 'var(--dex-gray-800)', wordBreak: 'break-word' }}>
                          {isDe ? `Sub-Event: ${it.title}` : `Sub-event: ${it.title}`}
                        </div>
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 3 }}>
                          {isDe ? 'Geändert: ' : 'Changed: '}{changedLabels}
                        </div>
                        <div style={{ fontSize: '0.76rem', color: '#8a6d3b', marginTop: 6, lineHeight: 1.45, background: '#fcf8e3', border: '1px solid #faebcc', borderRadius: 4, padding: '6px 8px' }}>
                          {isDe
                            ? <>Für dieses Sub-Event gibt es noch keinen Outlook-Termin. Wenn du den Haken setzt, wird das Sub-Event in der Eventverwaltung neu angelegt, damit der Outlook-Termin entsteht. <strong>Die bestehende Teilnehmerliste mit allen Anmeldungen bleibt erhalten</strong> — nur die DEX_Events-Zeile bekommt eine neue ID.</>
                            : <>This sub-event does not have an Outlook event yet. If you tick the box, the sub-event is re-created in the event admin so the Outlook event can be generated. <strong>The existing participant list with all registrations stays intact</strong> — only the DEX_Events row gets a new ID.</>}
                        </div>
                      </div>
                    </label>
                  );
                }
                // v15.3: leere changedFields-Liste = Item kommt aus dem
                // persistierten OutlookDirty-Flag (frühere Session,
                // wurde damals nicht synchronisiert). Klartext-Hinweis
                // statt leerer „Geändert:"-Zeile.
                const isFromPersistedDirty = it.changedFields.length === 0;
                return (
                  <label
                    key={it.eventId}
                    style={{
                      display: 'flex', alignItems: 'flex-start', gap: 12,
                      padding: '12px 14px',
                      borderBottom: isLast ? 'none' : '1px solid var(--dex-gray-200)',
                      cursor: 'pointer',
                      background: isFromPersistedDirty ? '#fff8e8' : '#fff',
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={e => {
                        const next = e.target.checked;
                        setOutlookConfirmChecks(prev => ({ ...prev, [it.eventId]: next }));
                      }}
                      style={{ width: 18, height: 18, marginTop: 2, cursor: 'pointer', flexShrink: 0 }}
                    />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '0.92rem', fontWeight: 600, color: 'var(--dex-gray-800)', wordBreak: 'break-word' }}>
                        {it.kind === 'top'
                          ? (isDe ? `Hauptevent: ${it.title}` : `Main event: ${it.title}`)
                          : (isDe ? `Sub-Event: ${it.title}` : `Sub-event: ${it.title}`)}
                      </div>
                      {isFromPersistedDirty ? (
                        <div style={{ fontSize: '0.78rem', color: '#8a6d3b', marginTop: 4, lineHeight: 1.45 }}>
                          {isDe
                            ? <>⏳ <strong>Frühere Änderung nicht synchronisiert</strong> — beim letzten Speichern dieses Events wurden Outlook-relevante Felder geändert, der Outlook-Sync wurde aber damals übersprungen. Haken setzen, um die Teilnehmer jetzt nachträglich per Outlook-Update zu informieren.</>
                            : <>⏳ <strong>Earlier change not yet synced</strong> — Outlook-relevant fields were changed in a previous save of this event, but the Outlook sync was skipped at the time. Tick the box to send the catch-up Outlook update to attendees now.</>}
                        </div>
                      ) : (
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 3 }}>
                          {isDe ? 'Geändert: ' : 'Changed: '}{changedLabels}
                        </div>
                      )}
                    </div>
                  </label>
                );
              })}
            </div>
            <p style={{
              margin: '0 0 12px', fontSize: '0.8rem', color: 'var(--dex-gray-500)',
              lineHeight: 1.5,
            }}>
              {isDe
                ? 'Bei angehakten Events bekommen die Teilnehmer eine „Aktualisierter Termin"-Benachrichtigung von Outlook. Nicht angehakte Termine werden für später als „ausstehender Outlook-Sync" markiert.'
                : 'Ticked events trigger an “updated meeting” notification from Outlook for attendees. Unticked invites are flagged as “pending Outlook sync” for later.'}
            </p>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 8 }}>
              <button
                className="btn btn-secondary"
                onClick={cancelOutlookSave}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                className="btn btn-primary"
                style={{ background: 'var(--dex-green, #86bc25)', borderColor: 'var(--dex-green, #86bc25)' }}
                onClick={() => confirmOutlookSave()}
              >
                {isDe ? 'Speichern' : 'Save'}
              </button>
            </div>
          </div>
        )}
      </Modal>

      {/* v17.3: Unsaved-Changes-Confirm-Modal. Erscheint, wenn der User
          auf „Zurueck" klickt und das Formular gegenueber dem Initial-
          Snapshot Aenderungen hat. */}
      {unsavedConfirmOpen && (
        <Modal
          open={true}
          onClose={() => { unsavedConfirmOpen.resolve(false); setUnsavedConfirmOpen(null); }}
          maxWidth={480}
          padding={24}
          ariaLabel={isDe ? 'Ungespeicherte Änderungen' : 'Unsaved changes'}
        >
          <h3 style={{ margin: '0 0 12px', fontSize: '1.1rem', color: 'var(--dex-orange-dark, #b35a00)' }}>
            {isDe ? 'Ungespeicherte Änderungen' : 'Unsaved changes'}
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: '0.9rem', lineHeight: 1.5, color: 'var(--dex-gray-700)' }}>
            {isDe
              ? <>Du hast Änderungen am Event vorgenommen, die noch <strong>nicht gespeichert</strong> sind. Wenn du jetzt zurückgehst, gehen sie verloren. Was möchtest du tun?</>
              : <>You have made changes to this event that are <strong>not saved yet</strong>. Going back now will discard them. What do you want to do?</>}
          </p>
          <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={() => { unsavedConfirmOpen.resolve(false); setUnsavedConfirmOpen(null); }}
              style={{ fontSize: '0.85rem' }}
            >
              {isDe ? 'Hier bleiben' : 'Stay here'}
            </button>
            <button
              type="button"
              className="btn btn-danger"
              onClick={() => { unsavedConfirmOpen.resolve(true); setUnsavedConfirmOpen(null); }}
              style={{ fontSize: '0.85rem' }}
            >
              {isDe ? 'Änderungen verwerfen' : 'Discard changes'}
            </button>
            {/* v17.7: Dritter Button — Speichern und zurueck zum Event.
                Wir blockieren die laufende Back-Nav (resolve(false)) und
                triggern attemptSubmit; nach erfolgreichem Save dispatched
                EventCreationPage selbst „dex-event-submit-success" und
                DexEventPlatform navigiert zum Organizer-Menue. */}
            <button
              type="button"
              className="btn btn-primary"
              onClick={() => {
                unsavedConfirmOpen.resolve(false);
                setUnsavedConfirmOpen(null);
                window.setTimeout(() => { attemptSubmit(); }, 0);
              }}
              style={{ fontSize: '0.85rem' }}
            >
              <Send size={14} /> {isDe ? 'Speichern und zurück zum Event' : 'Save and return to event'}
            </button>
          </div>
        </Modal>
      )}
    </div>
  );
}
