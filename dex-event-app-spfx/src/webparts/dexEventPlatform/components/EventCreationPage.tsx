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
import { EventService } from '../services/EventService';
import { eventCreatedEmail, buildOutlookBody, stripOutlookWrapper, parseOutlookHeadings, replacePlaceholders, getCachedOrbBase64 } from '../services/EmailTemplates';
import { EventType, AgendaItem } from '../types';
import { Trash2, Send, Plus, X, Users, Check } from './Icons';
import { RichText } from '@pnp/spfx-controls-react/lib/controls/richText';
import { HtmlEditorModal } from './HtmlEditorModal';
import { RegisterPreviewModal } from './RegisterPreviewModal';
import { InfoTooltip } from './InfoTooltip';
import BulkUserImportModal from './BulkUserImportModal';
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
  type: 'text' | 'select' | 'number' | 'checkbox' | 'user' | 'roommate';
  required: boolean;
  // Optionen als Array (incl. leerer Slots fuer "frisch hinzugefuegte" Eintraege)
  options: string[];
  visible: boolean;
  externalLinks?: Array<{ label: string; url: string }>;
  /** v7.11: Bei type=select erlaubt true Mehrfachauswahl (Checkbox-Liste statt
   *  Single-Dropdown). Wert wird " | "-getrennt gespeichert. */
  multi?: boolean;
  /** v7.20: Optionale Beschreibung — landet als "i"-Tooltip neben dem
   *  Feld-Label im Registrierungsformular. */
  helpText?: string;
  /** v7.21: Sichtbarkeitsbedingung — Feld nur anzeigen wenn das Quell-Feld
   *  einen der `values` als Antwort hat. */
  showIf?: { fieldId: string; values: string[] };
  /** v10.24: Bei aktiver Split-Capacity Feld nur fuer eine der zwei
   *  Gruppen sichtbar machen ('A' = Durchstarter / Gruppe A, 'B' =
   *  Funstarter / Gruppe B). 'all' / undefined = beide Gruppen. */
  onlyForGroup?: 'all' | 'A' | 'B';
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

export default function EventCreationPage(): React.ReactElement {
  const { goBack, selectedEventId, currentPage } = useNavigation();
  const { events, childEventsOf, createEvent, updateEvent, deleteEvent, refreshEvents } = useEvents();
  const { currentUser } = useCurrentUser();
  const { searchUsers, searchGroups, getGroupMembers, searchUsersByLocation } = useRoles();
  // Audience-Suche (Personen + Verteiler/Security-Groups)
  const [audienceSearch, setAudienceSearch] = React.useState('');
  const [audienceResults, setAudienceResults] = React.useState<Array<{ kind: 'user' | 'group'; email: string; displayName: string }>>([]);
  const [isSearchingAudience, setIsSearchingAudience] = React.useState(false);
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
  const [customFields, setCustomFields] = React.useState<CustomFieldInput[]>(
    editEvent ? editEvent.eventSpecificFields.map(f => ({
      id: f.id, label: f.label, type: f.type, required: f.required,
      options: f.options ? [...f.options] : [], visible: true,
      // v7.11: multi-Flag aus dem persistierten Feld uebernehmen, damit der
      // Editor den Status "Mehrfachauswahl erlauben" korrekt anzeigt.
      ...(f.multi ? { multi: true } : {}),
      // v7.20: helpText aus dem persistierten Feld uebernehmen.
      ...(f.helpText ? { helpText: f.helpText } : {}),
      // v7.21: showIf-Bedingung aus dem persistierten Feld uebernehmen.
      ...(f.showIf ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } } : {}),
      // v10.24: onlyForGroup-Constraint aus dem persistierten Feld uebernehmen.
      ...(f.onlyForGroup ? { onlyForGroup: f.onlyForGroup } : {}),
      // v11.15: externalLinks aus dem persistierten Feld uebernehmen
      // (z.B. AGB/Datenschutz-URLs auf b2run_datenschutz). Vorher
      // hat der Edit-Load die Links komplett gedroppt → beim ersten
      // Save eines bestehenden Events sind sie verloren gegangen.
      ...(f.externalLinks && f.externalLinks.length > 0 ? { externalLinks: f.externalLinks.map(x => ({ ...x })) } : {}),
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
  const [disableOutlook, setDisableOutlook] = React.useState(editEvent ? !!editEvent.disableOutlook : false);
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
  const [excludeResolvedUsers, setExcludeResolvedUsers] = React.useState<Array<{ email: string; displayName: string; firstName: string; lastName: string; jobTitle: string; location: string; source: string }>>([]);
  const [excludeResolving, setExcludeResolving] = React.useState(false);
  const [excludeSearch, setExcludeSearch] = React.useState('');
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
  // Nur im Edit-Modus: standardmaessig wird der Outlook-Termin NICHT angefasst,
  // damit bei kleinen Aenderungen (z.B. Description) nicht unnoetig eine
  // "Updated meeting"-Benachrichtigung an alle Teilnehmer geht. Der Organizer
  // muss die Checkbox aktiv setzen wenn er moechte dass Titel/Start/Ende im
  // Outlook-Termin aktualisiert werden.
  const [triggerOutlookUpdate, setTriggerOutlookUpdate] = React.useState(false);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [emailTemplates, setEmailTemplates] = React.useState<Array<{ id: number; templateType: string; language: string; subject: string; heading: string; headingColor: string; bodyHtml: string }>>([]);
  const [emailTemplateOverrides, setEmailTemplateOverrides] = React.useState<Record<string, { subject: string; heading: string; bodyHtml: string }>>(
    editEvent?.emailTemplateOverrides ? (() => {
      try {
        const parsed = JSON.parse(editEvent.emailTemplateOverrides);
        // _eventLogo wird separat ueber emailLogoPreview-State verwaltet —
        // nicht in emailTemplateOverrides doppelt halten (sonst laesst sich das
        // Custom-Logo nach Remove nicht mehr aus SP entfernen).
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const { _eventLogo, ...rest } = parsed;
        return rest;
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
  const [dragFieldId, setDragFieldId] = React.useState<string | null>(null);
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
    disableOutlook?: boolean;
    /** Per-Sub-Event Custom-Fields (v10.11+). Ersetzt die hardcoded Funstarter/
     *  Durchstarter-Frage bei B2Run — wer eine zusätzliche Auswahl-Frage pro
     *  Sub-Event will, definiert sie hier individuell. Default: leeres Array
     *  (= Sub-Event ohne zusätzliche Frage, Teilnehmer wählen nur den Termin).
     *  Wird unabhängig vom Hauptevent-customFields gespeichert; mit dem Button
     *  „Vom Hauptevent kopieren" kann der Organizer die Felder duplizieren. */
    customFields?: CustomFieldInput[];
  }
  const [subEvents, setSubEvents] = React.useState<SubEventDraft[]>(() => {
    if (!editEvent) return [];
    const kids = childEventsOf(editEvent.id);
    return kids.map(k => ({
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
      disableOutlook: k.disableOutlook,
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
        showIf: f.showIf,
      })),
    }));
  });
  // Snapshot der beim Edit-Start vorhandenen Sub-Event-DB-IDs, um beim Save
  // entfernte Sub-Events zu löschen.
  const [initialSubEventDbIds] = React.useState<string[]>(() => {
    if (!editEvent) return [];
    return childEventsOf(editEvent.id).map(k => k.id);
  });
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
      if (typeof init === 'number' && init >= 0 && init <= 6) return init;
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
  const [isSearchingEmails, setIsSearchingEmails] = React.useState(false);
  const [submitted, setSubmitted] = React.useState(false);
  const [error, setError] = React.useState('');
  const [imageUploadError, setImageUploadError] = React.useState('');
  const [iconPickerOpen, setIconPickerOpen] = React.useState<string | null>(null);
  const [iconSearch, setIconSearch] = React.useState('');
  const [showAllIcons, setShowAllIcons] = React.useState(false);

  const locationOptions = ['Berlin', 'Dresden', 'Düsseldorf', 'Frankfurt', 'Görlitz', 'Halle', 'Hamburg', 'Hannover', 'Köln', 'Leipzig', 'Magdeburg', 'Mannheim', 'München', 'Nürnberg', 'Stuttgart', 'Walldorf', 'All'];

  const addCustomField = (): void => {
    setCustomFields([...customFields, {
      id: `cf-${Date.now()}`, label: '', type: 'text',
      required: false, options: [], visible: true,
    }]);
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

  const fillDemo = (): void => {
    const now = new Date();
    // Demo-Events sollen standardmaessig am Wochenende stattfinden:
    // Naechster Samstag ab now+7 Tage (getDay() === 6 = Samstag).
    const eventStart = new Date(now);
    eventStart.setDate(now.getDate() + 7);
    while (eventStart.getDay() !== 6) {
      eventStart.setDate(eventStart.getDate() + 1);
    }
    eventStart.setHours(9, 0, 0, 0); // Samstag 09:00
    // Event laeuft bis Sonntag 17:00 (zwei-Tages-Weekend-Event)
    const eventEnd = new Date(eventStart);
    eventEnd.setDate(eventStart.getDate() + 1);
    eventEnd.setHours(17, 0, 0, 0);
    const deadline = new Date(eventStart);
    deadline.setDate(eventStart.getDate() - 2);
    const lastDereg = new Date(eventStart);
    lastDereg.setDate(eventStart.getDate() - 1);

    const toDatetime = (d: Date): string => {
      const pad = (n: number): string => (n < 10 ? '0' : '') + n;
      return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
    };
    const toDate = (d: Date): string => {
      const pad = (n: number): string => (n < 10 ? '0' : '') + n;
      return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    };

    const dateStr = toDate(eventStart).replace(/-/g, '');
    setTitle(`Test_${dateStr}`);
    setDescription('Testbeschreibung für ein Demo-Event.');
    setLocation('Köln, Testort');
    setLocationFilter('');
    setAudience('All');
    setStartDate(toDatetime(eventStart));
    setEndDate(toDatetime(eventEnd));
    setRegistrationDeadline(toDate(deadline));
    setLastDeregisterDate(toDate(lastDereg));
    // v10.26: Demo-Event aktiviert geteilte Kapazität als Vorlage. Zwei
    // Gruppen á 25 Plätze, gemeinsame Warteliste deaktiviert (= getrennt,
    // damit der typ-bewusste Promote-Pfad als Demo durchläuft).
    setMaxParticipants('50');
    setUseSplitCapacities(true);
    setSplitLabelA('Teilnehmergruppe 1');
    setSplitLabelB('Teilnehmergruppe 2');
    setDurchstarterCapacity('25');
    setFunstarterCapacity('25');
    setSplitSharedWaitlist(false);
    setWaitlistEnabled(true);
    setEventImageUrl('');
    // v10.26: Ansprechpartner als Demo befüllt — der Organizer sieht
    // direkt wie das Feld auf der Anmelde-Seite gerendert wird.
    setContactName('Heike Musterfrau');
    setContactEmail('Musterfirma@online.de');
    setContactInfo('Bei Rückfragen zu Anmeldung oder Ablauf direkt melden.');
    // v10.26: Custom-Fields. Zusätzlich zu den Standard-Beispielen kommt
    // eine Pflicht-Checkbox die NUR für Teilnehmergruppe 1 sichtbar ist
    // (onlyForGroup='A') — zeigt den neuen Pro-Gruppe-Field-Mechanismus
    // direkt am Demo-Event.
    // v11.3: deterministic IDs damit das roommate-Feld eine showIf-Referenz
    // auf das zimmerart-Feld bauen kann (showIf-Filter-Vergleich nutzt
    // exakte fieldId-Strings, daher muss die ID stabil sein).
    const tDemo = Date.now();
    const idTshirt = `cf-${tDemo}`;
    const idAllergien = `cf-${tDemo + 1}`;
    const idDiet = `cf-${tDemo + 2}`;
    const idHotel = `cf-${tDemo + 3}`;
    const idZimmerart = `cf-${tDemo + 4}`;
    const idRoommate = `cf-${tDemo + 5}`;
    const idGroup1 = `cf-${tDemo + 6}`;
    setCustomFields([
      { id: idTshirt, label: 'T-Shirt Größe', type: 'select', required: false, options: ['Habe bereits ein T-Shirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true },
      { id: idAllergien, label: 'Allergien', type: 'text', required: false, options: [], visible: true },
      { id: idDiet, label: 'Essenspräferenzen', type: 'select', required: false, options: ['Keine Präferenzen', 'Vegetarisch', 'Vegan', 'Pescetarisch'], visible: true },
      { id: idHotel, label: 'Hotel benötigt', type: 'checkbox', required: false, options: [], visible: true },
      { id: idZimmerart, label: 'Zimmerart (falls Hotel benötigt)', type: 'select', required: false, options: ['Keine Präferenz', 'Einzelzimmer', 'Doppelzimmer'], visible: true },
      // v11.3: Roommate-Feld nur sichtbar wenn Zimmerart='Doppelzimmer' —
      // verhindert dass Einzelzimmer-Bucher versehentlich einen Roommate
      // angeben. Demonstriert direkt den showIf-Mechanismus aus v7.21.
      { id: idRoommate, label: 'Bevorzugter Zimmerpartner (bei Doppelzimmer)', type: 'roommate', required: false, options: [], visible: true,
        showIf: { fieldId: idZimmerart, values: ['Doppelzimmer'] } },
      // Pflicht-Checkbox nur für Gruppe 1 (Demo für onlyForGroup='A')
      { id: idGroup1, label: 'Zustimmung Sonderkonditionen Gruppe 1', type: 'checkbox', required: true, options: [], visible: true, onlyForGroup: 'A' },
    ]);
    // v10.26: Zwei Sub-Events anlegen — Subevent 1 mit Dropdown-Frage,
    // Subevent 2 mit Freitext-Frage. Damit sieht der Organizer beim
    // Demo-Klick sofort, wie die Sub-Event-Struktur + Sub-Event-spezifische
    // Custom-Fields zusammenspielen.
    const subStart1 = new Date(eventStart);
    subStart1.setHours(10, 0, 0, 0);
    const subEnd1 = new Date(eventStart);
    subEnd1.setHours(12, 0, 0, 0);
    const subStart2 = new Date(eventStart);
    subStart2.setDate(eventStart.getDate() + 1);
    subStart2.setHours(13, 0, 0, 0);
    const subEnd2 = new Date(eventStart);
    subEnd2.setDate(eventStart.getDate() + 1);
    subEnd2.setHours(15, 0, 0, 0);
    // v10.26: Demo-Agenda mit drei Programmpunkten am ersten Tag
    const dateIso = toDate(eventStart);
    setAgenda([
      { id: `ag-${Date.now()}`, date: dateIso, time: '09:00', endTime: '10:00', icon: 'Hello', title: 'Begrüßung & Kennenlernen', description: 'Kurzer Auftakt mit Vorstellung der Tagesziele.' },
      { id: `ag-${Date.now() + 1}`, date: dateIso, time: '10:15', endTime: '12:00', icon: 'Lightbulb', title: 'Workshop-Block', description: 'Interaktive Sessions in Kleingruppen.' },
      { id: `ag-${Date.now() + 2}`, date: dateIso, time: '12:00', endTime: '13:30', icon: 'EatDrink', title: 'Mittagessen', description: 'Stärkung und Networking.' },
    ]);
    // v10.26: Demo-Transferzeiten — Bus von Frankfurt zum Veranstaltungsort
    setTransferTimes([
      {
        id: `tt-${Date.now()}`,
        location: 'Frankfurt am Main',
        meetingPoint: 'Hauptbahnhof, Gleis 1',
        address: 'Im Hauptbahnhof, 60329 Frankfurt am Main',
        date: dateIso,
        departureTime: '07:30',
        arrivalTime: '08:45',
        description: 'Reisebus-Transfer zum Veranstaltungsort.',
      },
    ]);
    setSubEvents([
      {
        id: `se-${Date.now()}`,
        title: 'Subevent 1',
        startDate: toDatetime(subStart1),
        endDate: toDatetime(subEnd1),
        registrationDeadline: '',
        location: 'Köln, Testort — Raum A',
        description: 'Subevent 1 mit Auswahlfrage zum Format.',
        maxParticipants: 30,
        disableEmails: false,
        disableOutlook: false,
        customFields: [
          {
            id: `sef-${Date.now()}`,
            label: 'Welches Format passt für dich?',
            type: 'select',
            required: false,
            options: ['Workshop (interaktiv)', 'Vortrag (Frontal)', 'Roundtable'],
            visible: true,
          },
        ],
      },
      {
        id: `se-${Date.now() + 1}`,
        title: 'Subevent 2',
        startDate: toDatetime(subStart2),
        endDate: toDatetime(subEnd2),
        registrationDeadline: '',
        location: 'Köln, Testort — Raum B',
        description: 'Subevent 2 mit Freitext-Frage.',
        maxParticipants: 30,
        disableEmails: false,
        disableOutlook: false,
        customFields: [
          {
            id: `sef-${Date.now() + 1}`,
            label: 'Welches Thema möchtest du gerne vertieft besprechen?',
            type: 'text',
            required: false,
            options: [],
            visible: true,
          },
        ],
      },
    ]);
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
    // Sub-Events erben Organizer + OrganizerEmail vom Parent. Einmal sanitisieren
    // statt pro Iteration, identisch für alle Children.
    const sanitizedOrgPair = sanitizeOrganizerPairs();
    for (const draft of subEvents) {
      if (!draft.title || !draft.title.trim()) continue; // leere Drafts ignorieren
      const childPayload = {
        title: draft.title.trim(),
        type: 'Other',
        status: 'Active',
        description: draft.description || '',
        location: draft.location || '',
        locationAddress: '',
        locationFilter: locationFilter,
        audience: audience,
        filterMode: filterMode,
        startDate: draft.startDate || '',
        endDate: draft.endDate || '',
        registrationDeadline: draft.registrationDeadline || '',
        lastDeregisterDate: '',
        maxParticipants: draft.maxParticipants || 0,
        waitlistEnabled: true,
        eventImageUrl: '',
        organizer: sanitizedOrgPair.orgString,
        organizerEmail: sanitizedOrgPair.orgEmailString,
        outlookEventId: '',
        outlookBody: '',
        agenda: '[]',
        transfers: '[]',
        documents: '[]',
        funZone: '[]',
        quizClusterSize: 1,
        emailLanguage: emailLanguage,
        emailTemplateOverrides: '',
        disableEmails: !!draft.disableEmails,
        disableOutlook: !!draft.disableOutlook,
        isFictive: isFictive,
        customFields: draft.customFields || [],
        parentEventId: parentEventId,
      };
      if (draft.dbId) {
        keptDbIds.add(draft.dbId);
        // Update bestehender Sub-Event: nur geänderte Felder patchen. CustomFields
        // werden als JSON-String serialisiert (gleiche Form wie bei regulären
        // Events; siehe handleSubmit-Edit-Pfad).
        const cfJson = JSON.stringify((draft.customFields || [])
          .filter(f => f.label && f.label.trim().length > 0)
          .map(f => ({
            id: f.id,
            label: f.label.trim(),
            type: f.type,
            required: f.required,
            visible: f.visible,
            ...(f.helpText && f.helpText.trim() ? { helpText: f.helpText.trim() } : {}),
            ...(f.showIf && f.showIf.fieldId && f.showIf.values && f.showIf.values.length > 0
              ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } }
              : {}),
            ...(f.type === 'select' ? { options: f.options.map(o => o.trim()).filter(Boolean), ...(f.multi ? { multi: true } : {}) } : {}),
            ...(f.onlyForGroup && f.onlyForGroup !== 'all' ? { onlyForGroup: f.onlyForGroup } : {}),
            // v11.15: externalLinks (AGB/Datenschutz-URLs etc.) beim Save
            // mit-persistieren — vorher haben alle drei Persist-Pfade
            // (Edit-Save, Create-Save, Sub-Event-Save) sie gedroppt.
            ...(f.externalLinks && f.externalLinks.length > 0
              ? { externalLinks: f.externalLinks.map(x => ({ label: x.label, url: x.url })) }
              : {}),
          })));
        await updateEvent(draft.dbId, {
          'Title': childPayload.title,
          'Description': childPayload.description,
          'Location': childPayload.location,
          'StartDate': childPayload.startDate || null,
          'EndDate': childPayload.endDate || null,
          'RegistrationDeadline': childPayload.registrationDeadline || null,
          'MaxParticipants': childPayload.maxParticipants,
          'DisableEmails': childPayload.disableEmails,
          'DisableOutlook': childPayload.disableOutlook,
          'CustomFields': cfJson,
        });
      } else {
        await createEvent(childPayload);
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
      // v11.17: Debug-Trace fuer helpText-Save-Pfad. Maintainer hat
      // berichtet, dass eingegebene Beschreibungen nach dem Speichern
      // wieder weg sind. Diese Log-Zeile zeigt im Browser-DevTools
      // direkt, ob die customFields in dem Moment des Save-Klicks die
      // helpText-Eintraege haben — wenn ja, ist die Persistierung das
      // Problem; wenn nein, faengt der State sie nicht.
      // eslint-disable-next-line no-console
      console.log('[DEX][edit-save] customFields state at save:', customFields.map(f => ({
        id: f.id, label: f.label, helpText: f.helpText, onlyForGroup: f.onlyForGroup, showIf: f.showIf,
      })));
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
        'LocationFilter': locationFilter,
        'Audience': audience,
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
        'CustomFields': JSON.stringify(customFields
          .filter(f => f.label && f.label.trim().length > 0)
          .map(f => ({
            id: f.id, label: f.label.trim(), type: f.type, required: f.required, visible: f.visible,
            ...(f.helpText && f.helpText.trim() ? { helpText: f.helpText.trim() } : {}),
            ...(f.showIf && f.showIf.fieldId && f.showIf.values && f.showIf.values.length > 0
              ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } }
              : {}),
            ...(f.type === 'select' ? { options: f.options.map(o => o.trim()).filter(Boolean), ...(f.multi ? { multi: true } : {}) } : {}),
            ...(f.onlyForGroup && f.onlyForGroup !== 'all' ? { onlyForGroup: f.onlyForGroup } : {}),
            // v11.15: externalLinks (AGB/Datenschutz-URLs etc.) beim Save
            // mit-persistieren — vorher haben alle drei Persist-Pfade
            // (Edit-Save, Create-Save, Sub-Event-Save) sie gedroppt.
            ...(f.externalLinks && f.externalLinks.length > 0
              ? { externalLinks: f.externalLinks.map(x => ({ label: x.label, url: x.url })) }
              : {}),
          }))),
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
      const defaultOutlookBody = emailLanguage === 'EN'
        ? `<p>You are registered for the event <strong>${escHtml(title)}</strong>.</p>`
          + `<p>If you are unable to attend, please cancel your registration in time via the <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">Event Experience Platform</a> (&bdquo;My Events&ldquo;).</p>`
          + `<p>For organizational questions please contact <strong>${escHtml(orgNames || 'the organizer')}</strong>.</p>`
        : `<p>Du bist für das Event <strong>${escHtml(title)}</strong> angemeldet.</p>`
          + `<p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">Event Experience Platform</a> (&bdquo;Meine Events&ldquo;) ab.</p>`
          + `<p>Bei organisatorischen Fragen wende dich bitte an <strong>${escHtml(orgNames || 'den Organizer')}</strong>.</p>`;
      const resolvedBody = outlookBody
        ? replacePlaceholders(outlookBody, outlookVars)
        : defaultOutlookBody;
      const resolvedOlHeading = outlookHeading ? replacePlaceholders(outlookHeading, outlookVars) : title;
      const resolvedOlSub = outlookSubheading ? replacePlaceholders(outlookSubheading, outlookVars) : undefined;
      const wrappedOutlook = buildOutlookBody(resolvedOlHeading, resolvedBody, resolvedOlSub);
      // Outlook-spezifisches Logo direkt in den Body einbetten (Flow ersetzt {{ORB_URL}}
      // nur mit EmailImageBase64 — das ist fuer Mails. Fuer Outlook haben wir ein eigenes
      // Logo, also hier resolven).
      // v7.5: Fallback auf den globalen DEX-Orb-Cache, wenn der Organizer kein
      // event-spezifisches Outlook-Logo hochgeladen hat — sonst rendert
      // <img src=""> ein leeres Bildkasten und der Termin sieht halbfertig aus.
      updates['OutlookBody'] = wrappedOutlook.replace(/\{\{ORB_URL\}\}/g, outlookLogoPreview || getCachedOrbBase64() || '');
      updates['Agenda'] = JSON.stringify(agenda);
      updates['Transfers'] = JSON.stringify(transferTimes);
      updates['FunZone'] = JSON.stringify(quiz);
      updates['QuizClusterSize'] = Math.min(Math.max(1, quizClusterSize || 1), 4);
      updates['EmailLanguage'] = emailLanguage;
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
      updates['EmailTemplateOverrides'] = (Object.keys(emailTemplateOverrides).length > 0 || emailLogoPreview || outlookLogoPreview || Object.keys(b2runExtraConfig).length > 0 || Object.keys(qrScannerConfig).length > 0 || Object.keys(coOrganizerConfig).length > 0 || Object.keys(testTeamConfig).length > 0)
        ? JSON.stringify({
            ...(emailLogoPreview ? { _eventLogo: emailLogoPreview } : {}),
            ...(outlookLogoPreview ? { _outlookLogo: outlookLogoPreview } : {}),
            ...b2runExtraConfig,
            ...qrScannerConfig,
            ...coOrganizerConfig,
            ...testTeamConfig,
            ...emailTemplateOverrides,
          })
        : '';
      // v9.21: ActiveFrom als SP-DateTime
      updates['ActiveFrom'] = activeFrom ? new Date(activeFrom).toISOString() : null;
      // Custom-Mail-Logo in EmailImageBase64 (SP-Spalte) — der Flow ersetzt
      // {{ORB_URL}} in Mails damit. Wenn leer: Flow faellt auf _Config
      // DefaultImageBase64 (DEX-Orb) zurueck.
      updates['EmailImageBase64'] = emailLogoPreview || '';
      updates['DisableEmails'] = disableEmails;
      updates['DisableOutlook'] = disableOutlook;
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

      setProgress(50);

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

      const success = await updateEvent(selectedEventId, updates);
      if (success) {
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

        // Sub-Events persistieren (create/update/delete pro Draft). Seit v6.4.
        try { await persistSubEventsForParent(selectedEventId); }
        catch (err) { console.warn('[DEX] Sub-Events persistieren fehlgeschlagen:', err); }

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
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const merged = cfForFix.map((f: any) => {
                const sp = fixResult.customFieldMap![f.id] || f.spInternalName || '';
                return { ...f, spInternalName: sp };
              });
              await updateEvent(selectedEventId, { 'CustomFields': JSON.stringify(merged) });
            }
          }
        } catch (err) { console.warn('[DEX] Auto-fix Teilnehmer-Columns fehlgeschlagen:', err); }

        // Bild als Attachment hochladen (falls neues Bild gewaehlt wurde)
        if (imageFile) {
          try {
            setProgressLabel('Bild wird hochgeladen...');
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
        // Outlook-Termin Update triggern — NUR wenn der Organizer das explizit
        // per Checkbox angefordert hat. Grund: auch kleine Aenderungen (z.B.
        // Description-Tippfix) loesen sonst eine "Updated meeting"-Mail an ALLE
        // Teilnehmer aus, weil der Flow DEX_Outlook_Einladungen PATCH auf den
        // Outlook-Termin macht und Outlook automatisch benachrichtigt.
        if (!disableOutlook && triggerOutlookUpdate) {
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ctx = (window as any).__dexSpfxContext;
            if (ctx && editEvent?.subsiteUrl) {
              const svc = new EventService(ctx);
              await svc.queueOutlookEvent('', selectedEventId, title, 'UpdateEvent');
            }
          } catch { /* Outlook-Update optional */ }
          // Sub-Event-Outlook-Updates laufen bereits über persistSubEventsForParent
          // (updateEvent auf jeden Draft mit dbId) — der bestehende DEX_Outlook_Einladungen-
          // Flow für Parent-Events wird pro Sub-Event ebenfalls getriggert, weil ein
          // Sub-Event dieselbe Updateevent-Logik nutzt wie ein Top-Level-Event.
        }
        setProgress(100);
        setProgressLabel('Änderungen gespeichert!');
        // v9.45: Soft-Refresh analog zum Create-Pfad. Wizard verlassen via
        // CustomEvent, DexEventPlatform navigiert + zeigt Banner. Refresh wird
        // beim Update sofort getriggert (loadEvents in updateEvent hat schon
        // gefeuert) — kein delayed refresh nötig wie beim Create.
        try {
          window.dispatchEvent(new CustomEvent('dex-event-submit-success', {
            detail: { title: title, eventId: String(selectedEventId), type: 'update' as const },
          }));
        } catch { /* */ }
        setIsSubmitting(false);
      } else {
        setIsSubmitting(false);
        setProgress(0);
        setError('Event konnte nicht aktualisiert werden.');
      }
    } else {
      // Neues Event erstellen – Progress-Animation parallel laufen lassen
      try {
      setProgressLabel('Subsite wird erstellt...');
      const progressSteps = [
        { at: 20, label: 'Subsite wird erstellt...' },
        { at: 35, label: 'Teilnehmerliste wird angelegt...' },
        { at: 50, label: 'Spalten werden konfiguriert...' },
        { at: 65, label: 'Berechtigungen werden gesetzt...' },
        { at: 80, label: 'Event wird gespeichert...' },
        { at: 90, label: 'Fast fertig...' },
      ];
      let stepIdx = 0;
      const progressTimer = setInterval(() => {
        if (stepIdx < progressSteps.length) {
          setProgress(progressSteps[stepIdx].at);
          setProgressLabel(progressSteps[stepIdx].label);
          stepIdx++;
        }
      }, 2000);

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
        locationFilter,
        audience,
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
          const defaultBody = emailLanguage === 'EN'
            ? `<p>You are registered for the event <strong>${escHtml(title)}</strong>.</p>`
              + `<p>If you are unable to attend, please cancel your registration in time via the <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">Event Experience Platform</a> (&bdquo;My Events&ldquo;).</p>`
              + `<p>For organizational questions please contact <strong>${escHtml(orgNames || 'the organizer')}</strong>.</p>`
            : `<p>Du bist für das Event <strong>${escHtml(title)}</strong> angemeldet.</p>`
              + `<p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">Event Experience Platform</a> (&bdquo;Meine Events&ldquo;) ab.</p>`
              + `<p>Bei organisatorischen Fragen wende dich bitte an <strong>${escHtml(orgNames || 'den Organizer')}</strong>.</p>`;
          const resolvedBody = outlookBody ? replacePlaceholders(outlookBody, vars) : defaultBody;
          const resolvedHeading = outlookHeading ? replacePlaceholders(outlookHeading, vars) : title;
          const resolvedSub = outlookSubheading ? replacePlaceholders(outlookSubheading, vars) : undefined;
          const wrapped = buildOutlookBody(resolvedHeading, resolvedBody, resolvedSub);
          // v7.5: gleicher Fallback wie im Update-Pfad — globaler DEX-Orb,
          // damit der Termin nie ohne Hero-Bild rauskommt.
          return wrapped.replace(/\{\{ORB_URL\}\}/g, outlookLogoPreview || getCachedOrbBase64() || '');
        })(),
        agenda: JSON.stringify(agenda),
        transfers: JSON.stringify(transferTimes),
        documents: '[]', // Dokumente werden nach erfolgreichem Upload gespeichert
        funZone: JSON.stringify(quiz),
        quizClusterSize: Math.min(Math.max(1, quizClusterSize || 1), 4),
        emailLanguage,
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
          const hasAny = Object.keys(emailTemplateOverrides).length > 0 || emailLogoPreview || outlookLogoPreview || Object.keys(b2runExtra).length > 0 || Object.keys(qrExtra).length > 0 || Object.keys(coExtra).length > 0 || Object.keys(ttExtra).length > 0;
          return hasAny
            ? JSON.stringify({
                ...(emailLogoPreview ? { _eventLogo: emailLogoPreview } : {}),
                ...(outlookLogoPreview ? { _outlookLogo: outlookLogoPreview } : {}),
                ...b2runExtra,
                ...qrExtra,
                ...coExtra,
                ...ttExtra,
                ...emailTemplateOverrides,
              })
            : '';
        })(),
        disableEmails,
        disableOutlook,
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
        customFields: customFields
          .filter(f => f.label && f.label.trim().length > 0)
          .map(f => ({
            id: f.id, label: f.label.trim(), type: f.type, required: f.required, visible: f.visible,
            ...(f.helpText && f.helpText.trim() ? { helpText: f.helpText.trim() } : {}),
            ...(f.showIf && f.showIf.fieldId && f.showIf.values && f.showIf.values.length > 0
              ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } }
              : {}),
            ...(f.type === 'select' ? { options: f.options.map(o => o.trim()).filter(Boolean), ...(f.multi ? { multi: true } : {}) } : {}),
            ...(f.onlyForGroup && f.onlyForGroup !== 'all' ? { onlyForGroup: f.onlyForGroup } : {}),
            // v11.15: externalLinks (AGB/Datenschutz-URLs etc.) beim Save
            // mit-persistieren — vorher haben alle drei Persist-Pfade
            // (Edit-Save, Create-Save, Sub-Event-Save) sie gedroppt.
            ...(f.externalLinks && f.externalLinks.length > 0
              ? { externalLinks: f.externalLinks.map(x => ({ label: x.label, url: x.url })) }
              : {}),
          })),
      });

      clearInterval(progressTimer);
      if (eventId) {
        // v11.4: nicht direkt auf 100% — der ganze Block darunter
        // (Sub-Events, Dokumente, Event-Bild, Organizer-Mails) läuft noch.
        // Progress wird Schritt für Schritt erhöht und bei 2288 final auf
        // 100% gesetzt.
        setProgress(92);
        setProgressLabel(isDe ? 'Sub-Events werden angelegt...' : 'Creating sub-events...');
        try { await persistSubEventsForParent(String(eventId)); }
        catch (err) { console.warn('[DEX] Sub-Events beim Create persistieren fehlgeschlagen:', err); }
        setProgress(95);
        setProgressLabel(isDe ? 'Dokumente und Bild werden hochgeladen...' : 'Uploading documents and image...');
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
          window.dispatchEvent(new CustomEvent('dex-event-submit-success', {
            detail: { title: title, eventId: String(eventId), type: 'create' as const },
          }));
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

  // Templates laden wenn Step 4 (Kommunikation) erreicht wird
  // WICHTIG: Dieser useEffect MUSS vor dem early return (if submitted) stehen,
  // da React die gleiche Anzahl Hooks bei jedem Render erwartet (Rules of Hooks).
  React.useEffect(() => {
    if (currentStep === 4 && emailTemplates.length === 0) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (ctx) {
        const svc = new EventService(ctx);
        svc.getAllEmailTemplates().then(setEmailTemplates).catch(() => { /* Templates nicht verfuegbar */ });
      }
    }
  }, [currentStep]);

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
      'Veranstaltungsort und Adresse erfassen',
      'Start- und End-Datum (mit Uhrzeit) festlegen',
      'Anmeldefrist setzen — nach diesem Datum keine neuen Registrierungen mehr',
      'Optional: Letzte Abmeldemöglichkeit — danach können sich Teilnehmer nicht mehr selbst abmelden',
    ],
    [
      'Maximale Teilnehmerzahl festlegen (oder Unbegrenzt)',
      'Warteliste aktivieren — voll besetzte Events nehmen weitere Anmeldungen auf, bis ein Platz frei wird',
      'Optional: Geteilte Kapazität — zwei frei benannte Gruppen mit eigener Platzzahl + eigener oder gemeinsamer Warteliste',
      'Pflichtfelder pro Gruppe in Schritt 4 (z.B. Leistungsnachweis nur für eine bestimmte Gruppe)',
    ],
    [
      'Feldtyp wählen: Text, Zahl, Dropdown, Checkbox, Personen-Suche oder Roommate (Doppelzimmer)',
      'Mehrfachauswahl bei Dropdowns (z.B. mehrere Allergien anhaken)',
      'Pflichtfeld setzen (rotes Sternchen, Anmeldung blockiert wenn leer)',
      'Beschreibung pro Feld — landet als „i"-Tooltip neben dem Feld-Label',
      'Sichtbarkeitsbedingung: Feld nur dann anzeigen wenn eine andere Frage einen bestimmten Wert hat (z.B. „Zimmerart nur fragen wenn Hotel = ja")',
      'Reihenfolge per Drag oder Pfeilen — die Nummerierung passt sich automatisch an',
    ],
    [
      'E-Mail-Sprache (DE/EN) für die automatischen Mails an die Teilnehmer wählen',
      'Pro Mail-Template (Anmeldung, Storno, Warteliste, Erinnerung, QR-Code…) den Subject/Heading/Body anpassen — mit Live-Vorschau',
      'Eigenes Logo / Header-Bild für Mail und Outlook-Termin hochladen',
      'Outlook-Termin-Body individuell gestalten (Live-Vorschau zeigt wie das Outlook-Element später aussieht)',
      'Benachrichtigungen optional komplett deaktivieren — z.B. für interne Entwurfs-Events',
    ],
    [
      'Programm / Agenda pflegen (mehrtägig möglich, Drag-Reihenfolge pro Tag)',
      'Transferzeiten — Bus / Shuttle / Bahn von/zum Veranstaltungsort',
      'Dokumente hochladen (PDF) — Teilnehmer sehen sie auf MyEvents als Inline-Vorschau oder Download',
    ],
    [
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
      'Set event location and address',
      'Set start and end date (incl. time)',
      'Set the registration deadline — no new registrations after this date',
      'Optional: last self-cancel date — after that self-cancel is locked (late cancel)',
    ],
    [
      'Set the maximum number of attendees (or Unlimited)',
      'Enable waitlist — full events accept new registrations and promote them once a spot frees up',
      'Optional: split capacity — two freely-named groups with own seat count + own or shared waitlist',
      'Required fields per group in step 4 (e.g. proof of performance only for one group)',
    ],
    [
      'Pick a field type: text, number, dropdown, checkbox, people search or roommate (double room)',
      'Multi-select for dropdowns (e.g. tick multiple allergies)',
      'Mark required (red asterisk, blocks submit when empty)',
      'Description per field — appears as „i" tooltip next to the field label',
      'Visibility condition: only show this field when another question has a specific value (e.g. „Only ask room type if Hotel = yes")',
      'Reordering via drag or arrows — numbering updates automatically',
    ],
    [
      'Pick the email language (DE/EN) for automatic emails to attendees',
      'Edit subject / heading / body per email template (registration, cancellation, waitlist, reminder, QR code…) — with live preview',
      'Upload a custom logo / header image for the email and Outlook event',
      'Customise the Outlook event body (live preview shows how the Outlook item will appear)',
      'Optionally disable notifications entirely — e.g. for internal draft events',
    ],
    [
      'Maintain the event programme / agenda (multi-day supported, drag-reorder per day)',
      'Transfer times — bus / shuttle / train to and from the venue',
      'Upload documents (PDF) — attendees see them on MyEvents as inline preview or download',
    ],
    [
      'Create quiz questions for the event — multiple choice with any number of answer options',
      'Optionally upload an image per question (logo, photo quiz, etc.)',
      'Multiple correct answers are supported — all of them must be picked for full points',
      'Control cluster size: how many questions per „play block" — attendees can save progress and continue later',
      'See live highscore + statistics in the admin center (which questions are most often answered incorrectly)',
    ],
  ];

  const steps = [
    { label: t('create.step.basics'), icon: '1' },
    { label: t('create.step.datetime'), icon: '2' },
    { label: t('create.step.capacity'), icon: '3' },
    { label: t('create.step.fields'), icon: '4' },
    { label: t('create.step.communication'), icon: '5' },
    { label: t('create.step.documents'), icon: '6' },
    { label: t('create.step.funzone'), icon: '7' },
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
        // Schritt 2 (Ort & Programm) ist ohne Pflicht-Validierung —
        // Adresse / Agenda / Transferzeiten / Sub-Events sind alle
        // optional. Datum-Checks laufen jetzt in case 0.
        break;
      case 2:
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
        {/* ===== Step Progress Bar ===== */}
        <div style={{ marginBottom: 32 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', position: 'relative' }}>
            {/* Verbindungslinie */}
            <div style={{ position: 'absolute', top: 20, left: '10%', right: '10%', height: 3, background: 'var(--dex-gray-200)', zIndex: 0 }} />
            <div style={{ position: 'absolute', top: 20, left: '10%', height: 3, background: 'var(--dex-green)', zIndex: 1, width: `${Math.min(currentStep / (steps.length - 1) * 80, 80)}%`, transition: 'width 0.4s ease' }} />
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
                        fontFamily: 'inherit',
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
                    onClick={fillDemo}
                    style={{ fontSize: '0.8rem', padding: '4px 12px' }}
                  >
                    Demo
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
                      <strong>Automatismen:</strong> Organizer bekommen je nach Einstellung in <strong>Schritt 5 (Kommunikation)</strong> eine BCC-Kopie der Anmelde-/Abmelde-Mails. Late-Cancel- und Roommate-Mails gehen ebenfalls an alle Organizer. Wenn ein Teilnehmer die Outlook-Einladung weiterleitet und der Empfänger nicht angemeldet ist, bekommen die Organizer eine FYI-Mail.<br /><br />
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
                        const results = await searchUsers(q);
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
                        const results = await searchUsers(q);
                        setTestTeamResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setTestTeamResults([]); }
                    }, 350);
                  }}
                  onBlur={() => {
                    setTimeout(() => { setTestTeamSearch(''); setTestTeamResults([]); }, 150);
                  }}
                  placeholder="Name oder E-Mail eingeben (alle Deloitte-User)"
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
                        const results = await searchUsers(q);
                        setQrScannerResults(results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location || '' })));
                      } catch { setQrScannerResults([]); }
                    }, 350);
                  }}
                  onBlur={() => {
                    setTimeout(() => { setQrScannerSearch(''); setQrScannerResults([]); }, 150);
                  }}
                  placeholder={t('create.qrscanners.placeholder') || 'Name oder E-Mail eingeben (alle Deloitte-User)'}
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

              {/* ===== Step 1: Ort & Programm ===== */}
              <div style={{ display: currentStep === 1 ? 'block' : 'none' }}>
              <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                {isDe ? 'Schritt 2 — Ort & Programm' : 'Step 2 — Location & Programme'}
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

              {/* ===== Agenda Editor ===== */}
              <div className="form-group" style={{ marginTop: 24 }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '1rem', fontWeight: 700 }}>
                  <StepBadge n={11} />
                  {t('create.agenda')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> den <strong>Programmablauf des Events</strong> als Liste — pro Punkt: Datum, Start- und Endzeit, Titel, optionale Beschreibung und ein Icon (z.B. Kaffee, Vortrag, Pause).<br /><br />
                      <strong>Anzeige in der App:</strong> erscheint als <strong>schöner Timeline-Block</strong> auf der Anmelde-Seite und in Meine Events — Punkte werden automatisch nach Datum + Uhrzeit sortiert. Mehrtägige Events werden tageweise gruppiert.<br /><br />
                      <strong>Automatismen:</strong> die Agenda landet <strong>nicht</strong> automatisch im Outlook-Termin-Body (dafür gibt es das eigene Feld <strong>Text im Outlook-Termin</strong> in Schritt 5).<br /><br />
                      <strong>Empfehlung:</strong> hilft Teilnehmern, sich auf den Tag einzustellen — bei Tagungen oder Auswärtsterminen sehr empfohlen, bei kurzen Office-Events optional.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>event programme</strong> as a list — per item: date, start/end time, title, optional description and an icon (e.g. coffee, talk, break).<br /><br />
                      <strong>Shown in the app:</strong> shown as a <strong>nice timeline block</strong> on the registration page and in My Events — items are auto-sorted by date + time. Multi-day events are grouped per day.<br /><br />
                      <strong>Automation:</strong> the agenda is <strong>not</strong> automatically pulled into the Outlook event body (there is a dedicated field <strong>Text in the Outlook event</strong> in step 5 for that).<br /><br />
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

                {/* ===== Sub-Events (z.B. Workshop-Tage, Networking-Dinner, Kick-off-Sessions) ===== */}
                <div className="form-group" style={{ marginTop: 32 }}>
                  <label className="form-label" style={{ fontSize: '1rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8 }}>
                    <StepBadge n={13} />
                    {t('create.subevents')}
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
                    const deadlineObj = se.registrationDeadline ? (() => {
                      const local = isoToLocal(se.registrationDeadline);
                      if (!local) return null;
                      const [dp, tp] = local.split('T');
                      const [y, mo, da] = dp.split('-').map(n => parseInt(n, 10));
                      const [h, mi] = (tp || '00:00').split(':').map(n => parseInt(n, 10));
                      return new Date(y, mo - 1, da, h, mi, 0, 0);
                    })() : null;
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

                        {/* Zeit-Zeile: Start + Ende + Anmeldeschluss + Max.
                            v10.23: 4. Spalte von 120 auf 170 erweitert — der
                            DE-Label "Max. Teilnehmer (0 = unbegrenzt)" ist
                            sonst zu lang und der Hinweis "0 = unbegrenzt"
                            wird ueber dem Eingabefeld in eine zweite Zeile
                            umgebrochen. */}
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 170px', gap: 10, marginBottom: 10 }}>
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
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.subevents.deadline')}</label>
                            <DatePicker
                              selected={deadlineObj}
                              onChange={(date: Date | null) => {
                                const iso = dateToBerlinIso(date);
                                setSubEvents(subEvents.map((x, i) => i === idx ? { ...x, registrationDeadline: iso } : x));
                              }}
                              showTimeSelect
                              timeFormat="HH:mm"
                              timeIntervals={15}
                              timeCaption="Uhrzeit"
                              dateFormat="dd.MM.yyyy, HH:mm"
                              locale="de"
                              placeholderText="Optional"
                              className="form-input"
                              wrapperClassName="dex-datepicker-wrapper"
                              calendarClassName="dex-datepicker-calendar"
                              popperPlacement="bottom-start"
                              maxDate={startDateObj || undefined}
                              isClearable
                              autoComplete="off"
                            />
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.subevents.max')}</label>
                            <input
                              type="number"
                              min={0}
                              className="form-input"
                              value={se.maxParticipants || 0}
                              onChange={e => {
                                const v = parseInt(e.target.value, 10) || 0;
                                setSubEvents(subEvents.map((x, i) => i === idx ? { ...x, maxParticipants: v } : x));
                              }}
                              style={{ padding: '6px 10px', fontSize: '0.85rem' }}
                            />
                          </div>
                        </div>

                        {/* Ort + Beschreibung */}
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: 10, marginBottom: 10 }}>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.subevents.location')}</label>
                            <input
                              type="text"
                              className="form-input"
                              value={se.location || ''}
                              placeholder={t('create.subevents.location.placeholder')}
                              onChange={e => {
                                const v = e.target.value;
                                setSubEvents(subEvents.map((x, i) => i === idx ? { ...x, location: v } : x));
                              }}
                              style={{ padding: '6px 10px', fontSize: '0.85rem' }}
                            />
                          </div>
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

                        {/* Toggles: Mails + Outlook */}
                        <div style={{ display: 'flex', gap: 20, flexWrap: 'wrap', fontSize: '0.8rem', color: 'var(--dex-gray-600)', paddingTop: 6, borderTop: '1px dashed var(--dex-gray-200)' }}>
                          <label style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer' }}>
                            <input
                              type="checkbox"
                              checked={!se.disableEmails}
                              onChange={e => {
                                const v = !e.target.checked;
                                setSubEvents(subEvents.map((x, i) => i === idx ? { ...x, disableEmails: v } : x));
                              }}
                              style={{ width: 14, height: 14, cursor: 'pointer' }}
                            />
                            {t('create.subevents.emails')}
                          </label>
                          <label style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer' }}>
                            <input
                              type="checkbox"
                              checked={!se.disableOutlook}
                              onChange={e => {
                                const v = !e.target.checked;
                                setSubEvents(subEvents.map((x, i) => i === idx ? { ...x, disableOutlook: v } : x));
                              }}
                              style={{ width: 14, height: 14, cursor: 'pointer' }}
                            />
                            {t('create.subevents.outlook')}
                          </label>
                        </div>
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
                      };
                      setSubEvents([...subEvents, newSub]);
                    }}
                  >
                    <Plus size={14} /> {t('create.subevents.add')}
                  </button>
                </div>
              </div>

              {/* ===== Step 2: Kapazität, Fristen & Sichtbarkeit ===== */}
              <div style={{ display: currentStep === 2 ? 'block' : 'none' }}>
              <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                {isDe ? 'Schritt 3 — Kapazität & Sichtbarkeit' : 'Step 3 — Capacity & Visibility'}
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
                            searchUsers(val.trim()),
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
                        <strong>Automatismen:</strong> je nach Einstellung in <strong>Schritt 5 (Kommunikation)</strong> bekommen die Organizer eine <strong>Late-Cancel-Mail</strong> mit Name + Mail des Abmeldenden — damit Hotel, Catering oder Bus angepasst werden können.<br /><br />
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
                      Stattdessen kann der Organizer in Schritt 4 (Felder) ein
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
                        Du möchtest für eine der zwei Gruppen ein zusätzliches Pflichtfeld einblenden — z.B. eine Checkbox &bdquo;Leistungsnachweis vorhanden&ldquo; nur für die Gruppe der schnellen Läufer? Lege das Feld in <strong>Schritt 4 (Felder)</strong> an und stelle dort den Selector <strong>&bdquo;Sichtbar für Teilnehmergruppe&ldquo;</strong> auf <strong>&bdquo;Nur {(splitLabelA || '').trim() || 'Gruppe A'}&ldquo;</strong> bzw. <strong>&bdquo;Nur {(splitLabelB || '').trim() || 'Gruppe B'}&ldquo;</strong>. Das Feld wird dann in der Anmeldung dynamisch ein- oder ausgeblendet, sobald der Teilnehmer eine der zwei Boxen anklickt.
                      </>
                    ) : (
                      <>
                        Want to show an extra required field only for one of the two groups — e.g. a checkbox &ldquo;Performance proof available&rdquo; just for the fast-runner group? Add the field in <strong>step 4 (Fields)</strong> and set the <strong>&ldquo;Visible for attendee group&rdquo;</strong> selector there to <strong>&ldquo;{(splitLabelA || '').trim() || 'Group A'} only&rdquo;</strong> or <strong>&ldquo;{(splitLabelB || '').trim() || 'Group B'} only&rdquo;</strong>. The field will then be shown or hidden dynamically as the attendee picks one of the two boxes.
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

              </div>

              {/* ===== Step 3: Registrierungsfelder ===== */}
              <div style={{ display: currentStep === 3 ? 'block' : 'none' }}>
              <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                {isDe ? 'Schritt 4 — Felder' : 'Step 4 — Fields'}
              </h2>
              <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                {isDe
                  ? <><strong>Optional</strong> — die Standard-Teilnehmerdaten (<strong>Anrede, Vorname, Nachname, E-Mail</strong>) werden bei jeder Anmeldung automatisch erfasst, dazu kommen aus dem Deloitte-Profil <strong>Job Title, Standort, Department und Telefonnummer</strong>. Hier in Schritt 4 ergänzt du <strong>nur zusätzliche Fragen</strong>, die du speziell für dieses Event brauchst — vom T-Shirt-Größen-Dropdown bis zur Pflicht-Checkbox für AGB / Datenschutz. Wenn dein Event keine Zusatzfragen braucht, kannst du diesen Schritt einfach leer lassen.</>
                  : <><strong>Optional</strong> — the standard attendee data (<strong>salutation, first name, last name, email</strong>) is captured automatically for every registration, plus <strong>job title, location, department and phone</strong> are pulled from the Deloitte profile. In step 4 you only add <strong>extra questions</strong> specific to this event — from a T-shirt size dropdown to a privacy / terms required checkbox. If your event needs no extra questions, you can simply leave this step empty.</>}
              </p>

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

              {/* Dynamische Felder */}
              <div>
                {/* Bereich-Header: trennt Hauptevent-Felder visuell vom
                    Sub-Event-Block weiter unten (v10.11+). */}
                <h3 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.15rem', fontWeight: 700 }}>
                  {isDe ? 'Felder für das Hauptevent' : 'Fields for the main event'}
                </h3>
                <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                  {isDe
                    ? 'Diese Felder werden bei jeder Anmeldung abgefragt — egal ob das Event Sub-Events hat oder nicht. Für Sub-Event-spezifische Fragen siehe Bereich „Felder pro Sub-Event" weiter unten.'
                    : 'These fields are asked at every registration — regardless of whether the event has sub-events. For sub-event-specific questions see „Fields per sub-event" below.'}
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

                {/* v7.20: Spalten-Header oberhalb der Feld-Karten — erklaert
                    auf einen Blick welche Spalte was bedeutet. Nur sichtbar
                    wenn es mindestens 1 Feld gibt, sonst overhead. */}
                {/* v11.1: alter Tabellen-Header (Nr / Typ / Frage) entfernt —
                    seit der Card-Restrukturierung in v10.25 sind die Spalten
                    in jeder Card selbsterklärend (Nummern-Bubble + Label-
                    Input mit Placeholder + grüner Typ-Selector + Pflicht-
                    Pill + Lösch-X). Der Header passte mit den alten Spalten-
                    Breiten nicht mehr und stiftete optisch Verwirrung. */}
                {customFields.map((field, idx) => (
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
                      <input
                        className="form-input"
                        value={field.label}
                        placeholder={isDe ? 'Feld-Name (z.B. T-Shirt Größe)' : 'Field name (e.g. T-shirt size)'}
                        onChange={e => updateCustomField(field.id, { label: e.target.value })}
                        style={{
                          /* v11.1: max-width begrenzen, damit Typ + Pflicht
                             + X immer in einer Zeile bleiben (vorher: flex
                             1 1 220px hat den Input bei breiten Cards bis
                             zur halben Card-Breite gestreckt). 320px reicht
                             für lange Feld-Namen, lässt aber genug Raum
                             für die rechten Aktions-Elemente. */
                          flex: '0 1 320px', minWidth: 180, maxWidth: 320,
                          fontSize: '1rem', fontWeight: 600,
                          padding: '8px 12px',
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
                        onClick={() => removeCustomField(field.id)}
                        title={isDe ? 'Feld löschen' : 'Delete field'}
                        style={{ background: 'none', border: 'none', color: 'var(--dex-red)', padding: 4, cursor: 'pointer', flexShrink: 0 }}
                      >
                        <X size={18} />
                      </button>
                    </div>

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

                    {/* v7.20: Beschreibung pro Feld — landet im Registrierungs-
                        formular als "i"-Tooltip neben dem Label. */}
                    <div style={{ marginLeft: 32, marginTop: 10 }}>
                      <input
                        className="form-input"
                        placeholder={isDe
                          ? 'Beschreibung (optional, erscheint als "i"-Tooltip neben dem Feld)'
                          : 'Description (optional, shown as "i" tooltip next to the field)'}
                        value={field.helpText || ''}
                        onChange={e => updateCustomField(field.id, { helpText: e.target.value })}
                        style={{ width: '100%', fontSize: '0.82rem', padding: '6px 10px' }}
                      />
                    </div>


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
                            <div key={optIdx} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
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
                                  updateCustomField(field.id, { options: opts });
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
                  </div>
                ))}
              </div>

              {/* === Bereich 2: Felder pro Sub-Event (v10.11+) ============
                  Jedes Sub-Event hat eine eigene Custom-Fields-Liste, die bei
                  der Anmeldung auf einem Sub-Event zusätzlich zu den Hauptevent-
                  Feldern gerendert wird. Ersetzt die hardcoded Fun-/Durchstarter-
                  Frage durch frei konfigurierbare Auswahl-Fragen. */}
              <div style={{ marginTop: 32, paddingTop: 24, borderTop: '2px solid var(--dex-gray-200)' }}>
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
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                              {seFields.map(field => (
                                <div
                                  key={field.id}
                                  style={{
                                    border: '1px solid var(--dex-gray-200)',
                                    borderRadius: 6,
                                    padding: '10px 12px',
                                    background: '#fff',
                                  }}
                                >
                                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', marginBottom: 8 }}>
                                    <select
                                      className="form-input"
                                      value={field.type}
                                      onChange={e => updateSubEventCustomField(se.id, field.id, { type: e.target.value as CustomFieldInput['type'] })}
                                      style={{ flex: '0 0 130px', fontSize: '0.82rem', padding: '4px 8px' }}
                                    >
                                      <option value="text">{isDe ? 'Text' : 'Text'}</option>
                                      <option value="select">{isDe ? 'Dropdown' : 'Dropdown'}</option>
                                      <option value="number">{isDe ? 'Zahl' : 'Number'}</option>
                                      <option value="checkbox">{isDe ? 'Checkbox' : 'Checkbox'}</option>
                                    </select>
                                    <input
                                      className="form-input"
                                      placeholder={isDe ? 'Frage / Feld-Label (z.B. „Welche Strecke?")' : 'Question / field label (e.g. „Which distance?")'}
                                      value={field.label}
                                      onChange={e => updateSubEventCustomField(se.id, field.id, { label: e.target.value })}
                                      style={{ flex: 2, fontSize: '0.85rem', padding: '4px 8px' }}
                                    />
                                    <button
                                      type="button"
                                      onClick={() => removeSubEventCustomField(se.id, field.id)}
                                      style={{ background: 'none', border: 'none', color: 'var(--dex-red)', padding: 4, cursor: 'pointer' }}
                                      title={isDe ? 'Feld entfernen' : 'Remove field'}
                                    >
                                      <X size={16} />
                                    </button>
                                  </div>

                                  {/* Required-Toggle + Mehrfachauswahl */}
                                  <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', alignItems: 'center', marginBottom: field.type === 'select' ? 10 : 0 }}>
                                    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.78rem', cursor: 'pointer' }}>
                                      <input
                                        type="checkbox"
                                        checked={field.required}
                                        onChange={e => updateSubEventCustomField(se.id, field.id, { required: e.target.checked })}
                                      />
                                      {isDe ? 'Pflichtfeld' : 'Required'}
                                    </label>
                                    {field.type === 'select' && (
                                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.78rem', cursor: 'pointer' }}>
                                        <input
                                          type="checkbox"
                                          checked={!!field.multi}
                                          onChange={e => updateSubEventCustomField(se.id, field.id, { multi: e.target.checked })}
                                        />
                                        {isDe ? 'Mehrfachauswahl' : 'Multi-select'}
                                      </label>
                                    )}
                                  </div>

                                  {/* v10.27: Optionale Beschreibung pro Sub-Event-Feld
                                      — landet bei der Anmeldung als „i"-Tooltip
                                      neben dem Feld (analog zu Hauptevent-Custom-
                                      Fields). */}
                                  <div style={{ marginBottom: field.type === 'select' ? 10 : 0 }}>
                                    <input
                                      className="form-input"
                                      placeholder={isDe
                                        ? 'Beschreibung (optional, erscheint als „i"-Tooltip neben dem Feld)'
                                        : 'Description (optional, shown as „i"-tooltip next to the field)'}
                                      value={field.helpText || ''}
                                      onChange={e => updateSubEventCustomField(se.id, field.id, { helpText: e.target.value })}
                                      style={{ width: '100%', fontSize: '0.8rem', padding: '4px 8px' }}
                                    />
                                  </div>

                                  {/* Optionen-Editor für Dropdown-Felder */}
                                  {field.type === 'select' && (
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginLeft: 12 }}>
                                      <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', fontWeight: 600 }}>
                                        {isDe ? 'Antwort-Optionen:' : 'Answer options:'}
                                      </span>
                                      {field.options.map((opt, idx) => (
                                        <div key={idx} style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                                          <input
                                            className="form-input"
                                            placeholder={isDe ? `Option ${idx + 1}` : `Option ${idx + 1}`}
                                            value={opt}
                                            onChange={e => {
                                              const next = field.options.slice();
                                              next[idx] = e.target.value;
                                              updateSubEventCustomField(se.id, field.id, { options: next });
                                            }}
                                            style={{ flex: 1, fontSize: '0.82rem', padding: '4px 8px' }}
                                          />
                                          {field.options.length > 1 && (
                                            <button
                                              type="button"
                                              onClick={() => {
                                                const next = field.options.filter((_, i) => i !== idx);
                                                updateSubEventCustomField(se.id, field.id, { options: next });
                                              }}
                                              style={{ background: 'none', border: 'none', color: 'var(--dex-gray-500)', padding: 2, cursor: 'pointer' }}
                                            >
                                              <X size={14} />
                                            </button>
                                          )}
                                        </div>
                                      ))}
                                      <button
                                        type="button"
                                        onClick={() => updateSubEventCustomField(se.id, field.id, { options: [...field.options, ''] })}
                                        style={{ alignSelf: 'flex-start', background: 'none', border: '1px dashed var(--dex-gray-300)', padding: '3px 10px', fontSize: '0.75rem', borderRadius: 4, cursor: 'pointer', color: 'var(--dex-gray-700)' }}
                                      >
                                        <Plus size={11} /> {isDe ? 'Option hinzufügen' : 'Add option'}
                                      </button>
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

              </div>{/* close Step 3 */}

              {/* ===== Step 4: Kommunikation ===== */}
              <div style={{ display: currentStep === 4 ? 'block' : 'none' }}>
                <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                  {isDe ? 'Schritt 5 — Kommunikation' : 'Step 5 — Communication'}
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

                {/* Benachrichtigungen abschalten — v9.39: collapsed by default. */}
                <details className="form-group" style={{ marginTop: 24, padding: 16, background: 'var(--dex-gray-50, #f8f9fa)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-gray-200)' }}>
                  <summary style={{ cursor: 'pointer', listStyle: 'none', display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', fontWeight: 600 }}>
                    <StepBadge n={21} />
                    {t('create.notifications')}
                    <span style={{ marginLeft: 'auto', fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                      {isDe ? 'Standard – empfohlen, klick zum Anpassen' : 'Default – recommended, click to adjust'}
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

              </div>{/* close Step 4 */}

              {/* ===== Step 5: Dokumente ===== */}
              <div style={{ display: currentStep === 5 ? 'block' : 'none' }}>
                <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                  {isDe ? 'Schritt 6 — Dokumente' : 'Step 6 — Documents'}
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
              </div>{/* close Step 5 */}

              {/* ===== Step 6: Fun-Zone ===== */}
              <div style={{ display: currentStep === 6 ? 'block' : 'none' }}>
                <h2 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.4rem', fontWeight: 700 }}>
                  {isDe ? 'Schritt 7 — Fun-Zone' : 'Step 7 — Fun Zone'}
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
                                  if (!window.confirm(`Bereich "${sec}" entfernen? Die Fragen bleiben erhalten und landen in "Ohne Bereich".`)) return;
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
              </div>{/* close Step 6 */}

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

              <div style={{ display: 'flex', gap: 8 }}>
                {/* Im Edit-Modus immer einen Speichern-Button anzeigen, damit man nicht
                    durch alle Steps klicken muss wenn man nur eine Sache aendert */}
                {isEditMode && currentStep < steps.length - 1 && (
                  <button
                    className="btn btn-primary"
                    disabled={!title}
                    onClick={handleSubmit}
                    style={{ opacity: !title ? 0.5 : 1 }}
                    title="Aenderungen sofort speichern, ohne weitere Schritte"
                  >
                    <Send size={16} /> {t('create.save')}
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
                    onClick={handleSubmit}
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
                onClick={() => { setShowPreview(false); handleSubmit(); }}
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
        const currentBody = isOutlook
          ? outlookBody
          : isDescription
            ? description
            : (override?.bodyHtml || defaultTpl?.bodyHtml || '');
        return (
          <HtmlEditorModal
            open={htmlEditorOpen}
            onClose={() => setHtmlEditorOpen(false)}
            title={isOutlook ? 'Outlook-Termin: Body bearbeiten' : isDescription ? (isDe ? 'Event-Beschreibung bearbeiten' : 'Edit event description') : `E-Mail-Template: ${tType}`}
            value={currentBody}
            onChange={(html) => {
              if (isOutlook) {
                setOutlookBody(html);
              } else if (isDescription) {
                setDescription(html);
              } else {
                setEmailTemplateOverrides({
                  ...emailTemplateOverrides,
                  [tType]: { subject: currentSubject, heading: currentHeading, bodyHtml: html },
                });
              }
            }}
            previewMode={(isOutlook || isDescription) ? 'outlook' : 'email'}
            emailSubject={(!isOutlook && !isDescription) ? currentSubject : undefined}
            onEmailSubjectChange={(!isOutlook && !isDescription) ? (s) => setEmailTemplateOverrides({
              ...emailTemplateOverrides,
              [tType]: { subject: s, heading: currentHeading, bodyHtml: currentBody },
            }) : undefined}
            emailHeading={(!isOutlook && !isDescription) ? currentHeading : undefined}
            onEmailHeadingChange={(!isOutlook && !isDescription) ? (h) => setEmailTemplateOverrides({
              ...emailTemplateOverrides,
              [tType]: { subject: currentSubject, heading: h, bodyHtml: currentBody },
            }) : undefined}
            emailHeadingColor={(!isOutlook && !isDescription) ? (defaultTpl?.headingColor || '#86bc25') : undefined}
            outlookHeading={isOutlook ? outlookHeading : undefined}
            onOutlookHeadingChange={isOutlook ? setOutlookHeading : undefined}
            outlookSubheading={isOutlook ? outlookSubheading : undefined}
            onOutlookSubheadingChange={isOutlook ? setOutlookSubheading : undefined}
            previewVars={{
              EventTitle: title || 'Event Title',
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
              { key: '{{Name}}', label: 'Name' },
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
            multi: f.multi,
            showIf: f.showIf,
          })),
          isFictive,
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
                    const found = await searchUsers(v.trim());
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
                      const results = await searchUsers(emailSearch);
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
                    const results = await searchUsers(emailSearch);
                    setEmailSearchResults(results);
                    setIsSearchingEmails(false);
                  }}
                >
                  {isSearchingEmails ? '...' : 'Suchen'}
                </button>
              </div>
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

      {/* v9.28: Modal — neuer Quiz-Bereich anlegen (statt window.prompt) */}
      {newSectionModalOpen && (
        <div
          role="dialog"
          aria-modal="true"
          onClick={() => setNewSectionModalOpen(false)}
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            zIndex: 9999, padding: 16,
          }}
        >
          <div
            onClick={e => e.stopPropagation()}
            style={{
              background: '#fff', borderRadius: 'var(--dex-radius-lg, 12px)',
              padding: '24px 28px', maxWidth: 460, width: '100%',
              boxShadow: '0 12px 40px rgba(0,0,0,0.22)',
            }}
          >
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
          </div>
        </div>
      )}

      {/* Modal: Vorgeschlagene Felder auswaehlen (Multi-Select) */}
      {showSuggestedModal && (
        <div
          role="dialog"
          aria-modal="true"
          onClick={() => setShowSuggestedModal(false)}
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            zIndex: 9999, padding: 16,
          }}
        >
          <div
            onClick={e => e.stopPropagation()}
            style={{
              background: '#fff', borderRadius: 16, padding: '24px 28px',
              maxWidth: 540, width: '100%', boxShadow: '0 12px 48px rgba(0,0,0,0.18)',
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
        </div>
      )}
    </div>
  );
}
