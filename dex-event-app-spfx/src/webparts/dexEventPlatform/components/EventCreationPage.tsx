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
import { eventCreatedEmail, buildOutlookBody } from '../services/EmailTemplates';
import { EventType, AgendaItem } from '../types';
import { Trash2, Send, Plus, X, Users, Check } from './Icons';
import { RichText } from '@pnp/spfx-controls-react/lib/controls/richText';
import { HtmlEditorModal } from './HtmlEditorModal';
import { InfoTooltip } from './InfoTooltip';
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
  type: 'text' | 'select' | 'number' | 'checkbox' | 'user';
  required: boolean;
  // Optionen als Array (incl. leerer Slots fuer "frisch hinzugefuegte" Eintraege)
  options: string[];
  visible: boolean;
  externalLinks?: Array<{ label: string; url: string }>;
}

export default function EventCreationPage(): React.ReactElement {
  const { navigate, goBack, selectedEventId, currentPage } = useNavigation();
  const { events, createEvent, updateEvent, refreshEvents } = useEvents();
  const { currentUser } = useCurrentUser();
  const { searchUsers, searchGroups, getGroupMembers, roles } = useRoles();
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
  // Modal: Massenimport fuer Audience
  const [bulkImportOpen, setBulkImportOpen] = React.useState(false);
  const [bulkImportText, setBulkImportText] = React.useState('');
  const [bulkImportRunning, setBulkImportRunning] = React.useState(false);
  const [bulkImportReport, setBulkImportReport] = React.useState<{
    added: Array<{ lastname: string; firstname: string; email: string; originalInput?: string }>;
    notFound: string[];
    alreadyIn: Array<{ lastname: string; firstname: string; email: string }>;
    ambiguous: Array<{ input: string; matches: Array<{ email: string; displayName: string }> }>;
  } | null>(null);
  // Audience-Chip-Pagination + Inline-Suche
  const [audienceShowAll, setAudienceShowAll] = React.useState(false);
  const [audienceChipSearch, setAudienceChipSearch] = React.useState('');

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

  /**
   * Massenimport: Text parsen (getrennt mit ',', ';' oder Zeilenumbruch, plus
   * Tab als Fallback fuer Excel-Copy-Paste). Fuer jedes Fragment:
   *   - Wenn schon eine Email (x@y.z): direkt als Audience uebernehmen
   *   - Sonst: searchUsers(fragment) → bei genau 1 Treffer: uebernehmen; bei
   *     mehreren Treffern: in 'ambiguous' zwischenspeichern (User entscheidet
   *     per Modal manuell); bei 0 Treffern: in 'notFound'.
   * Sequenziell + mit kleinen Pausen, um die Graph-API nicht zu hammer.
   */
  const runBulkImport = async (): Promise<void> => {
    const raw = bulkImportText;
    if (!raw || !raw.trim()) return;
    setBulkImportRunning(true);
    setBulkImportReport(null);
    const existing = audience.split(',').map((s: string) => s.trim()).filter(Boolean);
    const existingLc = new Set(existing.map(e => e.toLowerCase()));
    // Split auf ',', ';', Zeilenumbruch und Tab
    const fragments = raw
      .split(/[,;\n\r\t]/)
      .map(s => s.trim())
      .filter(Boolean);
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    // Hilfs-Parser: 'Nachname, Vorname' oder 'Vorname Nachname' -> {lastname, firstname}
    const parseDisplayName = (displayName: string, fallbackEmail: string): { lastname: string; firstname: string } => {
      const dn = (displayName || '').trim();
      if (!dn) return { lastname: fallbackEmail || '', firstname: '' };
      if (dn.indexOf(',') >= 0) {
        const parts = dn.split(',').map(s => s.trim());
        return { lastname: parts[0] || dn, firstname: parts[1] || '' };
      }
      // Space-separiert: letztes Wort = Nachname
      const words = dn.split(/\s+/);
      if (words.length >= 2) {
        return { lastname: words[words.length - 1], firstname: words.slice(0, -1).join(' ') };
      }
      return { lastname: dn, firstname: '' };
    };

    const added: Array<{ lastname: string; firstname: string; email: string; originalInput?: string }> = [];
    const notFound: string[] = [];
    const alreadyIn: Array<{ lastname: string; firstname: string; email: string }> = [];
    const ambiguous: Array<{ input: string; matches: Array<{ email: string; displayName: string }> }> = [];
    const newEntries = existing.slice();

    for (const f of fragments) {
      // Schon drin?
      if (existingLc.has(f.toLowerCase())) {
        // Versuch Name-Lookup auch fuer 'alreadyIn', damit Report lesbar ist
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const hits: any[] = emailRegex.test(f) ? await searchUsers(f) : [];
          if (hits.length > 0) {
            const { lastname, firstname } = parseDisplayName(hits[0].displayName, f);
            alreadyIn.push({ lastname, firstname, email: f });
          } else {
            alreadyIn.push({ lastname: f, firstname: '', email: f });
          }
        } catch {
          alreadyIn.push({ lastname: f, firstname: '', email: f });
        }
        continue;
      }
      // Direkt-Email: Name per searchUsers(email) nachladen
      if (emailRegex.test(f)) {
        newEntries.push(f);
        existingLc.add(f.toLowerCase());
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const hits: any[] = await searchUsers(f);
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const exact = hits.find((h: any) => h.email && h.email.toLowerCase() === f.toLowerCase());
          if (exact) {
            const { lastname, firstname } = parseDisplayName(exact.displayName, f);
            added.push({ lastname, firstname, email: exact.email });
          } else {
            added.push({ lastname: f, firstname: '', email: f });
          }
        } catch {
          added.push({ lastname: f, firstname: '', email: f });
        }
        await new Promise(res => setTimeout(res, 80));
        continue;
      }
      // Name / Teilstring → searchUsers
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const hits: any[] = await searchUsers(f);
        if (!hits || hits.length === 0) {
          notFound.push(f);
          continue;
        }
        if (hits.length === 1) {
          const em = hits[0].email;
          if (em && !existingLc.has(em.toLowerCase())) {
            newEntries.push(em);
            existingLc.add(em.toLowerCase());
            const { lastname, firstname } = parseDisplayName(hits[0].displayName, em);
            added.push({ lastname, firstname, email: em, originalInput: f });
          } else {
            const { lastname, firstname } = parseDisplayName(hits[0].displayName, em || f);
            alreadyIn.push({ lastname, firstname, email: em || f });
          }
          continue;
        }
        // Mehrere Treffer: User muss manuell entscheiden
        ambiguous.push({
          input: f,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          matches: hits.slice(0, 5).map((h: any) => ({ email: h.email, displayName: h.displayName })),
        });
      } catch {
        notFound.push(f);
      }
      await new Promise(res => setTimeout(res, 120));
    }

    // Alphabetisch sortieren: Nachname, dann Vorname
    const sortFn = (a: { lastname: string; firstname: string }, b: { lastname: string; firstname: string }): number => {
      const ln = a.lastname.localeCompare(b.lastname, 'de', { sensitivity: 'base' });
      if (ln !== 0) return ln;
      return a.firstname.localeCompare(b.firstname, 'de', { sensitivity: 'base' });
    };
    added.sort(sortFn);
    alreadyIn.sort(sortFn);
    notFound.sort();

    setAudience(newEntries.join(', '));
    setBulkImportReport({ added, notFound, alreadyIn, ambiguous });
    setBulkImportRunning(false);
  };

  /** Ambiguous-Eintrag manuell aufloesen: pickt eine Match-Email aus. */
  const resolveAmbiguous = (input: string, email: string, displayName: string): void => {
    const list = audience.split(',').map((s: string) => s.trim()).filter(Boolean);
    if (list.indexOf(email) < 0) list.push(email);
    setAudience(list.join(', '));
    if (bulkImportReport) {
      const dn = (displayName || '').trim();
      let lastname = email;
      let firstname = '';
      if (dn.indexOf(',') >= 0) {
        const parts = dn.split(',').map(s => s.trim());
        lastname = parts[0] || email;
        firstname = parts[1] || '';
      } else if (dn.indexOf(' ') >= 0) {
        const words = dn.split(/\s+/);
        lastname = words[words.length - 1];
        firstname = words.slice(0, -1).join(' ');
      } else if (dn) {
        lastname = dn;
      }
      const newAdded = [...bulkImportReport.added, { lastname, firstname, email, originalInput: input }];
      newAdded.sort((a, b) => {
        const ln = a.lastname.localeCompare(b.lastname, 'de', { sensitivity: 'base' });
        if (ln !== 0) return ln;
        return a.firstname.localeCompare(b.firstname, 'de', { sensitivity: 'base' });
      });
      setBulkImportReport({
        ...bulkImportReport,
        ambiguous: bulkImportReport.ambiguous.filter(a => a.input !== input),
        added: newAdded,
      });
    }
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
  // Mehrere Organizer werden mit '; ' getrennt gespeichert (innerhalb eines Namens kann ',' vorkommen, z.B. 'Maerzluft, Petra')
  const [organizer, setOrganizer] = React.useState(
    editEvent ? editEvent.organizers.join('; ') : `${currentUser.firstName} ${currentUser.surname}`
  );
  const [organizerResults, setOrganizerResults] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const [organizerSearch, setOrganizerSearch] = React.useState('');
  // Beim Edit: organizerEmails aus dem gespeicherten Event uebernehmen, nicht auf currentUser
  // zuruecksetzen. Sonst ueberschreibt ein Edit+Save die gesamte Organizer-Email-Liste mit
  // nur der Mail des aktuellen Editors — alle anderen Organizer wuerden stumm aus der
  // Late-Cancel- / Organizer-Mail-Verteilung rausfallen.
  const [organizerEmails, setOrganizerEmails] = React.useState<string[]>(
    editEvent && editEvent.organizerEmails && editEvent.organizerEmails.length > 0
      ? editEvent.organizerEmails.slice()
      : [currentUser.email]
  );
  // isSearchingOrganizer entfaellt seit v4.8.0 — Filter laeuft sync gegen den
  // bereits geladenen DEX_Roles-State, kein Async-Spinner mehr noetig.
  const isSearchingOrganizer = false;
  const organizerTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
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
  const [filterMode, setFilterMode] = React.useState<'AND' | 'OR'>(
    editEvent ? editEvent.filterMode : 'AND'
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
  const [maxParticipants, setMaxParticipants] = React.useState(
    editEvent && editEvent.maxParticipants ? editEvent.maxParticipants.toString() : ''
  );
  const [unlimitedParticipants, setUnlimitedParticipants] = React.useState(
    !editEvent || !editEvent.maxParticipants || editEvent.maxParticipants === 0
  );
  const [waitlistEnabled, setWaitlistEnabled] = React.useState(true);
  const [eventImageUrl, setEventImageUrl] = React.useState(editEvent ? (editEvent.imageUrl || '') : '');
  const [imageFile, setImageFile] = React.useState<File | null>(null);
  const [imagePreview, setImagePreview] = React.useState(editEvent ? (editEvent.imageUrl || '') : '');
  const [customFields, setCustomFields] = React.useState<CustomFieldInput[]>(
    editEvent ? editEvent.eventSpecificFields.map(f => ({
      id: f.id, label: f.label, type: f.type, required: f.required,
      options: f.options ? [...f.options] : [], visible: true,
    })) : []
  );
  const [outlookBody, setOutlookBody] = React.useState(editEvent ? (editEvent.outlookBody || '') : '');
  // Modal-State fuer den HTML-Editor (Outlook-Body + E-Mail-Templates)
  const [htmlEditorOpen, setHtmlEditorOpen] = React.useState(false);
  const [htmlEditorMode, setHtmlEditorMode] = React.useState<'outlook' | 'email'>('outlook');
  const [htmlEditorTemplateType, setHtmlEditorTemplateType] = React.useState<string>('');
  const [emailLanguage, setEmailLanguage] = React.useState(
    editEvent
      ? (editEvent.emailLanguage || (locale === 'de' ? 'DE' : 'EN'))
      : (locale === 'de' ? 'DE' : 'EN')
  );
  const [disableEmails, setDisableEmails] = React.useState(editEvent ? !!editEvent.disableEmails : false);
  const [disableOutlook, setDisableOutlook] = React.useState(editEvent ? !!editEvent.disableOutlook : false);
  const [isFictive, setIsFictive] = React.useState(editEvent ? !!editEvent.isFictive : false);
  // Nur im Edit-Modus: standardmaessig wird der Outlook-Termin NICHT angefasst,
  // damit bei kleinen Aenderungen (z.B. Description) nicht unnoetig eine
  // "Updated meeting"-Benachrichtigung an alle Teilnehmer geht. Der Organizer
  // muss die Checkbox aktiv setzen wenn er moechte dass Titel/Start/Ende im
  // Outlook-Termin aktualisiert werden.
  const [triggerOutlookUpdate, setTriggerOutlookUpdate] = React.useState(false);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [emailTemplates, setEmailTemplates] = React.useState<Array<{ id: number; templateType: string; language: string; subject: string; heading: string; headingColor: string; bodyHtml: string }>>([]);
  const [emailTemplateOverrides, setEmailTemplateOverrides] = React.useState<Record<string, { subject: string; heading: string; bodyHtml: string }>>(
    editEvent?.emailTemplateOverrides ? (() => { try { return JSON.parse(editEvent.emailTemplateOverrides); } catch { return {}; } })() : {}
  );
  // editingTemplate state entfaellt seit Modal-Migration v4.7.0
  const [emailLogoPreview, setEmailLogoPreview] = React.useState(() => {
    if (!editEvent?.emailTemplateOverrides) return '';
    try { const o = JSON.parse(editEvent.emailTemplateOverrides); return o._eventLogo || ''; } catch { return ''; }
  });
  const [dragFieldId, setDragFieldId] = React.useState<string | null>(null);
  const [dragOverFieldId, setDragOverFieldId] = React.useState<string | null>(null);
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
  const [quiz, setQuiz] = React.useState<Array<{id: string; question: string; options: string[]; correctIndices: number[]}>>(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    editEvent?.quiz?.map(q => ({...q, correctIndices: q.correctIndices || [(q as any).correctIndex || 0]})) || []
  );
  const [currentStep, setCurrentStep] = React.useState(0);
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
  const [showPreview, setShowPreview] = React.useState(false);
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
  const SUGGESTED_FIELDS_CATALOG: Array<{ key: string; label: string; description: string; build: (_now: number) => CustomFieldInput }> = isDe ? [
    {
      key: 'tshirt',
      label: 'T-Shirt Größe',
      description: 'Dropdown mit Kein T-Shirt / XS–XXL',
      build: (n) => ({ id: `cf-${n}`, label: 'T-Shirt Größe', type: 'select', required: false, options: ['Habe bereits ein T-Shirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'allergies',
      label: 'Allergien',
      description: 'Freitextfeld für Allergien/Unverträglichkeiten',
      build: (n) => ({ id: `cf-${n}`, label: 'Allergien', type: 'text', required: false, options: [], visible: true }),
    },
    {
      key: 'diet',
      label: 'Essenspräferenzen',
      description: 'Dropdown: Keine Präferenzen / Vegetarisch / Vegan / Pescetarisch',
      build: (n) => ({ id: `cf-${n}`, label: 'Essenspräferenzen', type: 'select', required: false, options: ['Keine Präferenzen', 'Vegetarisch', 'Vegan', 'Pescetarisch'], visible: true }),
    },
    {
      key: 'hotel',
      label: 'Hotel benötigt',
      description: 'Checkbox: Teilnehmer benötigt ein Hotel',
      build: (n) => ({ id: `cf-${n}`, label: 'Hotel benötigt', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'roomtype',
      label: 'Zimmerart',
      description: 'Dropdown: Keine Präferenz / Einzelzimmer / Doppelzimmer',
      build: (n) => ({ id: `cf-${n}`, label: 'Zimmerart (falls Hotel benötigt)', type: 'select', required: false, options: ['Keine Präferenz', 'Einzelzimmer', 'Doppelzimmer'], visible: true }),
    },
    {
      key: 'roommate',
      label: 'Bevorzugter Zimmerpartner',
      description: 'Personen-Suche; Match-Erkennung im Admin Center',
      build: (n) => ({ id: `cf-${n}`, label: 'Bevorzugter Zimmerpartner (bei Doppelzimmer)', type: 'user', required: false, options: [], visible: true }),
    },
  ] : [
    {
      key: 'tshirt',
      label: 'T-Shirt size',
      description: 'Dropdown: No t-shirt needed / XS–XXL',
      build: (n) => ({ id: `cf-${n}`, label: 'T-Shirt size', type: 'select', required: false, options: ['I already have one', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'allergies',
      label: 'Allergies',
      description: 'Free-text field for allergies / intolerances',
      build: (n) => ({ id: `cf-${n}`, label: 'Allergies', type: 'text', required: false, options: [], visible: true }),
    },
    {
      key: 'diet',
      label: 'Dietary preferences',
      description: 'Dropdown: No preference / Vegetarian / Vegan / Pescetarian',
      build: (n) => ({ id: `cf-${n}`, label: 'Dietary preferences', type: 'select', required: false, options: ['No preference', 'Vegetarian', 'Vegan', 'Pescetarian'], visible: true }),
    },
    {
      key: 'hotel',
      label: 'Hotel required',
      description: 'Checkbox: participant needs a hotel room',
      build: (n) => ({ id: `cf-${n}`, label: 'Hotel required', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'roomtype',
      label: 'Room type',
      description: 'Dropdown: No preference / Single room / Double room',
      build: (n) => ({ id: `cf-${n}`, label: 'Room type (if hotel needed)', type: 'select', required: false, options: ['No preference', 'Single room', 'Double room'], visible: true }),
    },
    {
      key: 'roommate',
      label: 'Preferred roommate',
      description: 'People search; match detection in the admin center',
      build: (n) => ({ id: `cf-${n}`, label: 'Preferred roommate (for double room)', type: 'user', required: false, options: [], visible: true }),
    },
  ];

  const [showSuggestedModal, setShowSuggestedModal] = React.useState(false);
  const [suggestedSelection, setSuggestedSelection] = React.useState<Record<string, boolean>>({});

  const openSuggestedModal = (): void => {
    // Standard: alle ausgewaehlt, User kann dann abwaehlen was er nicht braucht
    const init: Record<string, boolean> = {};
    for (const s of SUGGESTED_FIELDS_CATALOG) init[s.key] = true;
    setSuggestedSelection(init);
    setShowSuggestedModal(true);
  };

  const addSelectedSuggestedFields = (): void => {
    const selected = SUGGESTED_FIELDS_CATALOG.filter(s => suggestedSelection[s.key]);
    if (selected.length === 0) { setShowSuggestedModal(false); return; }
    const now = Date.now();
    const newFields: CustomFieldInput[] = selected.map((s, i) => s.build(now + i));
    // Haengen ans Ende an (keine Duplikate entfernen; User kann selbst loeschen wenn noetig)
    setCustomFields([...customFields, ...newFields]);
    setShowSuggestedModal(false);
  };

  const removeCustomField = (id: string): void => {
    setCustomFields(customFields.filter(f => f.id !== id));
  };

  const updateCustomField = (id: string, updates: Partial<CustomFieldInput>): void => {
    setCustomFields(customFields.map(f => f.id === id ? { ...f, ...updates } : f));
  };

  /**
   * Template-Auswahl: setzt EventType und Custom Fields automatisch.
   * B2Run: legt alle Pflichtfelder fuer die Anmeldung bei b2run.com an
   * (laut Excel "Deloitte_Teilnehmer_innen_B2Run_Koeln_2025_v4.xlsx").
   */
  const applyTemplate = (template: 'blank' | 'b2run'): void => {
    setSelectedTemplate(template);
    if (template === 'blank') {
      setCustomFields([]);
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
        { id: 'b2run_mobilnummer', label: 'Mobilnummer (nur bei aktiviertem Infoservice)', type: 'text', required: false, options: [], visible: true },
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
        { id: 'b2run_mobilnummer', label: 'Mobile number (only if info service is enabled)', type: 'text', required: false, options: [], visible: true },
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
      setCustomFields(fields);
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
  React.useEffect(() => {
    if (!isB2runTemplate) return;
    const d = parseInt(durchstarterCapacity, 10) || 0;
    const f = parseInt(funstarterCapacity, 10) || 0;
    const sum = d + f;
    if (sum > 0) setMaxParticipants(String(sum));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [durchstarterCapacity, funstarterCapacity, isB2runTemplate]);

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
    setMaxParticipants('50');
    setWaitlistEnabled(true);
    setEventImageUrl('');
    setCustomFields([
      { id: `cf-${Date.now()}`, label: 'T-Shirt Größe', type: 'select', required: false, options: ['Habe bereits ein T-Shirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true },
      { id: `cf-${Date.now() + 1}`, label: 'Allergien', type: 'text', required: false, options: [], visible: true },
      { id: `cf-${Date.now() + 2}`, label: 'Essenspräferenzen', type: 'select', required: false, options: ['Keine Präferenzen', 'Vegetarisch', 'Vegan', 'Pescetarisch'], visible: true },
      { id: `cf-${Date.now() + 3}`, label: 'Hotel benötigt', type: 'checkbox', required: false, options: [], visible: true },
      { id: `cf-${Date.now() + 4}`, label: 'Zimmerart (falls Hotel benötigt)', type: 'select', required: false, options: ['Keine Präferenz', 'Einzelzimmer', 'Doppelzimmer'], visible: true },
      { id: `cf-${Date.now() + 5}`, label: 'Bevorzugter Zimmerpartner (bei Doppelzimmer)', type: 'user', required: false, options: [], visible: true },
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
  const updateQuizQuestion = (id: string, updates: Partial<{question: string; options: string[]; correctIndices: number[]}>): void => {
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

  const handleSubmit = async (): Promise<void> => {
    if (!title || !description) return;
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
        'Organizer': organizer,
        'CustomFields': JSON.stringify(customFields
          .filter(f => f.label && f.label.trim().length > 0)
          .map(f => ({
            id: f.id, label: f.label.trim(), type: f.type, required: f.required, visible: f.visible,
            ...(f.type === 'select' ? { options: f.options.map(o => o.trim()).filter(Boolean) } : {}),
          }))),
      };

      // Optionale Felder - immer senden damit Loeschungen wirken
      updates['LastDeregisterDate'] = deadlineToEndOfDayIso(lastDeregisterDate);
      updates['OutlookBody'] = outlookBody ? buildOutlookBody(title, outlookBody) : '';
      updates['Agenda'] = JSON.stringify(agenda);
      updates['Transfers'] = JSON.stringify(transferTimes);
      updates['FunZone'] = JSON.stringify(quiz);
      updates['EmailLanguage'] = emailLanguage;
      updates['EmailTemplateOverrides'] = (Object.keys(emailTemplateOverrides).length > 0 || emailLogoPreview)
        ? JSON.stringify({ ...(emailLogoPreview ? { _eventLogo: emailLogoPreview } : {}), ...emailTemplateOverrides })
        : '';
      updates['DisableEmails'] = disableEmails;
      updates['DisableOutlook'] = disableOutlook;
      updates['IsFictive'] = isFictive;
      if (isB2runTemplate) {
        updates['DurchstarterCapacity'] = parseInt(durchstarterCapacity, 10) || 0;
        updates['FunstarterCapacity'] = parseInt(funstarterCapacity, 10) || 0;
      }

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
        }
        setProgress(100);
        setProgressLabel('Änderungen gespeichert!');
        setTimeout(() => { setIsSubmitting(false); setSubmitted(true); }, 500);
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
        organizer,
        organizerEmail: organizerEmails.join(';'),
        outlookEventId: '',
        outlookBody,
        agenda: JSON.stringify(agenda),
        transfers: JSON.stringify(transferTimes),
        documents: '[]', // Dokumente werden nach erfolgreichem Upload gespeichert
        funZone: JSON.stringify(quiz),
        emailLanguage,
        emailTemplateOverrides: (Object.keys(emailTemplateOverrides).length > 0 || emailLogoPreview)
          ? JSON.stringify({ ...(emailLogoPreview ? { _eventLogo: emailLogoPreview } : {}), ...emailTemplateOverrides })
          : '',
        disableEmails,
        disableOutlook,
        isFictive,
        durchstarterCapacity: isB2runTemplate ? (parseInt(durchstarterCapacity, 10) || 0) : undefined,
        funstarterCapacity: isB2runTemplate ? (parseInt(funstarterCapacity, 10) || 0) : undefined,
        customFields: customFields
          .filter(f => f.label && f.label.trim().length > 0)
          .map(f => ({
            id: f.id, label: f.label.trim(), type: f.type, required: f.required, visible: f.visible,
            ...(f.type === 'select' ? { options: f.options.map(o => o.trim()).filter(Boolean) } : {}),
          })),
      });

      clearInterval(progressTimer);
      if (eventId) {
        setProgress(100);
        setProgressLabel('Event erfolgreich erstellt!');
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
                  // Events neu laden, damit das gerade hochgeladene Bild sofort sichtbar ist
                  await refreshEvents();
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
        // Kurz 100% zeigen, dann zur Erfolgsseite
        setTimeout(() => {
          setIsSubmitting(false);
          setSubmitted(true);
        }, 500);
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
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center' }}>
            <button className="btn btn-primary" onClick={() => navigate('register')}>Events anzeigen</button>
            <button className="btn btn-secondary" onClick={() => { setSubmitted(false); setTitle(''); }}>Weiteres Event erstellen</button>
          </div>
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
                    {field.type === 'select' ? (
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


  const steps = [
    { label: t('create.step.basics'), icon: '1' },
    { label: t('create.step.datetime'), icon: '2' },
    { label: t('create.step.capacity'), icon: '3' },
    { label: t('create.step.fields'), icon: '4' },
    { label: t('create.step.communication'), icon: '5' },
    { label: t('create.step.documents'), icon: '6' },
    { label: t('create.step.funzone'), icon: '7' },
  ];

  const getStepErrors = (): string[] => {
    const errors: string[] = [];
    switch (currentStep) {
      case 0:
        if (!title) errors.push('title');
        if (!organizer) errors.push('organizer');
        if (!description) errors.push('description');
        break;
      case 1:
        if (!startDate) errors.push('startDate');
        if (!endDate) errors.push('endDate');
        if (startDate && endDate && new Date(endDate) <= new Date(startDate)) errors.push('endBeforeStart');
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

  return (
    <div className="page-container">
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
                }}>
                  {step.label}
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

              {!isEditMode && (
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

              {/* ===== Step 0: Grundlagen ===== */}
              <div style={{ display: currentStep === 0 ? 'block' : 'none' }}>

              <div className="form-group">
                <label className="form-label">
                  <span className="required">*</span> {t('create.eventtitle')}
                  <InfoTooltip text={t('create.eventtitle.hint')} />
                </label>
                <input className="form-input" value={title} onChange={e => setTitle(e.target.value)} placeholder="z.B. Sommerfest 2026" style={errorBorderStyle('title')} />
                {fieldHasError('title') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
              </div>

              {/* Test-Event-Flag */}
              <div className="form-group" style={{ marginTop: 0, padding: 14, background: isFictive ? 'rgba(237,139,0,0.06)' : 'var(--dex-gray-50, #f8f9fa)', borderRadius: 'var(--dex-radius, 12px)', border: `1px solid ${isFictive ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-200)'}` }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
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
              </div>

              <div className="form-group" style={{ position: 'relative' }}>
                <label className="form-label">
                  <span className="required">*</span> {t('create.organizer')}
                  <InfoTooltip text={t('create.organizer.hint')} />
                </label>
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
                    const nextNames = orgList.filter((_, i) => i !== idx);
                    setOrganizer(nextNames.join('; '));
                    setOrganizerEmails(prev => prev.filter((_, i) => i !== idx));
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
                    const lp = val.trim().toLowerCase();
                    const filtered = roles
                      .filter(r => r.role === 'Organizer' || r.role === 'Admin')
                      .filter(r => !lp || r.userEmail.toLowerCase().indexOf(lp) >= 0 || (r.userName || '').toLowerCase().indexOf(lp) >= 0)
                      .slice(0, 50)
                      .map(r => ({ email: r.userEmail, displayName: r.userName || r.userEmail, location: r.location || '' }));
                    setOrganizerResults(filtered);
                  }}
                  onFocus={() => {
                    // Bei Fokus direkt alle verfuegbaren Organizer/Admins anzeigen.
                    const filtered = roles
                      .filter(r => r.role === 'Organizer' || r.role === 'Admin')
                      .slice(0, 50)
                      .map(r => ({ email: r.userEmail, displayName: r.userName || r.userEmail, location: r.location || '' }));
                    setOrganizerResults(filtered);
                  }}
                  onBlur={() => {
                    // Freitext verwerfen — nur per Dropdown ausgewaehlte Organizer zaehlen.
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

              <div className="form-group">
                <label className="form-label">
                  Standort-Filter
                </label>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                  Standardmäßig können <strong>alle Mitarbeiter</strong> dieses Event sehen. Wenn du hier Standorte auswählst, wird das Event <strong>nur für Mitarbeiter dieser Standorte</strong> sichtbar.<br />
                  <em>Beispiel: Du wählst &bdquo;Köln&ldquo; und &bdquo;Düsseldorf&ldquo; → Nur Mitarbeiter mit Standort Köln oder Düsseldorf sehen das Event in ihrer Übersicht. Alle anderen sehen es nicht.</em>
                </p>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                  {locationOptions.map(loc => {
                    const isChecked = locationFilter.split(',').map(s => s.trim()).indexOf(loc) >= 0;
                    const toggle = (): void => {
                      const current = locationFilter.split(',').map(s => s.trim()).filter(Boolean);
                      if (!isChecked) setLocationFilter([...current, loc].join(', '));
                      else setLocationFilter(current.filter(l => l !== loc).join(', '));
                    };
                    return (
                      <button
                        key={loc}
                        type="button"
                        onClick={toggle}
                        aria-pressed={isChecked}
                        style={{
                          display: 'inline-flex', alignItems: 'center', gap: 6,
                          padding: '6px 14px',
                          borderRadius: 999,
                          border: `1.5px solid ${isChecked ? 'var(--dex-green)' : 'var(--dex-gray-300)'}`,
                          background: isChecked ? 'var(--dex-green)' : '#fff',
                          color: isChecked ? '#fff' : 'var(--dex-gray-700)',
                          fontSize: '0.85rem',
                          fontWeight: isChecked ? 500 : 400,
                          cursor: 'pointer',
                          transition: 'all 0.15s ease',
                          boxShadow: isChecked ? '0 1px 3px rgba(134,188,37,0.25)' : 'none',
                        }}
                      >
                        {isChecked && <Check size={14} />}
                        {loc}
                      </button>
                    );
                  })}
                </div>
                {!locationFilter && (
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-green)', marginTop: 8 }}>
                    Kein Standort ausgewählt → Event ist für alle sichtbar.
                  </p>
                )}
                {(locationFilter || audience) && (
                  <button
                    className="btn btn-outline mt-8"
                    style={{ fontSize: '0.8rem', padding: '4px 12px' }}
                    onClick={() => setShowEmailModal(true)}
                    type="button"
                  >
                    <Users size={14} /> Zielgruppe prüfen
                  </button>
                )}
              </div>

              <div className="form-group" style={{ position: 'relative' }}>
                <label className="form-label">
                  Zielgruppen-Filter
                </label>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                  Suche nach <strong>Personen oder Gruppen</strong> (Verteilerlisten / Security Groups aus Entra) — diese sehen das Event zusätzlich, unabhängig vom Standort.
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
                    onClick={() => { setBulkImportText(''); setBulkImportReport(null); setBulkImportOpen(true); }}
                  >
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}><Users size={12} /> Massenimport (Liste einfügen)</span>
                  </button>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', margin: 0, flex: 1 }}>
                    Klicke einen Treffer an um ihn hinzuzufügen. Bei Gruppen kannst du im Chip per <Users size={11} /> die Mitglieder einsehen.
                    Statt zu suchen kannst du auch direkt die Verteiler-Mail eintippen (z.B. SAPAlliance@deloitte.com) oder Sondergruppen wie <code>DEALL</code>, <code>DEKOELN</code>.
                  </p>
                </div>
              </div>

              {/* UND/ODER Verknüpfung */}
              {locationFilter && audience && (
                <div className="form-group">
                  <label className="form-label">
                    Filterverknüpfung
                  </label>
                  <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                    Bestimmt, wie Standort-Filter und Zielgruppen-Filter kombiniert werden.<br />
                    <em>Beispiel ODER: Standort = Köln, Zielgruppe = SAPALL → Jeder der in Köln sitzt ODER in der SAP-Gruppe ist, sieht das Event.<br />
                    Beispiel UND: Nur wer in Köln sitzt UND in der SAP-Gruppe ist, sieht das Event.</em>
                  </p>
                  <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.9rem', cursor: 'pointer' }}>
                      <input type="radio" name="filterMode" value="OR" checked={filterMode === 'OR'} onChange={() => setFilterMode('OR')} />
                      <strong>ODER</strong>
                      <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>– Einer der Filter reicht</span>
                    </label>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.9rem', cursor: 'pointer' }}>
                      <input type="radio" name="filterMode" value="AND" checked={filterMode === 'AND'} onChange={() => setFilterMode('AND')} />
                      <strong>UND</strong>
                      <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>– Beides muss zutreffen</span>
                    </label>
                  </div>
                </div>
              )}

              <div className="form-group">
                <label className="form-label">
                  <span className="required">*</span> {t('create.description')}
                  <InfoTooltip text={t('create.description.hint')} />
                </label>
                <textarea className="form-textarea" value={description} onChange={e => setDescription(e.target.value)} style={{ minHeight: 120, ...errorBorderStyle('description') }} />
                {fieldHasError('description') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
              </div>

              <div className="form-group">
                <label className="form-label">
                  Event-Bild
                  <InfoTooltip text={t('create.eventimage.hint')} />
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

              </div>

              {/* ===== Step 1: Zeit & Ort ===== */}
              <div style={{ display: currentStep === 1 ? 'block' : 'none' }}>
              <div className="form-group">
                <label className="form-label">
                  {t('create.location')}
                  <InfoTooltip text={t('create.location.hint')} />
                </label>
                <input className="form-input" value={location} onChange={e => setLocation(e.target.value)} placeholder="z.B. RheinEnergieStadion, Köln" />
              </div>
              <div className="form-group">
                <label className="form-label">
                  Adresse
                  <InfoTooltip text={t('create.address.hint')} />
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

              <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                <div className="form-group">
                  <label className="form-label">
                    <span className="required">*</span> {t('create.startdate')}
                    <InfoTooltip text={t('create.startdate.hint')} />
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
                    maxDate={endDate ? new Date(endDate) : undefined}
                    isClearable
                    autoComplete="off"
                  />
                  {fieldHasError('startDate') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
                </div>
                <div className="form-group">
                  <label className="form-label">
                    <span className="required">*</span> {t('create.enddate')}
                    <InfoTooltip text={t('create.enddate.hint')} />
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
              <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginTop: -8, marginBottom: 12 }}>
                Die Uhrzeit wird für den Outlook-Kalendereintrag der Teilnehmer verwendet.
              </p>

              {/* ===== Agenda Editor ===== */}
              <div className="form-group" style={{ marginTop: 24 }}>
                <label className="form-label" style={{ fontSize: '1rem', fontWeight: 700 }}>
                  {t('create.agenda')}
                  <InfoTooltip text={t('create.agenda.hint')} />
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
                <label className="form-label" style={{ fontSize: '1rem', fontWeight: 700 }}>
                  {t('create.transfers')}
                  <InfoTooltip text={t('create.transfers.hint')} />
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

              </div>

              {/* ===== Step 2: Kapazität & Fristen ===== */}
              <div style={{ display: currentStep === 2 ? 'block' : 'none' }}>
              <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                <div className="form-group">
                  <label className="form-label">
                    {t('create.deadline')}
                    <InfoTooltip text={t('create.deadline.hint')} />
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
                    maxDate={startDate ? new Date(startDate) : undefined}
                  />
                </div>
                <div className="form-group">
                  <label className="form-label">
                    {t('create.lastcancel')}
                    <InfoTooltip text={t('create.lastcancel.hint')} />
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
                    maxDate={startDate ? new Date(startDate) : undefined}
                  />
                </div>
              </div>
              {fieldHasError('deadlineAfterStart') && <p style={{ color: 'var(--dex-red)', fontSize: '0.8rem', marginTop: -4, marginBottom: 8 }}>{t('create.error.deadlineAfterStart')}</p>}
              {fieldHasError('deregAfterStart') && <p style={{ color: 'var(--dex-red)', fontSize: '0.8rem', marginTop: -4, marginBottom: 8 }}>{t('create.error.deregAfterStart')}</p>}

              {/* B2Run: Split-Kapazitaeten fuer Durchstarter + Funstarter */}
              {isB2runTemplate ? (
                <div style={{ padding: 16, background: 'var(--dex-green-light, #f0fdf4)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-green)', marginBottom: 16 }}>
                  <label className="form-label" style={{ marginBottom: 4 }}>
                    {t('create.b2runcap')}
                    <InfoTooltip text={t('create.b2runcap.hint')} />
                  </label>
                  <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                    {t('create.b2runcap.desc')}
                  </p>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                    <div>
                      <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                        <Icon iconName="Running" style={{ fontSize: 14, marginRight: 6, color: 'var(--dex-green-dark, #6b9a1e)' }} />
                        {t('create.b2runcap.durchstarter')}
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
                        <Icon iconName="Running" style={{ fontSize: 14, marginRight: 6, color: 'var(--dex-orange, #ff8c00)' }} />
                        {t('create.b2runcap.funstarter')}
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
                  <div style={{ marginTop: 12 }}>
                    <label className="form-label" style={{ marginBottom: 4 }}>
                      {t('create.waitlist')}
                      <InfoTooltip text={t('create.b2runcap.waitlist.hint')} />
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
                        <InfoTooltip text={t('create.waitlist.hint')} />
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

              </div>

              {/* ===== Step 3: Registrierungsfelder ===== */}
              <div style={{ display: currentStep === 3 ? 'block' : 'none' }}>

              {/* Vorlage: B2Run-Felder automatisch befuellen (als moderne Checkbox) */}
              {!isEditMode && (
                <div className="form-group" style={{ marginBottom: 20 }}>
                  <label className="form-label" style={{ marginBottom: 8 }}>
                    {t('create.template')}
                    <InfoTooltip text={t('create.template.hint')} />
                  </label>
                  <label
                    style={{
                      display: 'flex', alignItems: 'flex-start', gap: 12,
                      padding: 14, borderRadius: 'var(--dex-radius, 12px)',
                      border: `2px solid ${selectedTemplate === 'b2run' ? 'var(--dex-green)' : 'var(--dex-gray-200)'}`,
                      background: selectedTemplate === 'b2run' ? 'rgba(134,188,37,0.08)' : '#fff',
                      cursor: 'pointer', transition: 'all 0.15s ease',
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={selectedTemplate === 'b2run'}
                      onChange={e => applyTemplate(e.target.checked ? 'b2run' : 'blank')}
                      style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 2, accentColor: '#86bc25' }}
                    />
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: '0.95rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>
                        {t('create.template.b2run')}
                      </div>
                      <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, marginTop: 3 }}>
                        {t('create.template.b2run.desc')}
                      </div>
                    </div>
                  </label>
                </div>
              )}

              {/* B2Run Startbloecke - moderne Liste mit + Button */}
              {(selectedTemplate === 'b2run' || (isEditMode && customFields.some(f => f.id === 'b2run_startblock'))) && (
                <div className="form-group" style={{ marginBottom: 24, padding: 16, background: 'var(--dex-green-light, #f0fdf4)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-green)' }}>
                  <label className="form-label" style={{ marginBottom: 4 }}>
                    {t('create.startblocks')}
                    <InfoTooltip text={t('create.startblocks.hint')} />
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
                <div className="flex-between mb-16">
                  <label className="form-label" style={{ marginBottom: 0 }}>{t('create.customfields')}</label>
                  <div style={{ display: 'flex', gap: 8 }}>
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
                  </div>
                </div>
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
                    {/* Feld-Header: Drag + Name + Typ + Pflicht + Löschen */}
                    <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: field.type === 'select' ? 12 : 0 }}>
                      <div
                        style={{ cursor: 'grab', padding: '0 4px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 1 }}
                        title="Ziehen zum Verschieben"
                      >
                        <span style={{ fontSize: '1rem', color: 'var(--dex-gray-400)', userSelect: 'none', lineHeight: 1 }}>⠿</span>
                        <div style={{ display: 'flex', gap: 2 }}>
                          <button
                            onClick={() => moveCustomField(field.id, 'up')}
                            disabled={idx === 0}
                            style={{ background: 'none', border: 'none', padding: 0, color: idx === 0 ? 'var(--dex-gray-300)' : 'var(--dex-gray-600)', cursor: idx === 0 ? 'default' : 'pointer', fontSize: '0.7rem' }}
                          >▲</button>
                          <button
                            onClick={() => moveCustomField(field.id, 'down')}
                            disabled={idx === customFields.length - 1}
                            style={{ background: 'none', border: 'none', padding: 0, color: idx === customFields.length - 1 ? 'var(--dex-gray-300)' : 'var(--dex-gray-600)', cursor: idx === customFields.length - 1 ? 'default' : 'pointer', fontSize: '0.7rem' }}
                          >▼</button>
                        </div>
                      </div>
                      <input className="form-input" placeholder={t('create.fieldname')} value={field.label} onChange={e => updateCustomField(field.id, { label: e.target.value })} style={{ flex: 2 }} />
                      <select className="form-select" value={field.type} onChange={e => updateCustomField(field.id, { type: e.target.value as CustomFieldInput['type'] })} style={{ flex: 1, maxWidth: 140 }}>
                        <option value="text">Text</option>
                        <option value="select">Dropdown</option>
                        <option value="number">Zahl</option>
                        <option value="checkbox">Checkbox</option>
                        <option value="user">Person (Suche)</option>
                      </select>
                      <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: '0.85rem', whiteSpace: 'nowrap' }}>
                        <input type="checkbox" checked={field.required} onChange={e => updateCustomField(field.id, { required: e.target.checked })} />
                        {t('create.required')}
                      </label>
                      <button onClick={() => removeCustomField(field.id)} style={{ background: 'none', border: 'none', color: 'var(--dex-red)', padding: 4, cursor: 'pointer' }}>
                        <X size={18} />
                      </button>
                    </div>

                    {/* Dropdown-Optionen als Tag-Liste */}
                    {field.type === 'select' && (
                      <div style={{ marginLeft: 32, paddingTop: 8, borderTop: '1px solid var(--dex-gray-200)' }}>
                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginBottom: 8, fontWeight: 600 }}>Dropdown-Optionen:</div>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 8 }}>
                          {/* Alle Optionen aus dem Array rendern (auch leere Slots) */}
                          {(field.options || []).map((opt, optIdx) => (
                            <div key={optIdx} style={{
                              display: 'flex', alignItems: 'center', gap: 4,
                              background: '#fff', border: '1px solid var(--dex-gray-300)',
                              borderRadius: 20, padding: '4px 8px 4px 12px', fontSize: '0.85rem',
                            }}>
                              <input
                                value={opt}
                                placeholder={`Option ${optIdx + 1}`}
                                onChange={e => {
                                  const opts = [...(field.options || [])];
                                  opts[optIdx] = e.target.value;
                                  updateCustomField(field.id, { options: opts });
                                }}
                                style={{
                                  border: 'none', background: 'transparent', outline: 'none',
                                  width: Math.max(60, (opt.length + 2) * 8), fontSize: '0.85rem',
                                }}
                              />
                              <button
                                onClick={() => {
                                  const opts = [...(field.options || [])];
                                  opts.splice(optIdx, 1);
                                  updateCustomField(field.id, { options: opts });
                                }}
                                style={{ background: 'none', border: 'none', color: 'var(--dex-gray-400)', padding: 0, cursor: 'pointer', fontSize: '0.9rem', lineHeight: 1 }}
                              >
                                ×
                              </button>
                            </div>
                          ))}
                          <button
                            onClick={() => {
                              // Genau EINEN leeren Slot anhaengen
                              updateCustomField(field.id, { options: [...(field.options || []), ''] });
                            }}
                            style={{
                              display: 'flex', alignItems: 'center', gap: 4,
                              background: 'none', border: '2px dashed var(--dex-gray-300)',
                              borderRadius: 20, padding: '4px 12px', fontSize: '0.8rem',
                              color: 'var(--dex-gray-500)', cursor: 'pointer',
                            }}
                          >
                            + Option
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                ))}
              </div>
              </div>{/* close Step 3 */}

              {/* ===== Step 4: Kommunikation ===== */}
              <div style={{ display: currentStep === 4 ? 'block' : 'none' }}>
                <h3 className="mb-16">{t('create.step.communication')}</h3>

                <div className="form-group">
                  <label className="form-label">{t('create.emaillanguage')}</label>
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

                {/* Benachrichtigungen abschalten */}
                <div className="form-group" style={{ marginTop: 24, padding: 16, background: 'var(--dex-gray-50, #f8f9fa)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-gray-200)' }}>
                  <label className="form-label" style={{ marginBottom: 8 }}>Benachrichtigungen</label>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                    Bei deaktivierten Optionen werden bei An- oder Abmeldungen keine Eintraege in die DEX_Emails- bzw. DEX_Outlook-Warteschlange geschrieben.
                  </p>
                  <label style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer', marginBottom: 8 }}>
                    <input
                      type="checkbox"
                      checked={!disableEmails}
                      onChange={e => setDisableEmails(!e.target.checked)}
                      style={{ width: 18, height: 18, cursor: 'pointer' }}
                    />
                    <span style={{ fontSize: '0.9rem' }}>
                      <strong>E-Mails versenden</strong>
                      <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
                        Anmelde-, Abmelde-, Warteliste- und Nachrueck-E-Mails
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
                      <strong>Outlook-Kalendereintrag senden</strong>
                      <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
                        Einladung bei Anmeldung, Ausladung bei Abmeldung
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
                        <strong>Outlook-Termin aktualisieren (Titel/Start/Ende)</strong>
                        <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-500)', lineHeight: 1.4, marginTop: 2 }}>
                          Nur aktivieren wenn Titel, Startzeit oder Endzeit wirklich geändert wurden.
                          Bei Aktivierung erhalten <strong>alle angemeldeten Teilnehmer</strong> automatisch eine
                          &bdquo;Updated meeting&ldquo;-Benachrichtigung von Outlook.
                          Standardmäßig ausgeschaltet, damit Tippfixes in Description/Agenda keine Update-Mails auslösen.
                        </span>
                      </span>
                    </label>
                  )}
                </div>

                <div className="form-group" style={{ marginTop: 24 }}>
                  <label className="form-label">{t('create.eventlogo')}</label>
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 8 }}>
                    {t('create.eventlogo.hint')}
                  </p>
                  {emailLogoPreview && (
                    <div style={{ marginBottom: 8, display: 'flex', alignItems: 'center', gap: 12 }}>
                      <img src={emailLogoPreview} alt="Event-Logo" style={{ maxWidth: 180, maxHeight: 80, borderRadius: 4 }} />
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
                      const compressed = await compressImage(file, 200, 0.85);
                      const reader = new FileReader();
                      reader.onload = (ev) => setEmailLogoPreview(ev.target?.result as string || '');
                      reader.readAsDataURL(compressed);
                    }} />
                  </label>
                </div>

                <div className="form-group" style={{ marginTop: 24 }}>
                  <label className="form-label">{t('create.outlookdesc')}</label>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                    <button
                      type="button"
                      className="btn btn-secondary"
                      onClick={() => { setHtmlEditorMode('outlook'); setHtmlEditorOpen(true); }}
                      style={{ fontSize: '0.85rem' }}
                    >
                      Bearbeiten & Vorschau
                    </button>
                    <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)' }}>
                      {outlookBody
                        ? `${outlookBody.replace(/<[^>]+>/g, '').substring(0, 80)}${outlookBody.length > 80 ? '…' : ''}`
                        : t('create.outlookdesc.placeholder')}
                    </span>
                  </div>
                </div>

                <h4 style={{ marginTop: 24, marginBottom: 12 }}>{t('create.templates.title')} ({emailLanguage})</h4>
                <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 12 }}>
                  {t('create.templates.hint')}
                </p>

                {/* TemplateType in DEX_EmailTemplates ist ASCII 'Nachruecken' (Umlaut nicht erlaubt in Choice-Feld). */}
                {['Anmeldung', 'Warteliste', 'Abmeldung', 'Nachruecken'].map(tType => {
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
              </div>{/* close Step 4 */}

              {/* ===== Step 5: Dokumente ===== */}
              <div style={{ display: currentStep === 5 ? 'block' : 'none' }}>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', marginBottom: 16 }}>
                  {t('create.documents.hint')}
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
              </div>{/* close Step 5 */}

              {/* ===== Step 6: Fun-Zone ===== */}
              <div style={{ display: currentStep === 6 ? 'block' : 'none' }}>
                <h3 className="mb-16">{t('create.step.funzone')}</h3>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 16 }}>
                  {t('create.funzone.hint')}
                </p>

                {quiz.map((q, qi) => (
                  <div key={q.id} style={{
                    padding: 16, marginBottom: 12, background: 'var(--dex-gray-50, #fafafa)',
                    borderRadius: 12, border: '1px solid var(--dex-gray-200)',
                  }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                      <label style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
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
                    <label style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginBottom: 4, display: 'block' }}>
                      {t('create.funzone.options')}
                    </label>
                    {q.options.map((opt, oi) => (
                      <div key={oi} style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 6 }}>
                        <input
                          type="checkbox"
                          checked={q.correctIndices?.includes(oi) || false}
                          onChange={() => {
                            const indices = q.correctIndices || [];
                            const newIndices = indices.includes(oi) ? indices.filter(x => x !== oi) : [...indices, oi];
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
                            const newOpts = q.options.filter((_, i) => i !== oi);
                            const newCorrect = (q.correctIndices || []).filter(ci => ci !== oi).map(ci => ci > oi ? ci - 1 : ci);
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
                ))}

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
            <div className="registration-actions mt-24">
              {currentStep === 0 ? (
                <button className="btn btn-danger" onClick={() => goBack()}><Trash2 size={16} /> {t('create.cancel')}</button>
              ) : (
                <button className="btn btn-secondary" onClick={() => setCurrentStep(currentStep - 1)}>
                  {t('general.back')}
                </button>
              )}

              {/* Im Edit-Modus immer einen Speichern-Button anzeigen, damit man nicht
                  durch alle Steps klicken muss wenn man nur eine Sache aendert */}
              {isEditMode && currentStep < steps.length - 1 && (
                <button
                  className="btn btn-primary"
                  disabled={!title || !description}
                  onClick={handleSubmit}
                  style={{ opacity: !title || !description ? 0.5 : 1 }}
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
                <>
                  <button
                    className="btn btn-secondary"
                    disabled={!title}
                    onClick={() => setShowPreview(true)}
                  >
                    {t('create.preview')}
                  </button>
                  <button
                    className="btn btn-primary"
                    disabled={!title || !description}
                    onClick={handleSubmit}
                    style={{ opacity: !title || !description ? 0.5 : 1 }}
                  >
                    <Send size={16} /> {isEditMode ? t('create.save') : t('create.submit')}
                  </button>
                </>
              )}
            </div>
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
                disabled={!title || !description}
                onClick={() => { setShowPreview(false); handleSubmit(); }}
              >
                <Send size={16} /> {isEditMode ? t('create.save') : t('create.submit')}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* HTML-Editor-Modal mit Live-Preview (Outlook-Termin oder E-Mail-Template) */}
      {(() => {
        if (!htmlEditorOpen) return null;
        const isOutlook = htmlEditorMode === 'outlook';
        const tType = htmlEditorTemplateType;
        const defaultTpl = !isOutlook ? emailTemplates.find(tp => tp.templateType === tType && tp.language === emailLanguage) : undefined;
        const override = !isOutlook ? emailTemplateOverrides[tType] : undefined;
        const currentSubject = override?.subject || defaultTpl?.subject || '';
        const currentHeading = override?.heading || defaultTpl?.heading || '';
        const currentBody = isOutlook
          ? outlookBody
          : (override?.bodyHtml || defaultTpl?.bodyHtml || '');
        return (
          <HtmlEditorModal
            open={htmlEditorOpen}
            onClose={() => setHtmlEditorOpen(false)}
            title={isOutlook ? 'Outlook-Termin: Body bearbeiten' : `E-Mail-Template: ${tType}`}
            value={currentBody}
            onChange={(html) => {
              if (isOutlook) {
                setOutlookBody(html);
              } else {
                setEmailTemplateOverrides({
                  ...emailTemplateOverrides,
                  [tType]: { subject: currentSubject, heading: currentHeading, bodyHtml: html },
                });
              }
            }}
            previewMode={isOutlook ? 'outlook' : 'email'}
            emailSubject={!isOutlook ? currentSubject : undefined}
            onEmailSubjectChange={!isOutlook ? (s) => setEmailTemplateOverrides({
              ...emailTemplateOverrides,
              [tType]: { subject: s, heading: currentHeading, bodyHtml: currentBody },
            }) : undefined}
            emailHeading={!isOutlook ? currentHeading : undefined}
            onEmailHeadingChange={!isOutlook ? (h) => setEmailTemplateOverrides({
              ...emailTemplateOverrides,
              [tType]: { subject: currentSubject, heading: h, bodyHtml: currentBody },
            }) : undefined}
            emailHeadingColor={!isOutlook ? (defaultTpl?.headingColor || '#86bc25') : undefined}
            previewVars={{
              EventTitle: title || 'Event Title',
              Name: 'Max Mustermann',
              Organizer: organizer || 'Organisator',
              AppUrl: 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView',
              WaitlistPosition: '1',
              EventDate: startDate ? new Date(startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
            }}
            insertableVars={[
              { key: '{{Name}}', label: 'Name' },
              { key: '{{EventTitle}}', label: 'Event' },
              { key: '{{Organizer}}', label: 'Organizer' },
              { key: '{{AppUrl}}', label: 'App Link' },
              { key: '{{WaitlistPosition}}', label: 'Waitlist #' },
            ]}
            imageBase64={imagePreview || ''}
          />
        );
      })()}

      {/* Massenimport-Modal fuer Audience */}
      {bulkImportOpen && (
        <div style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 1100,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }} onClick={() => { if (!bulkImportRunning) setBulkImportOpen(false); }}>
          <div className="card" style={{ width: '90%', maxWidth: 720, maxHeight: '85vh', overflow: 'auto', padding: 24 }} onClick={e => e.stopPropagation()}>
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0 }}>Massenimport Zielgruppe</h3>
              <button
                type="button"
                onClick={() => { if (!bulkImportRunning) setBulkImportOpen(false); }}
                style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4 }}
                aria-label="Schließen"
              >
                <X size={18} />
              </button>
            </div>
            <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', marginTop: 0 }}>
              Füge eine Liste aus Namen und/oder Email-Adressen ein, getrennt mit
              <code style={{ margin: '0 4px' }}>,</code>
              <code style={{ margin: '0 4px' }}>;</code>
              <code style={{ margin: '0 4px' }}>Tab</code>
              oder <code style={{ marginLeft: 4 }}>Zeilenumbruch</code>.
              Email-Adressen werden direkt übernommen. Bei Namen wird per People-Picker gesucht — eindeutige
              Treffer werden automatisch hinzugefügt, mehrdeutige Namen musst du unten manuell auflösen.
            </p>
            <textarea
              className="form-input"
              style={{ width: '100%', minHeight: 160, fontFamily: 'monospace', fontSize: '0.8rem' }}
              placeholder="z.B.:&#10;max.mustermann@deloitte.de; erika.mustermann@deloitte.de&#10;Schmitz, Alexander, Kraus, Annika&#10;oder aus Excel kopiert (Tab-getrennt)"
              value={bulkImportText}
              onChange={e => setBulkImportText(e.target.value)}
              disabled={bulkImportRunning}
            />

            {bulkImportReport && (
              <div style={{ marginTop: 16, padding: 12, background: 'var(--dex-gray-50)', borderRadius: 'var(--dex-radius)', fontSize: '0.8rem' }}>
                {bulkImportReport.added.length > 0 && (
                  <div style={{ marginBottom: 10 }}>
                    <strong style={{ color: 'var(--dex-green-dark)' }}>✓ Hinzugefügt ({bulkImportReport.added.length}):</strong>
                    <ul style={{ margin: '4px 0 0 16px', padding: 0 }}>
                      {bulkImportReport.added.map((a, i) => (
                        <li key={`a-${i}`}>
                          <strong>{a.lastname}</strong>{a.firstname ? `, ${a.firstname}` : ''} <span style={{ color: 'var(--dex-gray-400)' }}>— {a.email}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                )}
                {bulkImportReport.alreadyIn.length > 0 && (
                  <div style={{ marginBottom: 10 }}>
                    <strong style={{ color: 'var(--dex-gray-500)' }}>— Bereits in Zielgruppe ({bulkImportReport.alreadyIn.length}):</strong>
                    <ul style={{ margin: '4px 0 0 16px', padding: 0, color: 'var(--dex-gray-500)' }}>
                      {bulkImportReport.alreadyIn.map((a, i) => (
                        <li key={`w-${i}`}>
                          <strong>{a.lastname}</strong>{a.firstname ? `, ${a.firstname}` : ''} <span>— {a.email}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                )}
                {bulkImportReport.notFound.length > 0 && (
                  <div style={{ marginBottom: 10 }}>
                    <strong style={{ color: 'var(--dex-red)' }}>✗ Nicht gefunden ({bulkImportReport.notFound.length}):</strong>
                    <ul style={{ margin: '4px 0 0 16px', padding: 0, color: 'var(--dex-red)' }}>
                      {bulkImportReport.notFound.map((a, i) => <li key={`n-${i}`}>{a} — bitte manuell suchen oder als Email eintragen</li>)}
                    </ul>
                  </div>
                )}
                {bulkImportReport.ambiguous.length > 0 && (
                  <div style={{ marginBottom: 10 }}>
                    <strong style={{ color: 'var(--dex-orange, #ed8b00)' }}>? Mehrdeutig — bitte auswählen ({bulkImportReport.ambiguous.length}):</strong>
                    {bulkImportReport.ambiguous.map((a, i) => (
                      <div key={`m-${i}`} style={{ marginTop: 6, padding: 8, background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 4 }}>
                        <div style={{ fontWeight: 600, marginBottom: 4 }}>&bdquo;{a.input}&ldquo;</div>
                        {a.matches.map((m, j) => (
                          <button
                            key={`mm-${i}-${j}`}
                            type="button"
                            onClick={() => resolveAmbiguous(a.input, m.email, m.displayName)}
                            style={{
                              display: 'block', width: '100%', textAlign: 'left', padding: '6px 8px',
                              marginTop: 4, background: '#fff', border: '1px solid var(--dex-gray-200)',
                              borderRadius: 4, cursor: 'pointer', fontSize: '0.8rem',
                            }}
                          >
                            <strong>{m.displayName}</strong> <span style={{ color: 'var(--dex-gray-400)' }}>{m.email}</span>
                          </button>
                        ))}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            <div style={{ display: 'flex', gap: 8, marginTop: 16, justifyContent: 'flex-end' }}>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => { if (!bulkImportRunning) setBulkImportOpen(false); }}
                disabled={bulkImportRunning}
              >
                Schließen
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={runBulkImport}
                disabled={bulkImportRunning || !bulkImportText.trim()}
              >
                {bulkImportRunning ? 'Suche läuft...' : 'Verarbeiten'}
              </button>
            </div>
          </div>
        </div>
      )}

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
                <Users size={18} /> Zielgruppe prüfen
              </h3>
              <button
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-600)' }}
                onClick={() => setShowEmailModal(false)}
              >
                <X size={20} />
              </button>
            </div>

            <div style={{ marginBottom: 16, padding: '10px 14px', background: 'var(--dex-gray-100)', borderRadius: 'var(--dex-radius)', fontSize: '0.85rem' }}>
              <div style={{ marginBottom: 6 }}>
                <strong>Standort-Filter:</strong>{' '}
                {locationFilter ? locationFilter.split(',').map(s => s.trim()).map(s => (
                  <span key={s} className="badge badge-green" style={{ marginRight: 6 }}>{s}</span>
                )) : <span style={{ color: 'var(--dex-gray-400)' }}>Keine</span>}
              </div>
              <div style={{ marginBottom: 6 }}>
                <strong>Zielgruppen-Filter:</strong>{' '}
                {audience ? audience.split(',').map(s => s.trim()).map(s => (
                  <span key={s} className="badge badge-orange" style={{ marginRight: 6 }}>{s}</span>
                )) : <span style={{ color: 'var(--dex-gray-400)' }}>Keine</span>}
              </div>
              {locationFilter && audience && (
                <div>
                  <strong>Verknüpfung:</strong>{' '}
                  <span className={`badge ${filterMode === 'AND' ? 'badge-red' : 'badge-green'}`}>
                    {filterMode === 'AND' ? 'UND (beide müssen zutreffen)' : 'ODER (eines reicht)'}
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
                      <th style={{ textAlign: 'center', padding: 6 }}>Sichtbar?</th>
                    </tr>
                  </thead>
                  <tbody>
                    {emailSearchResults.map(u => {
                      const filters = locationFilter ? locationFilter.split(',').map(s => s.trim().toLowerCase()) : [];
                      const isAll = filters.length === 0 || filters.indexOf('all') >= 0;
                      const loc = (u.location || '').toLowerCase();
                      const visible = isAll || filters.some(f => {
                        const norm = f.replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ä/g, 'ae');
                        return loc.indexOf(f) >= 0 || loc.indexOf(norm) >= 0;
                      });
                      return (
                        <tr key={u.email} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                          <td style={{ padding: 6 }}>{u.displayName}</td>
                          <td style={{ padding: 6, color: 'var(--dex-gray-600)' }}>{u.email}</td>
                          <td style={{ padding: 6, color: 'var(--dex-gray-500)' }}>{u.location || '-'}</td>
                          <td style={{ padding: 6, textAlign: 'center' }}>
                            {visible
                              ? <span style={{ color: '#22c55e', fontWeight: 600 }}>&#10003;</span>
                              : <span style={{ color: '#ef4444', fontWeight: 500 }}>&mdash;</span>}
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
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 4 }}>
              {SUGGESTED_FIELDS_CATALOG.map(s => (
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
                  <span>
                    <strong style={{ fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>{s.label}</strong>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>{s.description}</div>
                  </span>
                </label>
              ))}
            </div>
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
