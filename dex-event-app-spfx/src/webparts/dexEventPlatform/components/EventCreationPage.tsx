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
import { Trash2, Send, Plus, X, Users } from './Icons';
import { RichText } from '@pnp/spfx-controls-react/lib/controls/richText';
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
  type: 'text' | 'select' | 'number' | 'checkbox';
  required: boolean;
  options: string;
  visible: boolean;
}

export default function EventCreationPage(): React.ReactElement {
  const { navigate, goBack, selectedEventId, currentPage } = useNavigation();
  const { events, createEvent, updateEvent } = useEvents();
  const { currentUser } = useCurrentUser();
  const { searchUsers } = useRoles();
  const { t } = useLanguage();

  // Edit-Modus: wenn wir auf 'edit-event' sind und eine selectedEventId haben
  const isEditMode = currentPage === 'edit-event' && !!selectedEventId;
  const editEvent = isEditMode ? events.find(e => e.id === selectedEventId) : null;

  // ISO-Datum zu datetime-local Format konvertieren (YYYY-MM-DDThh:mm)
  const isoToLocal = (iso: string): string => {
    if (!iso) return '';
    const d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    const pad = (n: number): string => (n < 10 ? '0' : '') + n;
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  };

  const [title, setTitle] = React.useState(editEvent ? editEvent.title : '');
  const [organizer, setOrganizer] = React.useState(
    editEvent ? editEvent.organizers.join(', ') : `${currentUser.firstName} ${currentUser.surname}`
  );
  const [organizerResults, setOrganizerResults] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const [organizerEmails, setOrganizerEmails] = React.useState<string[]>([currentUser.email]);
  const [isSearchingOrganizer, setIsSearchingOrganizer] = React.useState(false);
  const organizerTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const [location, setLocation] = React.useState(editEvent ? editEvent.location : '');
  const [locationFilter, setLocationFilter] = React.useState(
    editEvent ? editEvent.locationAudience.join(', ') : ''
  );
  const [audience, setAudience] = React.useState(
    editEvent && editEvent.audienceFilter ? editEvent.audienceFilter.join(', ') : ''
  );
  const [filterMode, setFilterMode] = React.useState<'AND' | 'OR'>(
    editEvent ? editEvent.filterMode : 'OR'
  );
  const [description, setDescription] = React.useState(editEvent ? editEvent.description : '');
  const [eventType, setEventType] = React.useState<EventType>(editEvent ? editEvent.type : 'Other');
  const [startDate, setStartDate] = React.useState(editEvent ? isoToLocal(editEvent.startDate) : '');
  const [endDate, setEndDate] = React.useState(editEvent ? isoToLocal(editEvent.endDate) : '');
  const [registrationDeadline, setRegistrationDeadline] = React.useState(
    editEvent ? isoToLocal(editEvent.registrationDeadline) : ''
  );
  const [lastDeregisterDate, setLastDeregisterDate] = React.useState(editEvent ? isoToLocal(editEvent.lastDeregisterDate) : '');
  const [maxParticipants, setMaxParticipants] = React.useState(
    editEvent && editEvent.maxParticipants ? editEvent.maxParticipants.toString() : ''
  );
  const [waitlistEnabled, setWaitlistEnabled] = React.useState(true);
  const [eventImageUrl, setEventImageUrl] = React.useState(editEvent ? (editEvent.imageUrl || '') : '');
  const [imageFile, setImageFile] = React.useState<File | null>(null);
  const [imagePreview, setImagePreview] = React.useState(editEvent ? (editEvent.imageUrl || '') : '');
  const [customFields, setCustomFields] = React.useState<CustomFieldInput[]>(
    editEvent ? editEvent.eventSpecificFields.map(f => ({
      id: f.id, label: f.label, type: f.type, required: f.required,
      options: f.options ? f.options.join(', ') : '', visible: true,
    })) : []
  );
  const [outlookBody, setOutlookBody] = React.useState(editEvent ? (editEvent.outlookBody || '') : '');
  const [emailLanguage, setEmailLanguage] = React.useState(editEvent ? (editEvent.emailLanguage || 'EN') : 'EN');
  const [disableEmails, setDisableEmails] = React.useState(editEvent ? !!editEvent.disableEmails : false);
  const [disableOutlook, setDisableOutlook] = React.useState(editEvent ? !!editEvent.disableOutlook : false);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [emailTemplates, setEmailTemplates] = React.useState<Array<{ id: number; templateType: string; language: string; subject: string; heading: string; headingColor: string; bodyHtml: string }>>([]);
  const [emailTemplateOverrides, setEmailTemplateOverrides] = React.useState<Record<string, { subject: string; heading: string; bodyHtml: string }>>(
    editEvent?.emailTemplateOverrides ? (() => { try { return JSON.parse(editEvent.emailTemplateOverrides); } catch { return {}; } })() : {}
  );
  const [editingTemplate, setEditingTemplate] = React.useState<string | null>(null);
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
  const [quiz, setQuiz] = React.useState<Array<{id: string; question: string; options: string[]; correctIndices: number[]}>>(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    editEvent?.quiz?.map(q => ({...q, correctIndices: q.correctIndices || [(q as any).correctIndex || 0]})) || []
  );
  const [currentStep, setCurrentStep] = React.useState(0);
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

  const locationOptions = ['Berlin', 'Düsseldorf', 'Frankfurt', 'Hamburg', 'Hannover', 'Köln', 'München', 'Stuttgart', 'All'];

  const addCustomField = (): void => {
    setCustomFields([...customFields, {
      id: `cf-${Date.now()}`, label: '', type: 'text',
      required: false, options: '', visible: true,
    }]);
  };

  const removeCustomField = (id: string): void => {
    setCustomFields(customFields.filter(f => f.id !== id));
  };

  const updateCustomField = (id: string, updates: Partial<CustomFieldInput>): void => {
    setCustomFields(customFields.map(f => f.id === id ? { ...f, ...updates } : f));
  };

  const fillDemo = (): void => {
    const now = new Date();
    const eventStart = new Date(now);
    eventStart.setDate(now.getDate() + 7);
    eventStart.setHours(14, 0, 0, 0); // 14:00 Uhr
    const eventEnd = new Date(eventStart);
    eventEnd.setHours(18, 0, 0, 0); // 18:00 Uhr
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
    setEventType('Other');
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
      { id: `cf-${Date.now()}`, label: 'T-Shirt Größe', type: 'select', required: true, options: 'XS, S, M, L, XL, XXL', visible: true },
      { id: `cf-${Date.now() + 1}`, label: 'Notfallkontakt (Name & Telefon)', type: 'text', required: true, options: '', visible: true },
      { id: `cf-${Date.now() + 2}`, label: 'Ernährungsbesonderheiten', type: 'text', required: false, options: '', visible: true },
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

    // Schritt 1: Bild komprimieren und hochladen
    setProgress(5);
    setProgressLabel('Bild wird komprimiert und hochgeladen...');
    let imageUrl = eventImageUrl;
    if (imageFile) {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ctx = (window as any).__dexSpfxContext;
        if (ctx) {
          const svc = new EventService(ctx);
          const compressed = await compressImage(imageFile);
          const uploadedUrl = await svc.uploadEventImage(compressed, title);
          if (uploadedUrl) imageUrl = uploadedUrl;
        }
      } catch (err) {
        console.warn('[DEX] Bild-Upload fehlgeschlagen', err);
        setImageUploadError('Bild-Upload fehlgeschlagen. Das Event wird ohne Bild erstellt.');
      }
    }
    setProgress(15);

    if (isEditMode && selectedEventId) {
      setProgressLabel('Event wird aktualisiert...');
      // Event aktualisieren - nur bekannte Felder senden
      const updates: Record<string, unknown> = {
        'Title': title,
        'EventType': eventType,
        'Description': description,
        'Location': location,
        'LocationFilter': locationFilter,
        'Audience': audience,
        'FilterMode': filterMode,
        'StartDate': startDate ? new Date(startDate).toISOString() : null,
        'EndDate': endDate ? new Date(endDate).toISOString() : null,
        'RegistrationDeadline': registrationDeadline ? new Date(registrationDeadline).toISOString() : null,
        'MaxParticipants': Number(maxParticipants) || 0,
        'EventImageUrl': imageUrl,
        'Organizer': organizer,
        'CustomFields': JSON.stringify(customFields.map(f => ({
          id: f.id, label: f.label, type: f.type, required: f.required, visible: f.visible,
          ...(f.type === 'select' ? { options: f.options.split(',').map(o => o.trim()).filter(Boolean) } : {}),
        }))),
      };

      // Optionale Felder - immer senden damit Loeschungen wirken
      updates['LastDeregisterDate'] = lastDeregisterDate ? new Date(lastDeregisterDate).toISOString() : null;
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

      setProgress(50);

      // Neue Dokumente als Attachments hochladen (kein Documents-Feld noetig)
      if (selectedEventId && documents.some(d => d.file)) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ctx = (window as any).__dexSpfxContext;
        if (ctx) {
          const svc = new EventService(ctx);
          for (const doc of documents) {
            if (doc.file) {
              await svc.uploadEventDocument(Number(selectedEventId), doc.file);
            }
          }
        }
      }

      const success = await updateEvent(selectedEventId, updates);
      if (success) {
        // Outlook-Termin Update triggern (wenn CalendarLink vorhanden, nicht wenn Outlook deaktiviert)
        if (!disableOutlook) {
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
        locationFilter,
        audience,
        filterMode,
        startDate: startDate ? new Date(startDate).toISOString() : '',
        endDate: endDate ? new Date(endDate).toISOString() : '',
        registrationDeadline: registrationDeadline ? new Date(registrationDeadline).toISOString() : '',
        lastDeregisterDate: lastDeregisterDate ? new Date(lastDeregisterDate).toISOString() : '',
        maxParticipants: Number(maxParticipants) || 0,
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
        customFields: customFields.map(f => ({
          id: f.id, label: f.label, type: f.type, required: f.required, visible: f.visible,
          ...(f.type === 'select' ? { options: f.options.split(',').map(o => o.trim()).filter(Boolean) } : {}),
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
            // Event-Created Mail an alle Organizer senden
            const allOrgEmails = organizerEmails.length > 0 ? organizerEmails : [currentUser.email];
            for (const orgEmail of allOrgEmails) {
              const emailData = eventCreatedEmail(organizer, title, subsiteUrl);
              svc.queueEmail(
                emailData.subject, orgEmail, organizer, emailData.body,
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
                      <select className="form-select" disabled><option>Please select</option>{field.options.split(',').map(o => o.trim()).filter(Boolean).map(opt => <option key={opt}>{opt}</option>)}</select>
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
                  <span className="info-icon" title="Name des Events, z.B. 'B2Run Frankfurt 2026'" style={{ marginLeft: 8 }}>i</span>
                </label>
                <input className="form-input" value={title} onChange={e => setTitle(e.target.value)} placeholder="z.B. B2Run Frankfurt 2026" style={errorBorderStyle('title')} />
                {fieldHasError('title') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
              </div>

              <div className="form-group">
                <label className="form-label">
                  <span className="required">*</span> {t('create.eventtype')}
                  <span className="info-icon" title="Kategorie des Events – bestimmt das Design der Event-Karte" style={{ marginLeft: 8 }}>i</span>
                </label>
                <select className="form-select" value={eventType} onChange={e => setEventType(e.target.value as EventType)}>
                  <option value="Other">Sonstiges Deloitte Event</option>
                  <option value="B2Run">B2Run</option>
                  <option value="JPMorgan">JP Morgan Run</option>
                </select>
              </div>

              <div className="form-group" style={{ position: 'relative' }}>
                <label className="form-label">
                  <span className="required">*</span> {t('create.organizer')}
                  <span className="info-icon" title="Name der Person, die das Event organisiert" style={{ marginLeft: 8 }}>i</span>
                </label>
                <input
                  className="form-input"
                  value={organizer}
                  onChange={e => {
                    const val = e.target.value;
                    setOrganizer(val);
                    if (organizerTimerRef.current) clearTimeout(organizerTimerRef.current);
                    // Nur den Teil nach dem letzten Komma fuer die Suche nutzen
                    const lastPart = val.split(',').pop()?.trim() || '';
                    if (lastPart.length >= 2) {
                      organizerTimerRef.current = setTimeout(async () => {
                        setIsSearchingOrganizer(true);
                        const results = await searchUsers(lastPart);
                        setOrganizerResults(results);
                        setIsSearchingOrganizer(false);
                      }, 300);
                    } else {
                      setOrganizerResults([]);
                    }
                  }}
                  placeholder="Name oder E-Mail eingeben zum Suchen..."
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
                    maxHeight: 200, overflowY: 'auto',
                  }}>
                    {organizerResults.map(u => (
                      <div
                        key={u.email}
                        style={{
                          padding: '8px 12px', cursor: 'pointer', fontSize: '0.85rem',
                          borderBottom: '1px solid var(--dex-gray-100)',
                        }}
                        onMouseDown={() => {
                          // Letzten Teil (Suchbegriff) durch ausgewaehlten Namen ersetzen
                          const parts = organizer.split(',').map(s => s.trim());
                          parts.pop();
                          parts.push(u.displayName);
                          setOrganizer(parts.filter(Boolean).join(', '));
                          setOrganizerEmails(prev => [...prev.filter(e => e !== u.email), u.email]);
                          setOrganizerResults([]);
                        }}
                      >
                        <strong>{u.displayName}</strong>
                        <span style={{ color: 'var(--dex-gray-400)', marginLeft: 8 }}>{u.email}</span>
                      </div>
                    ))}
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
                  {locationOptions.map(loc => (
                    <label key={loc} style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: '0.9rem' }}>
                      <input
                        type="checkbox"
                        checked={locationFilter.split(',').map(s => s.trim()).indexOf(loc) >= 0}
                        onChange={e => {
                          const current = locationFilter.split(',').map(s => s.trim()).filter(Boolean);
                          if (e.target.checked) setLocationFilter([...current, loc].join(', '));
                          else setLocationFilter(current.filter(l => l !== loc).join(', '));
                        }}
                      />
                      {loc}
                    </label>
                  ))}
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

              <div className="form-group">
                <label className="form-label">
                  Zielgruppen-Filter
                </label>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                  Hier kannst du <strong>zusätzliche Personen oder Gruppen</strong> einladen, die das Event sehen sollen — unabhängig vom Standort.<br />
                  <em>Beispiel: Du trägst &bdquo;SAPALL&ldquo; ein → Alle Mitarbeiter der SAP-Abteilung sehen das Event, auch wenn ihr Standort nicht im Standort-Filter steht. Du kannst auch einzelne E-Mail-Adressen angeben (z.B. mmustermann@deloitte.de).</em>
                </p>
                <input
                  className="form-input"
                  value={audience}
                  onChange={e => setAudience(e.target.value)}
                  placeholder="z.B. DEALL, SAPALL, mmustermann@deloitte.de"
                />
                <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>
                  Kommasepariert. Gruppen: DEALL, DEKOELN, SAPALL. Einzelne E-Mails: name@deloitte.de
                </p>
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
                  <span className="info-icon" title="Beschreibung des Events – wird den Teilnehmern auf der Registrierungsseite angezeigt" style={{ marginLeft: 8 }}>i</span>
                </label>
                <textarea className="form-textarea" value={description} onChange={e => setDescription(e.target.value)} style={{ minHeight: 120, ...errorBorderStyle('description') }} />
                {fieldHasError('description') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
              </div>

              <div className="form-group">
                <label className="form-label">
                  Event-Bild
                  <span className="info-icon" title="Wird als Hintergrundbild auf der Event-Karte angezeigt. Empfohlen: 800x400px, max. 5MB." style={{ marginLeft: 8 }}>i</span>
                </label>
                {imagePreview && (
                  <div style={{ position: 'relative', marginBottom: 8 }}>
                    <img src={imagePreview} alt="Vorschau" style={{ width: '100%', maxHeight: 150, objectFit: 'cover', borderRadius: 'var(--dex-radius)' }} />
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
                  <span className="info-icon" title="Adresse oder Name des Veranstaltungsortes" style={{ marginLeft: 8 }}>i</span>
                </label>
                <input className="form-input" value={location} onChange={e => setLocation(e.target.value)} placeholder="z.B. RheinEnergieStadion, Köln" />
              </div>

              <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                <div className="form-group">
                  <label className="form-label">
                    <span className="required">*</span> {t('create.startdate')}
                    <span className="info-icon" title="Datum und Uhrzeit werden für den Outlook-Kalendereintrag verwendet" style={{ marginLeft: 8 }}>i</span>
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
                <div className="form-group">
                  <label className="form-label">
                    <span className="required">*</span> {t('create.enddate')}
                    <span className="info-icon" title="Datum und Uhrzeit werden für den Outlook-Kalendereintrag verwendet" style={{ marginLeft: 8 }}>i</span>
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
                  <span className="info-icon" title="Programmablauf / Timeline des Events" style={{ marginLeft: 8 }}>i</span>
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
                  <span className="info-icon" title="Abfahrtszeiten pro Standort" style={{ marginLeft: 8 }}>i</span>
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
                    <span className="info-icon" title="Bis wann können sich Teilnehmer anmelden?" style={{ marginLeft: 8 }}>i</span>
                  </label>
                  <DatePicker
                    selected={registrationDeadline ? new Date(registrationDeadline) : null}
                    onChange={(date: Date | null) => setRegistrationDeadline(date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}` : '')}
                    dateFormat="dd.MM.yyyy"
                    locale="de"
                    placeholderText="Anmelde-Deadline"
                    className="form-input"
                    isClearable
                    autoComplete="off"
                    maxDate={startDate ? new Date(startDate) : undefined}
                  />
                </div>
                <div className="form-group">
                  <label className="form-label">
                    {t('create.lastcancel')}
                    <span className="info-icon" title="Bis wann können sich Teilnehmer wieder abmelden?" style={{ marginLeft: 8 }}>i</span>
                  </label>
                  <DatePicker
                    selected={lastDeregisterDate ? new Date(lastDeregisterDate) : null}
                    onChange={(date: Date | null) => setLastDeregisterDate(date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}` : '')}
                    dateFormat="dd.MM.yyyy"
                    locale="de"
                    placeholderText="Abmeldefrist"
                    className="form-input"
                    isClearable
                    autoComplete="off"
                    maxDate={startDate ? new Date(startDate) : undefined}
                  />
                </div>
              </div>
              {fieldHasError('deadlineAfterStart') && <p style={{ color: 'var(--dex-red)', fontSize: '0.8rem', marginTop: -4, marginBottom: 8 }}>{t('create.error.deadlineAfterStart')}</p>}
              {fieldHasError('deregAfterStart') && <p style={{ color: 'var(--dex-red)', fontSize: '0.8rem', marginTop: -4, marginBottom: 8 }}>{t('create.error.deregAfterStart')}</p>}

              <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                <div className="form-group">
                  <label className="form-label">
                    {t('create.maxparticipants')}
                  </label>
                  <div className="toggle-wrapper" style={{ marginTop: 4, marginBottom: 8 }}>
                    <label className="toggle">
                      <input type="checkbox" checked={!maxParticipants || maxParticipants === '0'} onChange={e => { setMaxParticipants(e.target.checked ? '0' : '50'); if (e.target.checked) setWaitlistEnabled(false); }} />
                      <span className="toggle-slider" />
                    </label>
                    <span style={{ fontSize: '0.9rem' }}>{!maxParticipants || maxParticipants === '0' ? (t('create.maxparticipants') === 'Max. Participants' ? 'Unlimited' : 'Unbegrenzt') : maxParticipants + ' Plätze'}</span>
                  </div>
                  {maxParticipants && maxParticipants !== '0' && (
                    <input className="form-input" type="number" min={1} value={maxParticipants} onChange={e => setMaxParticipants(e.target.value)} placeholder="Anzahl" />
                  )}
                </div>
                {maxParticipants && maxParticipants !== '0' && (
                  <div className="form-group">
                    <label className="form-label">
                      {t('create.waitlist')}
                      <span className="info-icon" title="Wenn aktiviert, können sich Teilnehmer auch nach Erreichen der Max-Teilnehmer anmelden (Status: Warteliste)" style={{ marginLeft: 8 }}>i</span>
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

              </div>

              {/* ===== Step 3: Registrierungsfelder ===== */}
              <div style={{ display: currentStep === 3 ? 'block' : 'none' }}>
              {/* Dynamische Felder */}
              <div>
                <div className="flex-between mb-16">
                  <label className="form-label" style={{ marginBottom: 0 }}>{t('create.customfields')}</label>
                  <button className="btn btn-outline" onClick={addCustomField} style={{ fontSize: '0.85rem', padding: '6px 14px' }}>
                    <Plus size={14} /> {t('create.addfield')}
                  </button>
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
                          {(field.options ? field.options.split(',').filter(o => o.trim()) : []).map((opt, optIdx) => (
                            <div key={optIdx} style={{
                              display: 'flex', alignItems: 'center', gap: 4,
                              background: '#fff', border: '1px solid var(--dex-gray-300)',
                              borderRadius: 20, padding: '4px 8px 4px 12px', fontSize: '0.85rem',
                            }}>
                              <input
                                value={opt.trim()}
                                placeholder={`Option ${optIdx + 1}`}
                                onChange={e => {
                                  const opts = field.options ? field.options.split(',') : [];
                                  opts[optIdx] = e.target.value;
                                  updateCustomField(field.id, { options: opts.join(',') });
                                }}
                                style={{
                                  border: 'none', background: 'transparent', outline: 'none',
                                  width: Math.max(60, (opt.trim().length + 2) * 8), fontSize: '0.85rem',
                                }}
                              />
                              <button
                                onClick={() => {
                                  const opts = field.options.split(',').filter(o => o.trim());
                                  opts.splice(optIdx, 1);
                                  updateCustomField(field.id, { options: opts.join(',') });
                                }}
                                style={{ background: 'none', border: 'none', color: 'var(--dex-gray-400)', padding: 0, cursor: 'pointer', fontSize: '0.9rem', lineHeight: 1 }}
                              >
                                ×
                              </button>
                            </div>
                          ))}
                          <button
                            onClick={() => {
                              const opts = field.options ? field.options + ',' : '';
                              updateCustomField(field.id, { options: opts });
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
                  <textarea
                    className="form-input"
                    rows={4}
                    value={outlookBody}
                    onChange={e => setOutlookBody(e.target.value)}
                    placeholder={t('create.outlookdesc.placeholder')}
                  />
                </div>

                <h4 style={{ marginTop: 24, marginBottom: 12 }}>{t('create.templates.title')} ({emailLanguage})</h4>
                <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 12 }}>
                  {t('create.templates.hint')}
                </p>

                {['Anmeldung', 'Warteliste', 'Abmeldung', 'Nachrücken'].map(tType => {
                  const defaultTpl = emailTemplates.find(t => t.templateType === tType && t.language === emailLanguage);
                  const override = emailTemplateOverrides[tType];
                  const isEditing = editingTemplate === tType;
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
                            onClick={() => setEditingTemplate(isEditing ? null : tType)}
                          >
                            {isEditing ? t('create.templates.close') : t('create.templates.edit')}
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
                      {!isEditing && (
                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                          {t('create.templates.subject')}: {currentSubject.replace(/\{\{EventTitle\}\}/g, title || '...')}
                        </div>
                      )}
                      {isEditing && (
                        <div style={{ marginTop: 8 }}>
                          <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>{t('create.templates.subject')}</label>
                          <input
                            className="form-input"
                            value={currentSubject}
                            onChange={e => setEmailTemplateOverrides({
                              ...emailTemplateOverrides,
                              [tType]: { subject: e.target.value, heading: currentHeading, bodyHtml: currentBody },
                            })}
                            style={{ fontSize: '0.8rem', marginBottom: 8 }}
                          />
                          <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>{t('create.templates.heading')}</label>
                          <input
                            className="form-input"
                            value={currentHeading}
                            onChange={e => setEmailTemplateOverrides({
                              ...emailTemplateOverrides,
                              [tType]: { subject: currentSubject, heading: e.target.value, bodyHtml: currentBody },
                            })}
                            style={{ fontSize: '0.8rem', marginBottom: 8 }}
                          />
                          <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginBottom: 4, display: 'block' }}>{t('create.templates.content')}</label>
                          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginBottom: 6 }}>
                            <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', marginRight: 4, lineHeight: '24px' }}>
                              {t('create.templates.content') === 'Content' ? 'Insert variable:' : 'Variable einfügen:'}
                            </span>
                            {[
                              { key: '{{Name}}', label: 'Name' },
                              { key: '{{EventTitle}}', label: 'Event' },
                              { key: '{{Organizer}}', label: 'Organizer' },
                              { key: '{{AppUrl}}', label: 'App Link' },
                              { key: '{{WaitlistPosition}}', label: 'Waitlist #' },
                            ].map(v => (
                              <button
                                key={v.key}
                                type="button"
                                onClick={() => { navigator.clipboard.writeText(v.key); }}
                                style={{
                                  fontSize: '0.7rem', padding: '2px 8px', borderRadius: 10,
                                  border: '1px solid var(--dex-green)', background: 'rgba(134,188,37,0.08)',
                                  color: 'var(--dex-green-dark)', cursor: 'pointer', fontFamily: 'monospace',
                                }}
                                title={t('create.templates.content') === 'Content' ? `Click to copy ${v.key}` : `Klicken um ${v.key} zu kopieren`}
                              >
                                {v.label}
                              </button>
                            ))}
                          </div>
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
    </div>
  );
}
