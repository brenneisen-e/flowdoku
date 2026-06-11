// v17.21: A4-Zusammenfassung eines Events für Partner / Reviewer.
// Generiert eine druckfertige HTML-Repräsentation aller Wizard-Sektionen
// (Grundlagen, Ort & Programm, Kapazität & Sichtbarkeit, Team-Anmeldung,
// Felder, Kommunikation, Dokumente, Fun-Zone, Sub-Events). Exportiert
// wahlweise als PDF (per Browser-Druckdialog -> Save as PDF) oder als
// Word-kompatibles .doc (HTML-Blob mit application/msword-MIME).
//
// Bewusst kein neues Dependency — html-to-pdf / docx-Bibliotheken bringen
// 200-500 KB ins Bundle. Browser-Print + Word-HTML-Trick decken den
// Use-Case "an Partner zur Durchsicht schicken" zuverlässig ab.

export interface SummaryAgendaItem {
  time?: string;
  topic?: string;
  speaker?: string;
}

export interface SummaryTransferItem {
  time?: string;
  description?: string;
}

export interface SummaryCustomField {
  id: string;
  label: string;
  type: string;
  required: boolean;
  helpText?: string;
  confirmLabel?: string;
  options?: string[];
  multi?: boolean;
  onlyForGroup?: 'all' | 'A' | 'B';
  labelEn?: string;
  helpTextEn?: string;
  confirmLabelEn?: string;
  optionsEn?: string[];
  showIf?: { fieldId: string; values: string[] };
}

export interface SummarySubEvent {
  title: string;
  startDate?: string;
  endDate?: string;
  location?: string;
  description?: string;
  maxParticipants?: number;
  waitlistEnabled?: boolean;
}

export interface SummaryQuizQuestion {
  question: string;
  options?: string[];
  correctIndex?: number;
}

export interface SummaryData {
  // Grundlagen
  title: string;
  description?: string;
  imageDataUrl?: string;
  startDate?: string;
  endDate?: string;
  organizers?: string[];
  organizerEmails?: string[];
  contactName?: string;
  contactEmail?: string;
  contactInfo?: string;
  testTeam?: string[];
  qrScanners?: string[];
  isFictive?: boolean;
  activeFrom?: string;
  // Ort & Programm
  location?: string;
  address?: { street?: string; houseNo?: string; zip?: string; city?: string };
  agenda?: SummaryAgendaItem[];
  transfers?: SummaryTransferItem[];
  // Kapazität & Sichtbarkeit
  locationFilter?: string[];
  audience?: string[];
  filterMode?: 'AND' | 'OR';
  excludedUsers?: string[];
  registrationDeadline?: string;
  lastDeregisterDate?: string;
  maxParticipants?: number;
  unlimitedParticipants?: boolean;
  waitlistEnabled?: boolean;
  durchstarterCapacity?: number;
  funstarterCapacity?: number;
  splitLabelA?: string;
  splitLabelB?: string;
  splitSharedWaitlist?: boolean;
  // Team-Anmeldung
  teamRegistrationEnabled?: boolean;
  teamSize?: number;
  askTeamName?: boolean;
  teamPartialAllowed?: boolean;
  teamOpenSlotsVisible?: boolean;
  teamJoinRequiresApproval?: boolean;
  // Felder
  askSalutation?: boolean;
  bilingualFields?: boolean;
  customFields?: SummaryCustomField[];
  allowAttendeeUpload?: boolean;
  attendeeUploadHint?: string;
  attendeeUploadLabel?: string;
  // Kommunikation
  emailLanguage?: string;
  disableEmails?: boolean;
  disableOutlook?: boolean;
  outlookHeading?: string;
  outlookSubheading?: string;
  outlookBody?: string;
  notifyOrgRegisterMode?: string;
  notifyOrgRegisterFromDate?: string;
  notifyOrgCancelMode?: string;
  // Dokumente
  documents?: Array<{ name: string; size?: number }>;
  // Fun-Zone
  funZone?: SummaryQuizQuestion[];
  quizClusterSize?: number;
  // Sub-Events
  subEvents?: SummarySubEvent[];
  childTermSingular?: string;
  childTermPlural?: string;
  subEventsOnlyMode?: boolean;
  requireSubEventSelection?: boolean;
  // Meta
  generatedAt: string; // ISO
  locale: 'de' | 'en';
}

const escapeHtml = (s: string | undefined | null): string => {
  if (!s) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
};

// v17.22: Lightweight Rich-HTML-Sanitizer. Die Felder `description`,
// `outlookBody` und `se.description` kommen aus dem RichText-Editor und
// enthalten legitimes Formatierungs-HTML (<p>, <strong>, <ul>, <a> …) —
// die dürfen NICHT escaped werden, sonst sieht der Reviewer rohe Tags.
// Aber: ein Organizer (oder eine kompromittierte Quelle) könnte
// <script>, Event-Handler (onload=…), javascript:-URLs oder
// <iframe>/<object>/<embed> einschleusen, die im Export-Fenster
// (gleiche SharePoint-Origin) ausgeführt würden. Dieser Sanitizer
// entfernt genau diese gefährlichen Konstrukte und lässt harmloses
// Layout-HTML durch. Kein vollwertiger DOMPurify, aber deckt die
// relevanten Vektoren für den reinen Druck-/Word-Export ab.
const sanitizeRichHtml = (html: string | undefined | null): string => {
  if (!html) return '';
  let out = String(html);
  // 1. Komplette gefährliche Element-Blöcke (inkl. Inhalt) entfernen.
  out = out.replace(/<\s*(script|style|iframe|object|embed|link|meta|base|form)\b[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi, '');
  // 2. Selbstschliessende / unvollständige Varianten derselben Tags.
  out = out.replace(/<\s*(script|iframe|object|embed|link|meta|base|form)\b[^>]*\/?>/gi, '');
  // 3. Inline-Event-Handler (onload=, onclick=, onerror=, …) entfernen —
  //    sowohl mit doppelten/einfachen Quotes als auch ohne Quotes.
  out = out.replace(/\son[a-z]+\s*=\s*"[^"]*"/gi, '');
  out = out.replace(/\son[a-z]+\s*=\s*'[^']*'/gi, '');
  out = out.replace(/\son[a-z]+\s*=\s*[^\s>]+/gi, '');
  // 4. javascript:- und data:text/html-URLs in href/src neutralisieren.
  out = out.replace(/(href|src)\s*=\s*"(?:\s*javascript:|\s*data:text\/html)[^"]*"/gi, '$1="#"');
  out = out.replace(/(href|src)\s*=\s*'(?:\s*javascript:|\s*data:text\/html)[^']*'/gi, "$1='#'");
  return out;
};

const formatDate = (iso: string | undefined, locale: 'de' | 'en'): string => {
  if (!iso) return '—';
  try {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return iso;
    return d.toLocaleString(locale === 'de' ? 'de-DE' : 'en-GB', {
      day: '2-digit', month: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit',
    });
  } catch {
    return iso;
  }
};

const formatBytes = (b?: number): string => {
  if (!b || b <= 0) return '';
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / 1024 / 1024).toFixed(1)} MB`;
};

const tt = (de: string, en: string, locale: 'de' | 'en'): string =>
  locale === 'de' ? de : en;

// Eine semantische Sektion mit Ueberschrift + Inhalt.
const section = (heading: string, content: string): string => `
  <section class="dex-section">
    <h2>${escapeHtml(heading)}</h2>
    ${content}
  </section>`;

// Key-Value-Liste — kompakte Darstellung von Settings.
const kv = (label: string, value: string | number | boolean | undefined | null, locale: 'de' | 'en'): string => {
  let display: string;
  if (value === undefined || value === null || value === '') {
    display = '—';
  } else if (typeof value === 'boolean') {
    display = value ? tt('Ja', 'Yes', locale) : tt('Nein', 'No', locale);
  } else {
    display = String(value);
  }
  return `
    <div class="dex-kv">
      <div class="dex-kv-label">${escapeHtml(label)}</div>
      <div class="dex-kv-value">${escapeHtml(display)}</div>
    </div>`;
};

const list = (items: string[]): string => {
  if (!items || items.length === 0) return '<p class="dex-empty">—</p>';
  return '<ul class="dex-list">' + items.map(i => `<li>${escapeHtml(i)}</li>`).join('') + '</ul>';
};

export function buildSummaryHtml(d: SummaryData): string {
  const lc = d.locale;
  const T = (de: string, en: string): string => tt(de, en, lc);

  // ===== Grundlagen =====
  const grundlagenKv = [
    kv(T('Titel', 'Title'), d.title || '—', lc),
    kv(T('Beginn', 'Start'), formatDate(d.startDate, lc), lc),
    kv(T('Ende', 'End'), formatDate(d.endDate, lc), lc),
    kv(T('Sichtbar ab', 'Visible from'), d.activeFrom ? formatDate(d.activeFrom, lc) : T('sofort', 'immediately'), lc),
    kv(T('Test-Event (Entwurf)', 'Test event (draft)'), !!d.isFictive, lc),
  ].join('');

  const grundlagenImage = d.imageDataUrl
    ? `<div class="dex-image"><img src="${escapeHtml(d.imageDataUrl)}" alt="Event-Bild" /></div>`
    : '';

  const grundlagenDescription = d.description
    ? `<div class="dex-rich"><strong>${T('Beschreibung', 'Description')}:</strong><br/>${sanitizeRichHtml(d.description)}</div>`
    : '';

  const grundlagenOrganizers = (d.organizers && d.organizers.length > 0)
    ? `<div class="dex-subblock"><strong>${T('Organizer', 'Organizers')}:</strong>${list(d.organizers)}</div>`
    : '';

  const grundlagenContact = (d.contactName || d.contactEmail)
    ? `<div class="dex-subblock"><strong>${T('Ansprechpartner', 'Contact')}:</strong>
        <div class="dex-rich">
          ${d.contactName ? `${escapeHtml(d.contactName)}<br/>` : ''}
          ${d.contactEmail ? `${escapeHtml(d.contactEmail)}<br/>` : ''}
          ${d.contactInfo ? `<span style="color:#666">${escapeHtml(d.contactInfo)}</span>` : ''}
        </div>
       </div>`
    : '';

  const grundlagenTestTeam = (d.testTeam && d.testTeam.length > 0)
    ? `<div class="dex-subblock"><strong>${T('Test-Team', 'Test team')}:</strong>${list(d.testTeam)}</div>`
    : '';

  const grundlagenQrScanners = (d.qrScanners && d.qrScanners.length > 0)
    ? `<div class="dex-subblock"><strong>${T('Check-In-Team', 'Check-in team')}:</strong>${list(d.qrScanners)}</div>`
    : '';

  const sectionGrundlagen = section(
    T('1. Grundlagen', '1. Basics'),
    grundlagenImage + `<div class="dex-kv-grid">${grundlagenKv}</div>` +
      grundlagenDescription + grundlagenOrganizers + grundlagenContact +
      grundlagenTestTeam + grundlagenQrScanners
  );

  // ===== Ort & Programm =====
  const addr = d.address || {};
  const addrLine = [addr.street, addr.houseNo].filter(Boolean).join(' ').trim();
  const addrCityLine = [addr.zip, addr.city].filter(Boolean).join(' ').trim();
  const ortKv = [
    kv(T('Veranstaltungsort', 'Venue'), d.location || '—', lc),
    kv(T('Straße', 'Street'), addrLine || '—', lc),
    kv(T('PLZ / Stadt', 'ZIP / City'), addrCityLine || '—', lc),
  ].join('');

  const agendaList = (d.agenda && d.agenda.length > 0)
    ? `<div class="dex-subblock"><strong>${T('Agenda', 'Agenda')}:</strong>
       <table class="dex-table"><thead><tr>
         <th>${T('Uhrzeit', 'Time')}</th>
         <th>${T('Thema', 'Topic')}</th>
         <th>${T('Sprecher', 'Speaker')}</th>
       </tr></thead><tbody>
       ${d.agenda.map(a => `<tr>
         <td>${escapeHtml(a.time || '')}</td>
         <td>${escapeHtml(a.topic || '')}</td>
         <td>${escapeHtml(a.speaker || '')}</td>
       </tr>`).join('')}
       </tbody></table></div>`
    : '';

  const transferList = (d.transfers && d.transfers.length > 0)
    ? `<div class="dex-subblock"><strong>${T('Transferzeiten', 'Transfer times')}:</strong>
       <table class="dex-table"><thead><tr>
         <th>${T('Uhrzeit', 'Time')}</th>
         <th>${T('Beschreibung', 'Description')}</th>
       </tr></thead><tbody>
       ${d.transfers.map(t => `<tr>
         <td>${escapeHtml(t.time || '')}</td>
         <td>${escapeHtml(t.description || '')}</td>
       </tr>`).join('')}
       </tbody></table></div>`
    : '';

  const sectionOrt = section(
    T('2. Ort & Programm', '2. Location & programme'),
    `<div class="dex-kv-grid">${ortKv}</div>` + agendaList + transferList
  );

  // ===== Kapazität & Sichtbarkeit =====
  const sichtKv = [
    kv(T('Standort-Filter', 'Location filter'),
      (d.locationFilter && d.locationFilter.length > 0)
        ? d.locationFilter.join(', ')
        : T('alle Standorte', 'all locations'), lc),
    kv(T('Mailverteiler / User', 'Mail audience'),
      (d.audience && d.audience.length > 0)
        ? `${d.audience.length} ${T('Eintrag/Einträge', 'entry/entries')}`
        : T('kein Filter', 'no filter'), lc),
    kv(T('Filter-Verknüpfung', 'Filter combination'), d.filterMode || 'AND', lc),
    kv(T('Ausgeschlossene User', 'Excluded users'),
      (d.excludedUsers && d.excludedUsers.length > 0)
        ? `${d.excludedUsers.length}`
        : '—', lc),
    kv(T('Anmeldeschluss', 'Registration deadline'), formatDate(d.registrationDeadline, lc), lc),
    kv(T('Abmeldefrist', 'Cancel deadline'), formatDate(d.lastDeregisterDate, lc), lc),
    kv(T('Teilnehmer max.', 'Max participants'),
      d.unlimitedParticipants ? T('unbegrenzt', 'unlimited') : (d.maxParticipants ?? 0), lc),
    kv(T('Warteliste', 'Waitlist'), !!d.waitlistEnabled, lc),
  ];

  if (d.durchstarterCapacity && d.durchstarterCapacity > 0) {
    sichtKv.push(
      kv(T('Geteilte Kapazität', 'Split capacity'), true, lc),
      kv(`${T('Gruppe', 'Group')} A (${d.splitLabelA || 'Durchstarter'})`, d.durchstarterCapacity, lc),
      kv(`${T('Gruppe', 'Group')} B (${d.splitLabelB || 'Funstarter'})`, d.funstarterCapacity ?? 0, lc),
      kv(T('Gemeinsame Warteliste', 'Shared waitlist'), !!d.splitSharedWaitlist, lc),
    );
  }

  const audienceDetail = (d.audience && d.audience.length > 0)
    ? `<div class="dex-subblock"><strong>${T('Mailverteiler im Detail', 'Mail audience details')}:</strong>${list(d.audience.slice(0, 25))}${d.audience.length > 25 ? `<p class="dex-empty">${T('… plus', '… plus')} ${d.audience.length - 25} ${T('weitere', 'more')}</p>` : ''}</div>`
    : '';

  const sectionSichtbarkeit = section(
    T('3. Kapazität & Sichtbarkeit', '3. Capacity & visibility'),
    `<div class="dex-kv-grid">${sichtKv.join('')}</div>` + audienceDetail
  );

  // ===== Team-Anmeldung =====
  let sectionTeam = '';
  if (d.teamRegistrationEnabled) {
    const teamKv = [
      kv(T('Team-Anmeldung aktiv', 'Team registration enabled'), true, lc),
      kv(T('Team-Größe', 'Team size'), d.teamSize ?? '—', lc),
      kv(T('Team-Namen abfragen', 'Ask team name'), !!d.askTeamName, lc),
      kv(T('Teil-Teams erlaubt', 'Partial teams allowed'), !!d.teamPartialAllowed, lc),
      kv(T('Offene Slots oeffentlich sichtbar', 'Open slots publicly visible'), !!d.teamOpenSlotsVisible, lc),
      kv(T('Beitritt mit Lead-Bestätigung', 'Join requires lead approval'), !!d.teamJoinRequiresApproval, lc),
    ].join('');
    sectionTeam = section(
      T('4. Team-Anmeldung', '4. Team registration'),
      `<div class="dex-kv-grid">${teamKv}</div>`
    );
  } else {
    sectionTeam = section(
      T('4. Team-Anmeldung', '4. Team registration'),
      `<p class="dex-empty">${T('Team-Anmeldung ist für dieses Event deaktiviert.', 'Team registration is disabled for this event.')}</p>`
    );
  }

  // ===== Felder =====
  const fieldTypeLabel = (t: string): string => {
    switch (t) {
      case 'text': return T('Text', 'Text');
      case 'select': return T('Dropdown', 'Dropdown');
      case 'number': return T('Zahl', 'Number');
      case 'checkbox': return T('Checkbox', 'Checkbox');
      case 'user': return T('Personen-Picker', 'Person picker');
      case 'roommate': return T('Roommate', 'Roommate');
      default: return t;
    }
  };

  const renderCustomField = (f: SummaryCustomField, idx: number): string => {
    const optionsRow = (f.options && f.options.length > 0)
      ? `<tr><th>${T('Optionen', 'Options')}</th><td>${
          f.options.map((o, i) => {
            const en = f.optionsEn && f.optionsEn[i];
            return en ? `${escapeHtml(o)} <span class="dex-en">(EN: ${escapeHtml(en)})</span>` : escapeHtml(o);
          }).join('<br/>')
        }${f.multi ? `<br/><em>${T('Mehrfachauswahl', 'Multi-select')}</em>` : ''}</td></tr>`
      : '';
    const helpRow = f.helpText
      ? `<tr><th>${T('Beschreibung', 'Description')}</th><td>${escapeHtml(f.helpText)}${f.helpTextEn ? `<br/><span class="dex-en">EN: ${escapeHtml(f.helpTextEn)}</span>` : ''}</td></tr>`
      : '';
    const confirmRow = (f.type === 'checkbox' && f.confirmLabel)
      ? `<tr><th>${T('Bestätigungs-Text', 'Confirmation text')}</th><td>${escapeHtml(f.confirmLabel)}${f.confirmLabelEn ? `<br/><span class="dex-en">EN: ${escapeHtml(f.confirmLabelEn)}</span>` : ''}</td></tr>`
      : '';
    const groupRow = (f.onlyForGroup && f.onlyForGroup !== 'all')
      ? `<tr><th>${T('Sichtbar für', 'Visible for')}</th><td>${T('nur Gruppe', 'group')} ${f.onlyForGroup}</td></tr>`
      : '';
    const showIfRow = f.showIf
      ? `<tr><th>${T('Sichtbar wenn', 'Visible when')}</th><td>${escapeHtml(f.showIf.fieldId)} = ${escapeHtml(f.showIf.values.join(' | '))}</td></tr>`
      : '';
    const labelEnRow = f.labelEn
      ? `<tr><th>${T('Englisch', 'English')}</th><td>${escapeHtml(f.labelEn)}</td></tr>`
      : '';

    return `<div class="dex-field-card">
      <div class="dex-field-head">
        <span class="dex-field-no">${idx + 1}</span>
        <span class="dex-field-title">${escapeHtml(f.label)}</span>
        <span class="dex-field-type">${escapeHtml(fieldTypeLabel(f.type))}</span>
        ${f.required ? `<span class="dex-pill">${T('Pflicht', 'Required')}</span>` : ''}
      </div>
      <table class="dex-mini-table">
        ${labelEnRow}
        ${helpRow}
        ${confirmRow}
        ${optionsRow}
        ${groupRow}
        ${showIfRow}
      </table>
    </div>`;
  };

  const fieldsBody = (d.customFields && d.customFields.length > 0)
    ? d.customFields.map(renderCustomField).join('')
    : `<p class="dex-empty">${T('Keine eigenen Felder konfiguriert.', 'No custom fields configured.')}</p>`;

  const fieldsMeta = [
    kv(T('Anrede abfragen', 'Ask salutation'), !!d.askSalutation, lc),
    kv(T('Zweisprachig (DE/EN)', 'Bilingual (DE/EN)'), !!d.bilingualFields, lc),
    kv(T('Teilnehmer-Upload erlauben', 'Allow attendee upload'), !!d.allowAttendeeUpload, lc),
  ].join('');

  const uploadHint = d.allowAttendeeUpload && (d.attendeeUploadLabel || d.attendeeUploadHint)
    ? `<div class="dex-subblock">
        ${d.attendeeUploadLabel ? `<strong>${T('Upload-Bezeichnung', 'Upload label')}:</strong> ${escapeHtml(d.attendeeUploadLabel)}<br/>` : ''}
        ${d.attendeeUploadHint ? `<strong>${T('Upload-Hinweis', 'Upload hint')}:</strong> ${escapeHtml(d.attendeeUploadHint)}` : ''}
       </div>`
    : '';

  const sectionFelder = section(
    T('5. Felder', '5. Fields'),
    `<div class="dex-kv-grid">${fieldsMeta}</div>` + uploadHint + fieldsBody
  );

  // ===== Kommunikation =====
  const kommKv = [
    kv(T('Mail-Sprache', 'Email language'), d.emailLanguage || 'DE', lc),
    kv(T('Mails an Teilnehmer deaktiviert', 'Attendee mails disabled'), !!d.disableEmails, lc),
    kv(T('Outlook-Termine deaktiviert', 'Outlook invites disabled'), !!d.disableOutlook, lc),
    kv(T('Organizer-Mitteilung Anmeldungen', 'Organizer notify (register)'), d.notifyOrgRegisterMode || 'never', lc),
    kv(T('Organizer-Mitteilung Abmeldungen', 'Organizer notify (cancel)'), d.notifyOrgCancelMode || 'never', lc),
  ].join('');

  const outlookPreview = (d.outlookHeading || d.outlookSubheading || d.outlookBody)
    ? `<div class="dex-subblock">
        <strong>${T('Outlook-Termin-Body', 'Outlook calendar body')}:</strong>
        ${d.outlookHeading ? `<div style="font-weight:700;font-size:1.1em;margin-top:6px">${escapeHtml(d.outlookHeading)}</div>` : ''}
        ${d.outlookSubheading ? `<div style="color:#666;margin-bottom:6px">${escapeHtml(d.outlookSubheading)}</div>` : ''}
        ${d.outlookBody ? `<div class="dex-rich">${sanitizeRichHtml(d.outlookBody)}</div>` : ''}
       </div>`
    : '';

  const sectionKomm = section(
    T('6. Kommunikation', '6. Communication'),
    `<div class="dex-kv-grid">${kommKv}</div>` + outlookPreview
  );

  // ===== Dokumente =====
  const docsList = (d.documents && d.documents.length > 0)
    ? `<table class="dex-table"><thead><tr>
         <th>${T('Datei', 'File')}</th>
         <th>${T('Größe', 'Size')}</th>
       </tr></thead><tbody>
       ${d.documents.map(doc => `<tr>
         <td>${escapeHtml(doc.name)}</td>
         <td>${escapeHtml(formatBytes(doc.size))}</td>
       </tr>`).join('')}
       </tbody></table>`
    : `<p class="dex-empty">${T('Keine Dokumente hochgeladen.', 'No documents uploaded.')}</p>`;
  const sectionDocs = section(T('7. Dokumente', '7. Documents'), docsList);

  // ===== Fun-Zone =====
  const quizBody = (d.funZone && d.funZone.length > 0)
    ? `<p>${T('Quiz-Cluster-Größe', 'Quiz cluster size')}: ${d.quizClusterSize ?? 1}</p>` +
      d.funZone.map((q, i) => `
        <div class="dex-quiz">
          <div class="dex-quiz-q"><strong>${i + 1}.</strong> ${escapeHtml(q.question)}</div>
          ${q.options ? `<ol type="a" class="dex-quiz-opts">${q.options.map((o, oi) => `<li${oi === q.correctIndex ? ' class="correct"' : ''}>${escapeHtml(o)}${oi === q.correctIndex ? ` <span class="dex-correct-mark">✓</span>` : ''}</li>`).join('')}</ol>` : ''}
        </div>`).join('')
    : `<p class="dex-empty">${T('Keine Quiz-Fragen.', 'No quiz questions.')}</p>`;
  const sectionFunZone = section(T('8. Fun-Zone', '8. Fun zone'), quizBody);

  // ===== Sub-Events =====
  const subTermPlural = d.childTermPlural || T('Sub-Events', 'Sub-events');
  let sectionSubEvents = '';
  if (d.subEvents && d.subEvents.length > 0) {
    const subBody = d.subEvents.map((se, i) => `
      <div class="dex-subevent">
        <h3>${i + 1}. ${escapeHtml(se.title || T('ohne Titel', 'untitled'))}</h3>
        <div class="dex-kv-grid">
          ${kv(T('Beginn', 'Start'), formatDate(se.startDate, lc), lc)}
          ${kv(T('Ende', 'End'), formatDate(se.endDate, lc), lc)}
          ${kv(T('Ort', 'Location'), se.location || '—', lc)}
          ${kv(T('Plätze', 'Capacity'), se.maxParticipants ?? '—', lc)}
          ${kv(T('Warteliste', 'Waitlist'), !!se.waitlistEnabled, lc)}
        </div>
        ${se.description ? `<div class="dex-rich">${sanitizeRichHtml(se.description)}</div>` : ''}
      </div>`).join('');
    const subMeta = [
      kv(T('Nur Sub-Events-Modus', 'Sub-events-only mode'), !!d.subEventsOnlyMode, lc),
      kv(T('Sub-Event-Auswahl Pflicht', 'Sub-event selection required'), !!d.requireSubEventSelection, lc),
    ].join('');
    sectionSubEvents = section(
      `${subTermPlural} (${d.subEvents.length})`,
      `<div class="dex-kv-grid">${subMeta}</div>` + subBody
    );
  }

  // ===== Komplettes Dokument =====
  const generated = formatDate(d.generatedAt, lc);
  const docTitle = `${d.title || T('Event-Zusammenfassung', 'Event summary')}`;

  // Print-CSS: A4-Seitenbreite, 18mm Rand, Serif-Font für den Text. Section-
  // Headlines in Deloitte-Grün. Page-Break-Hinweise zwischen den grossen
  // Sektionen, damit der PDF-Druck pro Sektion sauber paginiert.
  const css = `
    @page { size: A4; margin: 18mm; }
    * { box-sizing: border-box; }
    body { font-family: 'Segoe UI', Calibri, Arial, sans-serif; color: #222; line-height: 1.45; font-size: 11pt; margin: 0; }
    h1 { color: #4a7c1f; font-size: 22pt; margin: 0 0 4px; }
    h2 { color: #4a7c1f; font-size: 14pt; margin: 16px 0 8px; padding-bottom: 4px; border-bottom: 1.5px solid #86bc25; page-break-after: avoid; }
    h3 { color: #4a7c1f; font-size: 12pt; margin: 12px 0 6px; }
    p { margin: 4px 0; }
    .dex-meta { color: #666; font-size: 9pt; margin-bottom: 12px; }
    .dex-section { page-break-inside: auto; margin-bottom: 12px; }
    .dex-image { text-align: center; margin: 8px 0 12px; }
    .dex-image img { max-width: 100%; max-height: 80mm; border-radius: 4px; }
    .dex-rich { background: #f7f8f9; padding: 8px 10px; border-radius: 4px; border-left: 3px solid #86bc25; margin: 6px 0; }
    .dex-kv-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 2px 12px; margin: 4px 0 8px; }
    .dex-kv { padding: 3px 0; border-bottom: 1px dotted #ddd; display: flex; gap: 8px; min-width: 0; }
    .dex-kv-label { font-weight: 600; color: #555; min-width: 38%; flex-shrink: 0; }
    .dex-kv-value { color: #222; word-break: break-word; }
    .dex-subblock { margin: 8px 0; }
    .dex-list { margin: 4px 0 4px 18px; padding: 0; }
    .dex-list li { margin: 2px 0; }
    .dex-empty { color: #999; font-style: italic; margin: 4px 0; }
    .dex-table { width: 100%; border-collapse: collapse; margin: 6px 0; font-size: 10pt; }
    .dex-table th, .dex-table td { border: 1px solid #ddd; padding: 4px 8px; text-align: left; vertical-align: top; }
    .dex-table thead { background: rgba(134,188,37,0.10); }
    .dex-mini-table { width: 100%; border-collapse: collapse; font-size: 10pt; margin: 4px 0 8px; }
    .dex-mini-table th { width: 30%; text-align: left; padding: 3px 8px; color: #555; font-weight: 600; vertical-align: top; }
    .dex-mini-table td { padding: 3px 8px; vertical-align: top; }
    .dex-mini-table tr { border-bottom: 1px dotted #eee; }
    .dex-field-card { border: 1px solid #ddd; border-radius: 4px; padding: 8px 10px; margin: 6px 0; page-break-inside: avoid; background: #fafafa; }
    .dex-field-head { display: flex; align-items: center; gap: 8px; margin-bottom: 4px; flex-wrap: wrap; }
    .dex-field-no { width: 22px; height: 22px; border-radius: 50%; background: #86bc25; color: #fff; display: inline-flex; align-items: center; justify-content: center; font-size: 9pt; font-weight: 700; }
    .dex-field-title { font-weight: 700; font-size: 11pt; flex: 1; }
    .dex-field-type { font-size: 9pt; color: #4a7c1f; background: rgba(134,188,37,0.10); padding: 2px 8px; border-radius: 8px; }
    .dex-pill { font-size: 9pt; color: #c00; background: rgba(192,0,0,0.08); padding: 2px 8px; border-radius: 8px; font-weight: 600; }
    .dex-en { color: #005a9c; font-size: 9pt; }
    .dex-quiz { border-left: 3px solid #86bc25; padding: 4px 10px; margin: 6px 0; background: #fafafa; page-break-inside: avoid; }
    .dex-quiz-q { font-weight: 600; }
    .dex-quiz-opts { margin: 4px 0 0 24px; }
    .dex-quiz-opts li.correct { color: #4a7c1f; font-weight: 600; }
    .dex-correct-mark { color: #4a7c1f; }
    .dex-subevent { border: 1px solid #ddd; border-radius: 4px; padding: 8px 10px; margin: 6px 0; page-break-inside: avoid; }
    @media print {
      .dex-section { page-break-after: auto; }
      .dex-no-print { display: none !important; }
    }
  `;

  return `<!doctype html>
<html lang="${lc}">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(docTitle)}</title>
<style>${css}</style>
</head>
<body>
  <h1>${escapeHtml(docTitle)}</h1>
  <p class="dex-meta">${T('Erstellt am', 'Generated on')} ${escapeHtml(generated)} &middot; ${T('Event Experience Platform', 'Event Experience Platform')}</p>
  ${sectionGrundlagen}
  ${sectionOrt}
  ${sectionSichtbarkeit}
  ${sectionTeam}
  ${sectionFelder}
  ${sectionKomm}
  ${sectionDocs}
  ${sectionFunZone}
  ${sectionSubEvents}
</body>
</html>`;
}

/**
 * Oeffnet die A4-Zusammenfassung in einem neuen Fenster und triggert sofort
 * den Druckdialog. Im Browser-Druckdialog kann der User „Als PDF speichern"
 * wählen und das Dokument als PDF ablegen.
 *
 * Hintergrund: jsPDF + html2canvas oder ein serverseitiger PDF-Renderer
 * wären saubere Alternativen, kosten aber ~200 KB im Bundle ODER eine
 * extra Power-Automate-Action. Der Browser-Druckdialog ist auf jedem
 * Endgerät verfügbar, kann nativ PDF speichern und respektiert das
 * @page A4-CSS oben. Trade-off: User muss bewusst „Als PDF speichern"
 * klicken — wir können es nicht still im Hintergrund machen.
 */
export function exportSummaryAsPdf(data: SummaryData): void {
  const html = buildSummaryHtml(data);
  const w = window.open('', '_blank', 'width=900,height=1100,scrollbars=yes');
  if (!w) {
    // Pop-up blockiert — Fallback ueber Data-URL als Download.
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${(data.title || 'event-summary').replace(/[^a-z0-9äöüßÄÖÜ_-]+/gi, '_')}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    return;
  }
  // v17.22: opener-Referenz kappen — das neue Fenster teilt die SharePoint-
  // Origin; ohne diese Zeile könnte darin laufendes Script (falls der
  // Sanitizer je umgangen würde) ueber window.opener auf die Host-Seite
  // zugreifen (Reverse-Tabnabbing / SP-REST im User-Kontext).
  try { w.opener = null; } catch { /* */ }
  w.document.open();
  w.document.write(html);
  w.document.close();
  // v17.22: Statt eines fixen 500ms-Timeouts (Race: Remote-SharePoint-Bild
  // beim Edit-Flow oft noch nicht geladen → leeres/kaputtes Bild im PDF)
  // warten wir, bis alle <img> fertig sind (complete ODER onload/onerror),
  // mit 4s-Sicherheitsdeadline. Erst dann der Druckdialog.
  const triggerPrint = (): void => {
    try { w.focus(); w.print(); } catch { /* */ }
  };
  const waitForImages = (): void => {
    let imgs: HTMLImageElement[] = [];
    try { imgs = Array.prototype.slice.call(w.document.images || []); } catch { /* */ }
    const pending = imgs.filter(img => !img.complete);
    if (pending.length === 0) { triggerPrint(); return; }
    let done = 0;
    let fired = false;
    const finish = (): void => { if (!fired) { fired = true; triggerPrint(); } };
    pending.forEach(img => {
      const onSettled = (): void => { done++; if (done >= pending.length) finish(); };
      img.addEventListener('load', onSettled);
      img.addEventListener('error', onSettled);
    });
    // Sicherheits-Deadline: spätestens nach 4s drucken, egal ob Bilder fertig.
    setTimeout(finish, 4000);
  };
  // Kleiner Initial-Delay, damit document.images nach write/close gefüllt ist.
  setTimeout(waitForImages, 150);
}

/**
 * Speichert die A4-Zusammenfassung als Word-kompatibles .doc.
 *
 * Trick: Word akzeptiert HTML-Dateien mit der MIME `application/msword` und
 * der Endung `.doc` direkt — keine eigentliche Word-Binary nötig. Das CSS
 * wird mit reingebacken, Word rendert das relativ getreu (Tabellen, Farben,
 * Listen, page-breaks). Das spart eine echte `docx`-Bibliothek (200 KB).
 *
 * Word hat ein paar Eigenheiten:
 *  - `@page` wird ignoriert; das Dokumentlayout kommt aus den Default-
 *    Druckereinstellungen des Users (i.d.R. A4, ist also OK).
 *  - `display:grid` wird nicht unterstützt — die KV-Grids brechen dann
 *    sauber als Liste um, was visuell weiterhin lesbar bleibt.
 */
export function exportSummaryAsDoc(data: SummaryData): void {
  const html = buildSummaryHtml(data);
  // v17.22: Robust statt String-Slice. Vorher wurde bei `<head>` geschnitten
  // (`substring(indexOf('<head>'))`) — fällt indexOf auf -1 (z.B. nach
  // einem Refactor ohne <head>), liefert substring(-1) den ganzen String
  // inkl. doppeltem <html>. Jetzt ersetzen wir das oeffnende <html …>-Tag
  // gezielt durch das Word-Prelude; fehlt es, wird das Prelude vorangestellt.
  const wordPrelude = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">`;
  let wordHtml: string;
  const htmlTagMatch = html.match(/<html[^>]*>/i);
  if (htmlTagMatch) {
    wordHtml = html.replace(/<!doctype[^>]*>/i, '').replace(/<html[^>]*>/i, wordPrelude);
  } else {
    wordHtml = `${wordPrelude}${html}`;
  }
  const blob = new Blob([wordHtml], { type: 'application/msword' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${(data.title || 'event-summary').replace(/[^a-z0-9äöüßÄÖÜ_-]+/gi, '_')}.doc`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
