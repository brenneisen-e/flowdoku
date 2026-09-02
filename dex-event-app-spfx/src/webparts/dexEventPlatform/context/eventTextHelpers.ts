/**
 * Reine Modul-Helfer des Event-Contexts: Mail-Templates, Text-Aufbereitung,
 * Audit-Diff und Bild-URLs.
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Alles hier ist zustandsfrei — kein React, kein EventService, kein Zugriff
 * auf den Provider-State. `EventContext.tsx` re-exportiert die oeffentlichen
 * Funktionen unveraendert, damit keine Aufrufstelle angefasst werden musste.
 */

/**
 * Organizer-Namen für Mail-Anreden sauber formatieren:
 *   Input:  ['Sathasivam, Philipp', 'Oesterle, Ines']
 *   Output: 'Philipp Sathasivam und Ines Oesterle'  (bei DE)
 *           'Philipp Sathasivam and Ines Oesterle'  (bei EN)
 *
 * - Namen können auch ';'-separiert als Einzel-String kommen, wird gesplittet.
 * - Nachname/Vorname-Pairs werden vorgetauscht (SP-Default ist "Nachname, Vorname").
 * - Bei 1 Name: nur der Name. Bei 2: "A und B" / "A and B". Bei 3+: "A, B und C" / "A, B and C".
 */
/**
 * Wendet Event-spezifische Template-Overrides auf die globale SP-Vorlage an.
 *
 * - Override-JSON-Format: { "Anmeldung": { subject, heading, bodyHtml }, ... }
 * - Pro Feld gilt: Override > globale SP-Vorlage. headingColor bleibt immer
 *   die globale (Overrides ändern keine Brand-Farben).
 * - Wenn weder Override noch SP-Template existieren, gibt die Funktion null
 *   zurück und der Caller fällt auf das Code-Default zurück.
 */
export function applyEventTemplateOverride(
  spTemplate: { subject: string; headingColor: string; heading: string; subheading?: string; bodyHtml: string } | null,
  overridesJson: string | undefined,
  templateType: string
): { subject: string; headingColor: string; heading: string; subheading: string; bodyHtml: string; headingFontSize?: string; headingBold?: boolean; headingItalic?: boolean; subheadingColor?: string; subheadingFontSize?: string; subheadingBold?: boolean; subheadingItalic?: boolean; imageWidth?: number; imagePaddingV?: number; imagePaddingH?: number } | null {
  // v15.19: Subheading-Override pro Event mitziehen. Color/Size bleiben
  // weiterhin aus dem Standard-Template (wrapTemplate-Layout fest), nur
  // die Text-Werte (Subject, Heading, Subheading, Body) sind editierbar.
  if (!overridesJson) {
    if (!spTemplate) return null;
    return {
      subject: spTemplate.subject,
      headingColor: spTemplate.headingColor || '#86bc25',
      heading: spTemplate.heading,
      subheading: spTemplate.subheading || '',
      bodyHtml: spTemplate.bodyHtml,
    };
  }
  try {
    const all = JSON.parse(overridesJson) as Record<string, { subject?: string; heading?: string; subheading?: string; bodyHtml?: string; headingColor?: string; headingFontSize?: string; headingBold?: boolean; headingItalic?: boolean; subheadingColor?: string; subheadingFontSize?: string; subheadingBold?: boolean; subheadingItalic?: boolean }>;
    // v18.73: globales Header-Bild-Layout (Breite + Innenabstand). Liegt unter
    // dem reservierten Piggyback-Key `_headerImageLayout` und gilt für ALLE
    // Template-Typen des Events — daher hier einmal gelesen und in jeden
    // Rückgabe-Zweig gespreadet (auch wenn der konkrete Typ keinen Text-
    // Override hat).
    const il = (all as unknown as { _headerImageLayout?: { width?: number; paddingV?: number; paddingH?: number } })._headerImageLayout || {};
    const imgSpread = {
      ...(typeof il.width === 'number' && il.width > 0 ? { imageWidth: il.width } : {}),
      ...(typeof il.paddingV === 'number' && il.paddingV >= 0 ? { imagePaddingV: il.paddingV } : {}),
      ...(typeof il.paddingH === 'number' && il.paddingH >= 0 ? { imagePaddingH: il.paddingH } : {}),
    };
    const o = all[templateType];
    if (!o || (!o.subject && !o.heading && o.subheading === undefined && !o.bodyHtml && !o.headingColor && !o.headingFontSize && o.headingBold === undefined && o.headingItalic === undefined && !o.subheadingColor && !o.subheadingFontSize && o.subheadingBold === undefined && o.subheadingItalic === undefined)) {
      if (!spTemplate) return null;
      return {
        subject: spTemplate.subject,
        headingColor: spTemplate.headingColor || '#86bc25',
        heading: spTemplate.heading,
        subheading: spTemplate.subheading || '',
        bodyHtml: spTemplate.bodyHtml,
        ...imgSpread,
      };
    }
    return {
      subject: o.subject || spTemplate?.subject || '',
      heading: o.heading || spTemplate?.heading || '',
      // Override.subheading ist „intentional set" — auch leerer String
      // soll respektiert werden, damit man die zweite Zeile abschalten kann.
      subheading: o.subheading !== undefined ? o.subheading : (spTemplate?.subheading || ''),
      bodyHtml: o.bodyHtml || spTemplate?.bodyHtml || '',
      // v18.19: Überschrift-Farbe + -Größe pro Event überschreibbar.
      headingColor: o.headingColor || spTemplate?.headingColor || '#86bc25',
      ...(o.headingFontSize ? { headingFontSize: o.headingFontSize } : {}),
      // v18.22: Fett/Kursiv (Überschrift) + Unter-Überschrift-Formatierung.
      ...(o.headingBold !== undefined ? { headingBold: o.headingBold } : {}),
      ...(o.headingItalic !== undefined ? { headingItalic: o.headingItalic } : {}),
      ...(o.subheadingColor ? { subheadingColor: o.subheadingColor } : {}),
      ...(o.subheadingFontSize ? { subheadingFontSize: o.subheadingFontSize } : {}),
      ...(o.subheadingBold !== undefined ? { subheadingBold: o.subheadingBold } : {}),
      ...(o.subheadingItalic !== undefined ? { subheadingItalic: o.subheadingItalic } : {}),
      ...imgSpread,
    };
  } catch {
    if (!spTemplate) return null;
    return {
      subject: spTemplate.subject,
      headingColor: spTemplate.headingColor || '#86bc25',
      heading: spTemplate.heading,
      subheading: spTemplate.subheading || '',
      bodyHtml: spTemplate.bodyHtml,
    };
  }
}

/**
 * Strip SharePoint-Note-Field-Wrapper.
 *
 * Seit der Migration der Felder Organizer + OrganizerEmail von Single-Line-Text
 * auf Note (Multi-Line-Text, Plain) — nötig wegen 255-Char-Limit bei 10+ Co-
 * Organizern — wickelt SharePoint die Werte beim REST-Read in einen
 * `<div class="ExternalClassXXXX">…</div>`-Container. Das passiert obwohl
 * `RichText: false` gesetzt ist und ist eine bekannte SP-Quirk.
 *
 * Folge ohne Strip: `(e.Organizer || '').split(';')` zerhackt den Wrapper an
 * den Semikolons, das erste und letzte Stück enthalten dann die Tag-Reste
 * `<div class="…">…` bzw. `…</div>` und landen so in den Chip-Labels.
 *
 * Idempotent: Eingaben ohne Wrapper bleiben unverändert.
 */
export function stripSpNoteWrapper(value: string | null | undefined): string {
  if (!value) return '';
  let v = value.trim();
  v = v.replace(/^<div\b[^>]*>/i, '');
  v = v.replace(/<\/div>\s*$/i, '');
  return v.trim();
}

export function formatOrganizerList(organizers: string[], lang: string): string {
  const names: string[] = [];
  for (const entry of organizers || []) {
    // Akzeptiere ';' UND ',' als Top-Level-Trenner zwischen Personen.
    // Wenn die Anzahl der Komma-Tokens gerade und >=2 ist, behandeln wir sie als
    // Paare ('Lastname, Firstname, Lastname, Firstname, ...'). Sonst fallen wir
    // zurück auf Semikolon-Split + 'Lastname, Firstname' pro Stück.
    const raw = (entry || '').trim();
    if (!raw) continue;
    const semiPieces = raw.split(';').map(p => p.trim()).filter(Boolean);
    const pieces: string[] = [];
    for (const sp of semiPieces) {
      const commaTokens = sp.split(',').map(s => s.trim()).filter(Boolean);
      if (commaTokens.length >= 4 && commaTokens.length % 2 === 0) {
        // Paarweise interpretieren: ['Last','First','Last','First',...]
        for (let i = 0; i < commaTokens.length; i += 2) {
          pieces.push(`${commaTokens[i]}, ${commaTokens[i + 1]}`);
        }
      } else {
        pieces.push(sp);
      }
    }
    for (const piece of pieces) {
      const commaParts = piece.split(',').map(s => s.trim());
      if (commaParts.length === 2 && commaParts[0] && commaParts[1]) {
        names.push(`${commaParts[1]} ${commaParts[0]}`);
      } else {
        names.push(piece);
      }
    }
  }
  if (names.length === 0) return '';
  if (names.length === 1) return names[0];
  const conj = (lang || 'EN').toUpperCase() === 'DE' ? ' und ' : ' and ';
  if (names.length === 2) return `${names[0]}${conj}${names[1]}`;
  return `${names.slice(0, -1).join(', ')}${conj}${names[names.length - 1]}`;
}

/** v18.41: Sammelt die E-Mail-Adressen aus People-Picker-Feldern (user/roommate),
 *  die der Organizer als „CC bei An-/Abmelde-Mail" markiert hat. Format des
 *  Feldwerts ist „Anzeigename <email>". Liefert einen ';'-getrennten CC-String
 *  (ohne den Teilnehmer selbst, dedupliziert). NUR für Mails — nicht Outlook. */
export function collectCcEmailsFromFields(
  fields: Array<{ id: string; type: string; ccOnEmails?: boolean }> | undefined,
  customData: Record<string, string>,
  excludeEmail?: string
): string {
  const seen = new Set<string>();
  const out: string[] = [];
  const exclude = (excludeEmail || '').toLowerCase();
  for (const f of (fields || [])) {
    if (!f.ccOnEmails) continue;
    if (f.type !== 'user' && f.type !== 'roommate') continue;
    const v = customData[f.id];
    if (!v) continue;
    const m = v.match(/<([^>]+@[^>]+)>/);
    const em = m ? m[1].trim() : '';
    const lc = em.toLowerCase();
    if (em && lc !== exclude && !seen.has(lc)) { seen.add(lc); out.push(em); }
  }
  return out.join(';');
}

/**
 * v28.28: Zwei CC-Listen (Semikolon-getrennt) zusammenführen — ohne Dubletten
 * und ohne die Empfängeradresse selbst. Gebraucht, seit die Organizer-
 * Mitlese-Kopie auf CC statt BCC läuft und sich damit eine Liste mit den
 * CC-Feldern der Anmeldung teilt.
 */
export function mergeCcLists(a: string | undefined, b: string | undefined, excludeEmail?: string): string | undefined {
  const seen = new Set<string>();
  const out: string[] = [];
  const exclude = (excludeEmail || '').trim().toLowerCase();
  for (const part of [a || '', b || '']) {
    for (const em of part.split(';').map(s => s.trim()).filter(Boolean)) {
      const lc = em.toLowerCase();
      if (lc === exclude || seen.has(lc)) continue;
      seen.add(lc);
      out.push(em);
    }
  }
  return out.length ? out.join(';') : undefined;
}

// v19.33: SP-Spaltennamen → lesbare Labels fürs Event-Audit-Log.
const EVENT_AUDIT_LABELS: Record<string, string> = {
  Title: 'Titel', Description: 'Beschreibung', Location: 'Ort', LocationAddress: 'Adresse',
  StartDate: 'Start', EndDate: 'Ende', RegistrationDeadline: 'Anmeldeschluss',
  LastDeregisterDate: 'Letzte Abmeldung', MaxParticipants: 'Teilnehmerzahl',
  WaitlistEnabled: 'Warteliste', DisableEmails: 'E-Mails',
  DisableRegistrationEmail: 'Anmelde-Bestätigung', DisableCancellationEmail: 'Abmelde-Bestätigung',
  AutoDeregisterOnDecline: 'Outlook-Absage = Abmeldung', InactiveHandling: 'Ex-Deloitte-Konto: Verhalten', DisableOutlook: 'Outlook-Termin',
  EmailLanguage: 'Mail-Sprache', RegistrationLanguage: 'Anmeldesprache',
  CustomFields: 'Eventfelder', Agenda: 'Agenda', Transfers: 'Transferzeiten', FunZone: 'Quiz',
  OutlookBody: 'Outlook-Text', OutlookSubject: 'Outlook-Betreff', OutlookLocation: 'Outlook-Ort',
  OutlookStart: 'Outlook-Start', OutlookEnd: 'Outlook-Ende',
  LocationFilter: 'Standortfilter', Audience: 'Mailverteiler', FilterMode: 'Filterverknüpfung',
  ExcludedUsers: 'Ausgeschlossene Personen', AudienceResolvedEmails: 'Sichtbarkeits-Cache',
  Organizer: 'Organizer', OrganizerEmail: 'Organizer-Mails',
  ContactName: 'Ansprechpartner', ContactEmail: 'Kontakt-Mail', ContactInfo: 'Kontakt-Info',
  EventImageUrl: 'Event-Bild', EmailImageBase64: 'Mail-Logo', EmailTemplateOverrides: 'Mail-Vorlagen',
  DurchstarterCapacity: 'Kapazität Gruppe A', FunstarterCapacity: 'Kapazität Gruppe B',
  SplitLabelA: 'Label Gruppe A', SplitLabelB: 'Label Gruppe B', SplitHelpText: 'Gruppen-Hinweistext', SplitSectionTitle: 'Gruppen-Überschrift', SplitSharedWaitlist: 'Gemeinsame Warteliste',
  TeamRegistrationEnabled: 'Team-Anmeldung', TeamSize: 'Teamgröße', AskTeamName: 'Team-Name abfragen',
  AskSalutation: 'Anrede abfragen', BilingualFields: 'Zweisprachig',
  ConfirmDialogEnabled: 'Bestätigungs-Dialog', SelfCheckInEnabled: 'Self-Check-in',
  ActiveFrom: 'Sichtbar ab', NotifyOrgRegisterMode: 'Organizer-Kopie Anmeldung',
  NotifyOrgCancelMode: 'Organizer-Kopie Abmeldung', AllowAttendeeUpload: 'Datei-Upload',
};

/**
 * v19.33: Diff zwischen altem Roh-SP-Item und dem Update-Payload (beide
 * SP-Spalten-Format → verlässlicher Vergleich). Liefert NUR die wirklich
 * geänderten Felder als `{ Label: { old, new } }`. Lange/JSON-Felder werden
 * nicht ausgeschrieben (nur „(geändert)"), Booleans als „an"/„aus".
 */
// v26.20: Lesbare Zusammenfassung der Custom-Felder fürs Audit-Log —
// „Label: Beschreibung | …". Dient als Vorher/Nachher-Verlauf der Felder UND
// als Sicherheitsnetz, falls Feld-Beschreibungen (helpText) verloren gehen
// (dann steht die alte Beschreibung noch im Audit-Eintrag). Akzeptiert sowohl
// das CustomFields-JSON (String) als auch ein bereits geparstes Array.
export function summarizeCustomFields(raw: unknown): string {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let arr: any[];
  try { arr = Array.isArray(raw) ? raw : JSON.parse(typeof raw === 'string' ? (raw || '[]') : '[]'); } catch { return ''; }
  if (!Array.isArray(arr) || arr.length === 0) return '(keine Felder)';
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parts = arr.map((f: any) => {
    const label = String((f && (f.label || f.id)) || '?').trim();
    const help = String((f && f.helpText) || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
    return help ? `${label}: ${help.length > 200 ? `${help.slice(0, 200)}…` : help}` : label;
  });
  let s = parts.join(' | ');
  if (s.length > 2000) s = `${s.slice(0, 2000)}…`;
  return s;
}

export function buildEventUpdateDiff(
  oldItem: Record<string, unknown>,
  updates: Record<string, unknown>,
): Record<string, { old: string; new: string }> {
  const norm = (v: unknown): string => {
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'true' : 'false';
    return String(v).trim();
  };
  const same = (a: unknown, b: unknown): boolean => {
    const na = norm(a), nb = norm(b);
    if (na === nb) return true;
    const da = Date.parse(na), db = Date.parse(nb);
    if (!isNaN(da) && !isNaN(db)) return da === db;
    return false;
  };
  const prettyBool = (v: string): string => v === 'true' ? 'an' : v === 'false' ? 'aus' : v;
  const trunc = (s: string): string => s.length > 80 ? `${s.slice(0, 77)}…` : s;
  const opaque = new Set(['CustomFields', 'Agenda', 'Transfers', 'FunZone', 'EmailTemplateOverrides', 'AudienceResolvedEmails', 'EmailImageBase64', 'OutlookBody', 'ExcludedUsers', 'Documents']);
  const out: Record<string, { old: string; new: string }> = {};
  for (const key of Object.keys(updates)) {
    if (same(oldItem[key], updates[key])) continue;
    const label = EVENT_AUDIT_LABELS[key] || key;
    if (key === 'CustomFields') {
      // v26.20: Statt nur „(geändert)" die Feld-Beschreibungen vorher/nachher
      // ausschreiben — Verlauf + Sicherheitsnetz für verlorene Beschreibungen.
      out[label] = { old: summarizeCustomFields(oldItem[key]) || '(leer)', new: summarizeCustomFields(updates[key]) || '(leer)' };
    } else if (opaque.has(key)) {
      out[label] = { old: '(vorher)', new: '(geändert)' };
    } else {
      out[label] = { old: trunc(prettyBool(norm(oldItem[key]))) || '(leer)', new: trunc(prettyBool(norm(updates[key]))) || '(leer)' };
    }
  }
  return out;
}

// v26.17: Cache-Buster für das Event-Bild. Der Browser-Bild-Cache (IndexedDB,
// utils/imageCache) speichert pro BILD-URL — ändert sich beim Bild-Wechsel die
// URL NICHT (z.B. Attachment mit gleichem Dateinamen überschrieben), zeigte die
// App dauerhaft das alte, gecachte Bild. Wir hängen daher die letzte
// Änderungszeit des Events als `?v=`-Parameter an die ANZEIGE-URL — ändert sich
// das Event (= auch beim Bild-Tausch), wird die URL neu und der Cache greift
// frisch. Nur für die Anzeige; das gespeicherte EventImageUrl bleibt unberührt.
export function buildDisplayImageUrl(rawUrl: string, modified?: string): string {
  const url = (rawUrl || '').trim();
  if (!url || url.indexOf('data:') === 0) return url;
  if (!modified) return url;
  let v = '';
  try { v = String(new Date(modified).getTime()); } catch { v = ''; }
  if (!v || v === 'NaN') return url;
  return `${url}${url.indexOf('?') >= 0 ? '&' : '?'}v=${v}`;
}
