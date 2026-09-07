// v30.66: Aus EventCreationPage.tsx ausgelagert. Diese Helfer stehen auf
// Modul-Ebene und werden von den ausgelagerten Logik-Modulen (wizardSubmit,
// persistSubEvents, outlookChanges) gebraucht — ein Import aus der Seite
// zurueck waere ein Modul-Zyklus.
import { CustomFieldInput } from './customFieldInput';
import { CustomField } from '../../services/EventService';

/**
 * v28.32: Wenn Mail- und Outlook-Kopfbild identisch sind (der Normalfall, seit
 * v28.29 erst recht: Sub-Events erben das Bild des Hauptevents), stand dasselbe
 * Base64 ZWEIMAL im selben Datensatz — einmal als `_eventLogo`, einmal als
 * `_outlookLogo`. Statt der zweiten Kopie wird jetzt nur noch ein Marker
 * gespeichert; beim Laden wird daraus wieder das Mail-Bild. Spart bei einem
 * Foto-Kopfbild rund ein Drittel des Speicher-Payloads.
 */
export const outlookLogoPiggyback = (emailLogo: string, outlookLogo: string): Record<string, unknown> => {
  if (!outlookLogo) return {};
  if (emailLogo && outlookLogo === emailLogo) return { _outlookLogoSameAsMail: true };
  return { _outlookLogo: outlookLogo };
};

/** Gegenstück zu `outlookLogoPiggyback` beim Laden. */
export const readOutlookLogo = (ov: Record<string, unknown> | null | undefined): string => {
  if (!ov) return '';
  if (ov._outlookLogoSameAsMail) return (ov._eventLogo as string) || '';
  return (ov._outlookLogo as string) || '';
};

// v17.22: Einziger Serializer für Custom-Fields → CustomFields-JSON.
// Vorher dreimal copy-paste (Create-Save, Edit-Save, Sub-Event-Save), was
// dazu führte, dass der Sub-Event-Pfad die v17.20-EN-Varianten nicht
// mitnahm. Zentral hier, damit alle drei Pfade identisch persistieren.
//
// Wichtig (v17.22-Fix): DE-Optionen UND EN-Optionen werden POSITIONAL
// gepaart gefiltert — vorher wurde `options` per `.filter(Boolean)` von
// Leereinträgen befreit, `optionsEn` aber nicht, wodurch das Index-Mapping
// zwischen DE und EN bei leeren Slots verrutschte (leere/falsche EN-Labels
// auf der Anmeldeseite).
export function serializeCustomFields(
  fields: CustomFieldInput[],
  bilingual: boolean
): CustomField[] {
  return fields
    .filter(f => f.label && f.label.trim().length > 0)
    .map(f => {
      let optionsOut: string[] | undefined;
      let optionsEnOut: string[] | undefined;
      // v26.75: Vorfilter-Kategorien POSITIONAL zu den (bereinigten) Optionen.
      let categoriesOut: string[] | undefined;
      if (f.type === 'select') {
        const pairs = (f.options || [])
          .map((o, i) => ({ de: (o || '').trim(), en: ((f.optionsEn || [])[i] || '').trim(), cat: (((f.optionCategories || [])[i]) || '').trim() }))
          .filter(p => p.de.length > 0);
        optionsOut = pairs.map(p => p.de);
        if (bilingual && pairs.some(p => p.en.length > 0)) {
          optionsEnOut = pairs.map(p => p.en);
        }
        if (!f.multi && pairs.some(p => p.cat.length > 0)) {
          categoriesOut = pairs.map(p => p.cat);
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
        // v26.74: Vorauswahl nur bei Single-Select persistieren, und nur wenn der
        // Wert eine der (bereinigten) Optionen ist.
        ...(f.type === 'select' && !f.multi && f.defaultValue && (optionsOut || []).indexOf(f.defaultValue) >= 0
          ? { defaultValue: f.defaultValue }
          : {}),
        // v26.75: Vorfilter-Kategorien + optionale Beschriftung (nur Single-Select).
        ...(categoriesOut
          ? { optionCategories: categoriesOut, ...(f.prefilterLabel && f.prefilterLabel.trim() ? { prefilterLabel: f.prefilterLabel.trim() } : {}) }
          : {}),
        // v24.25: Uhrzeit-Flag nur bei Datums-Feldern persistieren.
        ...(f.type === 'date' && f.withTime ? { withTime: true } : {}),
        // v28.63: Buchbares Fenster + Nächte-Limit nur beim Zeitraum-Feld.
        ...(f.type === 'daterange' ? {
          ...(f.rangeStart ? { rangeStart: f.rangeStart } : {}),
          ...(f.rangeEnd ? { rangeEnd: f.rangeEnd } : {}),
          ...(f.maxNights && f.maxNights > 0 ? { maxNights: f.maxNights } : {}),
        } : {}),
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
        // v26.60: Roommate-Benachrichtigung — nur das explizite ABSCHALTEN
        // persistieren (undefined = an, Bestandsverhalten bleibt unverändert).
        ...(f.type === 'roommate' && f.notifyRoommate === false ? { notifyRoommate: false } : {}),
        // v29.40: Verteiler-Begrenzung des Personen-Feldes mitschreiben.
        ...((f.type === 'user' || f.type === 'roommate') && f.audienceOnly ? { audienceOnly: true } : {}),
      } as CustomField;
    });
}

// v27.3: Der Outlook-Body wird beim Speichern mit fest aufgelöstem {{Organizer}}
// gespeichert. Kommen später Organizer dazu, blieb der alte Name eingebacken.
// Beim Edit-Laden mappen wir den eingebackenen Organizer-Namen wieder auf
// {{Organizer}} zurück — dann löst der nächste Save mit ALLEN aktuellen
// Organizern neu auf. Sicher: findet sich nichts, bleibt der Body unverändert.
//
// v30.75: Die Fassung bis v30.74 hat den Termin-Text bei JEDEM Speichern
// aufgebläht. Der Save backt „Vorname Nachname" und verbindet mit „und"/„and"
// (formatOrganizerList: „A, B und C"); die Rück-Suche verglich aber nur gegen
// „A; B" und „A, B" — bei zwei oder mehr Organizern also NIE ein Treffer. Dann
// griff der Notnagel „ersten enthaltenen Namen ersetzen" und machte aus
// „A, B und C" ein „A, B und {{Organizer}}"; der nächste Save setzte dort
// wieder „A, B und C" ein → „A, B und A, B und C". Nach zwanzig Saves stand
// im Outlook-Termin ein halber Bildschirm Namen (Befund 07.09.2026, DTP
// Basics Training). Seit v30.71 („Kommunikation gemeinsam") wurde dieser
// Body außerdem in jeden Termin kopiert.
//
// Jetzt: EIN regulärer Ausdruck erkennt einen ganzen LAUF aus Organizer-Namen
// (beide Schreibweisen), verbunden mit Komma, Semikolon, „und" oder „and", und
// ersetzt ihn komplett. Ein aufgeblähter Absatz enthält u.U. mehrere Läufe
// (dazwischen ein Name, der kein Organizer mehr ist) — dann wird alles vom
// ersten bis zum letzten Lauf zu EINEM Platzhalter. Damit heilt das Laden
// die kaputten Bodies, statt sie weiter zu füttern. Einen Einzelnamen
// irgendwo im Text ersetzt die Funktion nicht mehr teilweise.
function organizerNameCandidates(organizers: string[]): string[] {
  const names = (organizers || []).map(n => (n || '').trim()).filter(Boolean);
  const flipped = names.map(n => {
    const c = n.indexOf(',');
    return c > 0 ? `${n.slice(c + 1).trim()} ${n.slice(0, c).trim()}` : n;
  });
  const seen: Record<string, boolean> = {};
  const out: string[] = [];
  for (const n of [...flipped, ...names]) {
    if (n.length < 3 || seen[n]) continue;
    seen[n] = true;
    out.push(n);
  }
  // Längste zuerst, damit „Anna Berg-Meier" nicht an „Anna Berg" hängen bleibt.
  return out.sort((a, b) => b.length - a.length);
}
const escapeRegExp = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
function organizerRunRegex(candidates: string[]): RegExp {
  const name = `(?:${candidates.map(escapeRegExp).join('|')})`;
  const sep = '(?:\\s*[,;]\\s*|\\s+(?:und|and)\\s+)';
  return new RegExp(`${name}(?:${sep}${name})*`, 'g');
}
const ORGANIZER_PH = '{{Organizer}}';

export function reinsertOrganizerPlaceholder(body: string, organizers: string[]): string {
  if (!body || !organizers || organizers.length === 0) return body;
  if (body.indexOf(ORGANIZER_PH) >= 0) return body; // schon Platzhalter
  const candidates = organizerNameCandidates(organizers);
  if (candidates.length === 0) return body;
  const run = organizerRunRegex(candidates);
  return body.split('</p>').map(par => {
    const replaced = par.replace(run, ORGANIZER_PH);
    const first = replaced.indexOf(ORGANIZER_PH);
    const last = replaced.lastIndexOf(ORGANIZER_PH);
    if (first < 0 || first === last) return replaced;
    return replaced.slice(0, first) + ORGANIZER_PH + replaced.slice(last + ORGANIZER_PH.length);
  }).join('</p>');
}

/** v30.75: Trägt der gespeicherte Body die Aufblähung aus dem Fehler oben
 *  (mehr als ein Organizer-Lauf in einem Absatz)? Dann muss der Wizard beim
 *  Speichern ein Outlook-Update anbieten — der Vergleich „Text geändert?"
 *  sieht nach der Heilung auf beiden Seiten denselben Platzhalter. */
export function outlookBodyOrganizerBloated(body: string, organizers: string[]): boolean {
  if (!body || body.indexOf(ORGANIZER_PH) >= 0) return false;
  const candidates = organizerNameCandidates(organizers);
  if (candidates.length === 0) return false;
  const run = organizerRunRegex(candidates);
  // v30.77: Ein Lauf ist auch dann aufgebläht, wenn er einen Namen ZWEIMAL
  // enthält. Die erste Fassung zählte nur die Läufe je Absatz — sobald aber
  // alle Namen der Flut noch Organizer sind, ist die ganze Flut EIN Lauf, und
  // das Update fürs Hauptevent blieb aus (Befund 07.09.2026, DTP Basics:
  // „leider Outlook nicht geändert").
  const single = new RegExp(`(?:${candidates.map(escapeRegExp).join('|')})`, 'g');
  return body.split('</p>').some(par => {
    const runs = par.match(run) || [];
    if (runs.length >= 2) return true;
    return runs.some(r => {
      const names = r.match(single) || [];
      const distinct: Record<string, true> = {};
      names.forEach(n => { distinct[n] = true; });
      return names.length > Object.keys(distinct).length;
    });
  });
}

/**
 * v16.4: Audience-Liste (kommasepariert) in eine flache, ';'-separierte
 * Liste von Member-E-Mails auflösen. Jede '@'-Eintrag wird via
 * getGroupMembers (Graph) probiert — wenn die Auflösung eine
 * Mitglieder-Liste liefert, werden alle deren E-Mails übernommen, sonst
 * wird der Eintrag als direkte User-E-Mail behandelt. Lowercase + dedupliziert.
 *
 * Wird beim Event-Save aufgerufen und in die SP-Spalte
 * `AudienceResolvedEmails` geschrieben. matchesAudience im
 * EventListPage prüft Sichtbarkeit zur Laufzeit gegen diese Liste.
 */
export async function resolveAudienceMembersToCsv(
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
        // Keine Member zurückgeliefert → behandle als direkte User-Adresse.
        out.add(item.toLowerCase());
      }
    } catch {
      out.add(item.toLowerCase());
    }
  }
  return Array.from(out).join(';');
}
