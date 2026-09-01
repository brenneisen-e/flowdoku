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
export function reinsertOrganizerPlaceholder(body: string, organizers: string[]): string {
  if (!body || !organizers || organizers.length === 0) return body;
  if (body.indexOf('{{Organizer}}') >= 0) return body; // schon Platzhalter
  const names = organizers.map(n => (n || '').trim()).filter(Boolean);
  // v29.21 (Audit): Der Save backt den Platzhalter über formatOrganizerList
  // — aus „Nachname, Vorname" (People-Picker-Format der Organizer-Spalte)
  // wird dort „Vorname Nachname". Genau diese Form steht also im
  // gespeicherten Body; die rohen Spalten-Namen matchen beim Regelfall mit
  // Komma NIE, und der v27.3-Mechanismus („Organizer-Wechsel löst beim
  // nächsten Save neu auf") war wirkungslos. Beide Formen probieren.
  const flipped = names.map(n => {
    const c = n.indexOf(',');
    return c > 0 ? `${n.slice(c + 1).trim()} ${n.slice(0, c).trim()}` : n;
  });
  const seenCand: Record<string, boolean> = {};
  const candidates = [...flipped, ...names].filter(n => (seenCand[n] ? false : (seenCand[n] = true)));
  for (const joiner of ['; ', ', ']) {
    const full = candidates.length > 1 ? flipped.join(joiner) : '';
    if (full && body.indexOf(full) >= 0) return body.split(full).join('{{Organizer}}');
  }
  const fullRaw = names.join('; ');
  if (fullRaw && body.indexOf(fullRaw) >= 0) return body.split(fullRaw).join('{{Organizer}}');
  // Bereits „kaputte"/veraltete Bodies: den ersten enthaltenen Organizer-Namen
  // (längster zuerst, um Teil-Treffer zu vermeiden) auf den Platzhalter mappen.
  for (const n of [...candidates].sort((a, b) => b.length - a.length)) {
    if (n.length >= 3 && body.indexOf(n) >= 0) return body.split(n).join('{{Organizer}}');
  }
  return body;
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
