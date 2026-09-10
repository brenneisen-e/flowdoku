/* Wer von den Organizern eines Events wird ANGEZEIGT — und wer nicht.
 *
 * Die Regel selbst ist alt (v24.8) und steht an sieben Anzeige-Stellen
 * (`EventCard` zweimal, `EventListPage`, `RegistrationPage`, `MyEventCard`,
 * `MyEventsPage`, `registration/EventCard`), überall gleich formuliert:
 *
 *   hideOrganizer && !hideOrganizerIndividualOnly  → gar keiner
 *   hideOrganizer &&  hideOrganizerIndividualOnly  → alle ausser denen in
 *                                                    hiddenOrganizerEmails
 *   sonst                                          → alle
 *
 * Neu ist der ZWEITE Verbraucher dieser Regel: Seit v31.10 richtet sich auch
 * das Organizer-CC der Rundmail danach (Nutzer-Ansage 10.09.2026: „pack nur
 * die organizer auf cc die auch angezeigt werden. wenn keine angezeigt werden
 * dann auch keiner cc"). Damit ist es keine reine Anzeige-Frage mehr, sondern
 * eine Aussage über den Verteiler — und dann darf sie nicht in einer
 * Render-Funktion wohnen. Wer eine dritte Stelle baut, die etwas an
 * „sichtbarer Organizer" festmacht, ruft `visibleOrganizerEmails`.
 *
 * `lastnameTokens` und `pairNamesEmails` sind aus `components/OrganizerList`
 * hierher gezogen (unverändert, Zeichen für Zeichen) — die Komponente
 * importiert sie von hier. Grund: Ein Helfer, den auch der Versand braucht,
 * darf nicht in einer React-Datei liegen; die Richtung Komponente → Utility
 * ist die einzige, die keinen Modul-Zyklus riskiert.
 */

/** Die Felder, die über die Sichtbarkeit entscheiden. */
export interface OrganizerVisibilitySource {
  organizers?: string[];
  organizerEmails?: string[];
  hideOrganizer?: boolean;
  hideOrganizerIndividualOnly?: boolean;
  hiddenOrganizerEmails?: string[];
}

/**
 * Extrahiert Lastname-Tokens aus einem Display-Namen.
 *  "von Albedyll, Benedikt" → ["von", "albedyll"]
 *  "Benedikt von Albedyll" → ["von", "albedyll"]  (letzte 2 Worte als Lastname-Heuristik bei Adelspräfix)
 *  "Heymann, Thorsten"     → ["heymann"]
 *  "Eike Brenneisen"       → ["brenneisen"]
 */
export function lastnameTokens(name: string): string[] {
  const trimmed = (name || '').trim();
  if (!trimmed) return [];
  let lastnamePart: string;
  if (trimmed.indexOf(',') >= 0) {
    lastnamePart = trimmed.split(',')[0].trim();
  } else {
    // Heuristik: letztes Wort ist Lastname; falls vorletztes Wort ein Adelspräfix ist,
    // nimm die letzten beiden.
    const words = trimmed.split(/\s+/).filter(Boolean);
    if (words.length === 0) return [];
    if (words.length === 1) return [words[0].toLowerCase()];
    const PREFIXES = new Set(['von', 'van', 'de', 'der', 'den', 'di', 'da', 'du', 'la', 'le']);
    if (words.length >= 2 && PREFIXES.has(words[words.length - 2].toLowerCase())) {
      lastnamePart = words.slice(-2).join(' ');
    } else {
      lastnamePart = words[words.length - 1];
    }
  }
  return lastnamePart.toLowerCase().split(/\s+/).filter(t => t.length >= 3);
}

/**
 * Defensives Pairing: bevorzugt Index-Match wenn die Email-Local-Part den Lastname
 * der Person enthält, sonst sucht den ersten passenden Email-Eintrag aus dem Pool.
 *
 * Hintergrund: bei legacy oder closure-bug-betroffenen Events kann es vorkommen,
 * dass `organizers[]` und `organizerEmails[]` aus dem SP-Storage out-of-sync sind.
 * Reines `emails[i]`-Pairing zeigt dann das falsche Foto neben dem Namen. Mit dem
 * Lastname-Matcher korrigiert sich das visuell, auch wenn die Storage-Reihenfolge
 * gedreht ist.
 */
export function pairNamesEmails(names: string[], emails: string[]): Array<{ name: string; email: string }> {
  const used = new Set<number>();
  const result: Array<{ name: string; email: string }> = [];
  const localParts = emails.map(e => (e || '').toLowerCase().split('@')[0]);

  const matchesLastname = (emailIdx: number, tokens: string[]): boolean => {
    if (emailIdx < 0 || emailIdx >= localParts.length) return false;
    const local = localParts[emailIdx];
    if (!local) return false;
    return tokens.some(t => local.indexOf(t) >= 0);
  };

  for (let i = 0; i < names.length; i++) {
    const tokens = lastnameTokens(names[i]);
    let chosen = -1;
    if (!used.has(i) && matchesLastname(i, tokens)) {
      chosen = i;
    } else {
      for (let j = 0; j < emails.length; j++) {
        if (!used.has(j) && matchesLastname(j, tokens)) { chosen = j; break; }
      }
    }
    if (chosen < 0) {
      // Kein Lastname-Match — Index-Fallback wenn Slot frei, sonst nächster freier Slot.
      if (!used.has(i) && i < emails.length) chosen = i;
      else for (let j = 0; j < emails.length; j++) if (!used.has(j)) { chosen = j; break; }
    }
    if (chosen >= 0) used.add(chosen);
    result.push({ name: names[i].trim(), email: chosen >= 0 ? (emails[chosen] || '').trim() : '' });
  }
  return result;
}

/**
 * Die E-Mail-Adressen der Organizer, die dem Teilnehmer TATSÄCHLICH angezeigt
 * werden — in der Reihenfolge der Anzeige, ohne Dubletten.
 *
 * Bewusst dieselben zwei Filter wie `OrganizerList`, nicht nur der
 * Ausblenden-Schalter:
 *  - Ein Eintrag ohne Namen wird nicht gerendert (`filter(o => !!o.name)`) und
 *    zählt deshalb auch hier nicht. Folge: Ein Event, an dem nur Adressen und
 *    keine Namen hängen, zeigt keine Organizer — und bekommt dann auch kein
 *    Organizer-CC. Das ist die wörtliche Zusage („wenn keine angezeigt werden,
 *    dann auch keiner im CC"), und die Zusage ist nur so viel wert, wie beide
 *    Seiten dieselbe Rechnung machen.
 *  - Ein Name ohne zugeordnete Adresse liefert nichts, was man anschreiben
 *    könnte, und fällt hier zusätzlich weg.
 */
export function visibleOrganizerEmails(ev: OrganizerVisibilitySource | null | undefined): string[] {
  if (!ev) return [];
  // „Alle ausblenden" ist das Ende der Rechnung — nicht der Anfang einer
  // Ausnahmeliste.
  if (ev.hideOrganizer && !ev.hideOrganizerIndividualOnly) return [];
  const hidden = new Set(
    (ev.hideOrganizer && ev.hideOrganizerIndividualOnly ? (ev.hiddenOrganizerEmails || []) : [])
      .map(e => (e || '').toLowerCase()),
  );
  const seen = new Set<string>();
  const out: string[] = [];
  for (const o of pairNamesEmails(ev.organizers || [], ev.organizerEmails || [])) {
    if (!o.name || !o.email) continue;
    const lc = o.email.toLowerCase();
    if (hidden.has(lc) || seen.has(lc)) continue;
    seen.add(lc);
    out.push(o.email);
  }
  return out;
}
