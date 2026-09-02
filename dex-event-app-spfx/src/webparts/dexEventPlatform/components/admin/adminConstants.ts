/**
 * Modul-Ebene aus AdminPage.tsx ausgelagert (v30.66) — Konstanten und der
 * eine reine Helfer, die vor der Komponente standen. Sie kennen den State
 * nicht und lassen sich deshalb als Ganzes verschieben; der Rueckweg waere
 * ein Modul-Zyklus, deshalb liegen sie hier und nicht in der Seite.
 */

// v14.11 / v19.30: Aggregierte Zeile im konsolidierten View (Hauptevent mit
// Sub-Events), eine pro Person. Auf Modul-Ebene definiert, damit auch
// Handler außerhalb des Render-Bodys (z.B. das Abmelde-/Edit-Modal von
// Feature A/B) den Typ referenzieren können.
// v28.21: Zeilen-Status, die einen echten Platz belegen. Für die
// Doppel-Anmelde-Erkennung: „alles außer Abgemeldet" war zu weit gefasst.
export const DUP_ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];

/**
 * v30.33: Beispiel-Teilnehmer-ID für Vorschau, Mail-Editor und Test-Versand.
 * Dort gibt es keine echte Person — der Organizer soll aber sehen, dass unter
 * dem QR-Code eine ID mitgeschickt wird, sonst fällt beim Gestalten der Mail
 * niemandem auf, wenn sie fehlt.
 */
export const SAMPLE_QR_ID = 17;

/**
 * v30.37: Sentinel für `regLoadError` — „darf die Liste nicht lesen" statt
 * „Lesen fehlgeschlagen". Als Marke und nicht als fertiger Satz, weil die
 * Meldung zweisprachig gerendert wird und `regLoadError` nur einen String
 * trägt.
 */
export const ACCESS_DENIED_MSG = '__DEX_ACCESS_DENIED__';

/**
 * v29.37: Kopfbild-Layout EINES Events (Piggyback `_headerImageLayout`, im
 * Wizard eingestellt — seit v29.29 bei neuen Events die volle Mailbreite).
 * Die Mails aus dem Organizer Center (Einladung/Erinnerung, Massenmail)
 * starteten fest bei 180/30/30 und ignorierten diese Einstellung: Dasselbe
 * Event verschickte seine Anmeldebestätigung mit Vollbild-Kopf und die
 * Erinnerung mit kleinem, zentriertem Bild.
 *
 * Hat das Event GAR KEINE Einstellung (Alt-Event, der Schlüssel wird nur bei
 * Abweichung vom Alt-Default geschrieben), gilt hier die volle Breite — wie
 * bei neuen Events seit v29.29. Das kleine zentrierte Bild ist die Ausnahme,
 * nicht der Regelfall; wer es will, stellt es im Mail-Editor um.
 */
export function eventHeaderImageLayout(overridesJson: string | undefined): { width: number; paddingV: number; paddingH: number } {
  const fullWidth = { width: 600, paddingV: 0, paddingH: 0 };
  if (!overridesJson) return fullWidth;
  try {
    const il = (JSON.parse(overridesJson) || {})._headerImageLayout;
    if (!il || typeof il !== 'object') return fullWidth;
    return {
      width: typeof il.width === 'number' && il.width > 0 ? il.width : 180,
      paddingV: typeof il.paddingV === 'number' && il.paddingV >= 0 ? il.paddingV : 30,
      paddingH: typeof il.paddingH === 'number' && il.paddingH >= 0 ? il.paddingH : 30,
    };
  } catch { return fullWidth; }
}

// v30.52: `headerOptsFor` ist als `mailHeaderOpts` nach utils/mailHeaderImage
// gewandert — der Orb-Schutz (v29.37) wird auch beim automatischen QR-Versand
// gebraucht, wo es keine Oberfläche gibt.
