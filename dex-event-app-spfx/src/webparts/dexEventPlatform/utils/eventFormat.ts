// v18.34 — Hilfsfunktion: lesbaren Outlook-Termin-Ort aus Veranstaltungsort +
// Adresse bauen. `LocationAddress` liegt in DEX_Events als JSON-String
// ({ street, houseNo, zip, city }) vor — der Power-Automate-Flow kann das nicht
// sinnvoll parsen, deshalb baut die App den fertigen Orts-String und legt ihn
// in der Spalte `OutlookLocation` ab. Der Flow mappt diese Spalte 1:1 in das
// Location-Feld des Outlook-Termins.

/**
 * v22.22 — Liegt das Event vollständig in der Vergangenheit?
 *
 * Ende (Fallback: Ende des Start-Tages, 23:59:59) < jetzt. Ohne verwertbare
 * Datumsangaben immer `false` (fail-open: lieber eine Abmeldung zu viel
 * erlauben als eine legitime blockieren). Genutzt für die Abmelde-Sperre
 * vergangener Events (Meine Events / Auto-Cancel-Deep-Link) und die stille
 * Organizer-Abmeldung im Admin Center (keine Mail / kein Outlook / kein
 * Nachrücken mehr).
 */
export function isEventOver(ev: { startDate?: string; endDate?: string }): boolean {
  const endRaw = (ev.endDate || '').trim();
  let end = endRaw ? new Date(endRaw).getTime() : NaN;
  if (!Number.isFinite(end)) {
    const startRaw = (ev.startDate || '').trim();
    if (startRaw) {
      const s = new Date(startRaw);
      if (Number.isFinite(s.getTime())) {
        end = new Date(s.getFullYear(), s.getMonth(), s.getDate(), 23, 59, 59).getTime();
      }
    }
  }
  return Number.isFinite(end) && end < Date.now();
}

interface AddressParts {
  street?: string;
  houseNo?: string;
  zip?: string;
  city?: string;
}

/**
 * Baut einen lesbaren Ort-String, z.B.
 * „Rheinterrasse Düsseldorf, Schwannstraße 6, 40476 Düsseldorf".
 *
 * @param location     Veranstaltungsort (Freitext, z.B. Gebäude-/Raumname)
 * @param address      Adresse als JSON-String ODER Objekt ({street,houseNo,zip,city})
 */
export function buildOutlookLocation(
  location?: string,
  address?: string | AddressParts | null
): string {
  let addr: AddressParts = {};
  if (typeof address === 'string') {
    try { addr = address ? JSON.parse(address) : {}; } catch { addr = {}; }
  } else if (address && typeof address === 'object') {
    addr = address;
  }
  const streetLine = [addr.street, addr.houseNo].map(s => (s || '').trim()).filter(Boolean).join(' ');
  const cityLine = [addr.zip, addr.city].map(s => (s || '').trim()).filter(Boolean).join(' ');
  return [location, streetLine, cityLine].map(s => (s || '').trim()).filter(Boolean).join(', ');
}
