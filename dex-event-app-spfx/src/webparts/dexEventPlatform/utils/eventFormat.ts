// v18.34 — Hilfsfunktion: lesbaren Outlook-Termin-Ort aus Veranstaltungsort +
// Adresse bauen. `LocationAddress` liegt in DEX_Events als JSON-String
// ({ street, houseNo, zip, city }) vor — der Power-Automate-Flow kann das nicht
// sinnvoll parsen, deshalb baut die App den fertigen Orts-String und legt ihn
// in der Spalte `OutlookLocation` ab. Der Flow mappt diese Spalte 1:1 in das
// Location-Feld des Outlook-Termins.

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
