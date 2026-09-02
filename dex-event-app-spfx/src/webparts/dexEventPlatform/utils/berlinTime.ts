// v30.66: Aus dem Komponentenkoerper von EventCreationPage.tsx herausgezogen.
//
// Der Grund ist kein Aufraeumen, sondern ein Fehler: `isoToLocal` stand als
// `const`-Arrow im Funktionskoerper (Zeile 372), wurde aber schon in einem
// useState-Lazy-Initializer weiter oben (Zeile 205) aufgerufen. Der laeuft beim
// ERSTEN Render — also bevor die Deklaration erreicht ist. Bei target=es5 wird
// daraus `var isoToLocal`, der Aufruf trifft `undefined`; das umgebende
// `catch`, das nur kaputtes JSON abfangen sollte, verschluckte den TypeError
// still. Folge: das feste "Anmeldung ab"-Datum startete IMMER leer und wurde
// beim naechsten Speichern aus den Overrides geloescht.
//
// Die drei Funktionen haengen an keinem State — auf Modul-Ebene kann diese
// Falle nicht wiederkehren, egal an welcher Stelle sie kuenftig gebraucht werden.

/** Gibt den Offset von Europe/Berlin zu UTC an dem gegebenen Zeitpunkt in ms zurück.
 *  Im Winter: +3600000 (+1h). Im Sommer: +7200000 (+2h). */
export const berlinOffsetMs = (dateUtc: Date): number => {
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
export const berlinLocalToUtcIso = (localStr: string): string => {
  if (!localStr) return '';
  // Parse den String erstmal als ob er UTC wäre -> das sind UTC-Zahlen die den Berlin-Werten entsprechen
  const asUtc = new Date(localStr.length === 16 ? localStr + ':00Z' : localStr + 'Z');
  if (isNaN(asUtc.getTime())) return '';
  // Der echte UTC-Zeitpunkt ist asUtc minus Berlin-Offset an diesem Zeitpunkt
  const offset = berlinOffsetMs(asUtc);
  return new Date(asUtc.getTime() - offset).toISOString();
};

/** UTC-ISO ("2026-04-23T17:00:00.000Z") nach datetime-local in Europe/Berlin
 *  ("2026-04-23T19:00") konvertieren — für das Input-Feld. */
export const isoToLocal = (iso: string): string => {
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
