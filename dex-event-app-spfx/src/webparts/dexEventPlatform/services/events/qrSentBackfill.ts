/**
 * v31.4: Die gedruckten QR-Nummern aus der Mail-Warteschlange zurückholen.
 *
 * Ausgangslage (Befund 08.09.2026, laufendes Event): Die Spalte `QrSentId`
 * gibt es erst seit v31.4. Für alle Events, deren QR-Mails vorher rausgingen,
 * steht nirgends, welche Nummer in welcher Mail gedruckt war — und genau die
 * Events sind betroffen, weil dort schon Abmeldungen die `TeilnehmerID`
 * verschoben haben.
 *
 * Die Nummern stehen aber noch da: `DEX_Emails` (Haupt-Site) hält den
 * kompletten HTML-Body jeder verschickten Mail, und die Zeilen werden erst
 * rund einen Monat nach dem Event archiviert. Dieses Modul liest sie und
 * parst die Zahl aus dem Block, den `EmailTemplates.buildQrBlockHtml`
 * erzeugt. Wer dort das Markup ändert, zieht `parseQrIdFromBody` im selben
 * Commit nach.
 *
 * Bewusst NUR lesen und parsen: Was daraus geschrieben wird, entscheidet der
 * Aufrufer nach einer Vorschau (components/admin/QrSentIdBackfillModal).
 */
import { SPHttpClient } from '@microsoft/sp-http';
import { EventService } from '../EventService';

/** Eine gefundene QR-Mail — je Mail-Zeile eine. */
export interface QrMailHit {
  /** Id der `DEX_Emails`-Zeile. Die HÖCHSTE gewinnt: der spätere Versand ist
   *  der gültige (ein zweiter Versand ersetzt die erste Mail). */
  mailId: number;
  /** Adresse der Person, um die es geht — kleingeschrieben. Bei einer an den
   *  Organizer umgeleiteten Mail die externe Adresse aus dem Hinweiskasten,
   *  sonst der Empfänger. */
  email: string;
  /** Die Adresse, an die die Mail tatsächlich ging (kleingeschrieben). */
  recipient: string;
  /** true = umgeleitet („Eigentlich für …"), die Mail ging an den Organizer. */
  redirected: boolean;
  /** Die Nummer, die in dieser Mail gedruckt stand. */
  qrId: number;
  /**
   * v31.4 (Review): Versand-Status der Zeile ('Pending' | 'Sent' | 'Failed',
   * leer bei Alt-Zeilen). Eine Mail mit 'Failed' wurde NIE zugestellt — sie
   * hat also auch nie eine Nummer gedruckt und darf eine tatsächlich
   * verschickte nicht über „die höchste Id gewinnt" verdrängen.
   */
  status: string;
}

export interface QrMailScan {
  /** false = ABGEBROCHEN. Dann sagt `status`, woran es lag; `hits` ist
   *  unvollständig und darf NICHT als „nichts gefunden" gelesen werden
   *  (CLAUDE.md: „Ein Lesefehler ist keine Null"). */
  ok: boolean;
  /** HTTP-Status des Abbruchs, 0 bei Netz-/Parse-Fehler. */
  status: number;
  hits: QrMailHit[];
  /** Gelesene Mail-Zeilen insgesamt (auch die ohne erkennbare Nummer). */
  scanned: number;
  /** Zeilen, in denen keine gedruckte Nummer stand (alte Mails vor v30.35). */
  unparsed: number;
  /** v31.4.1: Gelesene Zeilen, deren `EventId` zum Event passt. */
  byId: number;
  /** v31.4.1: Gelesene Zeilen, die NUR über den Titel gefunden wurden. Ist die
   *  Zahl hoch, trägt die Queue eine andere EventId als das geöffnete Event —
   *  genau der Fall, an dem der Nachtrag am 08.09.2026 vorbeigelaufen ist. */
  byTitle: number;
}

/**
 * Die gedruckte Nummer aus einem Mail-Body.
 *
 * Erzeugt wird sie in `buildQrBlockHtml` als
 * `<span style="color:#63666A;">ID:</span> <strong style="…">012</strong>`.
 * Zuerst wird genau darauf geprüft (die Farbe macht den Treffer eindeutig);
 * schlägt das fehl — SharePoint darf Rich-Text-Felder umformatieren —, greift
 * ein toleranteres Muster auf „`ID:` … Ziffern vor `</strong>`".
 * `parseInt` wegen der führenden Nullen („012" ist 12, nicht oktal: parseInt
 * mit Basis 10).
 */
export function parseQrIdFromBody(body: string): number | null {
  const html = body || '';
  const strict = /color:#63666A;">\s*ID:\s*<\/span>\s*<strong[^>]*>\s*(\d{1,6})\s*<\/strong>/i.exec(html);
  const loose = strict || /ID:\s*(?:<\/?[a-zA-Z][^>]*>\s*)*?(\d{1,6})\s*<\/strong>/i.exec(html);
  if (!loose) return null;
  const n = parseInt(loose[1], 10);
  return (isFinite(n) && n > 0) ? n : null;
}

/**
 * Für wen war die Mail gedacht? Bei externen Teilnehmern geht die QR-Mail an
 * den Organizer, und der Body sagt oben „Eigentlich für <strong>adresse</strong>".
 * Ohne diese Auflösung landeten alle externen Nummern auf dem Organizer.
 */
export function parseIntendedRecipient(body: string): string {
  const m = /Eigentlich für\s*<strong>\s*([^<]*@[^<]*?)\s*<\/strong>/i.exec(body || '');
  return m ? m[1].trim().toLowerCase() : '';
}

/** Die erste Adresse eines Empfänger-Felds (mehrere Organizer stehen mit „;"). */
function firstAddress(recipient: string): string {
  return (recipient || '').split(/[;,]/)[0].trim().toLowerCase();
}

/**
 * Alle QR-Mails EINES Events aus `DEX_Emails` lesen und die gedruckten
 * Nummern parsen.
 *
 * `$top=20`, weil jeder Body den QR-Code als Base64 enthält und damit 20–50 KB
 * groß ist; der Body wird nach dem Parsen sofort fallen gelassen. Bei einem
 * HTTP-Fehler wird ABGEBROCHEN (`ok: false`) statt die halbe Liste
 * zurückzugeben: `DEX_Emails` hat `ReadSecurity=2` — wer die Zeilen nicht
 * sehen darf, bekommt eine LEERE Antwort ohne Fehler, und „leer" darf hier
 * niemals als „es gab keine Mails" durchgehen.
 */
export async function scanQrMailsForEvent(svc: EventService, eventId: string, eventTitle?: string): Promise<QrMailScan> {
  const out: QrMailScan = { ok: true, status: 0, hits: [], scanned: 0, unparsed: 0, byId: 0, byTitle: 0 };
  const safeId = (eventId || '').replace(/'/g, "''");
  if (!safeId) return out;
  /**
   * v31.4.1 (Live-Befund 08.09.2026): Der Filter lief NUR über `EventId` und
   * fand 12 von 93 Mails — während die Liste in SharePoint, nach `EventTitle`
   * gefiltert, rund neunzig zeigte. Bewiesen war das am eigenen Ergebnis: Der
   * Kasten „N Mails ohne gedruckte Nummer" blieb aus, es waren also ALLE
   * gelesenen Mails auswertbar — gelesen wurden schlicht nur zwölf.
   *
   * `EventId` steht auf einer Queue-Zeile so, wie sie beim Versand geschrieben
   * wurde; ältere Zeilen (anderer Versandweg, kopiertes oder neu angelegtes
   * Event, Zeile aus einer Zeit vor einer Umstellung) tragen dort etwas
   * anderes oder nichts. Der Titel ist die zweite, unabhängige Spur.
   *
   * Warum das trotzdem nicht die falsche Person trifft: Der Aufrufer ordnet
   * jeden Treffer über die E-Mail-Adresse einer Zeile DIESER Teilnehmerliste
   * zu. Eine Mail aus einem gleichnamigen Fremd-Event findet dort niemanden
   * und fällt weg. Getrennt gezählt wird trotzdem (`byId`/`byTitle`), damit
   * die Oberfläche sagen kann, worüber gefunden wurde.
   */
  const safeTitle = (eventTitle || '').replace(/'/g, "''");
  const scope = safeTitle
    ? `(EventId eq '${safeId}' or EventTitle eq '${safeTitle}')`
    : `EventId eq '${safeId}'`;
  /**
   * v31.4.3: `encodeURIComponent` um den GANZEN Filter — so macht es jede
   * andere gefilterte Abfrage im Projekt (`waitlist.ts:83`,
   * `registration.ts:588`, `eventsCrud.ts:632` …), und meine war die einzige,
   * die den Ausdruck roh in die URL geschrieben hat.
   *
   * Warum das genau hier weh tut: Der Event-Titel steht seit v31.4.2 im
   * Filter, und er heißt „B2Run Köln". Ein roher Umlaut in der Query-Zeichen-
   * kette wird nicht verlässlich als UTF-8 übertragen; SharePoint vergleicht
   * dann gegen einen anderen Text und findet still NICHTS. Kein Fehler, keine
   * Meldung — nur ein Zweig der ODER-Bedingung, der nie zutrifft. Genau das
   * Muster, das dieser Dialog sonst überall vermeidet.
   */
  const filter = `${scope} and EmailType eq 'QRCode'`;
  let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Emails')/items`
    + `?$select=Id,Recipient,Body,Status,EventId&$filter=${encodeURIComponent(filter)}`
    + `&$orderby=Id asc&$top=20`;
  while (url) {
    let resp;
    try {
      resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
    } catch {
      return { ...out, ok: false, status: 0 };
    }
    if (!resp.ok) return { ...out, ok: false, status: resp.status };
    let data;
    try { data = await resp.json(); } catch { return { ...out, ok: false, status: 0 }; }
    const items = data.value || data.d?.results || [];
    for (const it of items) {
      out.scanned++;
      if (String(it.EventId || '') === eventId) out.byId++; else out.byTitle++;
      const body = String(it.Body || '');
      const qrId = parseQrIdFromBody(body);
      if (qrId === null) { out.unparsed++; continue; }
      const recipient = firstAddress(String(it.Recipient || ''));
      const intended = parseIntendedRecipient(body);
      out.hits.push({
        mailId: Number(it.Id) || 0,
        email: intended || recipient,
        recipient,
        redirected: !!intended,
        qrId,
        status: String(it.Status || ''),
      });
    }
    url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
  }
  return out;
}
