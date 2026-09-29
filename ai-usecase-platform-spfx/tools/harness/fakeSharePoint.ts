/* Ein SharePoint im Speicher — genau so viel, wie die App davon anfasst.
 *
 * Die App spricht SharePoint nur über `context.spHttpClient.get/post` an
 * (`services/SharePointService.ts`, `utils/displayName.ts`). Hier steht die
 * Gegenseite: Listen, Spalten, Zeilen, Anhänge, Rollenzuweisungen, Personen.
 * Was die App schreibt, kann sie danach wieder lesen — Anlegen, Ändern,
 * Löschen, Rolle vergeben, Rechte prüfen laufen also wirklich durch, nicht
 * nur die Anzeige.
 *
 * Bewusst NICHT nachgebaut (steht in README.md unter „Was der Harness NICHT
 * prüft"): echte Rechte, Drosselung (429), Paging, Graph.
 *
 * Grundregel: Was die Attrappe nicht kennt, antwortet sie mit 404 und trägt
 * es in `window.__harness.unhandled` ein. Ein stilles 200 würde eine neue
 * REST-Abfrage der App als „läuft" durchwinken — und genau das ist der Fehler,
 * den dieser Harness finden soll.
 */

import { START_USE_CASES } from '../../src/webparts/aiUseCasePlatform/data/startUseCases';
import { LIST } from '../../src/webparts/aiUseCasePlatform/constants';

export type Rolle = 'admin' | 'organizer' | 'user' | 'first';
export type Zustand = 'ok' | 'forbidden' | 'error' | 'empty' | 'leer' | 'roles403' | 'logerror' | 'fresh';

export interface Params {
  role: Rolle;
  lang: 'de' | 'en';
  mobile: boolean;
  state: Zustand;
  /** Künstliche Antwortzeit je Anfrage in ms — damit Ladezustände sichtbar werden. */
  delay: number;
  /** `variety` mischt Status/Bilder unter die Startdaten, `start` lässt sie unverändert. */
  data: 'variety' | 'start';
}

export function parseParams(search: string): Params {
  const q = new URLSearchParams(search);
  const pick = <T extends string>(k: string, erlaubt: T[], vorgabe: T): T => {
    const v = (q.get(k) || '') as T;
    return erlaubt.indexOf(v) >= 0 ? v : vorgabe;
  };
  return {
    role: pick<Rolle>('role', ['admin', 'organizer', 'user', 'first'], 'admin'),
    lang: pick<'de' | 'en'>('lang', ['de', 'en'], 'de'),
    mobile: q.get('mobile') === '1',
    state: pick<Zustand>('state', ['ok', 'forbidden', 'error', 'empty', 'leer', 'roles403', 'logerror', 'fresh'], 'ok'),
    delay: Math.max(0, Math.min(20000, parseInt(q.get('delay') || '25', 10) || 0)),
    data: pick<'variety' | 'start'>('data', ['variety', 'start'], 'variety'),
  };
}

/* ------------------------------------------------------------------ Modell */

const SITE = 'https://harness.local/sites/AIUC';
const API = SITE + '/_api/';
const SITE_REL = '/sites/AIUC';

/** SharePoint-Standard-IDs der Rechtestufen (dieselben wie in `SharePointService`). */
const RD = { read: 1073741826, contribute: 1073741827, design: 1073741828, full: 1073741829, edit: 1073741830 };

type Row = Record<string, any>;
interface Anhang { FileName: string; ServerRelativeUrl: string }
interface Prinzipal { Id: number; Email: string; LoginName: string; Title: string; PrincipalType: number }
interface Liste {
  key: 'roles' | 'useCases' | 'log';
  title: string;
  exists: boolean;
  fields: Set<string>;
  items: Row[];
  nextId: number;
  unique: boolean;
  /** Prinzipal-Id → Rechtestufen. */
  assign: Map<number, Set<number>>;
  anhaenge: Record<number, Anhang[]>;
}

const iso = (msVorJetzt: number): string => new Date(Date.now() - msVorJetzt).toISOString();
const TAG = 86400000;

const ICH: Prinzipal = {
  Id: 11, Email: 'eike.brenneisen@harness.local', Title: 'Brenneisen, Eike', PrincipalType: 1,
  LoginName: 'i:0#.f|membership|eike.brenneisen@harness.local',
};

/** Personen des Verzeichnisses (People-Picker, ensureuser, Rollenliste). */
const LEUTE: Array<{ Id: number; Title: string; Email: string; Job: string; Dept: string }> = [
  { Id: 12, Title: 'Beispiel, Max', Email: 'max.beispiel@deloitte.de', Job: 'Senior Manager', Dept: 'Technology & Transformation' },
  { Id: 13, Title: 'Muster, Erika', Email: 'erika.muster@deloitte.de', Job: 'Manager', Dept: 'Financial Services' },
  { Id: 14, Title: 'Testmann, Tim', Email: 'tim.testmann@deloitte.de', Job: 'Consultant', Dept: 'Financial Services' },
  { Id: 15, Title: 'Schulz, Sabine', Email: 'sabine.schulz@deloitte.de', Job: 'Senior Consultant', Dept: 'Risk Advisory' },
  { Id: 16, Title: 'Klein, Julia', Email: 'julia.klein@deloitte.de', Job: 'Manager', Dept: 'Technology & Transformation' },
  { Id: 17, Title: 'Neumann, Peter', Email: 'peter.neumann@deloitte.de', Job: 'Director', Dept: 'Financial Services' },
  { Id: 18, Title: 'Krause, Lena', Email: 'lena.krause@deloitte.de', Job: 'Consultant', Dept: 'Risk Advisory' },
  { Id: 19, Title: 'Wagner, Sophie', Email: 'sophie.wagner@deloitte.de', Job: 'Senior Consultant', Dept: 'Technology & Transformation' },
  { Id: 20, Title: 'Vogel, Jonas', Email: 'jonas.vogel@deloitte.de', Job: 'Analyst', Dept: 'Financial Services' },
];

/** Die Spalten, die `ensure*List` der App anlegt — Schreiben in eine andere Spalte ist HTTP 400. */
const SPALTEN_UC = ['Kurzbeschreibung', 'Beschreibung', 'Bereich', 'UcStatus', 'SalesRelevanz', 'Machbarkeit', 'DemoTauglichkeit',
  'AufrufArt', 'LinkSourceCode', 'LinkDeployment', 'LinkGuide', 'LinkWiki', 'LinkVideo', 'BildUrl', 'Reihenfolge',
  'BetreuerEmails', 'BetreuerNamen', 'Schlagworte'];
const SPALTEN_ROLLEN = ['UserName', 'Role', 'AssignedBy', 'AssignedDate'];
const SPALTEN_LOG = ['UseCaseId', 'Aktion', 'Detail', 'Wer'];

/* ------------------------------------------------------- Beispiel-Zeilen -- */

function bildSvg(titel: string, farbe: string): string {
  const k = titel.split(/[\s/-]+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase();
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360">`
    + `<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="${farbe}"/><stop offset="1" stop-color="#0b2e4f"/></linearGradient></defs>`
    + `<rect width="640" height="360" fill="url(#g)"/>`
    + `<g fill="none" stroke="rgba(255,255,255,.25)" stroke-width="2"><circle cx="500" cy="90" r="120"/><circle cx="520" cy="110" r="70"/><path d="M0 300 L160 240 L260 270 L400 190 L640 250"/></g>`
    + `<text x="40" y="320" font-family="Arial" font-size="64" font-weight="700" fill="#fff">${k}</text></svg>`;
  return 'data:image/svg+xml;utf8,' + encodeURIComponent(svg);
}

/** Abweichungen von den Startdaten, damit die Zustände sichtbar werden, die beim Umbau kaputtgehen. */
const VARIANTE: Record<number, { UcStatus?: string; AufrufArt?: string; LinkDeployment?: string }> = {
  2: { AufrufArt: 'eingebettet' },
  5: { UcStatus: 'InArbeit', LinkDeployment: '' },
  9: { UcStatus: 'Geplant', LinkDeployment: '' },
  12: { LinkDeployment: '' },
  14: { UcStatus: 'Archiviert' },
};
const BILD_IDS: Record<number, string> = { 1: '#1f7a8c', 4: '#4a7c1f', 7: '#7a3e8c', 10: '#b35a00', 16: '#0076a8' };

/** Ein Use Case in der Form, in der SharePoint die Zeile führt (Spaltennamen wie in `SpUseCaseRow`). */
function zeileAusStart(uc: any, id: number, variety: boolean): Row {
  const r = uc.ressourcen || {};
  const z: Row = {
    Id: id,
    Title: uc.titel,
    Kurzbeschreibung: uc.kurzbeschreibung || '',
    Beschreibung: uc.beschreibung || '',
    Bereich: uc.bereich || '',
    UcStatus: uc.status || 'Geplant',
    SalesRelevanz: uc.salesRelevanz === 'unbewertet' ? null : (uc.salesRelevanz || null),
    Machbarkeit: uc.machbarkeit === 'unbewertet' ? null : (uc.machbarkeit || null),
    DemoTauglichkeit: uc.demoTauglichkeit === 'unbewertet' ? null : (uc.demoTauglichkeit || null),
    AufrufArt: uc.aufrufArt || 'fenster',
    LinkSourceCode: r.sourceCode || '',
    LinkDeployment: r.deployment || '',
    LinkGuide: r.deploymentGuide || '',
    LinkWiki: r.wiki || '',
    LinkVideo: r.video || '',
    BildUrl: uc.bildUrl || '',
    Reihenfolge: typeof uc.reihenfolge === 'number' ? uc.reihenfolge : id * 10,
    BetreuerEmails: (uc.betreuerEmails || []).join('; '),
    BetreuerNamen: (uc.betreuerNamen || []).join('; '),
    Schlagworte: (uc.schlagworte || []).join('; '),
    Created: iso(45 * TAG),
    Modified: iso((id % 9 + 1) * TAG),
  };
  if (variety) {
    const v = VARIANTE[id];
    if (v) Object.keys(v).forEach(k => { z[k] = (v as any)[k]; });
    if (BILD_IDS[id]) z.BildUrl = bildSvg(uc.titel, BILD_IDS[id]);
    if (id % 4 === 1) { z.BetreuerEmails = 'erika.muster@deloitte.de'; z.BetreuerNamen = 'Muster, Erika'; }
  }
  return z;
}

/** Was die App vom Client-Objekt in Spalten macht — dieselbe Abbildung wie `toRow` im Dienst, damit Schreiben und Lesen zusammenpassen. */
export function baueListen(p: Params): { lists: Record<string, Liste>; prinzipale: Prinzipal[] } {
  const fresh = p.state === 'fresh';
  const neu = (key: Liste['key'], title: string, felder: string[]): Liste => ({
    key, title, exists: !fresh, fields: new Set(fresh ? [] : felder), items: [], nextId: 1, unique: !fresh,
    assign: new Map(), anhaenge: {},
  });
  const uc = neu('useCases', LIST.useCases, SPALTEN_UC);
  const rollen = neu('roles', LIST.roles, SPALTEN_ROLLEN);
  const log = neu('log', LIST.log, SPALTEN_LOG);

  const prinzipale: Prinzipal[] = [
    { Id: 3, Email: '', LoginName: 'c:0(.s|true', Title: 'AIUC Owners', PrincipalType: 8 },
    { Id: 4, Email: '', LoginName: 'c:0(.s|true', Title: 'AIUC Members', PrincipalType: 8 },
    { Id: 5, Email: '', LoginName: 'c:0(.s|true', Title: 'AIUC Visitors', PrincipalType: 8 },
    ICH,
    ...LEUTE.map(l => ({ Id: l.Id, Email: l.Email, Title: l.Title, PrincipalType: 1, LoginName: `i:0#.f|membership|${l.Email}` })),
  ];

  const setze = (l: Liste, pid: number, ...stufen: number[]): void => { l.assign.set(pid, new Set(stufen)); };
  const add = (l: Liste, row: Row): Row => { const z = { Id: l.nextId++, Created: iso(0), Modified: iso(0), ...row }; l.items.push(z); return z; };

  if (!fresh) {
    // Die Rollenliste. `Kurator` ist der gespeicherte Wert für „Use Case Organizer" (utils/rollen.ts).
    const rz: Array<[string, string, string, string, number]> = [
      ['max.beispiel@deloitte.de', 'Beispiel, Max', 'Admin', 'System (Erstinstallation)', 40],
      ['erika.muster@deloitte.de', 'Muster, Erika', 'Kurator', 'Beispiel, Max', 21],
      ['tim.testmann@deloitte.de', 'Testmann, Tim', 'Kurator', 'Beispiel, Max', 9],
      // Steht in der Rollenliste unter einer anderen Schreibweise als im Konto (Alias).
      ['j.klein@deloitte.de', 'Klein, Julia', 'Kurator', 'Beispiel, Max', 6],
      ['sabine.schulz@deloitte.de', 'Schulz, Sabine', 'User', 'Muster, Erika', 3],
      // Kein Konto mehr im Verzeichnis — ensureuser antwortet 404.
      ['ausgeschieden.person@deloitte.de', 'Ehemalig, Erwin', 'Kurator', 'Beispiel, Max', 120],
    ];
    if (p.role !== 'first') {
      rz.forEach(r => add(rollen, { Title: r[0], UserName: r[1], Role: r[2], AssignedBy: r[3], AssignedDate: iso(r[4] * TAG) }));
      if (p.role === 'admin') add(rollen, { Title: ICH.Email, UserName: ICH.Title, Role: 'Admin', AssignedBy: 'Beispiel, Max', AssignedDate: iso(2 * TAG) });
      if (p.role === 'organizer') add(rollen, { Title: ICH.Email, UserName: ICH.Title, Role: 'Kurator', AssignedBy: 'Beispiel, Max', AssignedDate: iso(2 * TAG) });
    }

    // Rechte: die Rollenliste hat eigene Rechte (Owners), die Use-Case-Liste und das Protokoll auch.
    [rollen, uc, log].forEach(l => setze(l, 3, RD.full));
    setze(rollen, 12, RD.full); setze(uc, 12, RD.edit); setze(log, 12, RD.contribute);
    setze(rollen, 13, RD.read); setze(uc, 13, RD.edit); setze(log, 13, RD.contribute);
    // Tim: nur Lesen auf der Rollenliste — Use Cases und Protokoll fehlen (Lücke für „Rechte prüfen").
    setze(rollen, 14, RD.read);
    // Sabine ist „User", trägt aber noch Edit auf den Use Cases (Überschuss).
    setze(uc, 15, RD.edit);
    // Julia hat alles — unter dem Konto julia.klein@, die Rollenliste kennt sie als j.klein@.
    setze(rollen, 16, RD.read); setze(uc, 16, RD.edit); setze(log, 16, RD.contribute);
    if (p.role === 'admin') { setze(rollen, 11, RD.full); setze(uc, 11, RD.edit); setze(log, 11, RD.contribute); }
    if (p.role === 'organizer') { setze(rollen, 11, RD.read); setze(uc, 11, RD.edit); setze(log, 11, RD.contribute); }

    // Use Cases und Protokoll.
    const leerStart = p.state === 'empty' || p.state === 'leer';
    if (!leerStart) {
      START_USE_CASES.forEach((s, i) => add(uc, zeileAusStart(s, i + 1, p.data === 'variety')));
      uc.nextId = START_USE_CASES.length + 1;
    }
    const merker = (): void => { add(log, { Title: 'erstbefuellung', UseCaseId: 0, Aktion: 'erstbefuellung', Detail: `${START_USE_CASES.length} Start-Use-Cases angelegt`, Wer: 'max.beispiel@deloitte.de', Created: iso(40 * TAG) }); };
    if (p.state === 'leer') merker();
    if (p.state !== 'empty' && p.state !== 'leer') {
      merker();
      const e = (tage: number, id: number, aktion: string, detail: string, wer: string): void => {
        add(log, { Title: aktion, UseCaseId: id, Aktion: aktion, Detail: detail, Wer: wer, Created: iso(tage * TAG) });
      };
      e(21, 1, 'geaendert', 'Bild', 'erika.muster@deloitte.de');
      e(20, 5, 'geaendert', 'Status: Geplant → InArbeit; Deployment-Link', 'erika.muster@deloitte.de');
      e(14, 20, 'angelegt', 'Testeintrag', 'tim.testmann@deloitte.de');
      e(14, 20, 'geloescht', 'Testeintrag', 'tim.testmann@deloitte.de');
      e(9, 2, 'geaendert', 'Aufruf: fenster → eingebettet', 'max.beispiel@deloitte.de');
      e(6, 9, 'geaendert', 'Status: Live → Geplant', 'erika.muster@deloitte.de');
      e(4, 21, 'loeschen-fehlgeschlagen', 'Demo-Zwischenstand', 'tim.testmann@deloitte.de');
      e(2, 14, 'geaendert', 'Status: Live → Archiviert', 'max.beispiel@deloitte.de');
      e(0.2, 12, 'geaendert', 'Deployment-Link entfernt', 'erika.muster@deloitte.de');
    }
  }
  return { lists: { roles: rollen, useCases: uc, log }, prinzipale };
}

/* --------------------------------------------------------------- Antworten - */

const FEHLER_403 = { 'odata.error': { code: '-2147024891, System.UnauthorizedAccessException', message: { lang: 'en-US', value: 'Access denied. You do not have permission to perform this action or access this resource.' } } };
const fehler = (msg: string, code = '-1, Microsoft.SharePoint.Client.InvalidClientQueryException'): any => ({ 'odata.error': { code, message: { lang: 'en-US', value: msg } } });

interface Ergebnis { status: number; body?: any; kind?: 'single' | 'list' }

/** Erzeugt die Attrappe eines `WebPartContext` und hängt den Zustand an `window.__harness`. */
export function baueKontext(p: Params): any {
  const { lists, prinzipale } = baueListen(p);
  let naechstePrinzipalId = 100;

  const harness = {
    params: p,
    lists,
    prinzipale,
    requests: [] as Array<{ method: string; url: string; status: number }>,
    unhandled: [] as string[],
  };
  (window as any).__harness = harness;

  const listeNachTitel = (titel: string): Liste | undefined =>
    (Object.keys(lists) as Array<keyof typeof lists>).map(k => lists[k]).filter(l => l.title === titel)[0];

  // Wer welche Schreibrechte hat — eine grobe Attrappe, keine SharePoint-Rechte.
  const wirksameRolle: Rolle = (p.state === 'fresh' || p.role === 'first') ? 'admin' : p.role;
  const darfSchreiben = (l: Liste): boolean => {
    if (l.key === 'roles') return wirksameRolle === 'admin';
    return wirksameRolle === 'admin' || wirksameRolle === 'organizer';
  };
  const darfStruktur = wirksameRolle === 'admin';

  const typName = (l: Liste): string => `SP.Data.${l.title.replace(/_/g, '_x005f_')}ListItem`;

  const findePrinzipal = (logon: string): Prinzipal | null => {
    let mail = (logon || '').trim();
    const m = mail.match(/[^|]+@[^|\s]+$/);
    if (m) mail = m[0];
    mail = mail.toLowerCase();
    if (!mail || mail.indexOf('@') < 0 || /ausgeschieden|unbekannt/.test(mail)) return null;
    const vorhanden = prinzipale.filter(x => x.Email.toLowerCase() === mail)[0];
    if (vorhanden) return vorhanden;
    // Alias: j.klein@ gehört zum Konto julia.klein@.
    if (mail === 'j.klein@deloitte.de') return prinzipale.filter(x => x.Id === 16)[0];
    const neuer: Prinzipal = {
      Id: naechstePrinzipalId++, Email: mail, LoginName: `i:0#.f|membership|${mail}`, PrincipalType: 1,
      Title: mail.split('@')[0].split('.').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' '),
    };
    prinzipale.push(neuer);
    return neuer;
  };

  const zuweisungenJson = (l: Liste): Row[] => {
    const out: Row[] = [];
    l.assign.forEach((stufen, pid) => {
      const pr = prinzipale.filter(x => x.Id === pid)[0];
      if (!pr) return;
      out.push({
        PrincipalId: pid,
        Member: { Email: pr.Email, LoginName: pr.LoginName, PrincipalType: pr.PrincipalType },
        RoleDefinitionBindings: Array.from(stufen).map(id => ({ Id: id })),
      });
    });
    return out;
  };

  /** Die Spalten, die in einem Schreib-Body stehen, müssen auf der Liste existieren. */
  const pruefeSpalten = (l: Liste, body: Row): Ergebnis | null => {
    const typ = body.__metadata && body.__metadata.type;
    if (typ && typ !== typName(l)) {
      return { status: 400, body: fehler(`The property '__metadata.type' does not match: "${typ}" is not a valid entity type for ${l.title}. Expected ${typName(l)}.`) };
    }
    const fehlt = Object.keys(body).filter(k => k !== '__metadata' && k !== 'Title' && !l.fields.has(k));
    if (fehlt.length > 0) {
      return { status: 400, body: fehler(`The field or property '${fehlt[0]}' does not exist.`, '-1, System.ArgumentException') };
    }
    return null;
  };

  const itemsNachAnfrage = (l: Liste, q: URLSearchParams): Ergebnis => {
    let zeilen = l.items.slice();
    const filter = q.get('$filter');
    if (filter) {
      const m = filter.match(/^(\w+) eq '(.*)'$/);
      if (m) {
        if (m[1] !== 'Title' && m[1] !== 'Id' && !l.fields.has(m[1])) {
          return { status: 400, body: fehler(`Column '${m[1]}' does not exist. It may have been deleted by another user.`, '-1, System.ArgumentException') };
        }
        zeilen = zeilen.filter(z => String(z[m[1]]) === m[2]);
      }
    }
    const ord = q.get('$orderby');
    if (ord && /^Id desc$/i.test(ord)) zeilen.sort((a, b) => b.Id - a.Id);
    const top = parseInt(q.get('$top') || '100', 10);
    zeilen = zeilen.slice(0, top);
    const sel = q.get('$select');
    if (sel) {
      const felder = sel.split(',').map(s => s.trim());
      zeilen = zeilen.map(z => { const o: Row = {}; felder.forEach(f => { o[f] = z[f]; }); return o; });
    }
    return { status: 200, body: zeilen, kind: 'list' };
  };

  async function verarbeite(method: string, urlRoh: string, headers: Record<string, string>, body: any): Promise<Ergebnis> {
    if (urlRoh.indexOf(API) !== 0) {
      harness.unhandled.push(`${method} ${urlRoh} (außerhalb von /_api/)`);
      return { status: 404, body: fehler('Harness: Adresse außerhalb der Site.') };
    }
    const rest = urlRoh.slice(API.length);
    const qi = rest.indexOf('?');
    const pfad = decodeURIComponent(qi < 0 ? rest : rest.slice(0, qi));
    const q = new URLSearchParams(qi < 0 ? '' : rest.slice(qi + 1));
    let json: Row = {};
    if (typeof body === 'string' && body) { try { json = JSON.parse(body); } catch { json = {}; } }
    const m = (re: RegExp): RegExpMatchArray | null => pfad.match(re);
    let t: RegExpMatchArray | null;

    // ---- Person, Verzeichnis -------------------------------------------------
    if (method === 'GET' && /^SP\.UserProfiles\.PeopleManager\/GetMyProperties$/i.test(pfad)) {
      return { status: 200, kind: 'single', body: {
        DisplayName: ICH.Title,
        UserProfileProperties: [
          { Key: 'PreferredName', Value: ICH.Title }, { Key: 'FirstName', Value: 'Eike' }, { Key: 'LastName', Value: 'Brenneisen' },
        ],
      } };
    }
    if (method === 'GET' && pfad === 'web/currentuser') {
      return { status: 200, kind: 'single', body: { Id: ICH.Id, Title: ICH.Title, Email: ICH.Email, LoginName: ICH.LoginName } };
    }
    if (method === 'GET' && pfad === 'web/associatedownergroup') {
      return { status: 200, kind: 'single', body: { Id: 3, Title: 'AIUC Owners' } };
    }
    if (method === 'POST' && pfad === 'web/ensureuser') {
      const pr = findePrinzipal(String(json.logonName || ''));
      if (!pr) return { status: 500, body: fehler('User cannot be found.', '-2146232832, Microsoft.SharePoint.SPException') };
      return { status: 200, kind: 'single', body: { Id: pr.Id, Title: pr.Title, Email: pr.Email, LoginName: pr.LoginName } };
    }
    if (method === 'GET' && (t = m(/^web\/siteusers\/getbyemail\('(.+)'\)$/i))) {
      const pr = prinzipale.filter(x => x.Email.toLowerCase() === t![1].toLowerCase())[0];
      return pr ? { status: 200, kind: 'single', body: { Id: pr.Id, Title: pr.Title, Email: pr.Email, LoginName: pr.LoginName } }
        : { status: 404, body: fehler('User cannot be found.', '-2146232832, Microsoft.SharePoint.SPException') };
    }
    if (method === 'POST' && /^SP\.UI\.ApplicationPages\.ClientPeoplePickerWebServiceInterface\.clientPeoplePickerSearchUser$/i.test(pfad)) {
      const qs = String((json.queryParams && json.queryParams.QueryString) || '').toLowerCase();
      const tokens = qs.split(/[\s,]+/).filter(Boolean);
      const treffer = LEUTE.filter(l => {
        const heu = (l.Title + ' ' + l.Email).toLowerCase();
        return tokens.length > 0 && tokens.every(tk => heu.indexOf(tk) >= 0);
      });
      const eintraege = treffer.map(l => ({
        Key: `i:0#.f|membership|${l.Email}`, DisplayText: l.Title, Description: l.Email,
        EntityData: { Email: l.Email, Title: l.Job, Department: l.Dept },
      }));
      return { status: 200, kind: 'single', body: { ClientPeoplePickerSearchUser: JSON.stringify(eintraege) } };
    }

    // ---- Listen ----------------------------------------------------------------
    if (method === 'POST' && pfad === 'web/lists') {
      if (!darfStruktur) return { status: 403, body: FEHLER_403 };
      const l = listeNachTitel(String(json.Title || ''));
      if (!l) return { status: 400, body: fehler(`Harness: unbekannte Liste ${json.Title}`) };
      l.exists = true; l.unique = false;
      return { status: 201, kind: 'single', body: { Id: 'harness', Title: l.title } };
    }

    const lm = pfad.match(/^web\/lists\/getbytitle\('(.+?)'\)(?:\/(.*))?$/i);
    if (lm) {
      const l = listeNachTitel(lm[1]);
      const unter = lm[2] || '';
      if (!l || !l.exists) {
        return { status: 404, body: fehler(`List '${lm[1]}' does not exist at site with URL '${SITE}'.`, '-1, System.ArgumentException') };
      }

      if (unter === '') {
        if (method !== 'GET') { harness.unhandled.push(`${method} ${pfad}`); return { status: 404, body: fehler('Harness: nicht behandelt.') }; }
        const sel = q.get('$select') || '';
        const o: Row = { Title: l.title, Id: l.key };
        if (/ListItemEntityTypeFullName/.test(sel) || !sel) o.ListItemEntityTypeFullName = typName(l);
        if (/HasUniqueRoleAssignments/.test(sel) || !sel) o.HasUniqueRoleAssignments = l.unique;
        return { status: 200, kind: 'single', body: o };
      }

      // Spalten
      if (method === 'GET' && (t = unter.match(/^fields\/getbytitle\('(.+?)'\)$/i))) {
        return l.fields.has(t[1]) ? { status: 200, kind: 'single', body: { InternalName: t[1], Title: t[1] } }
          : { status: 404, body: fehler(`Column '${t[1]}' does not exist.`, '-1, System.ArgumentException') };
      }
      if (method === 'POST' && unter === 'fields') {
        if (!darfStruktur) return { status: 403, body: FEHLER_403 };
        const name = String(json.Title || '');
        if (!name || !json.__metadata || !/^SP\.Field/.test(json.__metadata.type || '')) {
          return { status: 400, body: fehler('Harness: Feldtyp fehlt oder unbekannt (__metadata.type).') };
        }
        l.fields.add(name);
        return { status: 201, kind: 'single', body: { Title: name, InternalName: name } };
      }
      if (method === 'POST' && /^DefaultView\/viewfields\/addviewfield\(/i.test(unter)) {
        return { status: 200, kind: 'single', body: {} };
      }

      // Zeilen
      if (unter === 'items') {
        if (method === 'GET') {
          if (l.key === 'useCases' && p.state === 'forbidden') return { status: 403, body: FEHLER_403 };
          if (l.key === 'useCases' && p.state === 'error') return { status: 500, body: fehler('Harness: simulierter Serverfehler.', '-2146232832, Microsoft.SharePoint.SPException') };
          if (l.key === 'roles' && p.state === 'roles403') return { status: 403, body: FEHLER_403 };
          if (l.key === 'log' && p.state === 'logerror') return { status: 500, body: fehler('Harness: simulierter Serverfehler.', '-2146232832, Microsoft.SharePoint.SPException') };
          return itemsNachAnfrage(l, q);
        }
        if (method === 'POST') {
          if (l.key === 'useCases' && p.state === 'forbidden') return { status: 403, body: FEHLER_403 };
          if (l.key === 'useCases' && p.state === 'error') return { status: 500, body: fehler('Harness: simulierter Serverfehler.') };
          if (!darfSchreiben(l)) return { status: 403, body: FEHLER_403 };
          const bad = pruefeSpalten(l, json);
          if (bad) return bad;
          const z: Row = { Id: l.nextId++, Created: new Date().toISOString(), Modified: new Date().toISOString() };
          Object.keys(json).forEach(k => { if (k !== '__metadata') z[k] = json[k]; });
          l.items.push(z);
          return { status: 201, kind: 'single', body: z };
        }
      }
      if ((t = unter.match(/^items\((\d+)\)$/i))) {
        const id = parseInt(t[1], 10);
        const z = l.items.filter(x => x.Id === id)[0];
        if (!z) return { status: 404, body: fehler('Item does not exist. It may have been deleted by another user.', '-2147024809, System.ArgumentException') };
        const via = (headers['x-http-method'] || method).toUpperCase();
        if (via === 'MERGE' || via === 'PATCH') {
          if (!darfSchreiben(l)) return { status: 403, body: FEHLER_403 };
          const bad = pruefeSpalten(l, json);
          if (bad) return bad;
          Object.keys(json).forEach(k => { if (k !== '__metadata') z[k] = json[k]; });
          z.Modified = new Date().toISOString();
          return { status: 204 };
        }
        if (via === 'DELETE') {
          if (!darfSchreiben(l)) return { status: 403, body: FEHLER_403 };
          l.items = l.items.filter(x => x.Id !== id);
          return { status: 200 };
        }
        if (via === 'GET') return { status: 200, kind: 'single', body: z };
      }
      if (method === 'POST' && (t = unter.match(/^items\((\d+)\)\/recycle$/i))) {
        if (!darfSchreiben(l)) return { status: 403, body: FEHLER_403 };
        const id = parseInt(t[1], 10);
        if (!l.items.some(x => x.Id === id)) return { status: 404, body: fehler('Item does not exist.') };
        l.items = l.items.filter(x => x.Id !== id);
        return { status: 200, kind: 'single', body: { Recycle: 'harness-papierkorb' } };
      }

      // Anhänge
      if (method === 'GET' && (t = unter.match(/^items\((\d+)\)\/AttachmentFiles$/i))) {
        return { status: 200, kind: 'list', body: l.anhaenge[parseInt(t[1], 10)] || [] };
      }
      if (method === 'POST' && (t = unter.match(/^items\((\d+)\)\/AttachmentFiles\/add\(FileName='(.+?)'\)$/i))) {
        if (!darfSchreiben(l)) return { status: 403, body: FEHLER_403 };
        const id = parseInt(t[1], 10);
        const name = decodeURIComponent(t[2]);
        const rel = `${SITE_REL}/Lists/${l.title}/Attachments/${id}/${name}`;
        (l.anhaenge[id] = l.anhaenge[id] || []).push({ FileName: name, ServerRelativeUrl: rel });
        // Beim Betrachten über `serve.js` liegt das Bild danach auch wirklich unter dieser Adresse.
        try {
          if (body && typeof body.arrayBuffer === 'function') {
            await fetch('/__harness/anhang?pfad=' + encodeURIComponent(rel), { method: 'PUT', body, headers: { 'Content-Type': body.type || 'application/octet-stream' } });
          }
        } catch { /* ohne serve.js gibt es keinen Ablageort — das Bild bleibt dann unsichtbar */ }
        return { status: 200, kind: 'single', body: { FileName: name, ServerRelativeUrl: rel } };
      }
      if (method === 'POST' && (t = unter.match(/^items\((\d+)\)\/AttachmentFiles\/getByFileName\('(.+?)'\)\/recycleObject$/i))) {
        if (!darfSchreiben(l)) return { status: 403, body: FEHLER_403 };
        const id = parseInt(t[1], 10);
        const name = decodeURIComponent(t[2]);
        l.anhaenge[id] = (l.anhaenge[id] || []).filter(a => a.FileName !== name);
        return { status: 200, kind: 'single', body: {} };
      }

      // Rechte
      if (method === 'POST' && /^breakroleinheritance\(/i.test(unter)) {
        if (!darfStruktur) return { status: 403, body: FEHLER_403 };
        const kopie = /copyRoleAssignments=true/i.test(unter);
        l.unique = true;
        if (!kopie) l.assign = new Map();
        else { l.assign.set(3, new Set([RD.full])); l.assign.set(4, new Set([RD.contribute])); l.assign.set(5, new Set([RD.read])); }
        return { status: 200, kind: 'single', body: {} };
      }
      if (method === 'GET' && unter === 'roleassignments') {
        return { status: 200, kind: 'list', body: zuweisungenJson(l) };
      }
      if (method === 'POST' && (t = unter.match(/^roleassignments\/addroleassignment\(principalid=(\d+),\s*roledefid=(\d+)\)$/i))) {
        if (!darfStruktur) return { status: 403, body: FEHLER_403 };
        if (!l.unique) return { status: 500, body: fehler('Harness: Die Liste erbt ihre Rechte von der Site — erst breakroleinheritance.', '-2146232832, Microsoft.SharePoint.SPException') };
        const pid = parseInt(t[1], 10);
        const set = l.assign.get(pid) || new Set<number>();
        set.add(parseInt(t[2], 10));
        l.assign.set(pid, set);
        return { status: 200, kind: 'single', body: {} };
      }
      if (method === 'POST' && (t = unter.match(/^roleassignments\/removeroleassignment\(principalid=(\d+),\s*roledefid=(\d+)\)$/i))) {
        if (!darfStruktur) return { status: 403, body: FEHLER_403 };
        const pid = parseInt(t[1], 10);
        const set = l.assign.get(pid);
        if (!set || !set.has(parseInt(t[2], 10))) return { status: 500, body: fehler('Harness: Zuweisung nicht vorhanden.', '-2146232832, Microsoft.SharePoint.SPException') };
        set.delete(parseInt(t[2], 10));
        if (set.size === 0) l.assign.delete(pid);
        return { status: 200, kind: 'single', body: {} };
      }
      if ((t = unter.match(/^roleassignments\/getbyprincipalid\((\d+)\)$/i)) && (headers['x-http-method'] || '').toUpperCase() === 'DELETE') {
        if (!darfStruktur) return { status: 403, body: FEHLER_403 };
        const pid = parseInt(t[1], 10);
        if (!l.assign.has(pid)) return { status: 404, body: fehler('Harness: Zuweisung nicht vorhanden.', '-2146232832, Microsoft.SharePoint.SPException') };
        l.assign.delete(pid);
        return { status: 200 };
      }
      if (method === 'GET' && (t = unter.match(/^roleassignments\/getbyprincipalid\((\d+)\)\/roledefinitionbindings$/i))) {
        const set = l.assign.get(parseInt(t[1], 10));
        if (!set) return { status: 404, body: fehler('Harness: Zuweisung nicht vorhanden.', '-2146232832, Microsoft.SharePoint.SPException') };
        return { status: 200, kind: 'list', body: Array.from(set).map(id => ({ Id: id })) };
      }
    }

    harness.unhandled.push(`${method} ${pfad}${qi < 0 ? '' : '?' + rest.slice(qi + 1)}`);
    return { status: 404, body: fehler(`Harness: ${method} ${pfad} ist in der Attrappe nicht vorgesehen.`) };
  }

  /** Aus dem Ergebnis eine echte `Response` machen — `ok`, `status`, `headers`, `json()`, `text()`, `clone()` gibt es dann von selbst. */
  function alsResponse(e: Ergebnis, verbose: boolean): Response {
    if (e.body === undefined || e.status === 204) return new Response(null, { status: e.status === 204 ? 204 : e.status });
    let out: any = e.body;
    if (e.status < 400) {
      if (e.kind === 'list') out = verbose ? { d: { results: e.body } } : { value: e.body };
      else if (e.kind === 'single') out = verbose ? { d: e.body } : e.body;
    }
    return new Response(JSON.stringify(out), { status: e.status, headers: { 'Content-Type': verbose ? 'application/json;odata=verbose' : 'application/json;odata=nometadata' } });
  }

  async function anfrage(method: string, url: string, options?: any): Promise<Response> {
    const headers: Record<string, string> = {};
    const h = (options && options.headers) || {};
    Object.keys(h).forEach(k => { headers[k.toLowerCase()] = String(h[k]); });
    const via = (headers['x-http-method'] || '').toUpperCase();
    const echt = via ? via : method;
    if (p.delay > 0) await new Promise(res => setTimeout(res, p.delay));
    let e: Ergebnis;
    try {
      e = await verarbeite(method, url, headers, options && options.body);
    } catch (err) {
      e = { status: 500, body: fehler('Harness-Fehler: ' + String(err)) };
    }
    harness.requests.push({ method: echt, url: url.replace(API, ''), status: e.status });
    return alsResponse(e, /odata=verbose/i.test(headers['accept'] || ''));
  }

  return {
    pageContext: {
      user: { displayName: ICH.Title, email: ICH.Email, loginName: ICH.LoginName },
      web: { absoluteUrl: SITE, serverRelativeUrl: SITE_REL, title: 'AI Use Case Platform' },
      site: { absoluteUrl: SITE },
      legacyPageContext: { userId: ICH.Id },
      cultureInfo: { currentCultureName: 'de-DE', currentUICultureName: 'de-DE' },
    },
    spHttpClient: {
      get: (url: string, _cfg: unknown, options?: any) => anfrage('GET', url, options),
      post: (url: string, _cfg: unknown, options?: any) => anfrage('POST', url, options),
      fetch: (url: string, _cfg: unknown, options?: any) => anfrage(((options && options.method) || 'GET').toUpperCase(), url, options),
    },
  };
}
