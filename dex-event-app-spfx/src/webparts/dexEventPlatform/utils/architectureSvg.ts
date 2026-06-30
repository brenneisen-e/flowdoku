// v26.29 — Schematische Architektur-Landkarte als SVG (C4-/Container-Stil).
//
// Bewusst nach modernen Doku-Best-Practices gebaut: wenige, gruppierte
// Bausteine statt 25 Einzelboxen, eine feste Farb-Legende, Alltagssprache
// statt Entwickler-Jargon. Dasselbe SVG wird on-screen gezeigt UND (als Bild)
// ins PDF exportiert — eine einzige Quelle.

export interface ArchLegendItem { name: string; desc: string }
export interface ArchLegendBlock { heading: string; items: ArchLegendItem[] }

const C = {
  actorBg: '#eef0f2', actorBd: '#b8c0c9', actorTx: '#3a4047',
  appBg: '#eef6dc', appBd: '#86bc25', appTx: '#3c6300',
  dataBg: '#eaf2fb', dataBd: '#3b7dd8', dataTx: '#1f5cad',
  autoBg: '#fdf1e0', autoBd: '#e8901f', autoTx: '#a85f06',
  extBg: '#f1ecf9', extBd: '#7a5fb0', extTx: '#4f3a80',
  line: '#7b828b', text: '#2b2b2b', sub: '#5b6168', zoneBd: '#cfd6de',
};

const escX = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// Greedy-Wortumbruch auf max. `maxChars` Zeichen pro Zeile.
function wrap(text: string, maxChars: number): string[] {
  const words = text.split(/\s+/);
  const lines: string[] = [];
  let cur = '';
  for (const w of words) {
    if (!cur) { cur = w; continue; }
    if ((cur + ' ' + w).length <= maxChars) cur += ' ' + w;
    else { lines.push(cur); cur = w; }
  }
  if (cur) lines.push(cur);
  return lines;
}

function tspans(lines: string[], x: number, lineH: number): string {
  return lines.map((l, i) => `<tspan x="${x}" dy="${i === 0 ? 0 : lineH}">${escX(l)}</tspan>`).join('');
}

function textBlock(s: string, x: number, y: number, opts: { size: number; color: string; weight?: number; anchor?: string; maxChars?: number; lineH?: number }): string {
  const lines = opts.maxChars ? wrap(s, opts.maxChars) : [s];
  const lineH = opts.lineH || opts.size * 1.25;
  return `<text x="${x}" y="${y}" font-size="${opts.size}" fill="${opts.color}" font-weight="${opts.weight || 400}" text-anchor="${opts.anchor || 'start'}" font-family="Helvetica, Arial, sans-serif">${tspans(lines, x, lineH)}</text>`;
}

function rrect(x: number, y: number, w: number, h: number, fill: string, stroke: string, rx = 10, sw = 1.4): string {
  return `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${rx}" ry="${rx}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}" />`;
}

// Pfeil zwischen zwei Punkten (gerade) mit optionalem Label.
function arrow(x1: number, y1: number, x2: number, y2: number, label: string, isDe: boolean, dashed = false): string {
  const ang = Math.atan2(y2 - y1, x2 - x1);
  const ah = 8;
  const hx1 = x2 - ah * Math.cos(ang - Math.PI / 6), hy1 = y2 - ah * Math.sin(ang - Math.PI / 6);
  const hx2 = x2 - ah * Math.cos(ang + Math.PI / 6), hy2 = y2 - ah * Math.sin(ang + Math.PI / 6);
  const mx = (x1 + x2) / 2, my = (y1 + y2) / 2;
  const lbl = label
    ? textBlock(label, mx + (Math.abs(x2 - x1) < 4 ? 10 : 0), my + (Math.abs(y2 - y1) < 4 ? -5 : -4), { size: 12.5, color: C.sub, anchor: Math.abs(x2 - x1) < 4 ? 'start' : 'middle' })
    : '';
  void isDe;
  return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${C.line}" stroke-width="1.6" ${dashed ? 'stroke-dasharray="5 4"' : ''} />`
    + `<polygon points="${x2},${y2} ${hx1},${hy1} ${hx2},${hy2}" fill="${C.line}" />` + lbl;
}

/** Baut die komplette Landkarte als SVG-String (+ intrinsische Maße). */
export function buildArchitectureSvg(isDe: boolean): { svg: string; width: number; height: number } {
  const W = 1160, H = 840;
  const t = (de: string, en: string): string => (isDe ? de : en);
  let s = '';

  // Hintergrund
  s += `<rect x="0" y="0" width="${W}" height="${H}" fill="#ffffff" />`;

  // Schicht 1 — Personen
  const a = { x: 330, y: 30, w: 500, h: 62 };
  s += rrect(a.x, a.y, a.w, a.h, C.actorBg, C.actorBd, 31);
  s += textBlock(t('Mitarbeitende · Organisator:innen · Admins', 'Employees · organizers · admins'), a.x + a.w / 2, a.y + 26, { size: 17, weight: 700, color: C.actorTx, anchor: 'middle' });
  s += textBlock(t('nutzen DEX im Browser, in Teams oder SharePoint', 'use DEX in the browser, Teams or SharePoint'), a.x + a.w / 2, a.y + 46, { size: 12.5, color: C.sub, anchor: 'middle' });

  s += arrow(W / 2, a.y + a.h, W / 2, a.y + a.h + 36, t('anmelden · Events anlegen · an-/abmelden · Fragen stellen', 'sign in · create events · register/cancel · ask questions'), isDe);

  // Schicht 2 — App
  const ap = { x: 300, y: 150, w: 560, h: 70 };
  s += rrect(ap.x, ap.y, ap.w, ap.h, C.appBg, C.appBd, 12, 1.8);
  s += textBlock('DEX Event Experience Platform', ap.x + ap.w / 2, ap.y + 28, { size: 17, weight: 700, color: C.appTx, anchor: 'middle' });
  s += textBlock(t('Die App — läuft direkt in SharePoint Online, Anmeldung über Microsoft 365', 'The app — runs inside SharePoint Online, sign-in via Microsoft 365'), ap.x + ap.w / 2, ap.y + 50, { size: 12.5, color: C.sub, anchor: 'middle', maxChars: 70 });

  s += arrow(W / 2, ap.y + ap.h, W / 2, ap.y + ap.h + 34, t('speichert & liest alle Daten', 'stores & reads all data'), isDe);

  // Schicht 3 — SharePoint (Datenspeicher) mit 4 Gruppen
  const z = { x: 40, y: 274, w: W - 80, h: 196 };
  s += rrect(z.x, z.y, z.w, z.h, '#f7faff', C.zoneBd, 12, 1.2);
  s += textBlock(t('SharePoint Online — der zentrale Datenspeicher', 'SharePoint Online — the central data store'), z.x + 16, z.y + 24, { size: 15, weight: 700, color: C.dataTx });

  const groups = [
    { title: t('Events & Rollen', 'Events & roles'), desc: t('Welche Events es gibt und wer was darf.', 'Which events exist and who may do what.'), foot: 'DEX_Events · DEX_Roles · DEX_EmailTemplates · DEX_Participants' },
    { title: t('Anmeldungen je Event', 'Registrations per event'), desc: t('Pro Event ein eigener Bereich mit der Teilnehmerliste — jede:r sieht nur die eigene Anmeldung.', 'A dedicated area per event with the attendee list — each person sees only their own entry.'), foot: t('Event-Bereiche (Subsites)', 'event subsites') },
    { title: t('Aufgaben-Warteschlangen', 'Task queues'), desc: t('Aufträge für die Automatik: Mails, Einladungen, Nummern, Freigaben.', 'Jobs for the automation: mails, invitations, numbers, approvals.'), foot: 'DEX_Emails · DEX_Outlook · DEX_IDReorder · DEX_Tickets …' },
    { title: t('Protokolle & Archiv', 'Logs & archive'), desc: t('Nachvollziehbarkeit (wer hat was getan) und automatisches Aufräumen.', 'Traceability (who did what) and automatic clean-up.'), foot: 'DEX_ChangeLog · DEX_Archive · DEX_WeeklyReports …' },
  ];
  const gGap = 16, gW = (z.w - 32 - gGap * 3) / 4, gY = z.y + 40, gH = 138;
  groups.forEach((g, i) => {
    const gx = z.x + 16 + i * (gW + gGap);
    s += rrect(gx, gY, gW, gH, C.dataBg, C.dataBd, 9, 1.3);
    s += `<rect x="${gx}" y="${gY}" width="${gW}" height="6" rx="3" fill="${C.dataBd}" />`;
    s += textBlock(g.title, gx + gW / 2, gY + 28, { size: 13.5, weight: 700, color: C.dataTx, anchor: 'middle', maxChars: 24 });
    s += textBlock(g.desc, gx + 12, gY + 50, { size: 11, color: C.text, maxChars: 34, lineH: 14 });
    s += textBlock(g.foot, gx + 12, gY + gH - 14, { size: 9.5, color: C.sub, maxChars: 38, lineH: 12 });
  });

  // Pfeil von Datenspeicher zur Automatik
  s += arrow(330, z.y + z.h, 330, z.y + z.h + 44, t('neue Aufgabe startet die Automatik', 'a new task starts the automation'), isDe);
  // Rückweg (Status)
  s += arrow(250, z.y + z.h + 44, 250, z.y + z.h, t('Status zurück', 'status back'), isDe, true);

  // Schicht 4 — Automatik (links) + Microsoft 365 (rechts)
  const au = { x: 40, y: z.y + z.h + 48, w: 640, h: 150 };
  s += rrect(au.x, au.y, au.w, au.h, C.autoBg, C.autoBd, 12, 1.6);
  s += textBlock(t('Power Automate — die Automatik im Hintergrund', 'Power Automate — the background automation'), au.x + 18, au.y + 28, { size: 15, weight: 700, color: C.autoTx });
  const bullets = [
    t('verschickt alle E-Mails (Bestätigung, Erinnerung, Berichte)', 'sends all emails (confirmation, reminder, reports)'),
    t('erstellt & aktualisiert Outlook-Termine, lädt Teilnehmer ein/aus', 'creates & updates Outlook events, invites/uninvites attendees'),
    t('vergibt Teilnehmer-Nummern und rückt von der Warteliste nach', 'assigns attendee numbers and promotes from the waitlist'),
    t('erteilt Freigaben/Berechtigungen für Anmeldungen', 'grants approvals/permissions for registrations'),
  ];
  bullets.forEach((b, i) => {
    const by = au.y + 52 + i * 22;
    s += `<circle cx="${au.x + 22}" cy="${by - 4}" r="2.6" fill="${C.autoBd}" />`;
    s += textBlock(b, au.x + 32, by, { size: 11.5, color: C.text, maxChars: 78 });
  });

  const ex = { x: 720, y: z.y + z.h + 48, w: W - 80 - 720, h: 150 };
  s += rrect(ex.x, ex.y, ex.w, ex.h, C.extBg, C.extBd, 12, 1.6);
  s += textBlock(t('Microsoft 365 (im Tenant)', 'Microsoft 365 (in tenant)'), ex.x + 18, ex.y + 28, { size: 15, weight: 700, color: C.extTx });
  s += textBlock('Outlook / Exchange', ex.x + 18, ex.y + 58, { size: 12.5, weight: 700, color: C.text });
  s += textBlock(t('Postfach für Mails & Kalender-Einladungen', 'mailbox for mails & calendar invitations'), ex.x + 18, ex.y + 76, { size: 11, color: C.sub, maxChars: 44 });
  s += textBlock('Microsoft Graph', ex.x + 18, ex.y + 104, { size: 12.5, weight: 700, color: C.text });
  s += textBlock(t('Profile, Gruppen & Standorte (Sichtbarkeit, Personensuche)', 'profiles, groups & locations (visibility, people search)'), ex.x + 18, ex.y + 122, { size: 11, color: C.sub, maxChars: 44 });

  s += arrow(au.x + au.w, au.y + au.h / 2, ex.x, au.y + au.h / 2, t('nutzt', 'uses'), isDe);

  // Legende (Farb-Schlüssel)
  const ly = au.y + au.h + 26;
  const legend = [
    { c: C.actorBd, l: t('Personen', 'People') },
    { c: C.appBd, l: t('App', 'App') },
    { c: C.dataBd, l: t('Daten (SharePoint)', 'Data (SharePoint)') },
    { c: C.autoBd, l: t('Automatik (Power Automate)', 'Automation (Power Automate)') },
    { c: C.extBd, l: t('Microsoft 365', 'Microsoft 365') },
  ];
  let lx = 40;
  legend.forEach(le => {
    s += `<rect x="${lx}" y="${ly - 10}" width="14" height="14" rx="3" fill="${le.c}" />`;
    s += textBlock(le.l, lx + 20, ly + 1, { size: 11.5, color: C.text });
    lx += 24 + le.l.length * 7 + 26;
  });
  s += textBlock(t('Alles innerhalb des Microsoft-365-Tenants von Deloitte — keine externen Dienste.', 'Everything within Deloitte’s Microsoft 365 tenant — no external services.'), W - 40, ly + 1, { size: 11, color: C.sub, anchor: 'end' });

  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" font-family="Helvetica, Arial, sans-serif">${s}</svg>`;
  return { svg, width: W, height: H };
}

/** Kompakte, alltagssprachliche Legende für die PDF-Detailseite. */
export function architectureLegend(isDe: boolean): ArchLegendBlock[] {
  const t = (de: string, en: string): string => (isDe ? de : en);
  return [
    {
      heading: t('Events & Rollen', 'Events & roles'),
      items: [
        { name: 'DEX_Events', desc: t('Alle Events und Sub-Events mit Datum, Ort, Sichtbarkeit, Kapazität und Kommunikations-Einstellungen.', 'All events and sub-events with date, location, visibility, capacity and communication settings.') },
        { name: 'DEX_Roles', desc: t('Wer welche Rolle hat: Nutzer, Organisator oder Admin.', 'Who has which role: user, organizer or admin.') },
        { name: 'DEX_EmailTemplates', desc: t('Anpassbare Mail-Vorlagen und zentrale Einstellungen (Logo, Standardbild).', 'Customizable mail templates and central settings (logo, default image).') },
        { name: 'DEX_Participants', desc: t('Übergreifendes Verzeichnis aller Teilnehmenden (für Statistik/Kennzahlen).', 'Cross-event directory of all attendees (for statistics/KPIs).') },
      ],
    },
    {
      heading: t('Anmeldungen je Event', 'Registrations per event'),
      items: [
        { name: t('Event-Bereiche (Subsites)', 'Event areas (subsites)'), desc: t('Jedes Event hat einen eigenen Bereich mit seiner Teilnehmerliste. Jede Person sieht nur die eigene Anmeldung.', 'Each event has its own area with its attendee list. Each person sees only their own registration.') },
      ],
    },
    {
      heading: t('Aufgaben-Warteschlangen (starten die Automatik)', 'Task queues (trigger the automation)'),
      items: [
        { name: 'DEX_Emails', desc: t('Alle zu versendenden Mails — die Automatik verschickt sie.', 'All mails to be sent — the automation sends them.') },
        { name: 'DEX_Outlook', desc: t('Kalender-Aufträge: ein-/ausladen, Termin aktualisieren/löschen.', 'Calendar jobs: add/remove attendees, update/delete the event.') },
        { name: 'DEX_IDReorder', desc: t('Teilnehmer-Nummern neu vergeben und von der Warteliste nachrücken.', 'Re-assign attendee numbers and promote from the waitlist.') },
        { name: 'DEX_Tickets', desc: t('Fragen & Antworten (Support) zu Events und zur App.', 'Questions & answers (support) for events and the app.') },
        { name: t('weitere (Freigaben & Zugriffe)', 'others (approvals & access)'), desc: t('DEX_OrganizerRequests, DEX_TeamJoinRequests, DEX_AccessFix, DEX_AssistantAccess — kleine Aufträge rund um Freigaben und Berechtigungen.', 'DEX_OrganizerRequests, DEX_TeamJoinRequests, DEX_AccessFix, DEX_AssistantAccess — small jobs around approvals and permissions.') },
      ],
    },
    {
      heading: t('Protokolle & Archiv', 'Logs & archive'),
      items: [
        { name: 'DEX_ChangeLog', desc: t('Wer hat wann was getan — die Nachvollziehbarkeit pro Event.', 'Who did what and when — the traceability per event.') },
        { name: 'DEX_Archive', desc: t('Alte Einträge wandern hierher, damit die Arbeitslisten schlank bleiben.', 'Old entries move here so the working lists stay lean.') },
        { name: t('weitere Protokolle', 'further logs'), desc: t('DEX_WeeklyReports, DEX_InactiveNotices, DEX_OrganizerArchived, DEX_TeilnehmerCounter.', 'DEX_WeeklyReports, DEX_InactiveNotices, DEX_OrganizerArchived, DEX_TeilnehmerCounter.') },
      ],
    },
    {
      heading: t('Automatik — Power Automate', 'Automation — Power Automate'),
      items: [
        { name: t('E-Mail-Versand', 'Mail sending'), desc: t('Verschickt Bestätigungen, Erinnerungen und Berichte über das Sammelpostfach.', 'Sends confirmations, reminders and reports via the shared mailbox.') },
        { name: t('Outlook-Termine', 'Outlook events'), desc: t('Erstellt und aktualisiert Termine und lädt Teilnehmende ein bzw. wieder aus.', 'Creates and updates events and invites/uninvites attendees.') },
        { name: t('Nummern & Warteliste', 'Numbers & waitlist'), desc: t('Vergibt fortlaufende Teilnehmer-Nummern und rückt von der Warteliste nach.', 'Assigns sequential attendee numbers and promotes from the waitlist.') },
        { name: t('Freigaben', 'Approvals'), desc: t('Setzt Berechtigungen für stellvertretende Anmeldungen und Co-Organisatoren.', 'Sets permissions for proxy registrations and co-organizers.') },
      ],
    },
    {
      heading: t('Microsoft-365-Dienste', 'Microsoft 365 services'),
      items: [
        { name: 'Outlook / Exchange', desc: t('Sammelpostfach für den Mail-Versand und die Kalender-Einladungen.', 'Shared mailbox for sending mails and calendar invitations.') },
        { name: 'Microsoft Graph', desc: t('Liefert Profile, Gruppen und Standorte — für Sichtbarkeit und Personensuche.', 'Provides profiles, groups and locations — for visibility and people search.') },
      ],
    },
  ];
}
