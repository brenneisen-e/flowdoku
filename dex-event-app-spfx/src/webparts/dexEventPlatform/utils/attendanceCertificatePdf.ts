/**
 * v30.96: Teilnahmebescheinigung (PDF) — je Person eine Seite mit dem Event,
 * dem Programm und den Punkten, an denen die Anwesenheit erfasst wurde
 * (Spalte `AgendaCheckIns`, utils/agendaCheckIns). Konzept:
 * docs/konzept-programmpunkte.md, Stufe 5d.
 *
 * Datenquelle ist ausschließlich die erfasste Anwesenheit — nichts wird
 * „angenommen": Ein Punkt ohne Erfassung steht als nicht besucht drin. Die
 * Bescheinigung nennt deshalb auch, WIE erfasst wurde (Check-in per QR-Code
 * bzw. durch das Check-in-Team), damit sie als Beleg taugt.
 *
 * jsPDF wird dynamisch geladen (kein statischer Import im Boot-Pfad), wie bei
 * Architektur- und Self-Check-in-PDF. Standard-Schrift Helvetica deckt
 * ä/ö/ü/ß ab; Häkchen werden gezeichnet statt als Zeichen gesetzt, weil die
 * Standard-Fonts kein ✓ kennen.
 */
import { AgendaItem } from '../types';
import { AgendaCheckIns, formatMarkTime } from './agendaCheckIns';
import { agendaGroups, groupLabel, groupDateLabel } from './agendaGroups';

export interface CertificateEvent {
  title: string;
  startDate?: string;
  endDate?: string;
  location?: string;
  organizers?: string[];
  agenda: AgendaItem[];
  agendaTermSingular?: string;
  agendaTermPlural?: string;
}

export interface CertificatePerson {
  name: string;
  email?: string;
  marks: AgendaCheckIns;
}

const GREEN = '#86bc25';
const GREEN_DARK = '#4a7c1f';
const GRAY_DARK = '#2b2b2b';
const GRAY_MED = '#666666';
const GRAY_LIGHT = '#bbbbbb';

function fmtDate(iso: string | undefined, isDe: boolean): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function periodOf(ev: CertificateEvent, isDe: boolean): string {
  const a = fmtDate(ev.startDate, isDe);
  const b = fmtDate(ev.endDate, isDe);
  return a && b && a !== b ? `${a} – ${b}` : (a || b);
}

function safeFile(s: string): string {
  return (s || '').replace(/[^\w\-äöüÄÖÜß ]+/g, '').trim().replace(/\s+/g, '_').slice(0, 60) || 'Bescheinigung';
}

/** Eine Seite je Person auf ein bestehendes Dokument zeichnen. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function drawPage(doc: any, ev: CertificateEvent, p: CertificatePerson, isDe: boolean, firstPage: boolean): void {
  if (!firstPage) doc.addPage();
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const L = 20; const R = pageW - 20; const W = R - L;
  const BOTTOM = pageH - 22;
  let y = 0;

  const header = (): void => {
    doc.setFillColor(GREEN);
    doc.rect(0, 0, pageW, 22, 'F');
    doc.setTextColor('#ffffff');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text(isDe ? 'Teilnahmebescheinigung' : 'Certificate of attendance', L, 14);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.text('DEX Event Experience Platform', R, 14, { align: 'right' });
    y = 34;
  };
  const footer = (): void => {
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor(GRAY_MED);
    const created = new Date().toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    doc.text(isDe
      ? `Erstellt am ${created} aus der DEX App. Die Anwesenheit wurde je ${termS} per Check-in erfasst (QR-Code der Person oder Check-in-Team).`
      : `Created ${created} from the DEX app. Attendance was recorded per ${termS.toLowerCase()} via check-in (the person's QR code or the check-in team).`, L, pageH - 12, { maxWidth: W });
  };
  const ensure = (h: number): void => {
    if (y + h > BOTTOM) { footer(); doc.addPage(); header(); }
  };

  const termS = (ev.agendaTermSingular || '').trim() || (isDe ? 'Programmpunkt' : 'Agenda item');
  const termP = (ev.agendaTermPlural || '').trim() || (isDe ? 'Programmpunkte' : 'Agenda items');
  const items = (ev.agenda || []).filter(a => a && (a.title || a.time));
  const visited = items.filter(a => !!p.marks[a.id]);

  header();

  // Person und Event
  doc.setTextColor(GRAY_DARK);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  const intro = isDe
    ? `Hiermit wird bestätigt, dass`
    : `This is to confirm that`;
  doc.text(intro, L, y); y += 9;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(18);
  doc.text(p.name || (p.email || '—'), L, y); y += 6;
  if (p.email) {
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.setTextColor(GRAY_MED);
    doc.text(p.email, L, y); y += 4;
  }
  y += 5;
  doc.setTextColor(GRAY_DARK);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  const period = periodOf(ev, isDe);
  const where = (ev.location || '').trim();
  const line2 = isDe
    ? `an der Veranstaltung`
    : `attended the event`;
  doc.text(line2, L, y); y += 8;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(14);
  doc.setTextColor(GREEN_DARK);
  const titleLines = doc.splitTextToSize(ev.title || '', W);
  doc.text(titleLines, L, y); y += titleLines.length * 6.2;
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10.5);
  doc.setTextColor(GRAY_DARK);
  const meta = [period, where].filter(Boolean).join(' · ');
  if (meta) { doc.text(meta, L, y); y += 6; }
  y += 2;
  const summary = isDe
    ? `teilgenommen hat und bei ${visited.length} von ${items.length} ${termP} als anwesend erfasst wurde.`
    : `and was recorded as present at ${visited.length} of ${items.length} ${termP.toLowerCase()}.`;
  const sumLines = doc.splitTextToSize(summary, W);
  doc.text(sumLines, L, y); y += sumLines.length * 5.2 + 6;

  // Programm mit Häkchen, nach Clustern
  doc.setFillColor(GREEN);
  doc.rect(L, y - 3.4, 2.2, 5, 'F');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.setTextColor(GRAY_DARK);
  doc.text(termP, L + 5, y); y += 7;

  const groups = agendaGroups(items);
  const colTime = L + 8; const colTitle = L + 34; const colMark = R - 26;
  groups.forEach((g, gi) => {
    ensure(9);
    if (groups.length > 1 || g.cluster) {
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(10);
      doc.setTextColor(GREEN_DARK);
      const gl = `${groupLabel(g, gi, isDe)}${g.dates.length ? ` · ${groupDateLabel(g, isDe)}` : ''}`;
      doc.text(gl, L, y); y += 5.5;
    }
    g.items.forEach(it => {
      const mark = p.marks[it.id];
      const title = `${it.title || (isDe ? '(ohne Titel)' : '(untitled)')}${it.location ? ` · ${it.location}` : ''}`;
      const titleLines2 = doc.splitTextToSize(title, colMark - colTitle - 4);
      const rowH = Math.max(5.2, titleLines2.length * 4.6) + 1.4;
      ensure(rowH);
      // Häkchen: gefüllter grüner Kreis mit weißem Haken; nicht erfasst: grauer Ring.
      const cx = L + 3; const cy = y - 1.6;
      if (mark) {
        doc.setFillColor(GREEN);
        doc.circle(cx, cy, 2.1, 'F');
        doc.setDrawColor('#ffffff');
        doc.setLineWidth(0.6);
        doc.line(cx - 1.1, cy, cx - 0.3, cy + 0.9);
        doc.line(cx - 0.3, cy + 0.9, cx + 1.2, cy - 1);
      } else {
        doc.setDrawColor(GRAY_LIGHT);
        doc.setLineWidth(0.4);
        doc.circle(cx, cy, 2.1, 'S');
      }
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9.5);
      doc.setTextColor(GRAY_MED);
      doc.text(`${it.time || ''}${it.endTime ? `–${it.endTime}` : ''}`, colTime, y);
      doc.setTextColor(mark ? GRAY_DARK : GRAY_MED);
      doc.setFont('helvetica', mark ? 'bold' : 'normal');
      doc.text(titleLines2, colTitle, y);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8.5);
      doc.setTextColor(mark ? GREEN_DARK : GRAY_LIGHT);
      doc.text(mark ? `${isDe ? 'anwesend' : 'present'} ${formatMarkTime(mark.at)}` : (isDe ? 'nicht erfasst' : 'not recorded'), R, y, { align: 'right' });
      y += rowH;
    });
    y += 2;
  });

  // Organizer
  const orgs = (ev.organizers || []).map(o => (o || '').trim()).filter(Boolean);
  if (orgs.length) {
    ensure(14);
    y += 4;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9.5);
    doc.setTextColor(GRAY_MED);
    doc.text(isDe ? 'Organisiert von' : 'Organised by', L, y); y += 5;
    doc.setTextColor(GRAY_DARK);
    doc.setFontSize(10.5);
    const orgLines = doc.splitTextToSize(orgs.join(', '), W);
    doc.text(orgLines, L, y); y += orgLines.length * 5;
  }
  footer();
}

/** Bescheinigung für EINE Person laden. */
export async function downloadAttendanceCertificate(ev: CertificateEvent, person: CertificatePerson, isDe: boolean): Promise<void> {
  const { jsPDF } = await import('jspdf');
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  drawPage(doc, ev, person, isDe, true);
  doc.save(`${isDe ? 'Teilnahmebescheinigung' : 'Certificate'}_${safeFile(ev.title)}_${safeFile(person.name)}.pdf`);
}

/** Bescheinigungen für MEHRERE Personen in einer Datei (eine Seite je Person). */
export async function downloadAttendanceCertificates(ev: CertificateEvent, persons: CertificatePerson[], isDe: boolean): Promise<number> {
  if (persons.length === 0) return 0;
  const { jsPDF } = await import('jspdf');
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  persons.forEach((p, i) => drawPage(doc, ev, p, isDe, i === 0));
  doc.save(`${isDe ? 'Teilnahmebescheinigungen' : 'Certificates'}_${safeFile(ev.title)}_${new Date().toISOString().slice(0, 10)}.pdf`);
  return persons.length;
}
