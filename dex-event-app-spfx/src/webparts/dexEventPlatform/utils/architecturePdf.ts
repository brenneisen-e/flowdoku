// v26.28 — Architektur-PDF der DEX-Plattform (clientseitig per jsPDF).
//
// Erzeugt ein mehrseitiges A4-Dokument, das die gesamte Architektur in Worten
// beschreibt: Schichten (App, Daten, Automatisierung, Dienste), ALLE
// SharePoint-Listen mit Funktion, die Power-Automate-Flows (Trigger → Aktion),
// die genutzten Microsoft-365-Dienste und die typischen Datenflüsse.
//
// jsPDF wird dynamisch geladen (kein statischer Import im Boot-Pfad), analog
// zum Self-Check-in-PDF.

const GREEN = '#86bc25';
const GRAY_DARK = '#2b2b2b';
const GRAY_MED = '#666666';

export interface ArchPdfItem { name: string; desc: string }
export interface ArchPdfBlock { heading: string; note?: string; items: ArchPdfItem[] }

export async function downloadArchitecturePdf(opts: {
  isDe: boolean;
  title: string;
  intro: string;
  version?: string;
  blocks: ArchPdfBlock[];
}): Promise<void> {
  const { jsPDF } = await import('jspdf');
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const L = 16;            // linker Rand
  const R = pageW - 16;    // rechter Rand
  const contentW = R - L;
  const BOTTOM = pageH - 16;

  let y = 0;
  let pageNo = 0;

  const footer = (): void => {
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor(GRAY_MED);
    const left = opts.isDe ? 'DEX Event Experience Platform — Architektur' : 'DEX Event Experience Platform — Architecture';
    doc.text(left + (opts.version ? `  ·  v${opts.version}` : ''), L, pageH - 9);
    doc.text(`${opts.isDe ? 'Seite' : 'Page'} ${pageNo}`, R, pageH - 9, { align: 'right' });
    const note = opts.isDe
      ? 'Intern · Läuft vollständig im Microsoft-365-Tenant von Deloitte — keine externen Dienste.'
      : 'Internal · Runs entirely within Deloitte’s Microsoft 365 tenant — no external services.';
    doc.text(note, L, pageH - 5.5);
  };

  const newPage = (withBar: boolean): void => {
    if (pageNo > 0) { footer(); doc.addPage(); }
    pageNo++;
    if (withBar) {
      doc.setFillColor(GREEN);
      doc.rect(0, 0, pageW, 24, 'F');
      doc.setTextColor('#ffffff');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(17);
      doc.text(opts.title, L, 15);
      y = 32;
    } else {
      y = 18;
    }
  };

  // Stellt sicher, dass `h` mm Platz vorhanden sind — sonst neue Seite.
  const ensure = (h: number): void => { if (y + h > BOTTOM) newPage(false); };

  newPage(true);

  // Intro
  doc.setTextColor(GRAY_MED);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  const introLines = doc.splitTextToSize(opts.intro, contentW);
  doc.text(introLines, L, y);
  y += introLines.length * 5 + 4;

  for (const block of opts.blocks) {
    ensure(16);
    // Abschnitts-Überschrift mit grünem Balken links
    doc.setFillColor(GREEN);
    doc.rect(L, y - 3.6, 2.4, 5.2, 'F');
    doc.setTextColor(GRAY_DARK);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(12.5);
    doc.text(block.heading, L + 5, y);
    y += 6;

    if (block.note) {
      doc.setFont('helvetica', 'italic');
      doc.setFontSize(8.5);
      doc.setTextColor(GRAY_MED);
      const noteLines = doc.splitTextToSize(block.note, contentW);
      ensure(noteLines.length * 4 + 2);
      doc.text(noteLines, L, y);
      y += noteLines.length * 4 + 2;
    }

    for (const it of block.items) {
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9.5);
      const nameLines = doc.splitTextToSize(it.name, contentW);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      const descLines = doc.splitTextToSize(it.desc, contentW - 4);
      const blockH = nameLines.length * 4.4 + descLines.length * 4.2 + 3;
      ensure(blockH);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9.5);
      doc.setTextColor(GREEN);
      doc.text(nameLines, L, y);
      y += nameLines.length * 4.4;
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.setTextColor(GRAY_DARK);
      doc.text(descLines, L + 4, y);
      y += descLines.length * 4.2 + 3;
    }
    y += 4;
  }

  footer();

  const fname = opts.isDe ? 'DEX-Architektur.pdf' : 'DEX-architecture.pdf';
  doc.save(fname);
}
