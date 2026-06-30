// v26.29 — Architektur-PDF: Seite 1 ist die schematische Landkarte (als Bild
// aus dem SVG gerendert → keine Schrift-/Encoding-Probleme), Seite 2+ eine
// kompakte, zweispaltige Legende in Alltagssprache.

import { ArchLegendBlock } from './architectureSvg';

const GREEN = '#86bc25';
const GREEN_DK = '#4a7c1f';
const GRAY_DARK = '#2b2b2b';
const GRAY_MED = '#666666';

// jsPDF-Helvetica kennt nur CP1252 — exotische Zeichen (→) auf ASCII abbilden.
const ascii = (s: string): string => (s || '')
  .replace(/→/g, '->').replace(/←/g, '<-').replace(/↔/g, '<->')
  .replace(/⇒/g, '=>').replace(/[↑↓]/g, '|');

function svgToPng(svg: string, w: number, h: number, scale: number): Promise<string> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      try {
        const canvas = document.createElement('canvas');
        canvas.width = Math.round(w * scale);
        canvas.height = Math.round(h * scale);
        const ctx = canvas.getContext('2d');
        if (!ctx) { reject(new Error('no 2d context')); return; }
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL('image/png'));
      } catch (e) { reject(e as Error); }
    };
    img.onerror = () => reject(new Error('svg image load failed'));
    img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
  });
}

export async function downloadArchitecturePdf(opts: {
  isDe: boolean;
  title: string;
  version?: string;
  svg: string;
  svgWidth: number;
  svgHeight: number;
  legend: ArchLegendBlock[];
}): Promise<void> {
  const { jsPDF } = await import('jspdf');
  const png = await svgToPng(opts.svg, opts.svgWidth, opts.svgHeight, 2.5);

  const doc = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'landscape' });
  const W = doc.internal.pageSize.getWidth();
  const H = doc.internal.pageSize.getHeight();
  const M = 10;
  let pageNo = 0;

  const titleBar = (txt: string, big: boolean): void => {
    doc.setFillColor(GREEN); doc.rect(0, 0, W, big ? 15 : 12, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(big ? 15 : 12); doc.setTextColor('#ffffff');
    doc.text(ascii(txt), M, big ? 10.5 : 8.4);
  };
  const footer = (): void => {
    doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(GRAY_MED);
    doc.text(ascii((opts.isDe ? 'DEX Event Experience Platform — Architektur' : 'DEX Event Experience Platform — Architecture')) + (opts.version ? `  ·  v${opts.version}` : ''), M, H - 6);
    doc.text(`${opts.isDe ? 'Seite' : 'Page'} ${pageNo}`, W - M, H - 6, { align: 'right' });
  };

  // ---- Seite 1: Landkarte als Bild ----
  pageNo++;
  titleBar(opts.title, true);
  const availW = W - 2 * M;
  const availH = H - 20 - 10;            // unter Titelbalken, über Footer
  const ratio = opts.svgWidth / opts.svgHeight;
  let drawW = availW, drawH = availW / ratio;
  if (drawH > availH) { drawH = availH; drawW = availH * ratio; }
  const ix = (W - drawW) / 2;
  const iy = 18 + (availH - drawH) / 2;
  doc.addImage(png, 'PNG', ix, iy, drawW, drawH, undefined, 'FAST');
  footer();

  // ---- Seite 2+: kompakte Legende (zweispaltig) ----
  doc.addPage('a4', 'landscape'); pageNo++;
  titleBar(opts.isDe ? 'Legende — was bedeutet was (in Kürze)' : 'Legend — what means what (in brief)', false);

  const colGap = 12;
  const colW = (W - 2 * M - colGap) / 2;
  const colX = [M, M + colW + colGap];
  const topY = 20, bottomY = H - 10;
  let col = 0, y = topY;

  const nextCol = (need: number): void => {
    if (y + need <= bottomY) return;
    if (col === 0) { col = 1; y = topY; }
    else { footer(); doc.addPage('a4', 'landscape'); pageNo++; titleBar(opts.isDe ? 'Legende (Fortsetzung)' : 'Legend (continued)', false); col = 0; y = topY; }
  };

  for (const blk of opts.legend) {
    doc.setFont('helvetica', 'bold'); doc.setFontSize(10);
    const hl = doc.splitTextToSize(ascii(blk.heading), colW) as string[];
    nextCol(hl.length * 5 + 3);
    doc.setFillColor(GREEN); doc.rect(colX[col], y - 3.4, 2, hl.length * 5, 'F');
    doc.setTextColor(GREEN_DK); doc.text(hl, colX[col] + 4, y);
    y += hl.length * 5 + 1.6;

    for (const it of blk.items) {
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8.4);
      const nl = doc.splitTextToSize(ascii(it.name), colW) as string[];
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8);
      const dl = doc.splitTextToSize(ascii(it.desc), colW - 2) as string[];
      nextCol(nl.length * 3.8 + dl.length * 3.6 + 3);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8.4); doc.setTextColor(GREEN_DK);
      doc.text(nl, colX[col], y); y += nl.length * 3.8;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(GRAY_DARK);
      doc.text(dl, colX[col] + 2, y); y += dl.length * 3.6 + 2.8;
    }
    y += 3;
  }
  footer();
  doc.save(opts.isDe ? 'DEX-Architektur.pdf' : 'DEX-architecture.pdf');
}
