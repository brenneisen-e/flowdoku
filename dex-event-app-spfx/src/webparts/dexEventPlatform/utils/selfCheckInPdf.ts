// v18.33 — Druckbares Self-Check-in-PDF (statischer QR + Anleitung)
//
// Erzeugt clientseitig per jsPDF ein A4-PDF mit dem event-spezifischen
// QR-Code und einer kurzen Schritt-für-Schritt-Anleitung. Der QR enthält den
// statischen Check-in-Link; gescannt mit der normalen Handy-Kamera öffnet er
// die App und checkt den eingeloggten Teilnehmer automatisch ein.

import { jsPDF } from 'jspdf';
import * as QRCode from 'qrcode';
import { buildStaticCheckInUrl } from './selfCheckIn';

const DELOITTE_GREEN = '#86bc25';
const GRAY_DARK = '#2b2b2b';
const GRAY_MED = '#666666';

/** Generiert das PDF und löst den Browser-Download aus. */
export async function downloadSelfCheckInPdf(opts: {
  eventTitle: string;
  eventDateLabel?: string;
  locationLabel?: string;
  token: string;
}): Promise<void> {
  const url = buildStaticCheckInUrl(opts.token);
  const qrDataUrl = await QRCode.toDataURL(url, { width: 900, margin: 1, errorCorrectionLevel: 'M' });

  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  const pageW = doc.internal.pageSize.getWidth();

  // Grüner Kopfbalken
  doc.setFillColor(DELOITTE_GREEN);
  doc.rect(0, 0, pageW, 26, 'F');
  doc.setTextColor('#ffffff');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(18);
  doc.text('Self-Check-in', 18, 17);

  // Event-Titel
  doc.setTextColor(GRAY_DARK);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(16);
  const titleLines = doc.splitTextToSize(opts.eventTitle || 'Event', pageW - 36);
  doc.text(titleLines, 18, 40);

  let cursorY = 40 + titleLines.length * 7;

  // Datum / Ort
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  doc.setTextColor(GRAY_MED);
  if (opts.eventDateLabel) {
    doc.text(opts.eventDateLabel, 18, cursorY + 2);
    cursorY += 7;
  }
  if (opts.locationLabel) {
    doc.text(opts.locationLabel, 18, cursorY + 2);
    cursorY += 7;
  }

  // QR-Code zentriert
  const qrSize = 90;
  const qrX = (pageW - qrSize) / 2;
  const qrY = cursorY + 8;
  doc.addImage(qrDataUrl, 'PNG', qrX, qrY, qrSize, qrSize);

  // Anleitung
  let y = qrY + qrSize + 14;
  doc.setTextColor(GRAY_DARK);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(13);
  doc.text('So checkst du dich ein:', 18, y);
  y += 9;

  const steps = [
    'Öffne die Kamera-App auf deinem Smartphone.',
    'Richte sie auf diesen QR-Code — es erscheint ein Link.',
    'Tippe auf den Link. Die DEX-App öffnet sich.',
    'Du wirst automatisch als anwesend eingecheckt — fertig!',
  ];
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  doc.setTextColor(GRAY_MED);
  steps.forEach((step, i) => {
    const lines = doc.splitTextToSize(`${i + 1}.  ${step}`, pageW - 40);
    doc.text(lines, 20, y);
    y += lines.length * 6 + 2;
  });

  // Hinweis-Fußzeile
  y += 4;
  doc.setDrawColor('#e0e0e0');
  doc.line(18, y, pageW - 18, y);
  y += 7;
  doc.setFontSize(9);
  doc.setTextColor(GRAY_MED);
  const note = doc.splitTextToSize(
    'Hinweis: Der Check-in funktioniert nur, wenn du für dieses Event angemeldet und am Veranstaltungstag vor Ort bist. ' +
    'Du checkst dich immer als du selbst ein — der Code lässt sich nicht für andere Personen verwenden.',
    pageW - 36
  );
  doc.text(note, 18, y);

  const safeTitle = (opts.eventTitle || 'Event').replace(/[^\w\s-]+/g, '').trim().replace(/\s+/g, '-').slice(0, 50);
  doc.save(`Self-Checkin-${safeTitle || 'Event'}.pdf`);
}
