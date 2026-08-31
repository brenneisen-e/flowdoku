/**
 * v30.36: QR-Code mit dem Deloitte-D. in der Mitte — an EINER Stelle.
 *
 * Vorher gab es zwei Erzeuger: den Massen-Versand im Organizer Center und den
 * Auto-Versand für Nachzügler im EventContext. Die liefen auseinander (der
 * Auto-Versand hatte bis v30.35 eine niedrigere Fehlerkorrektur), und das
 * fällt niemandem auf — es scannt einfach schlechter. Deshalb hier gebündelt.
 *
 * Zwei bewusste Entscheidungen:
 *
 * 1. **Fehlerkorrektur 'H' (30 %) statt Default 'M' (15 %).** Der Code wird am
 *    Einlass vom Handy-DISPLAY abfotografiert. Spiegelungen, Fingerabdrücke,
 *    Displayschutz und schräge Winkel fressen genau die Reserve, die vorher
 *    fehlte. Die Datenmenge (`DEX|<nr>|<mail>`) ist klein genug, dass 'H' das
 *    Modul-Raster kaum vergrößert.
 *
 * 2. **Das Zeichen wird gezeichnet, nicht als Bild geladen.** Ein Canvas-Glyph
 *    plus Kreis ist bei jeder Größe scharf und braucht kein Asset, das mit
 *    ausgeliefert und gepflegt werden müsste. Fällt das Zeichnen aus, kommt
 *    der Code trotzdem — ein QR ohne Logo ist besser als gar keiner.
 */

/** Deloitte-Grün. Identisch zum `--dex-green` der Oberfläche. */
const DELOITTE_GREEN = '#86BC25';

/**
 * Zeichnet das Deloitte-D. mittig auf ein quadratisches Canvas.
 *
 * Die Kantenlänge ist bewusst konservativ: 24 % Kantenlänge sind rund 6 % der
 * Fläche. Das liegt weit unter dem, was 'H' verkraftet — die restliche Reserve
 * gehört der realen Abnutzung und nicht dem Logo. Wer hier größer geht,
 * verbessert die Optik und verschlechtert genau das, wofür 'H' da ist.
 */
function drawDMark(canvas: HTMLCanvasElement): void {
  const ctx = canvas.getContext('2d');
  if (!ctx) return;
  const box = Math.round(canvas.width * 0.24);
  const x0 = Math.round((canvas.width - box) / 2);
  const y0 = x0; // Canvas ist quadratisch
  const pad = Math.round(box * 0.16);

  // Weißes Feld darunter, damit das Zeichen nicht auf Modulen klebt.
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(x0 - pad, y0 - pad, box + pad * 2, box + pad * 2);

  // D und Punkt als Gruppe zentrieren — sonst sitzt das Zeichen optisch links,
  // weil der Punkt rechts Platz braucht.
  ctx.textBaseline = 'alphabetic';
  ctx.textAlign = 'left';
  ctx.font = `bold ${Math.round(box * 0.98)}px Arial, Helvetica, sans-serif`;
  const dWidth = ctx.measureText('D').width;
  const dotR = box * 0.11;
  const gap = box * 0.07;
  const groupW = dWidth + gap + dotR * 2;
  const startX = x0 + (box - groupW) / 2;
  const baseY = y0 + box * 0.9;

  ctx.fillStyle = '#000000';
  ctx.fillText('D', startX, baseY);

  ctx.fillStyle = DELOITTE_GREEN;
  ctx.beginPath();
  ctx.arc(startX + dWidth + gap + dotR, baseY - dotR, dotR, 0, Math.PI * 2);
  ctx.fill();
}

/**
 * Erzeugt den Teilnehmer-QR als data:-URL — mit hoher Fehlerkorrektur und dem
 * Deloitte-D. in der Mitte.
 *
 * `qrcode` wird dynamisch geladen: Beide Aufrufer liegen auf Pfaden, die die
 * Bibliothek sonst ins Boot-Bundle ziehen würden.
 *
 * Wirft nicht. Schlägt die Erzeugung fehl, kommt ein leerer String zurück —
 * die Aufrufer haben dafür ihren Text-Fallback.
 */
export async function buildParticipantQrDataUrl(qrData: string, width = 300): Promise<string> {
  try {
    const QRCode = await import('qrcode');
    const canvas = document.createElement('canvas');
    await QRCode.toCanvas(canvas, qrData, { width, margin: 2, errorCorrectionLevel: 'H' });
    try {
      drawDMark(canvas);
    } catch { /* Logo ist Kür — der Code zählt */ }
    return canvas.toDataURL('image/png');
  } catch {
    return '';
  }
}
