// v30.66: Aus EventCreationPage.tsx ausgelagert, damit die ausgelagerten
// Wizard-Schritte (BasicsStep, CommunicationStep) dieselbe Kompression nutzen
// koennen, ohne aus der Seite zu importieren (das waere ein Modul-Zyklus).
/**
 * v28.31: `flattenToJpeg` erzwingt JPEG auf WEISSEM Grund. Hintergrund: Der
 * Zuschnitt-Dialog liefert PNG (für die transparenten Ecken des Kreis-
 * Zuschnitts), und ein 600px-Foto als PNG wiegt schnell ~800 KB. Als Mail-/
 * Outlook-Kopfbild steckt dieser Base64-String DREIMAL im selben Save-Payload
 * (OutlookBody, _eventLogo, _outlookLogo) — damit riss ein einziges Event die
 * SharePoint-Grenze von 2 MB, und „Speichern" tat nichts mehr. Der Mail-Kopf
 * steht ohnehin auf Weiss, transparente Ecken werden also korrekt weiß
 * gefüllt (ohne den weissen Grund wuerden sie bei JPEG schwarz).
 */
export async function compressImage(file: File, maxWidth: number = 1200, quality: number = 0.8, flattenToJpeg: boolean = false): Promise<File> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      let { width, height } = img;
      if (width > maxWidth) {
        height = Math.round((height * maxWidth) / width);
        width = maxWidth;
      }
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      if (!ctx) { resolve(file); return; }
      // v23.16: PNG-Quellen als PNG ausgeben — sonst gehen transparente Bereiche
      // (z.B. die per Kreis-Zuschnitt freigeschnittenen Ecken) bei der
      // JPEG-Konvertierung verloren und werden SCHWARZ gefüllt. Nur Nicht-PNG
      // (Fotos) werden zu JPEG komprimiert. v28.31: `flattenToJpeg` sticht das
      // — dann wird vorher weiß grundiert, die Ecken bleiben also sauber.
      const isPng = (file.type || '').toLowerCase() === 'image/png';
      const outType = (isPng && !flattenToJpeg) ? 'image/png' : 'image/jpeg';
      const outExt = outType === 'image/png' ? '.png' : '.jpg';
      if (outType === 'image/jpeg') {
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, width, height);
      }
      ctx.drawImage(img, 0, 0, width, height);
      canvas.toBlob(
        (blob) => {
          if (!blob || (blob.size >= file.size && !flattenToJpeg)) {
            // Komprimierung bringt nichts oder ist grösser → Original verwenden.
            // Bei flattenToJpeg NICHT zurückfallen: Ein PNG-Original mag kleiner
            // sein, taugt aber genau deshalb nicht — wir wollen hier zwingend
            // das JPEG, damit der Save-Payload klein bleibt.
            resolve(file);
            return;
          }
          const compressed = new File([blob], file.name.replace(/\.[^.]+$/, outExt), { type: outType });
          resolve(compressed);
        },
        outType,
        quality
      );
    };
    img.onerror = () => reject(new Error('Bild konnte nicht geladen werden'));
    img.src = URL.createObjectURL(file);
  });
}
