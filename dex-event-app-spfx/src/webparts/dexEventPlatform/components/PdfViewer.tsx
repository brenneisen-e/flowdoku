/**
 * PDF Viewer using react-pdf v6 (React 17 compatible).
 *
 * Warum eigene Komponente:
 * - Mobile Browser (iOS Safari, Android Chrome) rendern PDFs nicht in <iframe>
 * - react-pdf rendert PDF-Seiten per Canvas und funktioniert auf allen Plattformen
 * - Fake-Worker (pdf.worker.entry) läuft im Main-Thread, keine CSP-Probleme in SPFx
 */

import * as React from 'react';
import { Document, Page, pdfjs } from 'react-pdf';
// Fake Worker Loader - lädt pdfjs-Worker im Main-Thread (keine externen URLs, CSP-kompatibel)
// eslint-disable-next-line @typescript-eslint/no-var-requires
require('pdfjs-dist/build/pdf.worker.entry');
// Fallback: workerSrc leer lassen, damit pdfjs den fake worker (global) verwendet
if (!pdfjs.GlobalWorkerOptions.workerSrc) {
  pdfjs.GlobalWorkerOptions.workerSrc = '';
}

interface PdfViewerProps {
  blob: Blob;
  height?: number;
}

export default function PdfViewer(props: PdfViewerProps): React.ReactElement {
  const { blob, height = 500 } = props;
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [numPages, setNumPages] = React.useState<number>(0);
  const [error, setError] = React.useState<string>('');
  const [width, setWidth] = React.useState<number>(0);

  // Container-Breite ermitteln und bei Resize aktualisieren
  React.useEffect(() => {
    const update = (): void => {
      if (containerRef.current) {
        setWidth(containerRef.current.clientWidth);
      }
    };
    update();
    window.addEventListener('resize', update);
    return () => window.removeEventListener('resize', update);
  }, []);

  const onLoadSuccess = React.useCallback((pdf: { numPages: number }): void => {
    setNumPages(pdf.numPages);
    setError('');
  }, []);

  const onLoadError = React.useCallback((err: Error): void => {
    console.warn('[DEX] PDF load error:', err);
    setError(err.message || 'PDF konnte nicht geladen werden');
  }, []);

  return (
    <div
      ref={containerRef}
      style={{
        width: '100%',
        height,
        overflowY: 'scroll',
        overflowX: 'hidden',
        // touch-action: pan-y erlaubt vertikales Scrollen per Touch (sonst blockiert Canvas)
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        touchAction: 'pan-y' as any,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        WebkitOverflowScrolling: 'touch' as any,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        overscrollBehavior: 'contain' as any,
        background: '#525659',
        padding: 12,
      }}
    >
      {error ? (
        <div style={{ padding: 24, textAlign: 'center', color: '#fff' }}>
          <p style={{ margin: 0 }}>{error}</p>
        </div>
      ) : (
        <Document
          file={blob}
          onLoadSuccess={onLoadSuccess}
          onLoadError={onLoadError}
          loading={
            <div style={{ padding: 40, textAlign: 'center', color: '#fff' }}>
              Vorschau wird geladen...
            </div>
          }
          error={
            <div style={{ padding: 24, textAlign: 'center', color: '#fff' }}>
              PDF konnte nicht geladen werden
            </div>
          }
        >
          {Array.from({ length: numPages }, (_, i) => (
            <div
              key={i}
              style={{
                marginBottom: 12,
                boxShadow: '0 2px 8px rgba(0,0,0,0.4)',
                // touch-action auch auf Page-Wrapper, damit Canvas das Scrollen nicht blockiert
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                touchAction: 'pan-y' as any,
              }}
            >
              <Page
                pageNumber={i + 1}
                width={width > 24 ? width - 24 : undefined}
                renderTextLayer={false}
                renderAnnotationLayer={false}
              />
            </div>
          ))}
        </Document>
      )}
    </div>
  );
}
