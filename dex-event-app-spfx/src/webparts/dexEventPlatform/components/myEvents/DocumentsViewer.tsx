/* DocumentsViewer — aus MyEventsPage.tsx ausgelagert (Zeilen 156-166 und
 * 552-707 des urspruenglichen Stands, v30.65). Zeigt die Event-Dokumente in
 * „Meine Events" an und klappt PDFs inline auf; `getDocIconName` gehoert nur
 * hierher und ist deshalb mitgewandert. Der Code ist zeichengleich uebernommen.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';

// v20.0 (Audit): PdfViewer zieht react-pdf (+pdfjs) ins Bundle — lazy laden,
// der Viewer wird nur beim Öffnen eines Dokuments gebraucht.
const PdfViewer = React.lazy(() => import('../PdfViewer'));

function getDocIconName(name: string): string {
  const ext = name.split('.').pop()?.toLowerCase() || '';
  switch (ext) {
    case 'pdf': return 'PDF';
    case 'doc': case 'docx': return 'WordDocument';
    case 'xls': case 'xlsx': return 'ExcelDocument';
    case 'ppt': case 'pptx': return 'PowerPointDocument';
    case 'jpg': case 'jpeg': case 'png': case 'gif': return 'FileImage';
    default: return 'Page';
  }
}

export default function DocumentsViewer({ documents, t }: { documents: Array<{name: string; url: string; size?: number}>; t: (key: string) => string }): React.ReactElement {
  const [expandedDoc, setExpandedDoc] = React.useState<string | null>(null);
  const [blobUrl, setBlobUrl] = React.useState<string>('');
  const [pdfBlob, setPdfBlob] = React.useState<Blob | null>(null);
  const [loading, setLoading] = React.useState(false);
  // Mobile-Erkennung: Auf Mobile nutzen wir react-pdf (Canvas), auf Desktop bleibt iframe (bewährt)
  const [isMobile, setIsMobile] = React.useState<boolean>(
    typeof window !== 'undefined' && window.matchMedia ? window.matchMedia('(max-width: 768px)').matches : false
  );
  React.useEffect(() => {
    if (typeof window === 'undefined' || !window.matchMedia) return;
    const mq = window.matchMedia('(max-width: 768px)');
    const handler = (e: MediaQueryListEvent): void => setIsMobile(e.matches);
    if (mq.addEventListener) mq.addEventListener('change', handler);
    else mq.addListener(handler);
    return () => {
      if (mq.removeEventListener) mq.removeEventListener('change', handler);
      else mq.removeListener(handler);
    };
  }, []);

  const toggleDoc = async (doc: { url: string; name: string }): Promise<void> => {
    if (expandedDoc === doc.url) {
      setExpandedDoc(null);
      if (blobUrl) { URL.revokeObjectURL(blobUrl); setBlobUrl(''); }
      setPdfBlob(null);
      return;
    }
    setExpandedDoc(doc.url);
    setLoading(true);
    setBlobUrl('');
    setPdfBlob(null);

    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (!ctx) { setLoading(false); return; }

      // Datei per SPHttpClient REST API als Binary laden
      const siteUrl = ctx.pageContext.web.absoluteUrl;
      const origin = doc.url.match(/^https?:\/\/[^/]+/)?.[0] || '';
      const serverRelPath = decodeURIComponent(doc.url.replace(origin, ''));

      // Pfad-Segmente einzeln encoden (Leerzeichen, Klammern etc.)
      const encodedPath = serverRelPath.split('/').map(s => encodeURIComponent(s)).join('/');
      const apiUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodedPath}')/$value`;

      // XHR für Binary-Download (zuverlässiger als fetch für SharePoint)
      const blob = await new Promise<Blob | null>((resolve) => {
        const xhr = new XMLHttpRequest();
        xhr.open('GET', apiUrl, true);
        xhr.responseType = 'blob';
        xhr.withCredentials = true;
        xhr.setRequestHeader('Accept', '*/*');
        xhr.onload = () => {
          if (xhr.status === 200 && xhr.response) {
            resolve(xhr.response as Blob);
          } else {
            console.warn('[DEX] Doc XHR failed:', xhr.status, apiUrl);
            resolve(null);
          }
        };
        xhr.onerror = () => { console.warn('[DEX] Doc XHR error'); resolve(null); };
        xhr.send();
      });

      if (blob && blob.size > 0) {
        const ext = doc.name.split('.').pop()?.toLowerCase() || '';
        const mimeMap: Record<string, string> = { pdf: 'application/pdf', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' };
        const correctBlob = (mimeMap[ext] && blob.type !== mimeMap[ext]) ? new Blob([blob], { type: mimeMap[ext] }) : blob;
        if (ext === 'pdf' && isMobile) {
          // Mobile: PDF via react-pdf (Canvas) - funktioniert wo iframe versagt
          setPdfBlob(correctBlob);
        } else {
          // Desktop oder Bilder: Blob-URL + iframe (bewährt)
          setBlobUrl(URL.createObjectURL(correctBlob));
        }
      }
    } catch (err) { console.warn('[DEX] Doc viewer error:', err); }
    setLoading(false);
  };

  // Cleanup blob URLs bei Unmount
  // v30.67: `[blobUrl]` statt `[]` — mit leerer Liste hielt die Cleanup den
  // Wert des ERSTEN Renders ('') und gab nie etwas frei; jedes geöffnete PDF
  // blieb bis zum Tab-Ende im Speicher. Mit dem Wert als Dependency läuft sie
  // beim Wechsel mit der ALTEN URL (die ab da niemand mehr rendert — toggleDoc
  // setzt vorher '') und beim Unmount mit der letzten. Ein doppeltes revoke
  // (Zuklappen räumt selbst auf) ist folgenlos.
  React.useEffect(() => {
    return () => { if (blobUrl) URL.revokeObjectURL(blobUrl); };
  }, [blobUrl]);

  return (
    <div style={{ marginTop: 12 }}>
      <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 6 }}>
        {t('myevents.documents')}
      </div>
      {documents.map((doc, i) => {
        const isExpanded = expandedDoc === doc.url;

        return (
          <div key={i} style={{ marginBottom: 6 }}>
            <div style={{
              display: 'flex', alignItems: 'center', gap: 10, padding: '10px 14px',
              background: isExpanded ? 'var(--dex-green-light, #f0fdf4)' : 'var(--dex-gray-100)',
              borderRadius: isExpanded ? '8px 8px 0 0' : 8,
              cursor: 'pointer', fontSize: '0.85rem', color: 'var(--dex-gray-700)',
              transition: 'background 0.15s',
            }} onClick={() => toggleDoc(doc)}>
              <span style={{ width: 32, height: 32, borderRadius: '50%', background: 'var(--dex-green-dark, #6b9a1e)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                <Icon iconName={getDocIconName(doc.name)} style={{ fontSize: 16, color: '#fff' }} />
              </span>
              <span style={{ flex: 1, fontWeight: isExpanded ? 600 : 400 }}>{doc.name}</span>
              {doc.size ? <span style={{ color: 'var(--dex-gray-400)', fontSize: '0.75rem' }}>{(doc.size / 1024).toFixed(0)} KB</span> : null}
              {doc.url && doc.url.startsWith('http') && (
                <a href={doc.url} target="_blank" rel="noopener noreferrer" onClick={e => e.stopPropagation()} style={{ color: 'var(--dex-green-dark)', fontSize: '0.72rem', textDecoration: 'none' }}>
                  <Icon iconName="Download" style={{ fontSize: 14 }} />
                </a>
              )}
              <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-400)' }}>{isExpanded ? '▲' : '▼'}</span>
            </div>
            {isExpanded && (
              <div style={{
                border: '1px solid var(--dex-gray-200)', borderTop: 'none',
                borderRadius: '0 0 8px 8px', overflow: 'hidden', background: '#fff',
              }}>
                {loading ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--dex-gray-400)' }}>
                    {t('myevents.agenda') === 'Programm' ? 'Vorschau wird geladen...' : 'Loading preview...'}
                  </div>
                ) : pdfBlob ? (
                  /* PDF via react-pdf (Canvas) - funktioniert Desktop + Mobile, eigenes Scrolling.
                     v20.0: lazy Chunk — Suspense zeigt kurz den Lade-Hinweis. */
                  <React.Suspense fallback={<div style={{ padding: 40, textAlign: 'center', color: 'var(--dex-gray-400)' }}>…</div>}>
                    <PdfViewer blob={pdfBlob} height={600} />
                  </React.Suspense>
                ) : blobUrl ? (
                  /* Desktop-PDF + Bilder via iframe.
                     #view=FitH zwingt das Browser-PDF-Plugin in vertikalen Scroll-Modus
                     (sonst wird oft "Fit page" angenommen und der Scrollbalken fehlt). */
                  <iframe
                    src={doc.name.toLowerCase().endsWith('.pdf') ? `${blobUrl}#view=FitH&toolbar=1` : blobUrl}
                    scrolling="auto"
                    style={{ width: '100%', height: '75vh', minHeight: 600, border: 'none', display: 'block' }}
                    title={doc.name}
                  />
                ) : (
                  <div style={{ padding: 24, textAlign: 'center' }}>
                    <a href={doc.url} target="_blank" rel="noopener noreferrer" style={{ color: 'var(--dex-green-dark)' }}>
                      {t('myevents.agenda') === 'Programm' ? 'Im Browser öffnen' : 'Open in browser'}
                    </a>
                  </div>
                )}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
