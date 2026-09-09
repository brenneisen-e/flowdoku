/* DocumentsViewer — aus MyEventsPage.tsx ausgelagert (Zeilen 156-166 und
 * 552-707 des urspruenglichen Stands, v30.65). Zeigt die Event-Dokumente in
 * „Meine Events" an und klappt PDFs inline auf; `getDocIconName` gehoert nur
 * hierher und ist deshalb mitgewandert. Der Code ist zeichengleich uebernommen.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
// v31.8: Optik über die gemeinsamen dex-ui-Klassen. Die Dokumentzeile war ein
// `<div onClick>` mit Text-Pfeil: kein Tastaturfokus, kein Hover-Versprechen,
// und der Pfeil war ein Zeichen statt eines Symbols. `ensureDexUiStyles()`
// ruft die Seiten-Komponente (`MyEventsPage`), Unterkomponenten nie.
import { cx } from '../dexUi';
import { AlertCircle, ChevronDown } from '../Icons';

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

export default function DocumentsViewer({ documents, t, isDe }: { documents: Array<{name: string; url: string; size?: number}>; t: (key: string) => string; isDe: boolean }): React.ReactElement {
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
    <div className="dex-ui-section">
      <div className="dex-ui-section-title">{t('myevents.documents')}</div>
      {documents.map((doc, i) => {
        const isExpanded = expandedDoc === doc.url;

        return (
          <div key={i} style={{ marginBottom: 6 }}>
            {/* v31.8: Die Zeile ist jetzt ein echter <button> (Tastaturfokus,
                aria-expanded); der Download-Link steht als eigene Aktion
                daneben, statt im Klickfeld zu liegen — ein <a> im <button>
                wäre ungültiges Markup. */}
            {/* v31.8: `dex-ui-row` hat im Ruhezustand weder Grund noch Rahmen —
                die Zeile war nur beim Überfahren als Zeile zu erkennen, also auf
                dem Handy gar nicht (Leitfaden 6b). `--framed` gibt ihr den
                dauerhaften Rahmen und lässt Hover und `is-active` durch, weil es
                kein `background` setzt. Lokal bleibt nur der Radius: aufgeklappt
                schließt die Zeile bündig mit dem Panel darunter ab. */}
            <div
              className={cx('dex-ui-row', 'dex-ui-row--framed', isExpanded && 'is-active')}
              style={isExpanded ? { borderRadius: '10px 10px 0 0' } : undefined}
            >
              <button
                type="button"
                className="dex-ui-rowbtn"
                onClick={() => toggleDoc(doc)}
                aria-expanded={isExpanded}
                style={{ display: 'flex', alignItems: 'center', gap: 10, flex: 1, minWidth: 0 }}
              >
                <span className="dex-ui-avatar" style={{ background: 'var(--dex-green-dark, #6b9a1e)' }}>
                  <Icon iconName={getDocIconName(doc.name)} style={{ fontSize: 16, color: '#fff' }} />
                </span>
                <span className="dex-ui-row-main">
                  {/* v31.8: `dex-ui-row-title` kürzt mit Ellipse. Bei Namen wie
                      &bdquo;Agenda_DTP_Basics_Training_Oktober_2026_final_v3.pdf&ldquo;
                      stand danach nirgends mehr, welches Dokument man öffnet —
                      der volle Name lebte nur im `title`, und was nur beim
                      Überfahren erscheint, gibt es auf dem Handy nicht
                      (Leitfaden 6b). `--wrap` lässt ihn umbrechen. */}
                  <span
                    className="dex-ui-row-title dex-ui-row-title--wrap"
                    style={{ display: 'block', fontWeight: isExpanded ? 700 : 600 }}
                  >{doc.name}</span>
                  {doc.size ? <span className="dex-ui-row-sub" style={{ display: 'block' }}>{(doc.size / 1024).toFixed(0)} KB</span> : null}
                </span>
                <span className={cx('dex-ui-disclosure-chevron', isExpanded && 'is-open')}>
                  <ChevronDown size={16} />
                </span>
              </button>
              {/* Die Abblendung von `dex-ui-row-actions` auf 0.7 hebt das
                  Stylesheet seit v31.8 unter `@media (hover: none)` auf — ohne
                  Hover bliebe der Download-Knopf sonst dauerhaft blass. */}
              {doc.url && doc.url.startsWith('http') && (
                <span className="dex-ui-row-actions">
                  <a
                    href={doc.url}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="dex-ui-iconbtn"
                    title={isDe ? 'Herunterladen' : 'Download'}
                    aria-label={isDe ? `${doc.name} herunterladen` : `Download ${doc.name}`}
                    onClick={e => e.stopPropagation()}
                    style={{ color: 'var(--dex-green-dark)', textDecoration: 'none' }}
                  >
                    <Icon iconName="Download" style={{ fontSize: 14 }} />
                  </a>
                </span>
              )}
            </div>
            {isExpanded && (
              <div style={{
                border: '1px solid var(--dex-gray-200)', borderTop: 'none',
                borderRadius: '0 0 10px 10px', overflow: 'hidden', background: '#fff',
              }}>
                {loading ? (
                  <div className="dex-ui-muted" style={{ padding: 40, textAlign: 'center' }}>
                    {isDe ? 'Vorschau wird geladen …' : 'Loading preview …'}
                  </div>
                ) : pdfBlob ? (
                  /* PDF via react-pdf (Canvas) - funktioniert Desktop + Mobile, eigenes Scrolling.
                     v20.0: lazy Chunk — Suspense zeigt kurz den Lade-Hinweis.
                     v31.8: derselbe Satz wie oben statt eines nackten „…" —
                     drei Punkte allein sagen niemandem, was gerade passiert. */
                  <React.Suspense fallback={<div className="dex-ui-muted" style={{ padding: 40, textAlign: 'center' }}>{isDe ? 'Vorschau wird geladen …' : 'Loading preview …'}</div>}>
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
                  /* v31.8: Dieser Zweig heißt „Datei nicht geladen" (kein
                     SPFx-Kontext, HTTP-Fehler, leere Antwort). Vorher stand
                     dort nur ein Link — der Grund war nirgends zu lesen. */
                  <div style={{ padding: 16 }}>
                    <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 12 }}>
                      <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                      <span>
                        {isDe
                          ? 'Die Vorschau konnte nicht geladen werden. Du kannst das Dokument direkt im Browser öffnen.'
                          : 'The preview could not be loaded. You can open the document in your browser instead.'}
                      </span>
                    </div>
                    <a href={doc.url} target="_blank" rel="noopener noreferrer" className="btn btn-secondary dex-ui-btn-sm" style={{ textDecoration: 'none' }}>
                      {isDe ? 'Im Browser öffnen' : 'Open in browser'}
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
