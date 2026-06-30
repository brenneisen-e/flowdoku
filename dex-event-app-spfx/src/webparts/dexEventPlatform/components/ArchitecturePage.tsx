/**
 * v26.29: Architektur-Landkarte der DEX-Plattform (Admin-only).
 *
 * Zeigt eine saubere, schematische Skizze (C4-/Container-Stil, Alltagssprache):
 * Personen → App → SharePoint-Daten (gruppiert) → Automatik (Power Automate) →
 * Microsoft-365-Dienste, mit Verbindungspfeilen und Farb-Legende. Dieselbe
 * SVG-Grafik lässt sich als PDF herunterladen (Bild + kompakte Legende).
 *
 * Bewusst NICHT technisch: keine Feldnamen, keine Versionshinweise — gruppierte
 * Bausteine in verständlicher Sprache (siehe moderne Doku-Best-Practices).
 */
import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { FileText, Columns, Book } from './Icons';
import { buildArchitectureSvg, architectureLegend } from '../utils/architectureSvg';
import { downloadArchitecturePdf } from '../utils/architecturePdf';
import { APP_VERSION } from '../version';

export default function ArchitecturePage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { isAdmin, originalIsAdmin } = useRoles();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const adminLike = isAdmin || originalIsAdmin;

  React.useEffect(() => { if (!adminLike) navigate('start'); }, [adminLike, navigate]);
  if (!adminLike) return <div className="page-container" />;

  const { svg, width, height } = buildArchitectureSvg(isDe);
  const screenSvg = svg.replace('<svg ', '<svg style="width:100%;height:auto;display:block" ');

  const [pdfBusy, setPdfBusy] = React.useState(false);
  const downloadPdf = async (): Promise<void> => {
    if (pdfBusy) return;
    setPdfBusy(true);
    try {
      await downloadArchitecturePdf({
        isDe,
        version: APP_VERSION,
        title: isDe ? 'Architektur der DEX-Plattform' : 'DEX platform architecture',
        svg, svgWidth: width, svgHeight: height,
        legend: architectureLegend(isDe),
      });
    } catch (e) { console.warn('[DEX] Architektur-PDF fehlgeschlagen:', e); }
    finally { setPdfBusy(false); }
  };

  return (
    <div className="page-container">
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center', marginBottom: 14 }}>
        <button className="btn btn-secondary" onClick={() => navigate('admin-hub')}>
          {isDe ? '← Zurück zum Admin' : '← Back to admin'}
        </button>
        <button className="btn btn-primary" style={{ display: 'inline-flex', alignItems: 'center', gap: 6, marginLeft: 'auto' }} disabled={pdfBusy} onClick={() => { void downloadPdf(); }}>
          <FileText size={16} /> {pdfBusy ? (isDe ? 'PDF wird erstellt…' : 'Generating PDF…') : (isDe ? 'Als PDF herunterladen' : 'Download as PDF')}
        </button>
      </div>

      <h1 style={{ marginTop: 0 }}>{isDe ? 'Architektur der DEX-Plattform' : 'DEX platform architecture'}</h1>
      <p style={{ color: 'var(--dex-gray-600)', marginTop: 0, maxWidth: 860 }}>
        {isDe
          ? 'Eine schematische Landkarte der Plattform in Alltagssprache: wer sie nutzt, wo die Daten liegen, was die Automatik im Hintergrund erledigt und welche Microsoft-365-Dienste dabei genutzt werden. Alles läuft im Microsoft-365-Tenant von Deloitte — ohne externe Dienste.'
          : 'A schematic map of the platform in plain language: who uses it, where the data lives, what the background automation does, and which Microsoft 365 services are involved. Everything runs within Deloitte’s Microsoft 365 tenant — without external services.'}
      </p>

      <div
        style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 12, background: '#fff', boxShadow: '0 2px 10px rgba(0,0,0,0.06)', padding: 14, overflowX: 'auto' }}
      >
        <div style={{ minWidth: 720 }} dangerouslySetInnerHTML={{ __html: screenSvg }} />
      </div>

      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 18 }}>
        <button className="btn btn-secondary" style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={() => navigate('flowcharts')}>
          <Columns size={16} /> {isDe ? 'Detaillierte Prozessübersicht' : 'Detailed process overview'}
        </button>
        <button className="btn btn-secondary" style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={() => navigate('manual')}>
          <Book size={16} /> {isDe ? 'Handbuch' : 'Manual'}
        </button>
      </div>
    </div>
  );
}
