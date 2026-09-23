/**
 * GrantScannersHandler (v31.86)
 *
 * Deep-Link-Handler für `action=grantscanners&event=<Id>` aus der Admin-Mail
 * „Check-in-Team braucht Rechte". Die Mail entsteht, wenn ein Organizer
 * (kein Admin) im Assistenten ein Check-in-Team speichert und SharePoint die
 * Zuweisung auf mindestens einer Teilnehmerliste ablehnt oder eine Person
 * nicht auflösbar ist — typisch: Der Organizer wurde erst nachträglich
 * benannt und hat auf einem älteren Termin kein „Manage Permissions".
 *
 * Der Klick vergibt als Admin die Rechte DIREKT: Klammer plus alle Termine
 * des Events, Check-in-Team aus `_qrScanners` der Klammer-Zeile, Organizer
 * und Co-Organizer als „nie entziehen"-Liste. Ergebnis als Modal.
 *
 * Nur für Admins; ohne Admin-Rechte passiert nichts (analog grantaccess).
 */
import * as React from 'react';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import Modal from './Modal';
import { deepLinkParams } from '../utils/deepLink';
import { EventService } from '../services/EventService';
import { Check, X } from './Icons';

type Ergebnis = {
  title: string;
  sites: number;
  scanners: string[];
  granted: number;
  unresolved: string[];
  failed: Array<{ site: string; email: string; status: number }>;
  error?: string;
};

export default function GrantScannersHandler(): React.ReactElement | null {
  const { isAdmin, originalIsAdmin } = useRoles();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const adminLike = isAdmin || originalIsAdmin;

  const [running, setRunning] = React.useState(false);
  const [result, setResult] = React.useState<Ergebnis | null>(null);
  const [open, setOpen] = React.useState(false);
  const handledRef = React.useRef(false);

  React.useEffect(() => {
    if (!adminLike || handledRef.current) return;
    let eventId = '';
    try {
      const p = deepLinkParams();
      if (p.get('action') !== 'grantscanners') return;
      eventId = (p.get('event') || '').trim();
    } catch { return; }
    if (!eventId) return;
    handledRef.current = true;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    if (!ctx) return;
    const svc = new EventService(ctx);
    setRunning(true);
    setOpen(true);
    void (async () => {
      try {
        const rows = await svc.getEvents();
        const root = rows.find(r => String(r.Id) === eventId);
        if (!root) {
          setResult({ title: '', sites: 0, scanners: [], granted: 0, unresolved: [], failed: [], error: isDe ? 'Event nicht gefunden.' : 'Event not found.' });
          setRunning(false);
          return;
        }
        // Check-in-Team und Co-Organizer stehen im Piggyback-JSON der Klammer.
        let scanners: string[] = [];
        let coOrgs: string[] = [];
        try {
          const ov = JSON.parse(root.EmailTemplateOverrides || '{}');
          const qr = Array.isArray(ov._qrScanners) ? ov._qrScanners : [];
          scanners = qr.map((x: { email?: string }) => String(x?.email || '').trim()).filter(Boolean);
          const co = Array.isArray(ov._coOrganizers) ? ov._coOrganizers : [];
          coOrgs = co.map((x: { email?: string }) => String(x?.email || '').trim()).filter(Boolean);
        } catch { /* kein JSON → kein Team */ }
        const organizers = (root.OrganizerEmail || '').split(/[;,]/).map(s => s.trim()).filter(Boolean);
        const sites = Array.from(new Set(
          rows
            .filter(r => String(r.Id) === eventId || String(r.ParentEventId || '') === eventId)
            .map(r => (r.SubsiteUrl || '').trim())
            .filter(Boolean)));
        if (scanners.length === 0 || sites.length === 0) {
          setResult({
            title: root.Title || '', sites: sites.length, scanners, granted: 0, unresolved: [], failed: [],
            error: scanners.length === 0
              ? (isDe ? 'Am Event ist kein Check-in-Team hinterlegt — vielleicht wurde es inzwischen entfernt.' : 'No check-in team is stored on the event — it may have been removed since.')
              : (isDe ? 'Das Event hat noch keine Teilnehmerliste.' : 'The event has no participant list yet.'),
          });
          setRunning(false);
          return;
        }
        const r = await svc.ensureScannerListPermissions(sites, scanners, [], organizers.concat(coOrgs));
        setResult({
          title: root.Title || '', sites: sites.length, scanners, granted: r.granted, unresolved: r.unresolved,
          failed: r.failed.map(f => ({ site: f.site, email: f.email, status: f.status })),
        });
      } catch (e) {
        setResult({ title: '', sites: 0, scanners: [], granted: 0, unresolved: [], failed: [], error: e instanceof Error ? e.message : String(e) });
      }
      setRunning(false);
    })();
  }, [adminLike, isDe]);

  if (!open) return null;

  const siteName = (s: string): string => s.replace(/\/+$/, '').split('/').pop() || s;
  const ok = !!result && !result.error && result.failed.length === 0 && result.unresolved.length === 0;

  return (
    <Modal open={open} onClose={() => { if (!running) setOpen(false); }} maxWidth={560} padding={24} dismissable={!running} ariaLabel={isDe ? 'Check-in-Team berechtigen' : 'Grant check-in team access'}>
      <h2 style={{ marginTop: 0, marginBottom: 4, color: 'var(--dex-green-dark, #4a7c1f)' }}>
        {isDe ? 'Check-in-Team berechtigen' : 'Grant check-in team access'}
      </h2>
      {running || !result ? (
        <p style={{ fontSize: '0.9rem', color: 'var(--dex-gray-600)' }}>
          {isDe ? 'Rechte werden gesetzt …' : 'Granting rights …'}
        </p>
      ) : (
        <>
          {result.title && (
            <p style={{ marginTop: 0, fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
              {isDe ? 'Event: ' : 'Event: '}<strong>{result.title}</strong>
              {result.sites > 0 && (isDe ? ` · ${result.sites} Teilnehmerliste(n)` : ` · ${result.sites} participant list(s)`)}
            </p>
          )}
          {result.error ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', border: '1px solid var(--dex-gray-200)', borderRadius: 10, background: '#fff', color: 'var(--dex-red, #c00)', fontSize: '0.88rem' }}>
              <X size={14} /> {result.error}
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginTop: 8 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', border: '1px solid var(--dex-gray-200)', borderRadius: 10, background: '#fff', fontSize: '0.88rem' }}>
                <span style={{ color: ok ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-orange-dark, #b35a00)', display: 'inline-flex', flexShrink: 0 }}>{ok ? <Check size={14} /> : <X size={14} />}</span>
                <span style={{ flex: 1 }}>
                  {isDe
                    ? `${result.granted} Recht(e) für ${result.scanners.length} Person(en) gesetzt.`
                    : `${result.granted} grant(s) set for ${result.scanners.length} person(s).`}
                </span>
              </div>
              {result.unresolved.length > 0 && (
                <div style={{ padding: '10px 12px', border: '1px solid var(--dex-gray-200)', borderRadius: 10, background: '#fff', fontSize: '0.84rem', color: 'var(--dex-red, #c00)' }}>
                  {isDe ? 'Kein SharePoint-Konto gefunden: ' : 'No SharePoint account found: '}{result.unresolved.join(', ')}
                </div>
              )}
              {result.failed.length > 0 && (
                <div style={{ padding: '10px 12px', border: '1px solid var(--dex-gray-200)', borderRadius: 10, background: '#fff', fontSize: '0.84rem', color: 'var(--dex-red, #c00)' }}>
                  {isDe ? 'Abgelehnt: ' : 'Rejected: '}
                  {result.failed.slice(0, 6).map(f => `${siteName(f.site)} · ${f.email} · HTTP ${f.status}`).join(' — ')}
                  {result.failed.length > 6 ? ' …' : ''}
                </div>
              )}
            </div>
          )}
          <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 16 }}>
            <button type="button" className="btn btn-primary" onClick={() => setOpen(false)}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          </div>
        </>
      )}
    </Modal>
  );
}
