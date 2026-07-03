/**
 * GrantAccessHandler (v26.59)
 *
 * Deep-Link-Handler für `action=grantaccess&emails=a;b` aus der
 * „SharePoint-Zugriff benötigt"-Admin-Mail (internationale Zielgruppen-
 * Personen außerhalb @deloitte.de). Statt den Admin auf die rohe
 * SharePoint-Berechtigungsseite zu schicken, vergibt der Klick die
 * Leserechte DIREKT (Besucher-Gruppe der Site) und zeigt das Ergebnis
 * als Modal in der App — pro Person: vergeben / war schon berechtigt /
 * fehlgeschlagen.
 *
 * Nur für Admins; ohne Admin-Rechte passiert nichts (analog approveorg).
 */
import * as React from 'react';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import Modal from './Modal';
import { deepLinkParams } from '../utils/deepLink';
import { EventService } from '../services/EventService';
import { Check, X } from './Icons';

type GrantResult = { email: string; status: 'granted' | 'already' | 'failed' };

export default function GrantAccessHandler(): React.ReactElement | null {
  const { isAdmin, originalIsAdmin } = useRoles();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const adminLike = isAdmin || originalIsAdmin;

  const [running, setRunning] = React.useState(false);
  const [results, setResults] = React.useState<GrantResult[] | null>(null);
  const [open, setOpen] = React.useState(false);
  const handledRef = React.useRef(false);

  React.useEffect(() => {
    if (!adminLike || handledRef.current) return;
    let emails: string[] = [];
    try {
      const p = deepLinkParams();
      if (p.get('action') !== 'grantaccess') return;
      emails = (p.get('emails') || '').split(';').map(s => s.trim()).filter(s => s.indexOf('@') > 0);
    } catch { return; }
    if (emails.length === 0) return;
    handledRef.current = true;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    if (!ctx) return;
    const svc = new EventService(ctx);
    setRunning(true);
    setOpen(true);
    void (async () => {
      const out: GrantResult[] = [];
      for (const mail of emails) {
        try {
          const has = await svc.userHasSiteAccess(mail);
          if (has === true) { out.push({ email: mail, status: 'already' }); continue; }
          const ok = await svc.grantSiteReadAccess(mail);
          out.push({ email: mail, status: ok ? 'granted' : 'failed' });
        } catch { out.push({ email: mail, status: 'failed' }); }
      }
      setResults(out);
      setRunning(false);
    })();
  }, [adminLike]);

  if (!open) return null;

  const label = (s: GrantResult['status']): { txt: string; col: string; icon: React.ReactNode } => {
    if (s === 'granted') return { txt: isDe ? 'Leserechte vergeben' : 'Read access granted', col: 'var(--dex-green-dark, #4a7c1f)', icon: <Check size={14} /> };
    if (s === 'already') return { txt: isDe ? 'War bereits berechtigt' : 'Already had access', col: 'var(--dex-gray-600, #666)', icon: <Check size={14} /> };
    return { txt: isDe ? 'Fehlgeschlagen — bitte manuell über die Site-Berechtigungen vergeben' : 'Failed — please grant manually via site permissions', col: 'var(--dex-red, #c00)', icon: <X size={14} /> };
  };

  return (
    <Modal open={open} onClose={() => { if (!running) setOpen(false); }} maxWidth={560} padding={24} dismissable={!running} ariaLabel={isDe ? 'SharePoint-Zugriff vergeben' : 'Grant SharePoint access'}>
      <h2 style={{ marginTop: 0, marginBottom: 4, color: 'var(--dex-green-dark, #4a7c1f)' }}>
        {isDe ? 'SharePoint-Zugriff vergeben' : 'Grant SharePoint access'}
      </h2>
      {running ? (
        <p style={{ fontSize: '0.9rem', color: 'var(--dex-gray-600)' }}>
          {isDe ? 'Leserechte werden vergeben …' : 'Granting read access …'}
        </p>
      ) : (
        <>
          <p style={{ marginTop: 0, fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
            {isDe
              ? 'Ergebnis der automatischen Freigabe (Leserechte über die Besucher-Gruppe der Site):'
              : 'Result of the automatic grant (read access via the site visitors group):'}
          </p>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginTop: 8 }}>
            {(results || []).map(r => {
              const l = label(r.status);
              return (
                <div key={r.email} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', border: '1px solid var(--dex-gray-200)', borderRadius: 10, background: '#fff' }}>
                  <span style={{ color: l.col, display: 'inline-flex', flexShrink: 0 }}>{l.icon}</span>
                  <span style={{ flex: 1, minWidth: 0, fontSize: '0.88rem', overflow: 'hidden', textOverflow: 'ellipsis' }}>{r.email}</span>
                  <span style={{ fontSize: '0.8rem', fontWeight: 600, color: l.col, textAlign: 'right' }}>{l.txt}</span>
                </div>
              );
            })}
          </div>
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
