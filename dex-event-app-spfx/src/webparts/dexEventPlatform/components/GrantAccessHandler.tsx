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
 *
 * v31.6: Der Knopf fragte bis hierher `userHasSiteAccess` — „darf die Person
 * die Seite öffnen?" — und meldete bei „ja" ungeprüft „war bereits
 * berechtigt". Genau das trifft die Personen, um die es geht: Wer aus einer
 * Member Firm über eine genehmigte SharePoint-Zugriffsanfrage ein
 * PERSÖNLICHES Leserecht hat, erfüllt die Prüfung — und sieht trotzdem kein
 * Event, weil DEX jedes Teilnehmer-Recht an die BESUCHER-GRUPPE vergibt (nie
 * an eine einzelne Person). Der Knopf tat für sie also nichts und meldete
 * Erfolg. Geprüft wird deshalb jetzt die Gruppen-Mitgliedschaft; wer nicht
 * drin ist, wird aufgenommen, egal ob er die Seite schon öffnen kann. Ist die
 * Prüfung selbst nicht möglich, wird aufgenommen (idempotent) und dazugesagt,
 * dass der Vorher-Stand unbekannt war — behauptet wird nichts.
 */
import * as React from 'react';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import Modal from './Modal';
import { deepLinkParams } from '../utils/deepLink';
import { EventService } from '../services/EventService';
import { Check, X } from './Icons';

// v31.6: `unverified` = die Mitgliedschaft in der Besucher-Gruppe ließ sich
// vorher nicht prüfen (keine Gruppe gefunden, fehlende Rechte, Netzfehler).
// Dann wird trotzdem aufgenommen — aber die Zeile sagt, was offen blieb.
type GrantResult = { email: string; status: 'granted' | 'already' | 'failed'; unverified?: boolean };

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
          // v31.6: Nicht „darf sie die Seite öffnen", sondern „ist sie in der
          // Besucher-Gruppe" — daran hängen die Event-Liste und jede
          // Teilnehmerliste. null (nicht prüfbar) zählt bewusst NICHT als
          // „drin": dann wird aufgenommen, das ist idempotent.
          const inGroup = await svc.isInVisitorsGroup(mail);
          if (inGroup === true) { out.push({ email: mail, status: 'already' }); continue; }
          const ok = await svc.grantSiteReadAccess(mail);
          out.push({ email: mail, status: ok ? 'granted' : 'failed', unverified: inGroup === null });
        } catch { out.push({ email: mail, status: 'failed' }); }
      }
      setResults(out);
      setRunning(false);
    })();
  }, [adminLike]);

  if (!open) return null;

  // v31.6: Die Texte benennen jetzt die Besucher-Gruppe statt „berechtigt" —
  // ein persönliches Leserecht ist genau NICHT das, was hier zählt.
  const label = (s: GrantResult['status']): { txt: string; col: string; icon: React.ReactNode } => {
    if (s === 'granted') return { txt: isDe ? 'In die Besucher-Gruppe aufgenommen' : 'Added to the visitors group', col: 'var(--dex-green-dark, #4a7c1f)', icon: <Check size={14} /> };
    if (s === 'already') return { txt: isDe ? 'War schon in der Besucher-Gruppe' : 'Already in the visitors group', col: 'var(--dex-gray-600, #666)', icon: <Check size={14} /> };
    return { txt: isDe ? 'Fehlgeschlagen — bitte manuell in die Besucher-Gruppe der Site aufnehmen' : 'Failed — please add manually to the site visitors group', col: 'var(--dex-red, #c00)', icon: <X size={14} /> };
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
              ? 'Ergebnis der automatischen Freigabe. Aufgenommen wird in die Besucher-Gruppe der Site — nur darüber sieht jemand die Events und die Anmeldeseiten:'
              : 'Result of the automatic grant. People are added to the site visitors group — that group alone opens the events and registration pages:'}
          </p>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginTop: 8 }}>
            {(results || []).map(r => {
              const l = label(r.status);
              return (
                <div key={r.email} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', border: '1px solid var(--dex-gray-200)', borderRadius: 10, background: '#fff' }}>
                  <span style={{ color: l.col, display: 'inline-flex', flexShrink: 0 }}>{l.icon}</span>
                  <span style={{ flex: 1, minWidth: 0, fontSize: '0.88rem', overflow: 'hidden', textOverflow: 'ellipsis' }}>{r.email}</span>
                  <span style={{ minWidth: 0, textAlign: 'right' }}>
                    <span style={{ fontSize: '0.8rem', fontWeight: 600, color: l.col, display: 'block' }}>{l.txt}</span>
                    {/* v31.6: Ehrlich bleiben — wenn die Mitgliedschaft vorher
                        nicht zu lesen war, steht das hier, statt sie zu
                        behaupten. Die Aufnahme selbst schadet nicht. */}
                    {r.unverified && r.status === 'granted' && (
                      <span className="dex-ui-muted" style={{ display: 'block', fontSize: '0.72rem', marginTop: 2 }}>
                        {isDe
                          ? 'Vorheriger Stand war nicht prüfbar — sicherheitshalber aufgenommen.'
                          : 'Previous membership could not be checked — added to be safe.'}
                      </span>
                    )}
                  </span>
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
