/**
 * v23.37: Hinweis-Banner für Admins bei offenen „Organizer werden"-Anträgen.
 *
 * - Sichtbar nur für Admins (auch im Demo-Modus via originalIsAdmin), wenn es
 *   mindestens einen offenen Antrag gibt.
 * - Klick öffnet ein Modal mit allen offenen Anträgen; pro Antrag kann der
 *   Admin direkt „Freigeben" (vergibt die Organizer-Rolle) oder „Ablehnen".
 * - Deep-Link aus der Antrags-Mail (?action=approveorg&request=<id>) öffnet das
 *   Modal automatisch — funktioniert aber nur als Admin.
 */
import * as React from 'react';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useDialog } from '../context/DialogContext';
import Modal from './Modal';
import { deepLinkParams } from '../utils/deepLink';
import { hasOrganizerRights } from '../utils/roleRank';
import { Settings, Check, X } from './Icons';

type Req = { id: number; email: string; name: string; location: string; message: string; created: string };

export default function OrganizerRequestsBanner(): React.ReactElement | null {
  const { isAdmin, originalIsAdmin, addRole, refreshRoles, roles } = useRoles();
  const { getOpenOrganizerRequests, markOrganizerRequestDecided, getOrganizerRequestDetails } = useEvents();
  const { locale } = useLanguage();
  const { navigate } = useNavigation();
  const { showAlert, confirmDialog } = useDialog();
  const isDe = locale === 'de';
  const adminLike = isAdmin || originalIsAdmin;

  const [requests, setRequests] = React.useState<Req[]>([]);
  const [open, setOpen] = React.useState(false);
  const [busyId, setBusyId] = React.useState<number | null>(null);
  const [loaded, setLoaded] = React.useState(false);
  // v26.58: Deep-Link nur EINMAL behandeln (der Effekt läuft bei jedem
  // requests-Update erneut).
  const deepLinkHandledRef = React.useRef(false);

  const reload = React.useCallback(async (): Promise<void> => {
    if (!adminLike) return;
    try { const r = await getOpenOrganizerRequests(); setRequests(r); } catch { /* */ }
    setLoaded(true);
  }, [adminLike, getOpenOrganizerRequests]);

  React.useEffect(() => { void reload(); }, [reload]);

  // Deep-Link aus der Antrags-Mail: Modal automatisch öffnen (nur als Admin).
  // v26.58: Ist der verlinkte Antrag NICHT mehr offen (ein anderer Admin war
  // schneller), landete man vorher kommentarlos auf der Landing Page — jetzt
  // geht es in die Rollenverwaltung mit dem Hinweis, wer wann entschieden hat.
  React.useEffect(() => {
    if (!adminLike || !loaded || deepLinkHandledRef.current) return;
    try {
      const p = deepLinkParams();
      if (p.get('action') !== 'approveorg') return;
      deepLinkHandledRef.current = true;
      const reqId = Number(p.get('request') || 0);
      const stillOpen = reqId > 0 ? requests.some(r => r.id === reqId) : requests.length > 0;
      if (stillOpen) { setOpen(true); return; }
      void (async () => {
        let msg = isDe
          ? 'Dieser Organizer-Antrag wurde bereits bearbeitet.'
          : 'This organizer request has already been handled.';
        if (reqId > 0) {
          try {
            const det = await getOrganizerRequestDetails(reqId);
            if (det) {
              const who = det.decidedByEmail || (isDe ? 'einen anderen Admin' : 'another admin');
              const when = det.decidedDate
                ? new Date(det.decidedDate).toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })
                : '';
              const person = det.name || det.email;
              if (det.status === 'Approved') {
                msg = isDe
                  ? `Der Antrag von ${person} wurde bereits${when ? ` am ${when}` : ''} durch ${who} freigegeben — die Person ist Organizer.`
                  : `The request from ${person} was already approved${when ? ` on ${when}` : ''} by ${who} — the person is an organizer.`;
              } else if (det.status === 'Rejected') {
                msg = isDe
                  ? `Der Antrag von ${person} wurde bereits${when ? ` am ${when}` : ''} durch ${who} abgelehnt.`
                  : `The request from ${person} was already rejected${when ? ` on ${when}` : ''} by ${who}.`;
              }
            }
          } catch { /* generische Meldung reicht */ }
        }
        navigate('settings');
        showAlert(msg, { variant: 'info' });
      })();
    } catch { /* */ }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [adminLike, loaded, requests]);

  if (!adminLike || !loaded || requests.length === 0) return null;

  const decide = async (r: Req, approve: boolean): Promise<void> => {
    setBusyId(r.id);
    try {
      if (approve) {
        // v29.63: Erst prüfen, ob die Person die Rechte längst hat. `addRole`
        // tut das nicht — es legt eine ZWEITE Zeile in DEX_Roles an, und der
        // Antragsweg schickte danach die volle Onboarding-Mail (v28.44) an
        // jemanden, der seit Monaten Organizer ist. Der Admin erfuhr davon
        // nichts. Auch ein Admin/IT-Admin darf hier nicht auf Organizer
        // gesetzt werden — das wäre eine Herabstufung durch die Hintertür.
        const reqMail = (r.email || '').trim().toLowerCase();
        const existing = roles.filter(x => (x.userEmail || '').trim().toLowerCase() === reqMail)[0];
        // v30.67: Über EINE Ableitung (utils/roleRank) statt einer eigenen
        // Aufzählung — die kannte F&A nicht. Seit v30.60 ist F&A ein
        // Organizer-Superset; die Freigabe lief hier bis v30.66 in `addRole`
        // → `updateRole(existing.id, 'Organizer')` und nahm der Person still
        // das F&A Center und den Abrechnungs-Schritt. Weder Admin noch
        // Betroffene erfuhren davon.
        const alreadyEntitled = !!existing && hasOrganizerRights(existing.role);
        if (alreadyEntitled) {
          const roleLabel = existing.role === 'Organizer'
            ? 'Organizer'
            : (isDe ? `${existing.role} (schließt Organizer-Rechte ein)` : `${existing.role} (includes organizer rights)`);
          const proceed = await confirmDialog(
            isDe
              ? `${r.name || r.email} hat bereits die Rolle „${roleLabel}" — die Rechte sind also schon da.\n\nEs wird KEINE Rolle angelegt und KEINE Onboarding-Mail verschickt.\n\nDen Antrag jetzt als erledigt abhaken?`
              : `${r.name || r.email} already holds the role "${roleLabel}" — the rights are already in place.\n\nNo role will be created and NO onboarding email will be sent.\n\nMark the request as done now?`,
            { confirmLabel: isDe ? 'Als erledigt abhaken' : 'Mark as done' },
          );
          if (!proceed) { setBusyId(null); return; }
          // Antrag schliessen, aber ohne Mail — die Person weiss ja nichts von
          // einem Wechsel, weil es keinen gab.
          await markOrganizerRequestDecided(r.id, 'Approved', r.email, r.name || r.email, { suppressMail: true });
          setRequests(prev => prev.filter(x => x.id !== r.id));
          setBusyId(null);
          return;
        }
        const ok = await addRole(r.email, r.name || r.email, 'Organizer', r.location || '');
        if (!ok) { setBusyId(null); return; }
        await markOrganizerRequestDecided(r.id, 'Approved', r.email, r.name || r.email);
        try { await refreshRoles(); } catch { /* */ }
      } else {
        await markOrganizerRequestDecided(r.id, 'Rejected', r.email, r.name || r.email);
      }
      setRequests(prev => prev.filter(x => x.id !== r.id));
    } finally { setBusyId(null); }
  };

  const count = requests.length;

  return (
    <>
      <div
        onClick={() => setOpen(true)}
        style={{
          display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer',
          margin: '0 0 12px', padding: '10px 16px', borderRadius: 10,
          background: 'rgba(134,188,37,0.12)', border: '1px solid var(--dex-green, #86bc25)',
          color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600, fontSize: '0.9rem',
        }}
      >
        <Settings size={18} />
        <span>
          {count === 1
            ? (isDe ? '1 offener Antrag „Organizer werden"' : '1 open request to become an organizer')
            : (isDe ? `${count} offene Anträge „Organizer werden"` : `${count} open requests to become an organizer`)}
        </span>
        <span style={{ marginLeft: 'auto', textDecoration: 'underline' }}>
          {isDe ? 'Ansehen & bestätigen' : 'Review & approve'}
        </span>
      </div>

      <Modal open={open} onClose={() => setOpen(false)} maxWidth={620} padding={24} ariaLabel={isDe ? 'Organizer-Anträge' : 'Organizer requests'}>
        <h2 style={{ marginTop: 0, marginBottom: 4, color: 'var(--dex-green-dark, #4a7c1f)' }}>
          {isDe ? 'Anträge „Organizer werden"' : 'Requests to become an organizer'}
        </h2>
        <p style={{ marginTop: 0, fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
          {isDe
            ? 'Mit „Freigeben" bekommt die Person die Organizer-Rolle und kann eigene Events anlegen. Sie wird per Mail informiert.'
            : 'With “Approve” the person gets the organizer role and can create their own events. They are notified by email.'}
        </p>
        {requests.length === 0 ? (
          <p style={{ color: 'var(--dex-gray-500)' }}>{isDe ? 'Keine offenen Anträge mehr.' : 'No open requests left.'}</p>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 8 }}>
            {requests.map(r => (
              <div key={r.id} style={{ display: 'flex', alignItems: 'flex-start', gap: 12, padding: '12px 14px', border: '1px solid var(--dex-gray-200)', borderRadius: 10, background: '#fff' }}>
                {r.email ? (
                  <img
                    src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(r.email)}&size=M`}
                    alt={r.name}
                    onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                    style={{ width: 40, height: 40, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                  />
                ) : null}
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 700, color: 'var(--dex-gray-800)' }}>{r.name || r.email}</div>
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
                    {r.email}{r.location ? ` · ${r.location}` : ''}
                  </div>
                  {r.message ? (
                    <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)', marginTop: 4, whiteSpace: 'pre-wrap' }}>{r.message}</div>
                  ) : null}
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6, flexShrink: 0 }}>
                  <button
                    type="button"
                    className="btn btn-primary"
                    style={{ fontSize: '0.78rem', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                    disabled={busyId === r.id}
                    onClick={() => { void decide(r, true); }}
                  >
                    <Check size={14} /> {busyId === r.id ? (isDe ? 'Bitte warten…' : 'Please wait…') : (isDe ? 'Freigeben' : 'Approve')}
                  </button>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ fontSize: '0.78rem', padding: '6px 12px', color: 'var(--dex-gray-700)', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                    disabled={busyId === r.id}
                    onClick={() => { void decide(r, false); }}
                  >
                    <X size={14} /> {isDe ? 'Ablehnen' : 'Reject'}
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
        <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 16 }}>
          <button type="button" className="btn btn-secondary" onClick={() => setOpen(false)}>
            {isDe ? 'Schließen' : 'Close'}
          </button>
        </div>
      </Modal>
    </>
  );
}
