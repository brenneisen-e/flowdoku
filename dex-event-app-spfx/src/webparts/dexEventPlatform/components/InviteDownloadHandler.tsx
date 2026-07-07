/**
 * InviteDownloadHandler (v26.73)
 *
 * Deep-Link-Handler für
 *   `action=downloadinvite&event=<id>&item=<regId>&email=<extEmail>&name=<extName>`
 * aus der externen Instruktions-Mail. Der fertige .eml-Einladungs-Entwurf wurde
 * bei der Anmeldung als Attachment an der Teilnehmer-Zeile abgelegt (der Anhang
 * darf per Deloitte-Mail-Regel NICHT direkt mit der Mail verschickt werden).
 * Dieser Handler öffnet ein kleines Fenster mit EINEM „Herunterladen"-Button —
 * bewusst kein Auto-Download, weil Browser den beim Seiten-Öffnen (ohne Klick)
 * blocken. Der Anmelder lädt die Datei und leitet sie aus dem eigenen Postfach
 * an die externe Person weiter.
 */
import * as React from 'react';
import { useLanguage } from '../context/LanguageContext';
import { useEvents } from '../context/EventContext';
import Modal from './Modal';
import { deepLinkParams } from '../utils/deepLink';
import { EventService } from '../services/EventService';
import { downloadEml } from '../utils/emlDraft';
import { Download, Check } from './Icons';

export default function InviteDownloadHandler(): React.ReactElement | null {
  const { locale } = useLanguage();
  const { events, isEventsLoading } = useEvents();
  const isDe = locale === 'de';

  const [open, setOpen] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [done, setDone] = React.useState(false);
  const [failed, setFailed] = React.useState(false);
  const [info, setInfo] = React.useState<{ eventTitle: string; name: string; email: string; item: number; subsiteUrl: string } | null>(null);
  const handledRef = React.useRef(false);

  React.useEffect(() => {
    if (isEventsLoading || handledRef.current) return;
    let p: URLSearchParams;
    try { p = deepLinkParams(); } catch { return; }
    if (p.get('action') !== 'downloadinvite') return;
    const eventParam = (p.get('event') || '').trim();
    const item = parseInt(p.get('item') || '0', 10);
    const email = (p.get('email') || '').trim();
    const name = (p.get('name') || '').trim() || email;
    if (!eventParam || !item) return;
    handledRef.current = true;
    let evt = events.find(e => e.id === eventParam);
    if (!evt) { const n = parseInt(eventParam, 10); if (!isNaN(n)) evt = events.find(e => e.eventNumber === n); }
    setOpen(true);
    if (!evt || !evt.subsiteUrl) { setFailed(true); return; }
    setInfo({ eventTitle: evt.title, name, email, item, subsiteUrl: evt.subsiteUrl });
  }, [isEventsLoading, events]);

  if (!open) return null;

  const doDownload = async (): Promise<void> => {
    if (!info) return;
    setBusy(true);
    setFailed(false);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    if (!ctx) { setBusy(false); setFailed(true); return; }
    const svc = new EventService(ctx);
    const res = await svc.getInviteEmlByItem(info.subsiteUrl, info.item).catch(() => null);
    setBusy(false);
    if (!res || !res.content) { setFailed(true); return; }
    downloadEml('Einladung_' + (info.email || 'extern'), res.content);
    setDone(true);
  };

  return (
    <Modal open={open} onClose={() => setOpen(false)} maxWidth={520} padding={24} ariaLabel={isDe ? 'Einladungs-Entwurf herunterladen' : 'Download invitation draft'}>
      <h2 style={{ marginTop: 0, marginBottom: 6, color: 'var(--dex-green-dark, #4a7c1f)' }}>
        {isDe ? 'Einladungs-Entwurf herunterladen' : 'Download invitation draft'}
      </h2>
      {!info ? (
        <p style={{ fontSize: '0.9rem', color: 'var(--dex-red, #c00)' }}>
          {isDe
            ? 'Der Einladungs-Entwurf konnte nicht gefunden werden. Bitte lade ihn im Organizer Center in der Teilnehmerzeile herunter.'
            : 'The invitation draft could not be found. Please download it from the participant row in the Organizer Center.'}
        </p>
      ) : (
        <>
          <p style={{ marginTop: 0, fontSize: '0.88rem', color: 'var(--dex-gray-700)', lineHeight: 1.55 }}>
            {isDe
              ? <>Fertiger Einladungs-Entwurf für <strong>{info.name}</strong>{info.email ? <> ({info.email})</> : ''} zum Event <strong>{info.eventTitle}</strong>. Herunterladen, in Outlook öffnen und auf &bdquo;Senden&ldquo; (bzw. &bdquo;Weiterleiten&ldquo;) klicken.</>
              : <>Ready-made invitation draft for <strong>{info.name}</strong>{info.email ? <> ({info.email})</> : ''} for the event <strong>{info.eventTitle}</strong>. Download it, open it in Outlook and click &ldquo;Send&rdquo; (or &ldquo;Forward&rdquo;).</>}
          </p>
          {done ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 12px', borderRadius: 8, background: 'rgba(134,188,37,0.10)', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.86rem', fontWeight: 600 }}>
              <Check size={16} /> {isDe ? 'Heruntergeladen — im Download-Ordner. In Outlook öffnen und weiterleiten.' : 'Downloaded — check your downloads. Open in Outlook and forward.'}
            </div>
          ) : failed ? (
            <p style={{ fontSize: '0.84rem', color: 'var(--dex-red, #c00)' }}>
              {isDe ? 'Download fehlgeschlagen. Bitte über das Organizer Center herunterladen.' : 'Download failed. Please download via the Organizer Center.'}
            </p>
          ) : null}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 18 }}>
            <button type="button" className="btn btn-secondary" onClick={() => setOpen(false)}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
            {!done && (
              <button type="button" className="btn btn-primary" disabled={busy} onClick={() => { void doDownload(); }} style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                <Download size={16} /> {busy ? (isDe ? 'Lädt…' : 'Loading…') : (isDe ? 'Herunterladen' : 'Download')}
              </button>
            )}
          </div>
        </>
      )}
    </Modal>
  );
}
