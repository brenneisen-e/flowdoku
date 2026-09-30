/**
 * v32.51 — Kopf-Maße einer Event-Mail, auch wenn das Mail-Logo noch nicht im
 * State ist.
 *
 * Seit v32.0.6/v32.0.10 liest der Start `DEX_Events` OHNE EmailImageBase64,
 * und `_eventLogo` liegt nicht mehr im Overrides-JSON; beides kommt erst mit
 * dem Hintergrund-Nachlauf (`ensureOutlookBodies`). Bis dahin sehen
 * `eventHeaderImageOpts` und `applyEventTemplateOverride` „kein eigenes Bild"
 * und liefern den kleinen Standard-Kopf — Anmeldebestätigung, Abmeldung,
 * Nachrücken und Organizer-Mails zeigten das Event-Bild klein, obwohl der
 * Flow über {{ORB_URL}} das echte Logo einsetzt (Nutzer-Befunde 30.09.2026).
 * Der Merker `_logosAusgelagert` hilft nicht: Er wird bei jedem vollständigen
 * Kommunikations-Write gesetzt, auch ohne Bild.
 *
 * Deshalb: Fehlt das Logo im State, genau diese eine Spalte der Zeile
 * nachlesen (5 Minuten gemerkt — Nachrücken in Serie liest nicht je Person).
 * Scheitert das, bleibt es beim bisherigen Verhalten.
 */
import { SPHttpClient } from '@microsoft/sp-http';
import { eventHeaderImageOpts } from './mailHeaderImage';

const cache = new Map<string, { at: number; p: Promise<string> }>();
const FRISCH_MS = 5 * 60 * 1000;

/** Das Mail-Logo eines Events — aus dem State, sonst aus der Spalte. '' = keins/unlesbar. */
export function mailLogoFrisch(ev: { id?: string; mailImageBase64?: string } | null | undefined): Promise<string> {
  if (!ev) return Promise.resolve('');
  if (ev.mailImageBase64) return Promise.resolve(ev.mailImageBase64);
  const idNum = Number(ev.id);
  if (!isFinite(idNum) || idNum <= 0) return Promise.resolve('');
  const key = String(idNum);
  const hit = cache.get(key);
  if (hit && Date.now() - hit.at < FRISCH_MS) return hit.p;
  const p = (async (): Promise<string> => {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (!ctx) return '';
      const site = ctx.pageContext.web.absoluteUrl;
      const r = await ctx.spHttpClient.get(
        `${site}/_api/web/lists/getbytitle('DEX_Events')/items(${idNum})?$select=EmailImageBase64`,
        SPHttpClient.configurations.v1,
      );
      if (!r.ok) { cache.delete(key); return ''; }
      const j = await r.json();
      return (j && typeof j.EmailImageBase64 === 'string') ? j.EmailImageBase64 : '';
    } catch { cache.delete(key); return ''; }
  })();
  cache.set(key, { at: Date.now(), p });
  return p;
}

/** `eventHeaderImageOpts` mit nachgelesenem Logo. */
export async function eventHeaderImageOptsFrisch(
  ev: { id?: string; emailTemplateOverrides?: string; mailImageBase64?: string },
): Promise<ReturnType<typeof eventHeaderImageOpts>> {
  return eventHeaderImageOpts(ev.emailTemplateOverrides, await mailLogoFrisch(ev));
}
