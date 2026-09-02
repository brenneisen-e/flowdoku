/**
 * v30.66 — Modularisierung Stufe 2: Thema „Branding-Assets" — die
 * seitenweiten Bilder (Logo, Orb) und das Startvideo liegen in
 * SiteAssets/DEX_Assets und gelten für ALLE Events.
 *
 * Nicht zu verwechseln mit dem Event-Bild und dem Mail-Logo eines einzelnen
 * Events (siehe events/eventAssets.ts) — das sind zwei andere Uploads.
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';

// ==================== v26.50: Logo & Branding (Admin-Center) ====================
// Zentrale Ablage des Default-Logos (PNG) + Logo-Videos. Das PNG lebt in der
// _Config-Zeile von DEX_EmailTemplates (LogoBase64) — ALLE neu versendeten
// Mails nutzen es automatisch — und wird zusätzlich als
// SiteAssets/DEX_Logos/Deloitte_Logo.png gespiegelt (Fallback-Pfad von
// loadLogosAsBase64). Das Video liegt als SiteAssets/DEX_Logos/dex-logo-video.*.
// v30.66: war `private static readonly` an der Klasse — nur hier gebraucht.
const BRANDING_VIDEO_BASENAME = 'dex-logo-video';

// ==================== Bild-Upload ====================

/**
 * SiteAssets-Unterordner sicherstellen:
 * - DEX_EventImages (Event-Bilder)
 * - DEX_Logos (Deloitte-Logo für E-Mail-Templates, manuell hochgeladen)
 */
export async function ensureAssetsFolders(svc: EventService): Promise<void> {
  const baseUrl = svc.context.pageContext.web.serverRelativeUrl;
  const folders = ['DEX_EventImages', 'DEX_Logos'];

  for (const folder of folders) {
    const folderUrl = `${baseUrl}/SiteAssets/${folder}`;
    try {
      const check = await svc._sp.get(
        `${svc.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')`,
        SPHttpClient.configurations.v1
      );
      if (check.ok) continue;
    } catch { /* */ }

    try {
      await svc._post(`${svc.siteUrl}/_api/web/folders`, {
        '__metadata': { 'type': 'SP.Folder' },
        'ServerRelativeUrl': folderUrl,
      });
      // Ordner erstellt
    } catch {
      console.warn(`[DEX] Konnte ${folder} Ordner nicht erstellen`);
    }
  }
}

/** Aktuelles Branding: Deloitte-Logo + DEX-Orb (Data-URIs) + Video-URL
 *  (leer wenn keins da). v26.58: orbBase64 ergänzt — das eigentliche
 *  DEX-Logo (bunter Ring, _Config.DefaultImageBase64 bzw. dex-orb.png);
 *  LogoBase64 ist das Deloitte-Logo der E-Mail-Kopfzeile. */
export async function getBranding(svc: EventService): Promise<{ logoBase64: string; orbBase64: string; videoUrl: string; videoFileName: string }> {
  let logoBase64 = '';
  let orbBase64 = '';
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id,LogoBase64,DefaultImageBase64`,
      SPHttpClient.configurations.v1
    );
    if (resp.ok) {
      const d = await resp.json();
      const it = (d.value || d.d?.results || [])[0];
      if (it && it.LogoBase64) logoBase64 = String(it.LogoBase64);
      if (it && it.DefaultImageBase64) orbBase64 = String(it.DefaultImageBase64);
    }
  } catch { /* */ }
  if (!logoBase64) {
    try { logoBase64 = await svc.loadFileAsBase64('DEX_Logos/Deloitte_Logo.png'); } catch { /* */ }
  }
  if (!orbBase64) {
    try { orbBase64 = await svc.loadFileAsBase64('DEX_Logos/dex-orb.png'); } catch { /* */ }
  }
  // Video: feste Kandidaten-Namen prüfen (mp4 bevorzugt).
  let videoUrl = '';
  let videoFileName = '';
  const serverRel = svc.context.pageContext.web.serverRelativeUrl;
  for (const ext of ['mp4', 'webm', 'mov']) {
    const name = `${BRANDING_VIDEO_BASENAME}.${ext}`;
    try {
      const check = await svc._sp.get(
        `${svc.siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRel}/SiteAssets/DEX_Logos/${name}')?$select=Exists,ServerRelativeUrl`,
        SPHttpClient.configurations.v1
      );
      if (check.ok) {
        videoUrl = `${new URL(svc.siteUrl).origin}${serverRel}/SiteAssets/DEX_Logos/${name}`;
        videoFileName = name;
        break;
      }
    } catch { /* nächster Kandidat */ }
  }
  return { logoBase64, orbBase64, videoUrl, videoFileName };
}

/** v26.58: Neues DEX-Logo (Orb, PNG als Data-URI) speichern:
 *  _Config.DefaultImageBase64 (Default-Mail-Bild / {{ORB_URL}}-Fallback)
 *  + Spiegelung nach SiteAssets/DEX_Logos/dex-orb.png. */
export async function saveBrandingOrb(svc: EventService, orbDataUri: string): Promise<boolean> {
  if (!orbDataUri || orbDataUri.indexOf('data:image/') !== 0) return false;
  let ok = false;
  try {
    const listName = 'DEX_EmailTemplates';
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id`,
      SPHttpClient.configurations.v1
    );
    let cfgId = 0;
    if (resp.ok) {
      const d = await resp.json();
      const it = (d.value || d.d?.results || [])[0];
      if (it) cfgId = Number(it.Id);
    }
    if (cfgId > 0) {
      const m = await svc._merge(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${cfgId})`, { 'DefaultImageBase64': orbDataUri });
      ok = m.ok;
    } else {
      const c = await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EmailTemplatesListItem' },
        'Title': '_Config', 'TemplateType': '_Config', 'DefaultImageBase64': orbDataUri,
      });
      ok = c.ok;
    }
  } catch (err) { console.warn('[DEX] saveBrandingOrb (_Config) failed:', err); }
  // Spiegel nach SiteAssets (best-effort — Fallback-Pfad + Download-Quelle).
  try {
    await svc.ensureAssetsFolders();
    const b64 = orbDataUri.split(',')[1] || '';
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const serverRel = svc.context.pageContext.web.serverRelativeUrl;
    await svc._sp.post(
      `${svc.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRel}/SiteAssets/DEX_Logos')/Files/add(url='dex-orb.png',overwrite=true)`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' }, body: bytes.buffer as ArrayBuffer }
    );
  } catch (err) { console.warn('[DEX] saveBrandingOrb (SiteAssets mirror) failed:', err); }
  return ok;
}

/** Neues Default-Logo (PNG als Data-URI) speichern: _Config.LogoBase64 (Mails)
 *  + Spiegelung nach SiteAssets/DEX_Logos/Deloitte_Logo.png (Fallback). */
export async function saveBrandingLogo(svc: EventService, logoDataUri: string): Promise<boolean> {
  if (!logoDataUri || logoDataUri.indexOf('data:image/') !== 0) return false;
  let ok = false;
  try {
    // _Config-Zeile finden bzw. anlegen, dann MERGE.
    const listName = 'DEX_EmailTemplates';
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id`,
      SPHttpClient.configurations.v1
    );
    let cfgId = 0;
    if (resp.ok) {
      const d = await resp.json();
      const it = (d.value || d.d?.results || [])[0];
      if (it) cfgId = Number(it.Id);
    }
    if (cfgId > 0) {
      const m = await svc._merge(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${cfgId})`, { 'LogoBase64': logoDataUri });
      ok = m.ok;
    } else {
      const c = await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EmailTemplatesListItem' },
        'Title': '_Config', 'TemplateType': '_Config', 'LogoBase64': logoDataUri,
      });
      ok = c.ok;
    }
  } catch (err) { console.warn('[DEX] saveBrandingLogo (_Config) failed:', err); }
  // Spiegel nach SiteAssets (best-effort — Fallback-Pfad + Download-Quelle).
  try {
    await svc.ensureAssetsFolders();
    const b64 = logoDataUri.split(',')[1] || '';
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const serverRel = svc.context.pageContext.web.serverRelativeUrl;
    await svc._sp.post(
      `${svc.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRel}/SiteAssets/DEX_Logos')/Files/add(url='Deloitte_Logo.png',overwrite=true)`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' }, body: bytes.buffer as ArrayBuffer }
    );
  } catch (err) { console.warn('[DEX] saveBrandingLogo (SiteAssets mirror) failed:', err); }
  return ok;
}

/** Neues Logo-Video nach SiteAssets/DEX_Logos hochladen (fester Name,
 *  overwrite). Liefert die absolute URL oder '' bei Fehler. */
export async function uploadBrandingVideo(svc: EventService, file: File): Promise<string> {
  try {
    await svc.ensureAssetsFolders();
    const ext = (file.name.split('.').pop() || 'mp4').toLowerCase();
    const safeExt = ['mp4', 'webm', 'mov'].indexOf(ext) >= 0 ? ext : 'mp4';
    const name = `${BRANDING_VIDEO_BASENAME}.${safeExt}`;
    const serverRel = svc.context.pageContext.web.serverRelativeUrl;
    const buf = await file.arrayBuffer();
    const resp = await svc._sp.post(
      `${svc.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRel}/SiteAssets/DEX_Logos')/Files/add(url='${name}',overwrite=true)`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' }, body: buf }
    );
    if (!resp.ok) return '';
    return `${new URL(svc.siteUrl).origin}${serverRel}/SiteAssets/DEX_Logos/${name}`;
  } catch (err) { console.warn('[DEX] uploadBrandingVideo failed:', err); return ''; }
}
