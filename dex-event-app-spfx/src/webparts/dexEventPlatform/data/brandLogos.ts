/**
 * v24.13: Offizielles SCHWARZES Deloitte-Logo als Base64-PNG.
 *
 * Verwendet im Rollen-Matrix-PDF (oben links, weißer Grund). Das in der App
 * gecachte Deloitte-Logo ist WEISS (für dunkle Mail-Header) und auf weißem
 * Hintergrund unsichtbar.
 *
 * SO BEFÜLLEN: Lege das schwarze Logo als Datei unter
 *   dex-event-app-spfx/assets/brand/Deloitte_Logo_black.png
 * ab — der Maintainer/Build kodiert es dann hier als `data:image/png;base64,…`
 * hinein. Solange dieser String leer ist, fällt das PDF auf das (per Canvas auf
 * Schwarz umgefärbte) weiße Logo zurück.
 */
export const DELOITTE_LOGO_BLACK = '';
