/**
 * v32.17/v32.18: Ist das ein Einführungs-Event zu DEX selbst?
 *
 * v32.17 erkannte es am Titel; seit v32.18 setzt ein Admin den Haken
 * „Einführungsevent“ in Schritt 1 (Piggyback _dexIntro, Nutzer-Ansage
 * 29.09.2026). Ein Termin (Sub-Event) erbt die Einordnung vom Hauptevent.
 */
export const istDexEinfuehrung = (
  ev?: { dexIntro?: boolean } | null,
  hauptevent?: { dexIntro?: boolean } | null,
): boolean => !!ev?.dexIntro || !!hauptevent?.dexIntro;
