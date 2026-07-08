/**
 * v27.10 (Refactor): unverändert aus RegistrationPage.tsx extrahiert
 * (reine String-Helfer ohne Komponenten-Abhängigkeiten).
 */

// v18.74: Externe Adresse = kein Deloitte-Deutschland-Postfach (@deloitte.de).
export const isExternalEmailAddr = (e: string): boolean => {
  const v = (e || '').trim();
  return !!v && !/@(.*\.)?deloitte\.de$/i.test(v);
};

// v18.74: Strengere Plausibilitätsprüfung gegen Tippfehler bei externen
// Adressen — fängt fehlende/zu kurze TLD, doppelte Punkte, mehrere @, führende/
// abschließende Punkte und Whitespace/Kommas ab. Verifiziert NICHT die Existenz
// des Postfachs (das geht clientseitig nicht), aber blockt offensichtliche
// Vertipper.
export const isPlausibleEmail = (e: string): boolean => {
  const v = (e || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[a-zA-Z]{2,}$/.test(v)) return false;
  if ((v.match(/@/g) || []).length !== 1) return false;
  if (v.indexOf('..') >= 0) return false;
  if (/[\s,;]/.test(v)) return false;
  const [local, domain] = v.split('@');
  if (!local || !domain) return false;
  if (local.startsWith('.') || local.endsWith('.')) return false;
  if (domain.startsWith('.') || domain.endsWith('.') || domain.startsWith('-')) return false;
  return true;
};
