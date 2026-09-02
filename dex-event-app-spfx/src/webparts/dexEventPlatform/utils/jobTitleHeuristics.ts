/**
 * v30.67: Job-Titel-Heuristik „Assistenz" — an EINER Stelle.
 *
 * Dieselbe Frage („ist diese Person eine Assistenz?") wurde bis hierher
 * dreimal mit drei Mustern beantwortet: der Service (`canRegisterForOthers`,
 * v23.9) prüft `assisten`/`assistan`, die Event-Sichtbarkeit
 * (`isEventVisibleForUser`) `/assisten|assistant/i`, und die Anmeldeseite nur
 * das englische `includes('assistant')`. Folge: Wer im Profil „Assistenz",
 * „Teamassistenz" oder „Assistentin" steht, sah den Umschalter „Für eine
 * andere Person anmelden" nie — obwohl der Service die Anmeldung zugelassen
 * hätte. Alle drei Stellen lesen dieselbe Profil-Eigenschaft (`Title`), die
 * Abweichung lag rein im Muster.
 *
 * Das Muster ist bewusst tolerant (Assistenz, Assistant, Team Assistant,
 * Executive Assistant, Assistentin …) und entspricht dem des Service.
 */
export function looksLikeAssistantJobTitle(jobTitle: string | undefined | null): boolean {
  return /assisten|assistan/i.test(jobTitle || '');
}
