import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

/**
 * Anzeigenamen des angemeldeten Users absichern (v28.64)
 * ------------------------------------------------------
 * `pageContext.user.displayName` liefert bei einzelnen Personen nicht den
 * Namen, sondern das Claims-Login-Token: `0#.f|membership|user@deloitte.de`
 * (bzw. mit Präfix `i:0#.f|…`).
 *
 * Ursache: SPFx nimmt den Namen aus der versteckten Liste
 * „User Information List" der Site — nicht aus dem Benutzerprofil. In diese
 * Liste wird ein User in dem Moment gestempelt, in dem er der Site zum ersten
 * Mal begegnet (Login, `ensureUser`, Zuweisung einer Berechtigung). Steht zu
 * diesem Zeitpunkt im Profil noch kein Anzeigename — typisch bei frisch
 * angelegten Konten, bei Gast-/B2B-Konten und bei Personen, die per
 * `ensureUser` über die Claim-Adresse angelegt wurden —, trägt SharePoint das
 * Login-Token als `Title` ein. Der nachgelagerte Sync des
 * Benutzerprofildienstes korrigiert die Liste NICHT rückwirkend, deshalb
 * bleibt der Eintrag dauerhaft falsch, obwohl das Profil längst stimmt.
 * (Siehe u.a. Microsoft Learn „Get user identity and properties in
 * SharePoint" zum Claims-Format sowie die einschlägigen Community-Threads zu
 * „i:0#.f|membership|" in Personenfeldern.)
 *
 * Konsequenz für uns: Der Wert aus `pageContext` ist nicht vertrauenswürdig.
 * Sieht er aus wie ein Claim, holen wir den Namen aus dem Benutzerprofil
 * (`GetMyProperties` → `PreferredName`) — das ist die Quelle, die stimmt.
 */

/** Sieht der Wert nach einem Claims-Login-Token aus statt nach einem Namen? */
export const looksLikeClaimName = (s: string): boolean => {
  const v = (s || '').trim();
  if (!v) return false;
  return /\|membership\b|^i:0[#|]|^c:0|0#\.[a-z]\||^\d+#\./i.test(v);
};

/** Die E-Mail aus einem Claims-Token ziehen („…|membership|a@b.de" → „a@b.de"). */
export const emailFromClaim = (s: string): string => {
  const m = (s || '').match(/\|([^|]+@[^|\s]+)\s*$/);
  return m ? m[1].trim().toLowerCase() : '';
};

/**
 * Den echten Anzeigenamen des angemeldeten Users aus dem Benutzerprofil holen.
 * Liefert '' wenn nichts Brauchbares da ist — der Aufrufer behält dann seinen
 * bisherigen Wert.
 */
export async function resolveMyDisplayName(context: WebPartContext): Promise<string> {
  try {
    const siteUrl = context.pageContext.web.absoluteUrl;
    const resp = await context.spHttpClient.get(
      `${siteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=DisplayName,UserProfileProperties`,
      SPHttpClient.configurations.v1,
    );
    if (!resp.ok) return '';
    const data = await resp.json();
    const props: Array<{ Key: string; Value: string }> = data.UserProfileProperties || [];
    const get = (key: string): string => {
      const p = props.filter(x => x.Key === key)[0];
      return p && p.Value ? String(p.Value).trim() : '';
    };
    // PreferredName ist der gepflegte Anzeigename; DisplayName kommt aus
    // derselben Quelle wie pageContext und kann denselben Fehler haben.
    const candidates = [
      get('PreferredName'),
      String(data.DisplayName || '').trim(),
      [get('LastName'), get('FirstName')].filter(Boolean).join(', '),
    ];
    for (const c of candidates) {
      if (c && !looksLikeClaimName(c)) return c;
    }
    return '';
  } catch {
    return '';
  }
}

/**
 * Sofort verwendbarer Ersatz, solange (oder falls) das Profil nichts hergibt:
 * lieber die E-Mail-Adresse als ein Login-Token — die ist wenigstens lesbar
 * und identifiziert die Person eindeutig.
 */
export const safeDisplayName = (raw: string, email: string): string => {
  const v = (raw || '').trim();
  if (v && !looksLikeClaimName(v)) return v;
  return emailFromClaim(v) || (email || '').trim() || '';
};
