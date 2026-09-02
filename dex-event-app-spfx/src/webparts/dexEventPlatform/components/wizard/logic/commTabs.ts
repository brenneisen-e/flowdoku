import * as React from 'react';
import { SUB_TRANSFER_GROUPS } from '../../../data/wizardHints';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { EmailOverrideEntry } from '../../wizard/emailOverrideEntry';

/* switchCommTab — aus EventCreationPage.tsx ausgelagert (Zeilen 3276-3361 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface SwitchCommTabCtx {
  activeCommTabIdx: number;
  autoDeregisterOnDecline: boolean;
  disableCancellationEmail: boolean;
  disableEmails: boolean;
  disableOutlook: boolean;
  disableRegistrationEmail: boolean;
  emailLanguage: string;
  emailLogoPreview: string;
  emailTemplateOverrides: Record<string, EmailOverrideEntry>;
  inactiveHandling: "notify" | "autoderegister";
  locale: import("../../../context/LanguageContext").Locale;
  outlookBody: string;
  outlookHeading: string;
  outlookLogoPreview: string;
  outlookSubheading: string;
  outlookSubject: string;
  setActiveCommTabIdx: React.Dispatch<React.SetStateAction<number>>;
  setAutoDeregisterOnDecline: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableCancellationEmail: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableEmails: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableOutlook: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableRegistrationEmail: React.Dispatch<React.SetStateAction<boolean>>;
  setEmailLanguage: React.Dispatch<React.SetStateAction<string>>;
  setEmailLogoPreview: React.Dispatch<React.SetStateAction<string>>;
  setEmailTemplateOverrides: React.Dispatch<React.SetStateAction<Record<string, EmailOverrideEntry>>>;
  setInactiveHandling: React.Dispatch<React.SetStateAction<"notify" | "autoderegister">>;
  setOutlookBody: React.Dispatch<React.SetStateAction<string>>;
  setOutlookHeading: React.Dispatch<React.SetStateAction<string>>;
  setOutlookLogoPreview: React.Dispatch<React.SetStateAction<string>>;
  setOutlookSubheading: React.Dispatch<React.SetStateAction<string>>;
  setOutlookSubject: React.Dispatch<React.SetStateAction<string>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  subEvents: SubEventDraft[];
  subEventsRef: React.MutableRefObject<SubEventDraft[]>;
  topLevelCommSnapshot: React.MutableRefObject<{ emailLanguage: string; emailLogoBase64: string; outlookLogoBase64: string; outlookBody: string; outlookHeading: string; outlookSubheading: string; outlookSubject: string; disableEmails: boolean; disableRegistrationEmail: boolean; disableCancellationEmail: boolean; autoDeregisterOnDecline: boolean; inactiveHandling?: 'notify' | 'autoderegister'; disableOutlook: boolean; emailTemplateOverrides: Record<string, EmailOverrideEntry>; }>;
}

export function switchCommTabImpl(ctx: SwitchCommTabCtx, nextIdx: number): void {
  const { activeCommTabIdx, autoDeregisterOnDecline, disableCancellationEmail, disableEmails, disableOutlook, disableRegistrationEmail, emailLanguage, emailLogoPreview, emailTemplateOverrides, inactiveHandling, locale, outlookBody, outlookHeading, outlookLogoPreview, outlookSubheading, outlookSubject, setActiveCommTabIdx, setAutoDeregisterOnDecline, setDisableCancellationEmail, setDisableEmails, setDisableOutlook, setDisableRegistrationEmail, setEmailLanguage, setEmailLogoPreview, setEmailTemplateOverrides, setInactiveHandling, setOutlookBody, setOutlookHeading, setOutlookLogoPreview, setOutlookSubheading, setOutlookSubject, setSubEvents, subEvents, subEventsRef, topLevelCommSnapshot } = ctx;
    if (nextIdx === activeCommTabIdx) return;
    // 1) Aktuellen UI-State in das ausgehende Slot schreiben.
    if (activeCommTabIdx > 0) {
      const fromIdx = activeCommTabIdx - 1;
      // v11.60: synchron in den Ref schreiben (siehe flushActiveCommTabToState).
      const flushed = subEventsRef.current.map((s, i) => i === fromIdx ? {
        ...s,
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        autoDeregisterOnDecline,
        inactiveHandling,
        disableOutlook,
        // v14.4: Mail-Text-Overrides pro Sub-Event mitspiegeln.
        emailTemplateOverrides: { ...emailTemplateOverrides },
      } : s);
      subEventsRef.current = flushed;
      setSubEvents(flushed);
    } else {
      // Slot 0 = Top-Level. Der UI-State wird hier direkt vom Top-Level-State
      // gehalten — kein Snapshot nötig, weil setEmailLanguage etc. den Wert
      // schon dort hält. Beim Zurück-Wechsel auf Tab 0 setzen wir die
      // Top-Level-States aus dem `topLevelCommSnapshot`-Ref (siehe unten).
      topLevelCommSnapshot.current = {
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        autoDeregisterOnDecline,
        inactiveHandling,
        disableOutlook,
        emailTemplateOverrides: { ...emailTemplateOverrides },
      };
    }
    // 2) Werte aus dem Ziel-Slot in die Step-5-UI laden.
    if (nextIdx === 0) {
      const snap = topLevelCommSnapshot.current;
      if (snap) {
        setEmailLanguage(snap.emailLanguage);
        setEmailLogoPreview(snap.emailLogoBase64 || '');
        setOutlookLogoPreview(snap.outlookLogoBase64 || '');
        setOutlookBody(snap.outlookBody || '');
        setOutlookHeading(snap.outlookHeading || '');
        setOutlookSubheading(snap.outlookSubheading || '');
        setOutlookSubject(snap.outlookSubject || '');
        setDisableEmails(!!snap.disableEmails);
        setDisableRegistrationEmail(!!snap.disableRegistrationEmail);
        setDisableCancellationEmail(!!snap.disableCancellationEmail);
        setAutoDeregisterOnDecline(!!snap.autoDeregisterOnDecline);
        setInactiveHandling(snap.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify');
        setDisableOutlook(!!snap.disableOutlook);
        setEmailTemplateOverrides(snap.emailTemplateOverrides || {});
      }
    } else {
      const sub = subEvents[nextIdx - 1];
      if (sub) {
        setEmailLanguage(sub.emailLanguage || (locale === 'de' ? 'DE' : 'EN'));
        setEmailLogoPreview(sub.emailLogoBase64 || '');
        setOutlookLogoPreview(sub.outlookLogoBase64 || '');
        setOutlookBody(sub.outlookBody || '');
        setOutlookHeading(sub.outlookHeading || sub.title || '');
        setOutlookSubheading(sub.outlookSubheading || '');
        setOutlookSubject(sub.outlookSubject || '');
        setDisableEmails(!!sub.disableEmails);
        setDisableRegistrationEmail(!!sub.disableRegistrationEmail);
        setDisableCancellationEmail(!!sub.disableCancellationEmail);
        setAutoDeregisterOnDecline(!!sub.autoDeregisterOnDecline);
        setInactiveHandling(sub.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify');
        setDisableOutlook(!!sub.disableOutlook);
        setEmailTemplateOverrides(sub.emailTemplateOverrides || {});
      }
    }
    setActiveCommTabIdx(nextIdx);
}

/* flushActiveCommTabToState — aus EventCreationPage.tsx ausgelagert (Zeilen 3385-3411 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface FlushActiveCommTabToStateCtx {
  activeCommTabIdx: number;
  autoDeregisterOnDecline: boolean;
  disableCancellationEmail: boolean;
  disableEmails: boolean;
  disableOutlook: boolean;
  disableRegistrationEmail: boolean;
  emailLanguage: string;
  emailLogoPreview: string;
  emailTemplateOverrides: Record<string, EmailOverrideEntry>;
  inactiveHandling: "notify" | "autoderegister";
  outlookBody: string;
  outlookHeading: string;
  outlookLogoPreview: string;
  outlookSubheading: string;
  outlookSubject: string;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  subEventsRef: React.MutableRefObject<SubEventDraft[]>;
}

export function flushActiveCommTabToStateImpl(ctx: FlushActiveCommTabToStateCtx): void {
  const { activeCommTabIdx, autoDeregisterOnDecline, disableCancellationEmail, disableEmails, disableOutlook, disableRegistrationEmail, emailLanguage, emailLogoPreview, emailTemplateOverrides, inactiveHandling, outlookBody, outlookHeading, outlookLogoPreview, outlookSubheading, outlookSubject, setSubEvents, subEventsRef } = ctx;
    if (activeCommTabIdx > 0) {
      const fromIdx = activeCommTabIdx - 1;
      // v11.60: synchron in den Ref schreiben — sonst sieht die direkt
      // anschliessende Detect-/Persist-Logik noch die alte Array.
      const flushed = subEventsRef.current.map((s, i) => i === fromIdx ? {
        ...s,
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        autoDeregisterOnDecline,
        inactiveHandling,
        disableOutlook,
        emailTemplateOverrides: { ...emailTemplateOverrides },
      } : s);
      subEventsRef.current = flushed;
      setSubEvents(flushed);
    }
    // Slot 0 (Top-Level) wird ohnehin direkt von den State-Variablen gespeist
    // — kein Snapshot-Flush nötig (resolveTopLevelCommState liest auf Tab 0
    // direkt aus dem State, der Snapshot wird nur für Sub-Tab-Pfade benutzt).
}

/* resolveTopLevelCommState — aus EventCreationPage.tsx ausgelagert (Zeilen 3426-3479 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface ResolveTopLevelCommStateCtx {
  activeCommTabIdx: number;
  autoDeregisterOnDecline: boolean;
  disableCancellationEmail: boolean;
  disableEmails: boolean;
  disableOutlook: boolean;
  disableRegistrationEmail: boolean;
  emailLanguage: string;
  emailLogoPreview: string;
  emailTemplateOverrides: Record<string, EmailOverrideEntry>;
  inactiveHandling: "notify" | "autoderegister";
  outlookBody: string;
  outlookHeading: string;
  outlookLogoPreview: string;
  outlookSubheading: string;
  outlookSubject: string;
  topLevelCommSnapshot: React.MutableRefObject<{ emailLanguage: string; emailLogoBase64: string; outlookLogoBase64: string; outlookBody: string; outlookHeading: string; outlookSubheading: string; outlookSubject: string; disableEmails: boolean; disableRegistrationEmail: boolean; disableCancellationEmail: boolean; autoDeregisterOnDecline: boolean; inactiveHandling?: 'notify' | 'autoderegister'; disableOutlook: boolean; emailTemplateOverrides: Record<string, EmailOverrideEntry>; }>;
}

export function resolveTopLevelCommStateImpl(ctx: ResolveTopLevelCommStateCtx): { emailLanguage: string; emailLogoBase64: string; outlookLogoBase64: string; outlookBody: string; outlookHeading: string; outlookSubheading: string; outlookSubject: string; disableEmails: boolean; disableRegistrationEmail: boolean; disableCancellationEmail: boolean; autoDeregisterOnDecline: boolean; inactiveHandling?: 'notify' | 'autoderegister'; disableOutlook: boolean; emailTemplateOverrides: Record<string, EmailOverrideEntry>; } {
  const { activeCommTabIdx, autoDeregisterOnDecline, disableCancellationEmail, disableEmails, disableOutlook, disableRegistrationEmail, emailLanguage, emailLogoPreview, emailTemplateOverrides, inactiveHandling, outlookBody, outlookHeading, outlookLogoPreview, outlookSubheading, outlookSubject, topLevelCommSnapshot } = ctx;
    if (activeCommTabIdx === 0) {
      return {
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        autoDeregisterOnDecline,
        inactiveHandling,
        disableOutlook,
        emailTemplateOverrides,
      };
    }
    const snap = topLevelCommSnapshot.current;
    if (snap) return snap;
    // Fallback (sollte praktisch nicht eintreten): wir sind auf einem Sub-Tab,
    // hatten aber noch keinen Snapshot — verwenden die aktuellen State-Werte,
    // damit zumindest kein Crash entsteht.
    return {
      emailLanguage,
      emailLogoBase64: emailLogoPreview,
      outlookLogoBase64: outlookLogoPreview,
      outlookBody,
      outlookHeading,
      outlookSubheading,
      outlookSubject,
      disableEmails,
      disableRegistrationEmail,
      disableCancellationEmail,
      autoDeregisterOnDecline,
      inactiveHandling,
      disableOutlook,
      emailTemplateOverrides,
    };
}

/**
 * v30.71: Die Themen des Kommunikations-Schritts, die ein Termin eigenständig
 * haben kann — je Thema die Felder des Slots. Schritt 27 (Organizer in Kopie)
 * fehlt bewusst: Er gilt event-weit und hat keinen Sub-Event-Slot.
 * `labelKey` ist der Übersetzungsschlüssel der Schritt-Überschrift, damit die
 * Themenliste im Schalter dieselben Namen trägt wie die Abschnitte darunter.
 */
export const COMM_TOPICS: Array<{ key: string; step: number; labelKey: string; fields: string[] }> = [
  { key: 'language', step: 25, labelKey: 'create.emaillanguage', fields: ['emailLanguage'] },
  { key: 'switches', step: 26, labelKey: 'create.notifications', fields: ['disableEmails', 'disableRegistrationEmail', 'disableCancellationEmail', 'autoDeregisterOnDecline', 'inactiveHandling', 'disableOutlook'] },
  { key: 'mailLogo', step: 28, labelKey: 'create.eventlogo.mail', fields: ['emailLogoBase64'] },
  { key: 'outlookLogo', step: 29, labelKey: 'create.outlooklogo', fields: ['outlookLogoBase64'] },
  { key: 'outlookText', step: 30, labelKey: 'create.outlookdesc', fields: ['outlookBody', 'outlookHeading', 'outlookSubheading', 'outlookSubject'] },
  { key: 'templates', step: 31, labelKey: 'create.templates.title', fields: ['emailTemplateOverrides'] },
];

/**
 * v30.71: EIN Thema vom gerade offenen Reiter auf alle anderen Termine.
 * Quelle ist der aktive Reiter (Haupt-Event oder ein Termin) — nach dem Flush,
 * damit die letzte Bearbeitung mitkommt (CLAUDE.md). Objektwerte geklont.
 * Der offene Reiter ist die Quelle und braucht deshalb kein Neuladen; die
 * v30.67-Falle aus applyCommToAllSubEvents (sichtbarer Reiter zeigt alte
 * Werte) tritt hier nicht auf.
 */
export async function applyCommTopicToAllSubEventsImpl(ctx: ApplyCommToAllSubEventsCtx, topicKey: string): Promise<void> {
  const { activeCommTabIdx, childTermPlural, confirmDialog, flushActiveCommTabToState, isDe, resolveTopLevelCommState, setSubEvents, showAlert, subEventsRef } = ctx;
  const topic = COMM_TOPICS.filter(x => x.key === topicKey)[0];
  if (!topic) return;
  flushActiveCommTabToState();
  const srcIdx = activeCommTabIdx - 1;
  const src = (srcIdx >= 0 ? subEventsRef.current[srcIdx] : resolveTopLevelCommState()) as unknown as Record<string, unknown>;
  if (!src) return;
  const targets = subEventsRef.current
    .map((s, i) => ({ s, i }))
    .filter(x => x.i !== srcIdx && x.s.title && x.s.title.trim());
  if (targets.length === 0) return;
  const term = childTermPlural || (isDe ? 'Termine' : 'dates');
  const ok = await confirmDialog(isDe
    ? `„${topic.key === 'templates' ? 'E-Mail-Texte' : topic.key === 'switches' ? 'Versand-Schalter' : topic.key === 'language' ? 'Mail-Sprache' : topic.key === 'mailLogo' ? 'Mail-Logo' : topic.key === 'outlookLogo' ? 'Outlook-Bild' : 'Outlook-Text'}" von diesem Reiter auf ${targets.length} ${term} übertragen? Dort vorhandene eigene Werte für dieses Thema werden überschrieben — alle anderen Themen bleiben, wie sie sind.`
    : `Copy this topic from this tab to ${targets.length} ${term}? Existing own values for this topic there will be overwritten — every other topic stays as it is.`,
    { title: isDe ? 'Ein Thema für alle übernehmen' : 'Apply one topic to all' });
  if (!ok) return;
  const next = subEventsRef.current.map((s, i) => {
    if (i === srcIdx || !s.title || !s.title.trim()) return s;
    const patch: Record<string, unknown> = {};
    for (const f of topic.fields) {
      const v = src[f];
      patch[f] = (v && typeof v === 'object') ? JSON.parse(JSON.stringify(v)) : v;
    }
    return { ...s, ...(patch as unknown as Partial<SubEventDraft>) };
  });
  subEventsRef.current = next;
  setSubEvents(next);
  showAlert(isDe
    ? `Übernommen für ${targets.length} ${term}. Gespeichert wird es mit dem Event.`
    : `Applied to ${targets.length} ${term}. It is saved with the event.`);
}

/* applyCommToAllSubEvents — aus EventCreationPage.tsx ausgelagert (Zeilen 5021-5048 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface ApplyCommToAllSubEventsCtx {
  // v30.67: aktiver Reiter + die Step-6-Setter, damit der sichtbare
  // Sub-Reiter nach dem Kopieren neu geladen werden kann (s. unten).
  activeCommTabIdx: number;
  childTermPlural: string;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  flushActiveCommTabToState: () => void;
  isDe: boolean;
  locale: import("../../../context/LanguageContext").Locale;
  resolveTopLevelCommState: () => { emailLanguage: string; emailLogoBase64: string; outlookLogoBase64: string; outlookBody: string; outlookHeading: string; outlookSubheading: string; outlookSubject: string; disableEmails: boolean; disableRegistrationEmail: boolean; disableCancellationEmail: boolean; autoDeregisterOnDecline: boolean; inactiveHandling?: 'notify' | 'autoderegister'; disableOutlook: boolean; emailTemplateOverrides: Record<string, EmailOverrideEntry>; };
  setAutoDeregisterOnDecline: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableCancellationEmail: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableEmails: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableOutlook: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableRegistrationEmail: React.Dispatch<React.SetStateAction<boolean>>;
  setEmailLanguage: React.Dispatch<React.SetStateAction<string>>;
  setEmailLogoPreview: React.Dispatch<React.SetStateAction<string>>;
  setEmailTemplateOverrides: React.Dispatch<React.SetStateAction<Record<string, EmailOverrideEntry>>>;
  setInactiveHandling: React.Dispatch<React.SetStateAction<"notify" | "autoderegister">>;
  setOutlookBody: React.Dispatch<React.SetStateAction<string>>;
  setOutlookHeading: React.Dispatch<React.SetStateAction<string>>;
  setOutlookLogoPreview: React.Dispatch<React.SetStateAction<string>>;
  setOutlookSubheading: React.Dispatch<React.SetStateAction<string>>;
  setOutlookSubject: React.Dispatch<React.SetStateAction<string>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  subEventsRef: React.MutableRefObject<SubEventDraft[]>;
}

export async function applyCommToAllSubEventsImpl(ctx: ApplyCommToAllSubEventsCtx): Promise<void> {
  const { activeCommTabIdx, childTermPlural, confirmDialog, flushActiveCommTabToState, isDe, locale, resolveTopLevelCommState, setAutoDeregisterOnDecline, setDisableCancellationEmail, setDisableEmails, setDisableOutlook, setDisableRegistrationEmail, setEmailLanguage, setEmailLogoPreview, setEmailTemplateOverrides, setInactiveHandling, setOutlookBody, setOutlookHeading, setOutlookLogoPreview, setOutlookSubheading, setOutlookSubject, setSubEvents, showAlert, subEventsRef } = ctx;
    const named = subEventsRef.current.filter(x => x.title && x.title.trim());
    if (named.length === 0) return;
    const term = childTermPlural || (isDe ? 'Sub-Events' : 'sub-events');
    const ok = await confirmDialog(isDe
      ? `Die Kommunikations-Einstellungen des Haupt-Events (Mail-Sprache, Logo, Outlook-Text, Überschriften, Betreff und alle Mail-Schalter) werden auf ALLE ${named.length} ${term} übertragen.\n\nBereits einzeln gepflegte Werte werden dabei überschrieben. Fortfahren?`
      : `The main event's communication settings will be applied to ALL ${named.length} sub-events. Individually maintained values will be overwritten. Continue?`,
      { title: isDe ? 'Für alle Termine gleich einstellen' : 'Apply to all dates' });
    if (!ok) return;
    // Erst den aktiven Reiter sichern — sonst kopiert man den Stand vor der
    // letzten Bearbeitung (CLAUDE.md: „Kommunikationsfelder der Sub-Events
    // liegen nicht laufend im Draft").
    flushActiveCommTabToState();
    const src = resolveTopLevelCommState() as unknown as Record<string, unknown>;
    const commGroup = SUB_TRANSFER_GROUPS.filter(g => g.key === 'communication')[0];
    const fields = commGroup ? commGroup.fields : [];
    // v30.67: über den Ref rechnen und ihn synchron mitschreiben — wie
    // flushActiveCommTabToState. Ein direkt folgender Flush (Speichern,
    // Reiterwechsel) liest `subEventsRef.current`; ein reines
    // setSubEvents(prev => …) wäre dort noch nicht angekommen.
    const next = subEventsRef.current.map(s => {
      if (!s.title || !s.title.trim()) return s;
      const patch: Record<string, unknown> = {};
      for (const f of fields) {
        const v = src[f];
        patch[f] = (v && typeof v === 'object') ? JSON.parse(JSON.stringify(v)) : v;
      }
      return { ...s, ...(patch as unknown as Partial<SubEventDraft>) };
    });
    subEventsRef.current = next;
    setSubEvents(next);
    // v30.67: Den gerade SICHTBAREN Sub-Reiter neu laden. Die Kommunikations-
    // felder liegen nicht laufend im Draft, sondern im UI-State (CLAUDE.md);
    // der zeigte nach dem Kopieren weiter die ALTEN Werte dieses Termins, und
    // der nächste flushActiveCommTabToState() schrieb genau diesen Stand in
    // den Slot zurück — der offene Termin war der einzige, der die Übernahme
    // nicht bekam, ohne jede Rückmeldung. Geladen wird aus dem frisch
    // gepatchten Slot (nicht aus `src`), damit Felder, die die Gruppe nicht
    // überträgt, ihren Sub-Wert behalten — dieselben Setter wie in
    // switchCommTab.
    if (activeCommTabIdx > 0) {
      const cur = next[activeCommTabIdx - 1];
      if (cur) {
        setEmailLanguage(cur.emailLanguage || (locale === 'de' ? 'DE' : 'EN'));
        setEmailLogoPreview(cur.emailLogoBase64 || '');
        setOutlookLogoPreview(cur.outlookLogoBase64 || '');
        setOutlookBody(cur.outlookBody || '');
        setOutlookHeading(cur.outlookHeading || cur.title || '');
        setOutlookSubheading(cur.outlookSubheading || '');
        setOutlookSubject(cur.outlookSubject || '');
        setDisableEmails(!!cur.disableEmails);
        setDisableRegistrationEmail(!!cur.disableRegistrationEmail);
        setDisableCancellationEmail(!!cur.disableCancellationEmail);
        setAutoDeregisterOnDecline(!!cur.autoDeregisterOnDecline);
        setInactiveHandling(cur.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify');
        setDisableOutlook(!!cur.disableOutlook);
        setEmailTemplateOverrides({ ...(cur.emailTemplateOverrides || {}) });
      }
    }
    showAlert(isDe
      ? `Die Kommunikation des Haupt-Events gilt jetzt für alle ${named.length} ${term}. Nicht vergessen zu speichern.`
      : `The main event's communication now applies to all ${named.length} sub-events. Don't forget to save.`,
      { variant: 'success' });
}

