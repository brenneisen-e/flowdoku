/**
 * v32.1.0: Mitmach-Tutorial — „Wir legen zusammen ein Test-Event an".
 *
 * Nutzer-Ansage 28.09.2026: „das Tutorial muss viel besser sein … ein echtes
 * Test-Event … Name eingeben, dann Test-Datum … der User muss Dinge anklicken
 * und selbst eingeben und dann auf nächsten Schritt klicken." Konzept und
 * Entscheidungen: `docs/konzept-mitmach-tutorial.md`.
 *
 * Unterschied zur Blätter-Tour (`TutorialOverlay`): Dieses Overlay blockiert
 * NICHTS. Das Abdunkeln ist ein Box-Shadow-Loch mit `pointer-events: none`,
 * die Person tippt und klickt im echten Assistenten. Der Coach sieht nicht
 * auf Klicks, sondern auf ZUSTÄNDE: `EventCreationPage` meldet bei jeder
 * Änderung einen flachen Schnappschuss (`CoachSnapshot`), und jede Station
 * sagt, wann sie erfüllt ist. Dadurch ist es egal, wie die Person dorthin
 * kommt (Tastatur, Maus, anderer Knopf).
 *
 * Zwei Arten von Stationen:
 *  - „input": Der Knopf „Weiter" AUF DER KARTE wird erst frei, wenn der
 *    Zustand erfüllt ist (Titel getippt, Datum gewählt).
 *  - „click": Die Karte hat keinen eigenen Knopf; die Person klickt den
 *    echten Knopf, der Coach rückt von selbst vor.
 *
 * Das Test-Event wird doppelt abgesichert (Formular + `tutorialTestSicher` in
 * wizardSubmit): Präfix „TEST – ", Entwurf, Organizer und Zielgruppe nur die
 * Person selbst. Mails und Outlook laufen echt, aber nur an sie.
 */
import * as React from 'react';
import { GraduationCap, X, Check, Info, Trash2 } from '../Icons';

export const COACH_Z_INDEX = 10800;

/** Was der Assistent dem Coach meldet — flach, damit Vergleiche billig sind. */
export interface CoachSnapshot {
  currentStep: number;
  tcAccepted: boolean;
  /** Abrechnungsfrage nach den Nutzungsbedingungen (nur wer Schritt 10 hat). */
  billingPromptOpen: boolean;
  title: string;
  startDate: string;
  endDate: string;
  isFictive: boolean;
  organizerEmails: string[];
  location: string;
  maxParticipants: string;
  unlimitedParticipants: boolean;
  customFieldCount: number;
  lastFieldLabel: string;
  subEventCount: number;
  countdown: number | null;
  isSubmitting: boolean;
  summaryOpen: boolean;
  createdEventId: string;
}

/** Was der Coach außer dem Schnappschuss über die App wissen muss. */
export interface CoachEnv {
  page: string;
  selectedEventId: string | null;
  testEventId: string;
  myEmail: string;
  isDe: boolean;
}

type StationKind = 'info' | 'input' | 'click';

interface Station {
  id: string;
  kind: StationKind;
  /** CSS-Selektoren; der erste SICHTBARE Treffer bekommt den Spotlight.
   *  Als Funktion, wenn das Ziel vom Stand abhängt (erst Knopf, dann Feld). */
  targets?: string[] | ((s: CoachSnapshot | null, base: StationBase) => string[]);
  /** Zusätzliche Flächen, die mit ins Loch gehören (z.B. offener Datepicker). */
  extra?: string[];
  /** Erwarteter Wizard-Schritt (0-basiert) — sonst Hinweis „zurück zu …". */
  wizardStep?: number;
  /** Seite, auf der die Station spielt. */
  page?: 'create-event' | 'admin';
  titleDe: string; titleEn: string;
  bodyDe: string; bodyEn: string;
  /** Erfüllt? */
  done?: (s: CoachSnapshot | null, env: CoachEnv, base: StationBase) => boolean;
  /** Zusätzliche Warnung (z.B. Datum in der Vergangenheit). */
  warn?: (s: CoachSnapshot | null, env: CoachEnv) => { de: string; en: string } | null;
  /** Darf übersprungen werden (optionales Feld). */
  skippable?: boolean;
}

/** Stand beim Betreten einer Station — z.B. Anzahl Fragen vorher. */
export interface StationBase { customFieldCount: number }

const TITLE = '[data-tour="wizard-title"]';
const NEXT = '[data-tour="wizard-next"]';

function endNachStart(s: CoachSnapshot): boolean {
  const a = new Date(s.startDate).getTime();
  const b = new Date(s.endDate).getTime();
  return !!s.startDate && !!s.endDate && !isNaN(a) && !isNaN(b) && b > a;
}

export const COACH_STATIONS: Station[] = [
  {
    id: 'intro', kind: 'info',
    titleDe: 'Wir legen zusammen ein Test-Event an', titleEn: "Let's create a test event together",
    bodyDe: 'Du tippst und klickst selbst — ich zeige dir, wo. Das Event ist echt, aber abgesichert: Nur du siehst es, Mail und Outlook-Termin gehen nur an dich, und am Ende kannst du es mit einem Klick löschen. Dauer: etwa fünf Minuten.',
    bodyEn: 'You type and click yourself — I show you where. The event is real but safe: only you can see it, the email and Outlook invite go only to you, and you can delete it with one click at the end. About five minutes.',
  },
  {
    id: 'terms', kind: 'click', page: 'create-event',
    titleDe: 'Kurz die Nutzungsbedingungen', titleEn: 'The terms of use',
    bodyDe: 'Lies die Punkte, setz beide Haken und klick auf „Akzeptieren & weiter". Kommt danach die Frage, ob das Event abrechnungsrelevant ist: Für das Test-Event „Nein".',
    bodyEn: 'Read the points, tick both boxes and click “Accept & continue”. If you are then asked whether the event is billing-relevant: “No” for the test event.',
    // Die Abrechnungsfrage (Pilot, nur mit Schritt 10) folgt direkt auf die
    // Bedingungen und liegt als Dialog über Schritt 1 — erst wenn sie zu ist,
    // kann die Person den Titel überhaupt erreichen (Harness-Befund v32.1.0).
    done: s => !!s && s.tcAccepted && !s.billingPromptOpen,
  },
  {
    id: 'title', kind: 'input', page: 'create-event', wizardStep: 0, targets: [TITLE],
    titleDe: 'Wie heißt dein Test-Event?', titleEn: 'What is your test event called?',
    bodyDe: 'Klick ins Feld und tipp einen Namen, zum Beispiel „Team-Frühstück". Beim Anlegen setzt DEX „TEST – " davor, damit jeder sofort sieht, dass es ein Test ist.',
    bodyEn: 'Click into the field and type a name, e.g. “Team breakfast”. When creating it, DEX adds “TEST – ” in front so everyone can see it is a test.',
    done: s => !!s && s.title.trim().length >= 3,
  },
  {
    id: 'dates', kind: 'input', page: 'create-event', wizardStep: 0,
    targets: ['[data-tour="wizard-dates"]'], extra: ['.react-datepicker-popper'],
    titleDe: 'Wann findet es statt?', titleEn: 'When does it take place?',
    bodyDe: 'Wähl ein Test-Datum, am besten nächste Woche: erst „Beginn", dann „Ende". Das Datum landet später 1:1 im Outlook-Termin.',
    bodyEn: 'Pick a test date, ideally next week: first “Start”, then “End”. It later goes 1:1 into the Outlook invite.',
    done: s => !!s && endNachStart(s),
    warn: s => {
      if (!s || !s.startDate) return null;
      const t = new Date(s.startDate).getTime();
      if (!isNaN(t) && t < Date.now()) return { de: 'Der Beginn liegt in der Vergangenheit — nimm besser einen Tag in der Zukunft.', en: 'The start is in the past — better pick a day in the future.' };
      if (s.startDate && s.endDate && !endNachStart(s)) return { de: 'Das Ende muss nach dem Beginn liegen.', en: 'The end must be after the start.' };
      return null;
    },
  },
  {
    id: 'draft', kind: 'input', page: 'create-event', wizardStep: 0, targets: ['[data-tour="wizard-draft"]'],
    titleDe: 'Entwurf bleibt an', titleEn: 'Keep it a draft',
    bodyDe: 'Ganz unten in Schritt 1 steht, wann das Event sichtbar wird. Der Haken „Entwurf" bleibt beim Test an — so sieht es außer dir niemand.',
    bodyEn: 'At the bottom of step 1 you decide when the event becomes visible. Keep “Draft” ticked for the test — nobody else will see it.',
    done: s => !!s && s.isFictive,
    warn: s => (s && !s.isFictive ? { de: 'Bitte den Haken bei „Entwurf" wieder setzen.', en: 'Please tick “Draft” again.' } : null),
  },
  {
    id: 'next1', kind: 'click', page: 'create-event', wizardStep: 0, targets: [NEXT],
    titleDe: 'Weiter zu Schritt 2', titleEn: 'On to step 2',
    bodyDe: 'Klick jetzt auf den grünen Knopf „Weiter".', bodyEn: 'Now click the green “Next” button.',
    done: s => !!s && s.currentStep >= 1,
  },
  {
    id: 'organizer', kind: 'click', page: 'create-event', wizardStep: 1, targets: ['[data-tour="wizard-organizer"]'],
    titleDe: 'Du bist der Organizer', titleEn: 'You are the organizer',
    bodyDe: 'Hier stehen die Organizer — du bist schon eingetragen. Beim Test bleibst du allein, sonst bekämen andere Mails. Klick dann unten auf „Weiter".',
    bodyEn: 'These are the organizers — you are already listed. Stay alone for the test, otherwise others would get emails. Then click “Next” below.',
    done: s => !!s && s.currentStep >= 2,
    warn: (s, env) => {
      if (!s) return null;
      const andere = s.organizerEmails.filter(e => (e || '').toLowerCase() !== env.myEmail.toLowerCase());
      return andere.length > 0 ? { de: 'Beim Speichern bleibst nur du Organizer — weitere Personen werden für das Test-Event nicht übernommen.', en: 'Only you will stay organizer when saving — other people are not taken over for the test event.' } : null;
    },
  },
  {
    id: 'location', kind: 'input', page: 'create-event', wizardStep: 2, targets: ['[data-tour="wizard-location"]'], skippable: true,
    titleDe: 'Wo findet es statt?', titleEn: 'Where does it take place?',
    bodyDe: 'Tipp einen Ort, zum Beispiel „Büro Köln, Raum 3". Er steht später auf der Anmeldeseite und im Outlook-Termin.',
    bodyEn: 'Type a location, e.g. “Cologne office, room 3”. It will show on the registration page and in the Outlook invite.',
    done: s => !!s && s.location.trim().length > 0,
  },
  {
    id: 'next3', kind: 'click', page: 'create-event', wizardStep: 2, targets: [NEXT],
    titleDe: 'Weiter', titleEn: 'Next', bodyDe: 'Und wieder auf „Weiter".', bodyEn: 'And “Next” again.',
    done: s => !!s && s.currentStep >= 3,
  },
  {
    // Vorgabe ist „Unbegrenzt" — dann gibt es kein Zahlenfeld. Erst die
    // Kachel „Begrenzt", dann das Feld (Harness-Befund v32.1.0: die Station
    // galt sonst sofort als erledigt, und niemand hatte etwas getan).
    id: 'capacity', kind: 'input', page: 'create-event', wizardStep: 3, skippable: true,
    targets: s => (s && !s.unlimitedParticipants ? ['[data-tour="wizard-capacity"]'] : ['[data-tour="wizard-capacity-mode"]']),
    titleDe: 'Wie viele Plätze?', titleEn: 'How many seats?',
    bodyDe: 'Wähl „Begrenzt" und tipp eine Zahl, zum Beispiel 10. Wer sich danach anmeldet, landet auf der Warteliste und rückt nach, sobald jemand absagt. Die Anmeldefrist hat DEX schon aus deinem Datum vorgeschlagen.',
    bodyEn: 'Choose “Limited” and type a number, e.g. 10. Anyone registering after that goes onto the waitlist and moves up when someone cancels. DEX has already suggested a registration deadline from your date.',
    done: s => !!s && !s.unlimitedParticipants && Number(s.maxParticipants) > 0,
  },
  {
    id: 'next4', kind: 'click', page: 'create-event', wizardStep: 3, targets: [NEXT],
    titleDe: 'Weiter zum Anmeldeformular', titleEn: 'On to the registration form', bodyDe: 'Klick auf „Weiter".', bodyEn: 'Click “Next”.',
    done: s => !!s && s.currentStep >= 4,
  },
  {
    id: 'question', kind: 'input', page: 'create-event', wizardStep: 4,
    // Erst der Knopf „Frage hinzufügen"; sobald es eine Frage MEHR gibt, das
    // Textfeld dieser neuen Frage (der Anker sitzt immer am letzten Feld).
    targets: (s, base) => (s && s.customFieldCount > base.customFieldCount
      ? ['[data-tour="wizard-question-label"]'] : ['[data-tour="wizard-add-question"]']), skippable: true,
    titleDe: 'Stell eine eigene Frage', titleEn: 'Ask your own question',
    bodyDe: 'Klick auf „Frage hinzufügen" und schreib eine Frage, zum Beispiel „Kaffee oder Tee?". Sie erscheint im Anmeldeformular, die Antworten siehst du später in der Teilnehmerliste.',
    bodyEn: 'Click “Add question” and write one, e.g. “Coffee or tea?”. It appears in the registration form; you will see the answers in the participant list.',
    done: (s, _e, base) => !!s && s.customFieldCount > base.customFieldCount && s.lastFieldLabel.trim().length > 0,
  },
  {
    id: 'next5', kind: 'click', page: 'create-event', wizardStep: 4, targets: [NEXT],
    titleDe: 'Weiter zur Kommunikation', titleEn: 'On to communication', bodyDe: 'Klick auf „Weiter".', bodyEn: 'Click “Next”.',
    done: s => !!s && s.currentStep >= 5,
  },
  {
    id: 'channel', kind: 'info', page: 'create-event', wizardStep: 5, targets: ['[data-tour="wizard-channel"]'],
    titleDe: 'Mail und Outlook-Termin', titleEn: 'Email and Outlook invite',
    bodyDe: 'Hier legst du fest, was Teilnehmer bekommen. Lass „Mail + Outlook-Termin" an — beim Test-Event geht beides nur an dich. Weiter unten zeigt dir die Vorschau, wie die Mail aussieht.',
    bodyEn: 'This is where you decide what participants get. Leave “Email + Outlook invite” on — for the test event both go only to you. Further down, the preview shows the email.',
  },
  {
    id: 'create', kind: 'click', page: 'create-event', targets: ['[data-tour="wizard-create-early"]', '[data-tour="wizard-submit"]'],
    titleDe: 'Jetzt anlegen', titleEn: 'Create it now',
    bodyDe: 'Team, Dokumente und Fun-Zone sind optional — für den Test brauchst du sie nicht. Klick auf „Event erstellen".',
    bodyEn: 'Team, documents and fun zone are optional — you do not need them for the test. Click “Create event”.',
    done: s => !!s && (s.countdown !== null || s.isSubmitting || !!s.createdEventId),
    warn: s => (s && s.subEventCount > 0 ? { de: 'Du hast Sub-Events angelegt — für den Test besser wieder entfernen, jedes bekäme eine eigene Teilnehmerliste.', en: 'You added sub-events — better remove them for the test, each would get its own participant list.' } : null),
  },
  {
    id: 'creating', kind: 'click',
    titleDe: 'DEX legt dein Event an', titleEn: 'DEX is creating your event',
    bodyDe: 'Erst zehn Sekunden zum Abbrechen — lass sie laufen. Dann entstehen Teilnehmerliste und Rechte, das dauert 30 bis 60 Sekunden. Danach zeigt dir DEX eine Zusammenfassung; schließ sie, um zu deinem Event zu kommen.',
    bodyEn: 'First ten seconds to cancel — let them run. Then the participant list and rights are created, which takes 30 to 60 seconds. DEX then shows a summary; close it to go to your event.',
    done: (_s, env) => env.page === 'admin' && !!env.testEventId && env.selectedEventId === env.testEventId,
  },
  {
    id: 'admin', kind: 'info', page: 'admin', targets: ['[data-tour="admin-next-steps"]'],
    titleDe: 'Da ist dein Test-Event', titleEn: 'There is your test event',
    bodyDe: 'Das ist das Organizer Center deines Events, Status „Entwurf". Der Kasten zeigt die nächsten Schritte eines echten Events. „Live schalten" und „Einladungsmail" bitte beim Test nicht nutzen. Schau in dein Postfach: Die Mail „Event angelegt" ist an dich gegangen.',
    bodyEn: 'This is the Organizer Center of your event, status “Draft”. The box shows the next steps of a real event. Please do not use “Go live” or “Invitation email” for the test. Check your inbox: the “Event created” email went to you.',
  },
  {
    id: 'finish', kind: 'info',
    titleDe: 'Geschafft!', titleEn: 'Done!',
    bodyDe: 'Du hast dein erstes Event angelegt. Soll das Test-Event wieder weg? Es landet im Papierkorb; der Outlook-Termin wird abgesagt.',
    bodyEn: 'You created your first event. Should the test event go away again? It goes to the recycle bin; the Outlook invite is cancelled.',
  },
];

interface Rect { top: number; left: number; width: number; height: number }

function sichtbar(el: Element): boolean {
  const h = el as HTMLElement;
  const r = h.getBoundingClientRect();
  if (r.width <= 0 || r.height <= 0) return false;
  let n: HTMLElement | null = h;
  while (n) {
    const cs = window.getComputedStyle(n);
    if (cs.display === 'none' || cs.visibility === 'hidden' || cs.opacity === '0') return false;
    n = n.parentElement;
  }
  return true;
}

function ersterSichtbarer(selectors: string[]): HTMLElement | null {
  for (const sel of selectors) {
    const all = Array.prototype.slice.call(document.querySelectorAll(sel)) as HTMLElement[];
    const hit = all.find(sichtbar);
    if (hit) return hit;
  }
  return null;
}

function vereinigung(els: HTMLElement[]): Rect | null {
  if (els.length === 0) return null;
  let t = Infinity, l = Infinity, b = -Infinity, r = -Infinity;
  for (const e of els) {
    const x = e.getBoundingClientRect();
    t = Math.min(t, x.top); l = Math.min(l, x.left); b = Math.max(b, x.bottom); r = Math.max(r, x.right);
  }
  return { top: t, left: l, width: r - l, height: b - t };
}

/** Ist gerade ein App-Dialog offen (nicht der Coach selbst)? */
function appDialogOffen(): boolean {
  const all = Array.prototype.slice.call(document.querySelectorAll('[role="dialog"]')) as HTMLElement[];
  return all.some(d => !d.closest('[data-coach]') && sichtbar(d));
}

export interface WizardCoachProps {
  snapshot: CoachSnapshot | null;
  env: CoachEnv;
  stationIdx: number;
  setStationIdx: (i: number) => void;
  base: StationBase;
  onClose: () => void;
  onGoWizardStep: (step: number) => void;
  onGoToWizard: () => void;
  onDeleteTest: () => void;
  onKeepTest: () => void;
  onRealEvent: () => void;
  deleteState: { busy: boolean; blockedReason: string | null; error: string | null; done: boolean };
}

export function WizardCoach(p: WizardCoachProps): React.ReactElement | null {
  const { snapshot: s, env, stationIdx, setStationIdx, base, onClose } = p;
  const isDe = env.isDe;
  const st = COACH_STATIONS[Math.min(stationIdx, COACH_STATIONS.length - 1)];
  const [rect, setRect] = React.useState<Rect | null>(null);
  const [dialogOpen, setDialogOpen] = React.useState(false);
  const scrolledForRef = React.useRef<string>('');

  const erfuellt = !!st.done && st.done(s, env, base);
  const targets: string[] = typeof st.targets === 'function' ? st.targets(s, base) : (st.targets || []);
  const targetsKey = targets.join('|');

  // Klick-Stationen rücken von selbst vor, sobald ihr Zustand erfüllt ist.
  // Die Nutzungsbedingungen werden übersprungen, wenn sie schon bestätigt sind.
  React.useEffect(() => {
    if (st.kind === 'click' && erfuellt) {
      const t = window.setTimeout(() => setStationIdx(stationIdx + 1), 350);
      return () => window.clearTimeout(t);
    }
    return undefined;
  }, [st.id, erfuellt, stationIdx, setStationIdx, st.kind]);

  // Spotlight: laufend nachmessen (Seiten laden lazy, Felder klappen auf,
  // Datepicker öffnet sich). 250 ms reichen fürs Auge und kosten nichts.
  React.useEffect(() => {
    const messen = (): void => {
      setDialogOpen(appDialogOffen());
      if (targets.length === 0) { setRect(null); return; }
      const el = ersterSichtbarer(targets);
      if (!el) { setRect(null); return; }
      if (scrolledForRef.current !== st.id + targetsKey) {
        scrolledForRef.current = st.id + targetsKey;
        const r = el.getBoundingClientRect();
        if (r.top < 120 || r.bottom > window.innerHeight - 40) {
          try { el.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch { el.scrollIntoView(); }
        }
      }
      const extras = (st.extra || []).map(sel => ersterSichtbarer([sel])).filter((x): x is HTMLElement => !!x);
      const r = vereinigung([el].concat(extras));
      // Verdeckt ein fremder Dialog das Ziel? Nicht jeder Dialog trägt
      // role="dialog" (die Abrechnungsfrage z.B. nicht) — deshalb fragen wir
      // den Browser, was an der Mitte des Ziels oben liegt. Der Coach selbst
      // zählt nicht (pointer-events: none bzw. data-coach).
      const er = el.getBoundingClientRect();
      const py = er.top + Math.min(er.height / 2, 20);
      const px = er.left + er.width / 2;
      const imBild = py > 0 && py < window.innerHeight && px > 0 && px < window.innerWidth;
      const oben = imBild ? document.elementFromPoint(px, py) as HTMLElement | null : null;
      const verdeckt = !!oben && !el.contains(oben) && !oben.contains(el) && !oben.closest('[data-coach]');
      if (verdeckt) setDialogOpen(true);
      setRect(r);
    };
    messen();
    const iv = window.setInterval(messen, 250);
    window.addEventListener('scroll', messen, true);
    window.addEventListener('resize', messen);
    return () => {
      window.clearInterval(iv);
      window.removeEventListener('scroll', messen, true);
      window.removeEventListener('resize', messen);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [st.id, targetsKey]);

  // Abgebrochener Countdown: Steht der Coach schon auf „DEX legt an", die
  // Person hat aber „Abbrechen" geklickt, zurück auf „Jetzt anlegen". Mit
  // Verzögerung, weil zwischen Countdown-Ende und `isSubmitting` ein Render
  // liegt, in dem beides leer ist.
  React.useEffect(() => {
    if (st.id !== 'creating' || !s || env.page !== 'create-event') return undefined;
    if (s.countdown !== null || s.isSubmitting || s.createdEventId || s.summaryOpen) return undefined;
    const t = window.setTimeout(() => setStationIdx(stationIdx - 1), 1500);
    return () => window.clearTimeout(t);
  }, [st.id, s, env.page, stationIdx, setStationIdx]);

  // ESC tut bewusst nichts: Die Person arbeitet im echten Formular, und ESC
  // schließt dort Datepicker und Dialoge — ein Tutorial, das dabei mit
  // verschwindet, wäre die falsche Antwort. Beenden geht über die Karte.

  const vw = window.innerWidth;
  const vh = window.innerHeight;
  const cardW = Math.min(380, vw - 32);
  const pad = 8;
  const falscheSeite = !!st.page && env.page !== st.page;
  const falscherSchritt = !falscheSeite && typeof st.wizardStep === 'number' && !!s && s.currentStep !== st.wizardStep && !erfuellt;
  const warnung = st.warn ? st.warn(s, env) : null;

  // Karte: bei offenem App-Dialog unten rechts andocken, damit sie nichts
  // verdeckt; sonst unter oder über dem Ziel; ohne Ziel mittig.
  let cardStyle: React.CSSProperties;
  const mobil = vw < 640;
  if (mobil) {
    cardStyle = { position: 'fixed', left: 12, right: 12, bottom: 12 };
  } else if (dialogOpen) {
    cardStyle = { position: 'fixed', right: 20, bottom: 20, width: cardW };
  } else if (rect && !falscheSeite) {
    const below = vh - (rect.top + rect.height);
    const left = Math.min(Math.max(16, rect.left + rect.width / 2 - cardW / 2), Math.max(16, vw - cardW - 16));
    if (below >= 300) cardStyle = { position: 'fixed', top: Math.max(0, rect.top + rect.height) + pad + 14, left, width: cardW };
    else if (rect.top >= 300) cardStyle = { position: 'fixed', bottom: vh - rect.top + pad + 14, left, width: cardW };
    else cardStyle = { position: 'fixed', right: 20, bottom: 20, width: cardW };
  } else {
    cardStyle = { position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', width: Math.min(440, vw - 32) };
  }

  const zentral = targets.length === 0;
  const dimmen = !dialogOpen && (zentral || !!rect) && !falscheSeite;
  const total = COACH_STATIONS.length;
  const letzte = stationIdx >= total - 1;

  const weiterKnopf = (): React.ReactNode => {
    if (st.id === 'finish') return null;
    if (st.kind === 'click') return null;
    const frei = st.kind === 'info' || erfuellt;
    return (
      <button type="button" className="btn btn-primary" disabled={!frei} onClick={() => setStationIdx(stationIdx + 1)}
        style={{ fontSize: '0.84rem', padding: '8px 18px', opacity: frei ? 1 : 0.5, cursor: frei ? 'pointer' : 'not-allowed' }}>
        {stationIdx === 0 ? (isDe ? 'Los geht’s' : 'Let’s go') : (isDe ? 'Weiter' : 'Next')}
      </button>
    );
  };

  return (
    <div data-coach="1" aria-live="polite" style={{ position: 'fixed', inset: 0, zIndex: COACH_Z_INDEX, pointerEvents: 'none' }}>
      {dimmen && (rect && !zentral ? (
        // Außen: das Loch im Abdunkeln (Box-Shadow). Innen: der pulsierende
        // Rahmen — getrennt, weil dexUiPulse selbst box-shadow animiert und
        // sonst das Abdunkeln überschreibt (Harness-Befund v32.1.0).
        <div style={{
          position: 'fixed', top: rect.top - pad, left: rect.left - pad,
          width: rect.width + pad * 2, height: rect.height + pad * 2, borderRadius: 12,
          boxShadow: '0 0 0 100000px rgba(15,23,42,0.4)',
          pointerEvents: 'none',
          transition: 'top 0.2s ease, left 0.2s ease, width 0.2s ease, height 0.2s ease',
        }}>
          <div style={{ position: 'absolute', inset: 0, borderRadius: 12, border: '2px solid var(--dex-green, #86bc25)', animation: 'dexUiPulse 1.6s ease-out infinite' }} />
        </div>
      ) : (
        <div style={{ position: 'absolute', inset: 0, background: 'rgba(15,23,42,0.45)', pointerEvents: 'none' }} />
      ))}
      {/* Ohne Abdunkeln, aber mit Ziel: nur der pulsierende Rahmen. */}
      {!dimmen && rect && !falscheSeite && !dialogOpen && (
        <div style={{
          position: 'fixed', top: rect.top - pad, left: rect.left - pad,
          width: rect.width + pad * 2, height: rect.height + pad * 2, borderRadius: 12,
          border: '2px solid var(--dex-green, #86bc25)', animation: 'dexUiPulse 1.6s ease-out infinite', pointerEvents: 'none',
        }} />
      )}

      <div role="dialog" aria-label={isDe ? st.titleDe : st.titleEn} style={{
        ...cardStyle, pointerEvents: 'auto', background: '#fff', borderRadius: 16,
        padding: '16px 18px 14px', boxShadow: '0 16px 48px rgba(0,0,0,0.3)', fontFamily: 'inherit',
        border: '1px solid var(--dex-gray-200)',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, gap: 8 }}>
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.7rem', fontWeight: 700, letterSpacing: '0.06em', textTransform: 'uppercase', color: 'var(--dex-green-dark, #4a7c1f)' }}>
            <GraduationCap size={14} /> {isDe ? 'Test-Event anlegen' : 'Create a test event'} · {Math.min(stationIdx + 1, total)}/{total}
          </span>
          <button type="button" onClick={onClose} aria-label={isDe ? 'Tutorial beenden' : 'End tutorial'}
            style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-400)', padding: 4, display: 'inline-flex' }}>
            <X size={16} />
          </button>
        </div>
        <div className="dex-ui-progress" style={{ marginBottom: 12 }}>
          <div className="dex-ui-progress-bar" style={{ width: `${Math.round(((stationIdx + 1) / total) * 100)}%` }} />
        </div>
        <h3 style={{ margin: '0 0 6px', fontSize: '1.02rem', color: 'var(--dex-gray-800)' }}>{isDe ? st.titleDe : st.titleEn}</h3>
        <p style={{ margin: '0 0 12px', fontSize: '0.86rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>{isDe ? st.bodyDe : st.bodyEn}</p>

        {/* Stand anzeigen: erfüllt / noch offen — die Person sieht, worauf der Coach wartet. */}
        {st.kind === 'input' && !!st.done && !falscheSeite && !falscherSchritt && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.8rem', marginBottom: 10, color: erfuellt ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-500)' }}>
            <span style={{ display: 'inline-flex', width: 18, height: 18, borderRadius: 999, alignItems: 'center', justifyContent: 'center', background: erfuellt ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-100)', color: '#fff' }}>
              {erfuellt && <Check size={12} />}
            </span>
            {erfuellt ? (isDe ? 'Erledigt — weiter geht’s.' : 'Done — on we go.') : (isDe ? 'Warte auf deine Eingabe …' : 'Waiting for your input …')}
          </div>
        )}
        {st.kind === 'click' && st.id !== 'creating' && !falscheSeite && !falscherSchritt && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.8rem', marginBottom: 10, color: 'var(--dex-gray-500)' }}>
            <Info size={14} /> {isDe ? 'Ich warte, bis du geklickt hast.' : 'I will wait until you click.'}
          </div>
        )}
        {warnung && (
          <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 10, fontSize: '0.8rem' }}>
            <span>{isDe ? warnung.de : warnung.en}</span>
          </div>
        )}
        {falscheSeite && st.page === 'create-event' && (
          <div className="dex-ui-callout dex-ui-callout--info" style={{ marginBottom: 10, fontSize: '0.8rem', alignItems: 'center' }}>
            <span style={{ flex: 1 }}>{isDe ? 'Du bist gerade nicht im Event-Assistenten.' : 'You are not in the event wizard right now.'}</span>
            <button type="button" className="btn btn-secondary" style={{ fontSize: '0.78rem', padding: '6px 12px' }} onClick={p.onGoToWizard}>
              {isDe ? 'Zurück zum Assistenten' : 'Back to the wizard'}
            </button>
          </div>
        )}
        {falscherSchritt && s && typeof st.wizardStep === 'number' && (
          <div className="dex-ui-callout dex-ui-callout--info" style={{ marginBottom: 10, fontSize: '0.8rem', alignItems: 'center' }}>
            <span style={{ flex: 1 }}>{isDe ? `Du bist gerade in Schritt ${s.currentStep + 1}.` : `You are in step ${s.currentStep + 1}.`}</span>
            <button type="button" className="btn btn-secondary" style={{ fontSize: '0.78rem', padding: '6px 12px' }} onClick={() => p.onGoWizardStep(st.wizardStep as number)}>
              {isDe ? `Zu Schritt ${st.wizardStep + 1}` : `Go to step ${st.wizardStep + 1}`}
            </button>
          </div>
        )}

        {st.id === 'finish' ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {p.deleteState.done ? (
              <div className="dex-ui-callout dex-ui-callout--success" style={{ fontSize: '0.84rem' }}>
                <span>{isDe ? 'Test-Event gelöscht. Du kannst jetzt dein echtes Event anlegen.' : 'Test event deleted. You can now create your real event.'}</span>
              </div>
            ) : (
              <>
                {p.deleteState.blockedReason && (
                  <div className="dex-ui-muted" style={{ fontSize: '0.78rem' }}>{p.deleteState.blockedReason}</div>
                )}
                {p.deleteState.error && (
                  <div className="dex-ui-callout dex-ui-callout--danger" style={{ fontSize: '0.8rem' }}><span>{p.deleteState.error}</span></div>
                )}
                <button type="button" className="btn btn-primary" onClick={p.onDeleteTest}
                  disabled={p.deleteState.busy || !!p.deleteState.blockedReason || !env.testEventId}
                  style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 6, opacity: (p.deleteState.busy || !!p.deleteState.blockedReason) ? 0.6 : 1 }}>
                  <Trash2 size={14} /> {p.deleteState.busy ? (isDe ? 'Wird gelöscht …' : 'Deleting …') : (isDe ? 'Test-Event löschen' : 'Delete test event')}
                </button>
              </>
            )}
            <div style={{ display: 'flex', gap: 8, justifyContent: 'space-between' }}>
              {!p.deleteState.done && (
                <button type="button" className="btn btn-secondary" onClick={p.onKeepTest} style={{ fontSize: '0.82rem' }}>
                  {isDe ? 'Als Entwurf behalten' : 'Keep as draft'}
                </button>
              )}
              <button type="button" className="btn btn-outline" onClick={p.onRealEvent} style={{ fontSize: '0.82rem', marginLeft: 'auto' }}>
                {isDe ? 'Echtes Event anlegen' : 'Create a real event'}
              </button>
            </div>
          </div>
        ) : (
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
            <button type="button" onClick={onClose} style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '6px 0', color: 'var(--dex-gray-400)', fontSize: '0.78rem', fontFamily: 'inherit' }}>
              {isDe ? 'Tutorial beenden' : 'End tutorial'}
            </button>
            <div style={{ display: 'flex', gap: 8 }}>
              {st.skippable && !erfuellt && (
                <button type="button" className="btn btn-secondary" onClick={() => setStationIdx(stationIdx + 1)} style={{ fontSize: '0.84rem', padding: '8px 14px' }}>
                  {isDe ? 'Überspringen' : 'Skip'}
                </button>
              )}
              {!letzte && weiterKnopf()}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
