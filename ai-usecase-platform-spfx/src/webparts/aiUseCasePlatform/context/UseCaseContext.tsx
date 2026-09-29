/**
 * Die Use Cases — laden, anlegen, aendern, loeschen.
 *
 * Der `ladeStatus` ist hier kein Beiwerk, sondern der Kern: Eine leere
 * Kachelwand und eine nicht lesbare Liste sehen gleich aus, und die
 * Kachelwand IST die App. In DEX hat genau diese Verwechslung eine
 * Organizerin ein Event mit 77 Anmeldungen als leer sehen lassen (v30.37).
 * Deshalb drei Zustaende: `laedt`, `ok`, `fehler` — und die Oberflaeche
 * benennt den dritten, statt „keine Use Cases" zu behaupten.
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { UseCase, BildAenderung, SpeicherErgebnis, StartbestandErgebnis } from '../types';
import { useRoles } from './RoleContext';
import { START_USE_CASES } from '../data/startUseCases';
import { geaendertText, nurGeaendertes } from '../utils/aenderungen';

export type LadeStatus = 'laedt' | 'ok' | 'fehler';

interface UseCaseContextType {
  useCases: UseCase[];
  ladeStatus: LadeStatus;
  /** HTTP-Status des letzten Lesens — fuer die Meldung („403" heisst Rechte, nicht kaputt). */
  letzterStatus: number;
  /** Was SharePoint im Klartext geantwortet hat. Ohne das raet man beim naechsten Mal wieder. */
  letzterFehler: string;
  /** Spalten, die beim Anlegen der Liste nicht entstanden sind. */
  fehlendeSpalten: string[];
  /**
   * Das Nachladen ist fehlgeschlagen, aber es gibt einen früheren Stand — der bleibt
   * sichtbar (statt die ganze Kachelwand durch eine Fehlermeldung zu ersetzen), und
   * dieser Text sagt, dass er veraltet sein kann. Leer = alles aktuell.
   */
  aktualisierungFehler: string;
  /** Die Start-Use-Cases sind nur teilweise angelegt (oder gar nicht) — mit Zahlen und Grund. */
  startbestandTeilweise: StartbestandErgebnis | null;
  reload: () => Promise<void>;
  create: (uc: Partial<UseCase>) => Promise<number | null>;
  update: (id: number, uc: Partial<UseCase>) => Promise<boolean>;
  remove: (id: number) => Promise<boolean>;
  /**
   * v1.3: Einen Use Case samt Kachelbild speichern (`id === null` = neu).
   *
   * Die Reihenfolge steht HIER und nirgends sonst, weil sie das Bild schützt:
   * erst das neue Bild hochladen, dann die Zeile speichern, und das alte Bild
   * erst NACH erfolgreichem Speichern entfernen. Bricht irgendein Schritt ab,
   * bleibt der alte Stand vollständig erhalten — „erst anlegen, dann löschen"
   * (DEX v30.67). Beim Anlegen gibt es die Zeile erst, wenn sie ein Bild
   * aufnehmen kann; scheitert der Upload, ist der Use Case trotzdem da und
   * `bildFehler` sagt es.
   */
  saveUseCase: (id: number | null, uc: Partial<UseCase>, bild: BildAenderung) => Promise<SpeicherErgebnis>;
  /** Alle vorkommenden Bereiche, alphabetisch — fuer den Filter. */
  bereiche: string[];
  /**
   * Die Start-Use-Cases anlegen — nur die, deren Titel es noch nicht gibt.
   *
   * Das ist der EINE Weg dafür: Die automatische Erstbefüllung, der Knopf auf der
   * Kachelwand und der Knopf im Studio rufen alle diese Funktion (bis v1.3 gab es
   * drei Wege mit drei Verhalten: einer ohne Merker, einer mit stummem Teilerfolg).
   * Wiederholbar ohne Doppelte; läuft schon ein Lauf, kommt `fehler: 'laeuft'`
   * zurück statt eines zweiten (sonst 36 Kacheln bei zwei Klicks).
   *
   * Der Grund eines Misserfolgs steht im Ergebnis, nicht in einem State: Der State
   * wäre beim Lesen der Meldung noch der von VOR dem Lauf.
   */
  seedStartUseCases: () => Promise<StartbestandErgebnis>;
}

const UseCaseContext = React.createContext<UseCaseContextType | undefined>(undefined);

export function UseCaseProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const { service, isOrganizer, isRolesLoading } = useRoles();
  const [useCases, setUseCases] = React.useState<UseCase[]>([]);
  const [ladeStatus, setLadeStatus] = React.useState<LadeStatus>('laedt');
  const [letzterStatus, setLetzterStatus] = React.useState(0);
  const [letzterFehler, setLetzterFehler] = React.useState('');
  const [aktualisierungFehler, setAktualisierungFehler] = React.useState('');
  const [startbestandTeilweise, setStartbestandTeilweise] = React.useState<StartbestandErgebnis | null>(null);

  // Zähler: Jeder Ladevorgang bekommt eine Nummer, und nur der ZULETZT gestartete
  // darf den State setzen. Parallele Nachladevorgänge (nach jedem Speichern, „Erneut
  // versuchen", Startbestand) überholen sich sonst — zwei schnelle Löschungen, und
  // die ältere Antwort (nur A gelöscht) kam zuletzt und ließ B wieder auftauchen.
  const reloadNr = React.useRef(0);
  const hatteDaten = React.useRef(false);

  const reload = React.useCallback(async (): Promise<void> => {
    const nr = ++reloadNr.current;
    const rows = await service.getUseCases();
    // Sofort nach dem await lesen: Die Felder gehören dem Service und werden vom
    // nächsten Lesen überschrieben.
    const status = service.lastUseCasesReadStatus;
    const fehler = service.lastReadError;
    if (nr !== reloadNr.current) return;
    setLetzterStatus(status);
    setLetzterFehler(fehler);
    if (rows === null) {
      // NICHT auf [] setzen. Ein Lesefehler ist keine Aussage ueber die Daten.
      // Gibt es schon einen Stand, bleibt er stehen — die Kachelwand durch eine
      // Fehlermeldung zu ersetzen, obwohl das Speichern gerade geklappt hat (429
      // beim Nachladen), sperrte auch „Neuer Use Case" für nichts.
      console.warn('[AIUC] Use Cases nicht lesbar — bestehender Stand bleibt.');
      if (hatteDaten.current) setAktualisierungFehler(fehler || `HTTP ${status}`);
      else setLadeStatus('fehler');
      return;
    }
    hatteDaten.current = true;
    setAktualisierungFehler('');
    setUseCases(rows);
    setLadeStatus('ok');
  }, [service]);

  /**
   * Die Start-Use-Cases anlegen — der EINE Weg (siehe `UseCaseContextType`).
   *
   * Nacheinander, nicht parallel: 18 gleichzeitige POSTs gegen dieselbe Liste sind
   * genau das Muster, das SharePoint drosselt. Der Merker „schon befüllt" kommt nach
   * dem ersten angelegten Eintrag ins Protokoll; ohne ihn käme der Bestand zurück,
   * sobald jemand alle Einträge absichtlich löscht.
   */
  const seedLaeuft = React.useRef(false);
  const seedStartUseCases = React.useCallback(async (): Promise<StartbestandErgebnis> => {
    const ohne = (fehler: string): StartbestandErgebnis => ({ angelegt: 0, fehlend: 0, fehler, merker: true });
    // Wie `create`/`update`/`remove`: ohne Organizer-Rolle wird nichts geschrieben.
    if (!isOrganizer) return ohne('rechte');
    if (seedLaeuft.current) return ohne('laeuft');
    seedLaeuft.current = true;
    try {
      const aktuell = await service.getUseCases();
      if (aktuell === null) return ohne(service.lastReadError || 'lesefehler');
      const norm = (t: string | undefined): string => (t || '').toLowerCase().replace(/\s+/g, ' ').trim();
      const vorhanden: { [titel: string]: boolean } = {};
      aktuell.forEach(u => { vorhanden[norm(u.titel)] = true; });
      const offen = START_USE_CASES.filter(u => !vorhanden[norm(u.titel)]);
      let angelegt = 0;
      let fehler = '';
      for (const uc of offen) {
        // eslint-disable-next-line no-await-in-loop
        if (await service.createUseCase(uc)) angelegt++;
        else fehler = service.lastReadError || 'Der Eintrag wurde von SharePoint abgelehnt.';
      }
      const merker = angelegt > 0 ? await service.merkeErstbefuellung(angelegt) : true;
      const erg: StartbestandErgebnis = { angelegt, fehlend: offen.length, fehler, merker };
      setStartbestandTeilweise(angelegt < offen.length ? erg : null);
      await reload();
      return erg;
    } finally {
      seedLaeuft.current = false;
    }
  }, [service, reload, isOrganizer]);

  /*
   * Start: erst die Rollen, dann die Listen.
   *
   * v1.3 (Sichtprüfung/Review): Bis dahin lief bei JEDEM Start jeder Person zuerst
   * `ensureUseCaseList` und `ensureLogList` — rund 24 nacheinander gestellte GETs
   * (Spalten-Proben), bevor die erste Kachel gelesen wurde. Angelegt wird aber nur
   * von Organizern und Admins; für alle anderen war das reine Wartezeit. Jetzt
   * wartet der Start auf die Rollen (die braucht die Oberfläche ohnehin) und
   * stellt die Listen nur für Organizer sicher.
   *
   * Erstbefüllung: Bedingung ist der ZUSTAND, nicht der Zeitpunkt — leer UND noch
   * nie befüllt (v1.1: „Liste ist gerade erst entstanden" trug genau einen Start
   * lang). Der Merker steht im Protokoll. Ist eine der beiden Listen nicht
   * lesbar, wird NICHT befüllt — ein Lesefehler ist keine leere Plattform (DEX
   * v30.37). Sie läuft NACH dem ersten Lesen im selben Ablauf, nicht in einem
   * eigenen Effekt: Zwei Effekte rennen gegeneinander, und auf einer frischen
   * Installation las die Erstbefüllung, bevor die Liste angelegt war.
   */
  const startRef = React.useRef(false);
  React.useEffect(() => {
    if (isRolesLoading || startRef.current) return;
    startRef.current = true;
    (async (): Promise<void> => {
      if (isOrganizer) {
        try {
          await service.ensureUseCaseList();
          await service.ensureLogList();
        } catch (e) {
          console.warn('[AIUC] Listen konnten nicht sichergestellt werden:', e);
        }
      }
      await reload();
      if (isOrganizer && await service.darfErstbefuellen()) await seedStartUseCases();
    })().catch(e => {
      console.warn('[AIUC] Start:', e);
      if (!hatteDaten.current) setLadeStatus('fehler');
    });
  }, [isRolesLoading, isOrganizer, service, reload, seedStartUseCases]);

  const create = React.useCallback(async (uc: Partial<UseCase>): Promise<number | null> => {
    if (!isOrganizer) return null;
    const id = await service.createUseCase(uc);
    if (id) {
      await service.log(id, 'angelegt', uc.titel || '');
      await reload();
    }
    return id;
  }, [service, reload, isOrganizer]);

  const update = React.useCallback(async (id: number, uc: Partial<UseCase>): Promise<boolean> => {
    if (!isOrganizer) return false;
    const alt = useCases.filter(u => u.id === id)[0];
    const ok = await service.updateUseCase(id, uc);
    if (ok) {
      // Was sich WIRKLICH geändert hat (v1.3) — vorher stand hier bei jedem
      // Speichern die Liste aller Schlüssel. Keine Änderung, keine Zeile.
      const text = geaendertText(alt, uc);
      if (text) await service.log(id, 'geaendert', text);
      await reload();
    }
    return ok;
  }, [service, reload, isOrganizer, useCases]);

  const saveUseCase = React.useCallback(async (
    id: number | null,
    uc: Partial<UseCase>,
    bild: BildAenderung,
  ): Promise<SpeicherErgebnis> => {
    if (!isOrganizer) return { ok: false, id, bildFehler: false, grund: 'rechte' };

    // --- Neu anlegen ---------------------------------------------------
    if (id === null) {
      // Ohne Bild-URL anlegen: Das Bild braucht die Zeile (der Anhang hängt an
      // ihr) und kommt danach.
      const neueId = await service.createUseCase({ ...uc, bildUrl: '' });
      if (!neueId) return { ok: false, id: null, bildFehler: false, grund: 'speichern' };
      await service.log(neueId, 'angelegt', uc.titel || '');
      let bildFehler = false;
      if (bild.art === 'neu') {
        const hoch = await service.uploadBild(neueId, bild.datei);
        if (hoch && await service.updateUseCase(neueId, { bildUrl: hoch.url })) {
          await service.log(neueId, 'geaendert', 'Bild');
        } else {
          bildFehler = true;
          // Ein Anhang ohne Verweis darauf ist Ballast — wegräumen.
          if (hoch) await service.entferneBilder(neueId, { nur: hoch.name });
        }
      }
      await reload();
      return { ok: true, id: neueId, bildFehler };
    }

    // --- Ändern --------------------------------------------------------
    const alt = useCases.filter(u => u.id === id)[0];

    // 1. Das neue Bild ZUERST. Scheitert es, ist noch nichts verändert.
    let neuesBild: { url: string; name: string } | null = null;
    if (bild.art === 'neu') {
      neuesBild = await service.uploadBild(id, bild.datei);
      if (!neuesBild) return { ok: false, id, bildFehler: true, grund: 'bild' };
    }

    // 2. Die Zeile speichern — mit der neuen Adresse (oder leer beim Entfernen).
    //    Geschrieben wird nur, was sich gegenüber dem Stand geändert hat — der Entwurf
    //    ist der Stand vom Öffnen des Dialogs, und ihn ganz zurückzuschreiben
    //    überschriebe, was jemand inzwischen geändert hat (u. a. das Bild: `bildUrl`
    //    stand dann wieder auf einer längst recycelten Datei).
    const daten: Partial<UseCase> = nurGeaendertes(alt, uc);
    if (neuesBild) daten.bildUrl = neuesBild.url;
    else if (bild.art === 'entfernen') daten.bildUrl = '';
    const ok = Object.keys(daten).length === 0 ? true : await service.updateUseCase(id, daten);
    if (!ok) {
      // Rückbau: das gerade hochgeladene Bild gehört jetzt niemandem.
      if (neuesBild) await service.entferneBilder(id, { nur: neuesBild.name });
      return { ok: false, id, bildFehler: false, grund: 'speichern' };
    }

    const text = geaendertText(alt, daten);
    if (text) await service.log(id, 'geaendert', text);

    // 3. Erst jetzt das alte Bild — der unumkehrbare Schritt kommt zuletzt.
    if (bild.art !== 'unveraendert') {
      await service.entferneBilder(id, neuesBild ? { behalte: neuesBild.name } : {});
    }
    await reload();
    return { ok: true, id, bildFehler: false };
  }, [service, reload, isOrganizer, useCases]);

  const remove = React.useCallback(async (id: number): Promise<boolean> => {
    if (!isOrganizer) return false;
    // Protokoll VOR dem Loeschen — danach ist der Titel weg, und das
    // Protokoll waere die Nebenbuchhaltung, die nichts mehr belegt.
    // (Dieselbe Reihenfolge wie in DEX: pruefbare Nebenbuchhaltung zuerst,
    // der unumkehrbare Schritt zuletzt.) Das Löschen geht in den Papierkorb und ist
    // damit umkehrbar — ein Protokollfehler blockiert es deshalb nicht, aber die
    // Zeile wird danach nachgeholt statt still zu fehlen.
    const uc = useCases.filter(u => u.id === id)[0];
    const titel = uc ? uc.titel : `Id ${id}`;
    const vorher = await service.log(id, 'geloescht', titel);
    const ok = await service.deleteUseCase(id);
    if (ok) {
      if (!vorher) await service.log(id, 'geloescht', titel);
      await reload();
    // Das Protokoll steht VOR dem Löschen und behauptet „gelöscht". Schlägt
    // das Löschen fehl, muss die Gegenbuchung dastehen — sonst sagt das
    // Protokoll etwas, das nicht stimmt.
    } else if (vorher) {
      await service.log(id, 'loeschen-fehlgeschlagen', titel);
    }
    return ok;
  }, [service, reload, useCases, isOrganizer]);

  const bereiche = React.useMemo(() => {
    const set: Record<string, true> = {};
    useCases.forEach(u => { if (u.bereich) set[u.bereich] = true; });
    return Object.keys(set).sort((a, b) => a.localeCompare(b, 'de'));
  }, [useCases]);

  const value = React.useMemo<UseCaseContextType>(() => ({
    useCases, ladeStatus, letzterStatus, letzterFehler, fehlendeSpalten: service.fehlendeSpalten,
    aktualisierungFehler, startbestandTeilweise,
    reload, create, update, remove, saveUseCase, bereiche, seedStartUseCases,
  }), [useCases, ladeStatus, letzterStatus, letzterFehler, aktualisierungFehler, startbestandTeilweise, reload, create, update, remove, saveUseCase, bereiche, seedStartUseCases, service]);

  return React.createElement(UseCaseContext.Provider, { value }, props.children);
}

export function useUseCases(): UseCaseContextType {
  const ctx = React.useContext(UseCaseContext);
  if (!ctx) throw new Error('useUseCases muss innerhalb des UseCaseProvider stehen');
  return ctx;
}
